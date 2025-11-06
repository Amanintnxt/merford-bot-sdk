# ──────────────────────────────────────────────────────────────────────────────
# Minimal bot.py — SSO + magic code + simple CLARIFY handling (text-only)
# ──────────────────────────────────────────────────────────────────────────────
import os
import re
import time
import asyncio
import logging
import requests
import openai
from dotenv import load_dotenv
from flask import Flask, request, Response
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter, TurnContext
from botbuilder.schema import Activity, Attachment, CardAction, ActionTypes, OAuthCard

# ─────────────── ENV & OPENAI ───────────────
load_dotenv()

APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_EP = os.getenv("AZURE_OPENAI_ENDPOINT")
OAUTH_CONNECTION = os.getenv("OAUTH_CONNECTION_NAME", "TeamsSSO")

openai.api_type = "azure"
openai.api_version = "2024-05-01-preview"
openai.api_key = AZURE_OPENAI_KEY
openai.azure_endpoint = AZURE_OPENAI_EP.rstrip("/")

# Map access levels to assistants (keep updated)
ASSISTANT_MAP = {
    "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
    "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
    "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
    "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1",
}

# ─────────────── Flask & Bot adapter ───────────────
app = Flask(__name__)
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# ─────────────── Simple state ───────────────
thread_map = {}              # user_id:assistant_id → thread_id
awaiting_clarification = set()  # users expecting to answer a CLARIFY

# ─────────────── Helpers ───────────────


def get_user_group_level(token: str) -> str | None:
    """Return Level 1..4 from Graph groups or None."""
    url = "https://graph.microsoft.com/v1.0/me/memberOf?$select=displayName"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        r = requests.get(url, headers=headers, timeout=10)
    except requests.RequestException:
        return None
    if r.status_code != 200:
        return None
    for g in r.json().get("value", []):
        name = (g.get("displayName") or "")
        if name == "Level1Access":
            return "Level 1"
        if name == "Level2Access":
            return "Level 2"
        if name == "Level3Access":
            return "Level 3"
        if name == "Level4Access":
            return "Level 4"
    return None


async def try_get_token(turn_context: TurnContext, magic_code: str | None = None):
    try:
        return await adapter.get_user_token(turn_context, OAUTH_CONNECTION, magic_code)
    except Exception:
        return None


async def ensure_token(turn_context: TurnContext) -> str | None:
    """Silent SSO → magic code in value.state → send OAuth card."""
    magic = None
    if turn_context.activity.value and isinstance(turn_context.activity.value, dict):
        magic = turn_context.activity.value.get("state")

    token_resp = await try_get_token(turn_context, magic)
    if token_resp and token_resp.token:
        return token_resp.token

    # No token: send sign-in card
    url = await adapter.get_oauth_sign_in_link(turn_context, OAUTH_CONNECTION)
    card = OAuthCard(
        text="Please sign in to continue.",
        connection_name=OAUTH_CONNECTION,
        buttons=[CardAction(type=ActionTypes.signin,
                            title="Sign In", value=url)],
    )
    await turn_context.send_activity(Activity(
        attachments=[Attachment(
            content_type="application/vnd.microsoft.card.oauth", content=card)]
    ))
    return None


def is_magic_code(text: str) -> bool:
    """Six to eight digits only."""
    return bool(re.fullmatch(r"\d{6,8}", (text or "").strip()))


def is_clarify(text: str) -> bool:
    """Assistant uses CLARIFY: prefix per instructions."""
    return bool(text and text.strip().upper().startswith("CLARIFY:"))


def strip_clarify(text: str) -> str:
    return text[len("CLARIFY:"):].strip()

# ─────────────── Core handler ───────────────


async def handle_activity(turn_context: TurnContext):
    a = turn_context.activity
    user_id = a.from_property.id

    # Conversation start
    if a.type == "conversationUpdate":
        for m in a.members_added or []:
            if m.id == a.recipient.id:
                await turn_context.send_activity("✅ Connected. Please sign in to continue.")
        return

    if a.type != "message":
        return

    user_text = (a.text or "").strip()

    # Magic code in message (Web Chat/Direct Line)
    if is_magic_code(user_text):
        token_resp = await try_get_token(turn_context, user_text)
        if token_resp and token_resp.token:
            await turn_context.send_activity("🔓 Sign-in successful! You can now ask your question.")
        else:
            await turn_context.send_activity("⚠️ Sign-in failed. Please click Sign In again.")
        return

    # Ensure token (Teams silent SSO or OAuth card)
    access_token = await ensure_token(turn_context)
    if not access_token:
        return  # waiting for sign-in

    # Resolve access level → assistant
    level = get_user_group_level(access_token)
    assistant_id = ASSISTANT_MAP.get(level)
    if not assistant_id:
        await turn_context.send_activity("❌ No assistant mapped to your access level.")
        return

    # Thread per user+assistant
    key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(key)
    if not thread_id:
        thread_id = openai.beta.threads.create().id
        thread_map[key] = thread_id

    # Add message (tag if this is a clarification response)
    if user_text:
        content = user_text
        if user_id in awaiting_clarification:
            content = f"(User clarification) {user_text}"
            awaiting_clarification.discard(user_id)
        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=content)

    # Run assistant (file_search enforced; set temperature=0 on assistant side)
    run = openai.beta.threads.runs.create(
        assistant_id=assistant_id,
        thread_id=thread_id,
        tool_choice={"type": "file_search"}
    )

    start = time.time()
    while run.status not in ("completed", "failed", "cancelled"):
        if time.time() - start > 60:
            await turn_context.send_activity("⏳ Still working… please try again in a moment.")
            return
        await asyncio.sleep(1)
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)

    # Get assistant reply
    msgs = openai.beta.threads.messages.list(
        thread_id=thread_id, order="desc", limit=5)
    reply = next(
        (m.content[0].text.value for m in msgs.data if m.role == "assistant"), None)

    # Simple clarify handling
    if reply and is_clarify(reply):
        question = strip_clarify(reply)
        awaiting_clarification.add(user_id)
        await turn_context.send_activity(Activity(type="message", text=question))
        return

    # Normal answer
    await turn_context.send_activity(reply or "❌ No response from assistant.")

# ─────────────── Routes ───────────────


@app.route("/api/messages", methods=["POST"])
def messages():
    try:
        if "application/json" not in request.headers.get("Content-Type", ""):
            return Response("Unsupported Media Type", 415)
        activity = Activity().deserialize(request.json)
        auth_hdr = request.headers.get("Authorization", "")

        async def _proc():
            return await adapter.process_activity(activity, auth_hdr, handle_activity)

        asyncio.run(_proc())
        return Response(status=200)
    except Exception as ex:
        logging.exception("Exception in /api/messages: %s", ex)
        return Response("Internal Server Error", 500)


@app.route("/", methods=["GET"])
def health():
    return "Bot is running."


# ─────────────── Main ───────────────
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app.run(host="0.0.0.0", port=3978)
