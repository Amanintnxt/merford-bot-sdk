import os
import time
import asyncio
import logging
import requests
import openai
from dotenv import load_dotenv
from flask import Flask, request, Response, jsonify, send_from_directory
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter, TurnContext
from botbuilder.schema import Activity
from openai import AzureOpenAI

# ──────────────── ENV & OPENAI CONFIG ────────────────
load_dotenv()

APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_EP = os.getenv("AZURE_OPENAI_ENDPOINT")
OAUTH_CONNECTION = os.getenv("OAUTH_CONNECTION_NAME", "TeamsSSO")
DIRECT_LINE_SECRET = os.getenv("DIRECT_LINE_SECRET", "")

openai.api_type = "azure"
openai.api_version = "2024-05-01-preview"
openai.api_key = AZURE_OPENAI_KEY
openai.azure_endpoint = AZURE_OPENAI_EP.rstrip("/")

client = AzureOpenAI(
    api_key=AZURE_OPENAI_KEY,
    azure_endpoint=AZURE_OPENAI_EP,
    api_version="2024-05-01-preview"
)

# ──────────────── ASSISTANT MAP ────────────────
ASSISTANT_MAP = {
    "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
    "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
    "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
    "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
}

# ──────────────── FLASK & BOT ADAPTER ────────────────
app = Flask(__name__)
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

thread_map = {}
awaiting_clarification = set()

# ──────────────── HELPERS ────────────────


def _is_clarify_message(text: str) -> bool:
    """Detect if assistant is asking a clarifying question."""
    if not text:
        return False
    t = text.strip().lower()
    return t.startswith("clarify:") or (t.endswith("?") and len(t) < 200)


def get_user_group_level(token: str) -> str | None:
    """Get user group (Level 1–4) via Microsoft Graph."""
    url = "https://graph.microsoft.com/v1.0/me/memberOf?$select=displayName"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code != 200:
            return None
        for g in resp.json().get("value", []):
            name = g.get("displayName", "")
            if "Level1Access" in name:
                return "Level 1"
            if "Level2Access" in name:
                return "Level 2"
            if "Level3Access" in name:
                return "Level 3"
            if "Level4Access" in name:
                return "Level 4"
    except Exception:
        return None
    return None


async def try_get_token(turn_context: TurnContext, magic_code=None):
    try:
        return await adapter.get_user_token(turn_context, OAUTH_CONNECTION, magic_code)
    except Exception:
        return None


async def ensure_token(turn_context: TurnContext):
    magic = None
    if turn_context.activity.value and isinstance(turn_context.activity.value, dict):
        magic = turn_context.activity.value.get("state")
    token_resp = await try_get_token(turn_context, magic)
    if token_resp and token_resp.token:
        return token_resp.token
    url = await adapter.get_oauth_sign_in_link(turn_context, OAUTH_CONNECTION)
    await turn_context.send_activity(f"Please sign in here to continue: {url}")
    return None

# ──────────────── MAIN MESSAGE HANDLER ────────────────


async def handle_activity(turn_context: TurnContext):
    a = turn_context.activity
    user_id = a.from_property.id
    user_text = (a.text or "").strip()

    # 1️⃣ On join
    if a.type == "conversationUpdate":
        for m in a.members_added or []:
            if m.id == a.recipient.id:
                await turn_context.send_activity("✅ Connected. Please sign in to continue.")
        return

    if a.type != "message":
        return

    # 2️⃣ Handle OAuth magic code
    if user_text.isdigit() and len(user_text) <= 10:
        token = await try_get_token(turn_context, user_text)
        if token and token.token:
            await turn_context.send_activity("🔓 Sign-in successful! You can now start asking questions.")
        else:
            await turn_context.send_activity("⚠️ Sign-in failed. Please click Sign In again.")
        return

    # 3️⃣ Ensure access token
    access_token = await ensure_token(turn_context)
    if not access_token:
        return

    # 4️⃣ Get user level
    level = get_user_group_level(access_token)
    assistant_id = ASSISTANT_MAP.get(level)
    if not assistant_id:
        await turn_context.send_activity("❌ No assistant assigned for your access level.")
        return

    # 5️⃣ Thread management
    key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(key)
    if not thread_id:
        thread_id = openai.beta.threads.create().id
        thread_map[key] = thread_id

    # 6️⃣ Add message
    if user_text:
        if user_id in awaiting_clarification:
            user_text = f"(Clarification) {user_text}"
            awaiting_clarification.discard(user_id)
        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=user_text)

    # 7️⃣ Run assistant
    run = openai.beta.threads.runs.create(
        assistant_id=assistant_id, thread_id=thread_id)
    start = time.time()
    while run.status not in ("completed", "failed", "cancelled"):
        if time.time() - start > 60:
            await turn_context.send_activity("⏳ Still processing... please try again.")
            return
        await asyncio.sleep(1)
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)

    # 8️⃣ Get reply
    msgs = openai.beta.threads.messages.list(
        thread_id=thread_id, order="desc", limit=5)
    reply = next(
        (m.content[0].text.value for m in msgs.data if m.role == "assistant"), None)

    # 9️⃣ Handle clarify messages
    if reply and _is_clarify_message(reply):
        awaiting_clarification.add(user_id)
        await turn_context.send_activity(reply)
        return

    # 🔟 Final normal reply
    await turn_context.send_activity(reply or "⚠️ No response received.")

# ──────────────── FLASK ROUTES ────────────────


@app.route("/api/messages", methods=["POST"])
def messages():
    try:
        if "application/json" not in request.headers.get("Content-Type", ""):
            return Response("Unsupported Media Type", 415)
        activity = Activity().deserialize(request.json)
        auth = request.headers.get("Authorization", "")

        async def run_task():
            return await adapter.process_activity(activity, auth, handle_activity)
        asyncio.run(run_task())
        return Response(status=200)
    except Exception as e:
        logging.error(f"Exception in /api/messages: {e}")
        return Response("Internal Server Error", 500)


@app.route("/", methods=["GET"])
def health():
    return "Bot is running."


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app.run(host="0.0.0.0", port=3978)
