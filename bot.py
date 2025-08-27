import os
import asyncio
import time
import logging
import requests
import openai

from dotenv import load_dotenv
from flask import Flask, request, Response, jsonify, send_from_directory
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter, TurnContext
from botbuilder.schema import Activity, Attachment, CardAction, ActionTypes, OAuthCard

# ───────────────────────────  ENV / OPENAI CONFIG  ───────────────────────────
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

# ───────────────────────────  LOGGING CONFIG  ────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s"
)

# ───────────────────────────  FLASK + BOT ADAPTER  ───────────────────────────
app = Flask(__name__, static_folder="static")
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# ───────────────────────────  IN-MEMORY STATE  ───────────────────────────────
thread_map = {}   # f"{user_id}:{assistant_id}" → thread_id
signed_in_users = {}   # user_id → True

# ───────────────────────────  ASSISTANT MAP  ─────────────────────────────────
ASSISTANT_MAP = {
    "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
    "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
    "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
    "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
}

# ───────────────────────────  GRAPH GROUP LOOK-UP  ───────────────────────────


def get_user_group_level(token: str) -> str | None:
    url = "https://graph.microsoft.com/v1.0/me/memberOf?$select=displayName"
    headers = {"Authorization": f"Bearer {token}"}

    resp = requests.get(url, headers=headers, timeout=10)
    if resp.status_code != 200:
        logging.warning("Group lookup failed %s – %s",
                        resp.status_code, resp.text)
        return None

    for grp in resp.json().get("value", []):
        name = grp.get("displayName")
        if name == "Level1Access":
            return "Level 1"
        if name == "Level2Access":
            return "Level 2"
        if name == "Level3Access":
            return "Level 3"
        if name == "Level4Access":
            return "Level 4"
    return None

# ───────────────────────────  MAIN BOT HANDLER  ──────────────────────────────


async def handle_message(turn_context: TurnContext):
    user_id = turn_context.activity.from_property.id

    # Initial connection greeting
    if turn_context.activity.type == "conversationUpdate":
        for m in turn_context.activity.members_added or []:
            if m.id == turn_context.activity.recipient.id:
                await turn_context.send_activity("✅ You are now connected to the bot.")
        return

    if turn_context.activity.type != "message" or not turn_context.activity.text.strip():
        return

    # Detect OAuth magic code (Web Chat)
    magic_code = None
    if turn_context.activity.value and "state" in turn_context.activity.value:
        magic_code = turn_context.activity.value["state"]
    elif turn_context.activity.text.strip().isdigit():
        magic_code = turn_context.activity.text.strip()

    # Acquire Teams SSO token
    token_resp = await adapter.get_user_token(turn_context, OAUTH_CONNECTION, magic_code)
    if not token_resp or not token_resp.token:
        sign_in_url = await adapter.get_oauth_sign_in_link(turn_context, OAUTH_CONNECTION)
        card = OAuthCard(
            text="Please sign in to continue.",
            connection_name=OAUTH_CONNECTION,
            buttons=[CardAction(type=ActionTypes.signin,
                                title="Sign In", value=sign_in_url)]
        )
        await turn_context.send_activity(Activity(
            attachments=[Attachment(
                content_type="application/vnd.microsoft.card.oauth", content=card)]
        ))
        logging.info("Sent sign-in card to user %s", user_id)
        return

    access_token = token_resp.token

    # First post-sign-in turn
    if user_id not in signed_in_users:
        signed_in_users[user_id] = True
        await turn_context.send_activity("🔐 Sign-in successful! You can now ask questions.")
        logging.info("User %s signed in", user_id)
        return

    # Determine user level and assistant
    level = get_user_group_level(access_token)
    logging.info("User %s level determined: %s", user_id, level)

    assistant_id = ASSISTANT_MAP.get(level)
    if not assistant_id:
        await turn_context.send_activity("❌ Assistant not found for your access level.")
        logging.warning(
            "Assistant not found for user %s (level %s)", user_id, level)
        return

    logging.info("User %s assigned to assistant: %s", user_id, assistant_id)

    # Thread handling (per-user, per-assistant)
    thread_key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(thread_key)
    if not thread_id:
        thread_id = openai.beta.threads.create().id
        thread_map[thread_key] = thread_id
        logging.info("Created new thread %s for %s", thread_id, thread_key)
    else:
        logging.info("Re-using thread %s for %s", thread_id, thread_key)

    # Add user message
    openai.beta.threads.messages.create(
        thread_id=thread_id,
        role="user",
        content=turn_context.activity.text
    )
    logging.info("Posted user message to thread %s", thread_id)

    # Start assistant run
    try:
        run = openai.beta.threads.runs.create(
            assistant_id=assistant_id, thread_id=thread_id)
        logging.info("Run %s created for thread %s", run.id, thread_id)
    except Exception as e:
        logging.error("Assistant run failed: %s", e)
        await turn_context.send_activity(f"❌ Assistant run failed: {e}")
        return

    # Poll until completed
    while run.status not in ("completed", "failed", "cancelled"):
        await asyncio.sleep(1)
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)
        logging.info("Run %s status: %s", run.id, run.status)

    # Fetch newest assistant message
    msgs = openai.beta.threads.messages.list(
        thread_id=thread_id, order="desc", limit=1)
    reply = next((m.content[0].text.value for m in msgs.data if m.role == "assistant"),
                 "❌ No reply from assistant.")

    logging.info("Assistant reply sent to user %s", user_id)
    await turn_context.send_activity(reply)

# ───────────────────────────  FLASK ENDPOINTS  ───────────────────────────────


@app.route("/api/messages", methods=["POST"])
def messages():
    try:
        if "application/json" not in request.headers.get("Content-Type", ""):
            return Response("Unsupported Media Type", 415)
        activity = Activity().deserialize(request.json)
        auth_hdr = request.headers.get("Authorization", "")

        async def _proc():
            return await adapter.process_activity(activity, auth_hdr, handle_message)
        asyncio.run(_proc())
        return Response(status=200)
    except Exception as exc:
        logging.exception("Exception in /api/messages: %s", exc)
        return Response("Internal Server Error", 500)


@app.route("/directline/token", methods=["POST"])
def directline_token():
    if not DIRECT_LINE_SECRET:
        return jsonify({"error": "DIRECT_LINE_SECRET not set"}), 500
    url = "https://directline.botframework.com/v3/directline/tokens/generate"
    headers = {"Authorization": f"Bearer {DIRECT_LINE_SECRET}"}
    resp = requests.post(url, headers=headers, timeout=10)
    if resp.status_code != 200:
        logging.error("Direct Line token generation failed: %s", resp.text)
        return jsonify({"error": "Failed to generate token"}), 500
    logging.info("Direct Line token generated")
    return jsonify({"token": resp.json().get("token")})


@app.route("/chat", methods=["GET"])
def serve_chat():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/", methods=["GET"])
def health():
    return "Bot is running."


# ───────────────────────────  MAIN  ──────────────────────────────────────────
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=3978)
