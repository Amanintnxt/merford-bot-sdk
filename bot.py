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

# ──────────────────────────────────────────────────────────────────────────────
#  Environment &  OpenAI client
# ──────────────────────────────────────────────────────────────────────────────
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

# ──────────────────────────────────────────────────────────────────────────────
#  Flask   +   Bot Framework adapter
# ──────────────────────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder="static")
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# ──────────────────────────────────────────────────────────────────────────────
#  In-memory state
# ──────────────────────────────────────────────────────────────────────────────
thread_map = {}   # key = f"{user_id}:{assistant_id}"  →  thread_id
signed_in_users = {}   # user_id → True once token validated

# ──────────────────────────────────────────────────────────────────────────────
#  Static mapping  Level → Assistant ID  (update to your real IDs)
# ──────────────────────────────────────────────────────────────────────────────
ASSISTANT_MAP = {
    "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
    "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
    "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
    "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
}

# ──────────────────────────────────────────────────────────────────────────────
#  Helper – get user level via Microsoft Graph
# ──────────────────────────────────────────────────────────────────────────────


def get_user_group_level(access_token: str) -> str | None:
    url = "https://graph.microsoft.com/v1.0/me/memberOf?$select=displayName"
    headers = {"Authorization": f"Bearer {access_token}"}

    resp = requests.get(url, headers=headers, timeout=10)
    if resp.status_code != 200:
        logging.warning("Group lookup failed %s – %s",
                        resp.status_code, resp.text)
        return None

    for grp in resp.json().get("value", []):
        name = grp.get("displayName")
        if name == "Level1Access":
            return "Level 1"
        elif name == "Level2Access":
            return "Level 2"
        elif name == "Level3Access":
            return "Level 3"
        elif name == "Level4Access":
            return "Level 4"
    return None

# ──────────────────────────────────────────────────────────────────────────────
#  Core bot handler
# ──────────────────────────────────────────────────────────────────────────────


async def handle_message(turn_context: TurnContext):
    user_id = turn_context.activity.from_property.id

    # 1️⃣  Conversation start → greeting
    if turn_context.activity.type == "conversationUpdate":
        for m in turn_context.activity.members_added or []:
            if m.id == turn_context.activity.recipient.id:
                await turn_context.send_activity("✅ You are now connected to the bot.")
        return

    # Ignore empty / non-text
    if turn_context.activity.type != "message" or not turn_context.activity.text.strip():
        return

    # 2️⃣  Magic-code detection for WebChat OAuth
    magic_code = None
    if turn_context.activity.value and "state" in turn_context.activity.value:
        magic_code = turn_context.activity.value["state"]
    elif turn_context.activity.text.strip().isdigit():
        magic_code = turn_context.activity.text.strip()

    # 3️⃣  Acquire Teams SSO token
    token_resp = await adapter.get_user_token(
        turn_context, OAUTH_CONNECTION, magic_code
    )
    if not token_resp or not token_resp.token:
        sign_in_url = await adapter.get_oauth_sign_in_link(
            turn_context, OAUTH_CONNECTION
        )
        card = OAuthCard(
            text="Please sign in to continue.",
            connection_name=OAUTH_CONNECTION,
            buttons=[CardAction(type=ActionTypes.signin,
                                title="Sign In",
                                value=sign_in_url)]
        )
        attachment = Attachment(
            content_type="application/vnd.microsoft.card.oauth",
            content=card
        )
        await turn_context.send_activity(Activity(attachments=[attachment]))
        return

    access_token = token_resp.token

    # First turn after sign-in → confirmation only
    if user_id not in signed_in_users:
        signed_in_users[user_id] = True
        await turn_context.send_activity("🔐 Sign-in successful! You can now ask questions.")
        return

    # 4️⃣  Determine level  → assistant
    level = get_user_group_level(access_token)

    logging.info(f"User is at: {level}")

    assistant_id = ASSISTANT_MAP.get(level)

    logging.info(f"User assigned to assistant: {assistant_id}")
    if not assistant_id:
        await turn_context.send_activity("❌ Assistant not found for your access level")
        return

    # 5️⃣  Create / reuse a thread dedicated to this assistant+user
    thread_key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(thread_key)
    if not thread_id:
        thread_id = openai.beta.threads.create().id
        thread_map[thread_key] = thread_id

    # 6️⃣  Add user message
    openai.beta.threads.messages.create(
        thread_id=thread_id,
        role="user",
        content=turn_context.activity.text
    )

    # 7️⃣  Run the dedicated assistant
    try:
        run = openai.beta.threads.runs.create(
            assistant_id=assistant_id,
            thread_id=thread_id
        )
    except Exception as e:                # includes wrong-ID error
        logging.error("Assistant run failed: %s", e)
        await turn_context.send_activity(f"❌ Assistant run failed: {e}")
        return

    # 8️⃣  Wait for completion (simple polling)
    while run.status not in ("completed", "failed", "cancelled"):
        await asyncio.sleep(1)
        run = openai.beta.threads.runs.retrieve(thread_id=thread_id,
                                                run_id=run.id)

    # 9️⃣  Fetch only the newest assistant message from this run
    msgs = openai.beta.threads.messages.list(
        thread_id=thread_id,
        order="desc",
        limit=1
    )
    reply = next(
        (m.content[0].text.value for m in msgs.data if m.role == "assistant"),
        "❌ No reply from assistant."
    )

    await turn_context.send_activity(reply)

# ──────────────────────────────────────────────────────────────────────────────
#  Flask endpoints
# ──────────────────────────────────────────────────────────────────────────────


@app.route("/api/messages", methods=["POST"])
def messages():
    try:
        if "application/json" not in request.headers.get("Content-Type", ""):
            return Response("Unsupported Media Type", 415)
        activity = Activity().deserialize(request.json)
        auth_hdr = request.headers.get("Authorization", "")

        async def _process():  # wrap so we can call await from sync route
            return await adapter.process_activity(activity, auth_hdr, handle_message)
        asyncio.run(_process())
        return Response(status=200)
    except Exception as exc:
        logging.error("Exception in /api/messages: %s", exc)
        return Response("Internal Server Error", 500)


@app.route("/directline/token", methods=["POST"])
def directline_token():
    if not DIRECT_LINE_SECRET:
        return jsonify({"error": "DIRECT_LINE_SECRET not set"}), 500
    url = "https://directline.botframework.com/v3/directline/tokens/generate"
    headers = {"Authorization": f"Bearer {DIRECT_LINE_SECRET}"}
    resp = requests.post(url, headers=headers, timeout=10)
    if resp.status_code != 200:
        return jsonify({"error": "Failed to generate token", "details": resp.text}), 500
    return jsonify({"token": resp.json().get("token")})


@app.route("/chat", methods=["GET"])
def serve_chat():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/", methods=["GET"])
def health():
    return "Bot is running."


# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app.run(host="0.0.0.0", port=3978)
