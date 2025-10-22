# ──────────────────────────────────────────────────────────────────────────────
# bot.py – Clarify-Enhanced Azure Bot (Teams + Direct Line + SSO)
# ──────────────────────────────────────────────────────────────────────────────
import os
import asyncio
import logging
import time
import requests
import openai
from dotenv import load_dotenv
from flask import Flask, request, Response, jsonify, send_from_directory, render_template
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter, TurnContext
from botbuilder.schema import Activity, Attachment, CardAction, ActionTypes, OAuthCard, SuggestedActions
from PyPDF2 import PdfReader
from openai import AzureOpenAI

# ─────────────────────────  ENV & CONFIG  ─────────────────────────────────────
load_dotenv()

APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_EP = os.getenv("AZURE_OPENAI_ENDPOINT")
OAUTH_CONNECTION = os.getenv("OAUTH_CONNECTION_NAME", "TeamsSSO")
DIRECT_LINE_SECRET = os.getenv("DIRECT_LINE_SECRET", "")
ADMIN_SECRET = os.getenv("ADMIN_SECRET")

openai.api_type = "azure"
openai.api_version = "2024-05-01-preview"
openai.api_key = AZURE_OPENAI_KEY
openai.azure_endpoint = AZURE_OPENAI_EP.rstrip("/")

client = AzureOpenAI(
    api_key=AZURE_OPENAI_KEY,
    azure_endpoint=AZURE_OPENAI_EP,
    api_version="2024-05-01-preview"
)

# ─────────────────────────  LOGGING  ─────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")

# ─────────────────────────  ASSISTANTS / VECTORS  ─────────────────────────────
ASSISTANT_MAP = {
    "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
    "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
    "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
    "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
}

VECTOR_STORES = {
    "Level 1": "vs_ICYlowKd3PPqtSp4m4wPzD47",
    "Level 2": "vs_FeOttDiAigZaxb8fjp1rAOIF",
    "Level 3": "vs_tO6kScvWu6oBn5R8YqeDkIX1",
    "Level 4": "vs_PJIPiZ91ojScAfJmKSCHrvx2"
}

# ─────────────────────────  FLASK / BOT ADAPTER  ──────────────────────────────
app = Flask(__name__, static_folder="static", template_folder="templates")
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# ─────────────────────────  IN-MEMORY STATE  ─────────────────────────────────
thread_map = {}
clarify_state = {}  # user_id → {"question": str, "pending": bool}

# ─────────────────────────  CLARIFY UTILITIES  ────────────────────────────────


def _is_clarify(text: str) -> bool:
    return text.strip().upper().startswith("CLARIFY:")


def _looks_like_clarify(text: str) -> bool:
    """Detect implicit clarifications."""
    if not text or "?" not in text:
        return False
    text_l = text.lower()
    return any(text_l.startswith(s) for s in (
        "what ", "which ", "can you", "could you", "do you", "please specify", "clarify"
    ))


def _strip_clarify(text: str) -> str:
    return text[len("CLARIFY:"):].strip() if _is_clarify(text) else text.strip()


def _clarify_actions(question_text: str):
    """Dynamic suggested answers."""
    lower = question_text.lower()
    if "model" in lower:
        opts = ["Model: M-series", "Model: Unknown", "Not sure"]
    elif "configuration" in lower:
        opts = ["Single leaf", "Double leaf", "Not sure"]
    elif "zone" in lower or "atex" in lower:
        opts = ["Zone 1", "Zone 2", "Not sure"]
    else:
        opts = ["I’ll specify", "Please repeat question", "Cancel"]
    return [CardAction(type=ActionTypes.im_back, title=o, value=o) for o in opts]

# ─────────────────────────  GRAPH LOOKUP  ────────────────────────────────────


def get_user_group_level(token: str) -> str | None:
    url = "https://graph.microsoft.com/v1.0/me/memberOf?$select=displayName"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
    except requests.RequestException:
        return None
    if resp.status_code != 200:
        return None
    for g in resp.json().get("value", []):
        name = g.get("displayName")
        if not name:
            continue
        if name == "Level1Access":
            return "Level 1"
        if name == "Level2Access":
            return "Level 2"
        if name == "Level3Access":
            return "Level 3"
        if name == "Level4Access":
            return "Level 4"
    return None

# ─────────────────────────  TOKEN HANDLING  ──────────────────────────────────


async def try_get_token(turn_context, magic=None):
    try:
        return await adapter.get_user_token(turn_context, OAUTH_CONNECTION, magic)
    except Exception:
        return None


async def ensure_token(turn_context):
    magic = None
    if turn_context.activity.value and isinstance(turn_context.activity.value, dict):
        magic = turn_context.activity.value.get("state")
    if not magic and turn_context.activity.text and turn_context.activity.text.strip().isdigit():
        magic = turn_context.activity.text.strip()

    token_resp = await try_get_token(turn_context, magic)
    if token_resp and token_resp.token:
        return token_resp.token

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

# ─────────────────────────  MAIN BOT HANDLER  ────────────────────────────────


async def handle_activity(turn_context: TurnContext):
    a = turn_context.activity
    user_id = a.from_property.id
    user_text = (a.text or "").strip()

    # Conversation start
    if a.type == "conversationUpdate":
        for m in a.members_added or []:
            if m.id == a.recipient.id:
                await turn_context.send_activity("✅ Connected. Please sign in to continue.")
        return

    # Only handle message type
    if a.type != "message":
        return

    # Acquire token
    token = await ensure_token(turn_context)
    if not token:
        return

    # Get user level
    level = get_user_group_level(token)
    if not level:
        await turn_context.send_activity("You do not have permission to access this bot.")
        return

    assistant_id = ASSISTANT_MAP.get(level)
    if not assistant_id:
        await turn_context.send_activity("Assistant not mapped for your access level.")
        return

    key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(key)
    if not thread_id:
        thread_id = openai.beta.threads.create().id
        thread_map[key] = thread_id

    # Typing indicator
    await turn_context.send_activity(Activity(type="typing"))

    # Check if awaiting clarification
    state = clarify_state.get(user_id, {"pending": False})
    if state["pending"]:
        logging.info(f"↩️ Clarification received from {user_id}: {user_text}")
        openai.beta.threads.messages.create(
            thread_id=thread_id,
            role="user",
            content=f"(User clarification) {user_text}"
        )
        clarify_state[user_id]["pending"] = False
    else:
        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=user_text)

    # Run assistant
    try:
        run = openai.beta.threads.runs.create(
            assistant_id=assistant_id, thread_id=thread_id,
            tool_choice={"type": "file_search"}
        )
    except Exception as e:
        await turn_context.send_activity(f"Assistant error: {e}")
        return

    # Wait efficiently
    for _ in range(40):
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)
        if run.status in ("completed", "failed", "cancelled"):
            break
        await asyncio.sleep(1)

    # Retrieve reply
    msgs = openai.beta.threads.messages.list(
        thread_id=thread_id, order="desc", limit=5)
    reply = next(
        (m.content[0].text.value for m in msgs.data if m.role == "assistant"), None)

    # Detect clarify
    if reply and (_is_clarify(reply) or _looks_like_clarify(reply)):
        question = _strip_clarify(reply)
        logging.info(f"🟡 CLARIFY triggered for {user_id}: {question}")
        clarify_state[user_id] = {"pending": True, "question": question}
        await turn_context.send_activity(Activity(
            type="message",
            text=question,
            suggested_actions=SuggestedActions(
                actions=_clarify_actions(question))
        ))
        return

    await turn_context.send_activity(reply or "No reply received.")

# ─────────────────────────  ROUTES  ──────────────────────────────────────────


@app.route("/api/messages", methods=["POST"])
def messages():
    activity = Activity().deserialize(request.json)
    auth_hdr = request.headers.get("Authorization", "")
    asyncio.run(adapter.process_activity(activity, auth_hdr, handle_activity))
    return Response(status=200)


@app.route("/", methods=["GET"])
def health():
    return "Bot is running with Clarify logic."


@app.route("/chat", methods=["GET"])
def chat():
    return send_from_directory(app.static_folder, "index.html")


# ─────────────────────────  MAIN  ────────────────────────────────────────────
if __name__ == "__main__":
    logging.info("🚀 Bot started with dynamic CLARIFY logic.")
    app.run(host="0.0.0.0", port=3978)
