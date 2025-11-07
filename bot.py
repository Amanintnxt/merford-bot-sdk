# ──────────────────────────────────────────────────────────────────────────────
# bot.py – Azure Teams/Direct Line bridge to OpenAI Assistants (multi-turn clarify + source-aware)
# ──────────────────────────────────────────────────────────────────────────────
import os
import asyncio
import logging
import time
import re
import requests
from dotenv import load_dotenv
from flask import Flask, request, Response, jsonify, render_template, send_from_directory
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter, TurnContext
from botbuilder.schema import Activity, Attachment, CardAction, ActionTypes, OAuthCard
from PyPDF2 import PdfReader
from openai import AzureOpenAI

# ──────────────── ENV CONFIG ────────────────
load_dotenv()

APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_EP = os.getenv("AZURE_OPENAI_ENDPOINT")
OAUTH_CONNECTION = os.getenv("OAUTH_CONNECTION_NAME", "TeamsSSO")
DIRECT_LINE_SECRET = os.getenv("DIRECT_LINE_SECRET", "")
ADMIN_SECRET = os.getenv("ADMIN_SECRET")

client = AzureOpenAI(
    api_key=AZURE_OPENAI_KEY,
    azure_endpoint=AZURE_OPENAI_EP,
    api_version="2024-05-01-preview"
)

# ──────────────── ASSISTANT & VECTOR STORE MAP ────────────────
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

# ──────────────── FLASK & BOT ADAPTER ────────────────
app = Flask(__name__, static_folder="static", template_folder="templates")
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# ──────────────── STATE TRACKING ────────────────
thread_map = {}
awaiting_clarification = set()

# ──────────────── HELPER FUNCTIONS ────────────────


def _is_question(text: str) -> bool:
    """Detects if text is a question."""
    if not text:
        return False
    q = text.lower().strip()
    if "?" in q:
        return True
    question_words = ["what", "which", "how",
                      "is", "can", "does", "are", "should"]
    return any(q.startswith(w) for w in question_words)


def _is_clarify(text: str) -> bool:
    return text.strip().upper().startswith("CLARIFY:")


def _strip_clarify(text: str) -> str:
    return text[len("CLARIFY:"):].strip() if _is_clarify(text) else text


def _looks_generic(reply: str) -> bool:
    if not reply:
        return False
    generic_words = ["may", "depends", "varies", "can be", "recommended"]
    return any(w in reply.lower() for w in generic_words)


def _conflict_detected(text: str) -> bool:
    """Detect conflicting model/spec terms (e.g., M-series + RC3)."""
    patterns = [("m-series", "rc3"), ("m-series", "rc2"), ("rc3", "rc2")]
    t = text.lower()
    return any(a in t and b in t for a, b in patterns)


def _extract_source(reply: str) -> str:
    """Detect simple source pattern."""
    match = re.search(r"Source:\s?([A-Za-z0-9_\-]+\.pdf)", reply)
    return match.group(1) if match else ""


def _clarify_needed(user_text: str, reply: str) -> bool:
    """Decide whether to ask clarifying question again."""
    if not reply:
        return False
    if _conflict_detected(user_text):
        return True
    if _is_question(user_text):
        return True
    if _looks_generic(reply):
        return True
    return False

# ──────────────── GRAPH HELPERS ────────────────


def get_user_group_level(token: str) -> str | None:
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

# ──────────────── PDF UPLOAD (unchanged) ────────────────


def is_pdf_text_based(path, min_len=10):
    try:
        text = "".join([p.extract_text() or "" for p in PdfReader(path).pages])
        return len(text.strip()) > min_len
    except Exception:
        return False


@app.route("/upload", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        if ADMIN_SECRET and request.form.get("secret") != ADMIN_SECRET:
            return "Unauthorized", 403
        level = request.form["level"]
        file = request.files["file"]
        if not file.filename.endswith(".pdf"):
            return "❌ Only PDFs allowed", 400

        os.makedirs("uploads", exist_ok=True)
        path = os.path.join("uploads", file.filename)
        file.save(path)

        if not is_pdf_text_based(path):
            os.remove(path)
            return "❌ Invalid (image-only) PDF", 400

        targets = {
            "Level 1": ["Level 1", "Level 2", "Level 3", "Level 4"],
            "Level 2": ["Level 2", "Level 3", "Level 4"],
            "Level 3": ["Level 3", "Level 4"],
            "Level 4": ["Level 4"],
        }[level]

        try:
            with open(path, "rb") as f:
                new_file = client.files.create(file=f, purpose="assistants")
            for tgt in targets:
                client.vector_stores.files.create(
                    vector_store_id=VECTOR_STORES[tgt],
                    file_id=new_file.id
                )
            return f"✅ Uploaded {file.filename} to {', '.join(targets)}"
        finally:
            os.remove(path)
    return render_template("upload.html")

# ──────────────── TOKEN HELPERS ────────────────


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

# ──────────────── MAIN HANDLER ────────────────


async def handle_activity(turn_context: TurnContext):
    a = turn_context.activity
    user_id = a.from_property.id
    user_text = (a.text or "").strip()

    if a.type == "conversationUpdate":
        for m in a.members_added or []:
            if m.id == a.recipient.id:
                await turn_context.send_activity("✅ Connected. Please sign in to continue.")
        return

    if a.type != "message":
        return

    # Magic code
    if user_text.isdigit() and len(user_text) <= 10:
        token = await try_get_token(turn_context, user_text)
        if token and token.token:
            await turn_context.send_activity("🔓 Sign-in successful! You can now ask your question.")
        else:
            await turn_context.send_activity("⚠️ Sign-in failed. Please click Sign In again.")
        return

    access_token = await ensure_token(turn_context)
    if not access_token:
        return

    level = get_user_group_level(access_token)
    assistant_id = ASSISTANT_MAP.get(level)
    if not assistant_id:
        await turn_context.send_activity("❌ No assistant assigned for your access level.")
        return

    key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(key)
    if not thread_id:
        thread_id = client.beta.threads.create().id
        thread_map[key] = thread_id

    # User clarification handling
    if user_text:
        if user_id in awaiting_clarification:
            user_text = f"(User clarification) {user_text}"
            awaiting_clarification.discard(user_id)
        client.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=user_text)

    run = client.beta.threads.runs.create(
        assistant_id=assistant_id,
        thread_id=thread_id,
        tool_choice={"type": "file_search"}
    )

    start = time.time()
    while run.status not in ("completed", "failed", "cancelled"):
        if time.time() - start > 60:
            await turn_context.send_activity("⏳ Still processing... please try again shortly.")
            return
        await asyncio.sleep(1)
        run = client.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)

    msgs = client.beta.threads.messages.list(
        thread_id=thread_id, order="desc", limit=5)
    reply = next(
        (m.content[0].text.value for m in msgs.data if m.role == "assistant"), None)

    # ─ Clarify / follow-up handling ─
    if reply:
        if _is_clarify(reply) or _clarify_needed(user_text, reply):
            question = _strip_clarify(reply) if _is_clarify(reply) else reply
            if not question.endswith("?"):
                question += "?"
            awaiting_clarification.add(user_id)
            await turn_context.send_activity(f"CLARIFY: {question}")
            return

        # Append source tag if missing
        if not _extract_source(reply):
            reply += "\n\n(Source: internal documentation)"

    await turn_context.send_activity(reply or "❌ No response from assistant.")

# ──────────────── ROUTES ────────────────


@app.route("/api/messages", methods=["POST"])
def messages():
    if "application/json" not in request.headers.get("Content-Type", ""):
        return Response("Unsupported Media Type", 415)
    activity = Activity().deserialize(request.json)
    auth_hdr = request.headers.get("Authorization", "")

    async def _proc():
        return await adapter.process_activity(activity, auth_hdr, handle_activity)

    asyncio.run(_proc())
    return Response(status=200)


@app.route("/chat", methods=["GET"])
def chat():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/", methods=["GET"])
def health():
    return "Bot is running."


# ──────────────── MAIN ────────────────
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    logging.info("🚀 Starting bot with global clarify + source logic...")
    app.run(host="0.0.0.0", port=3978)
