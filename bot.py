# ──────────────────────────────────────────────────────────────────────────────
# bot.py – Azure Bot + Teams/Direct Line integration with Clarify Logic
# Updated version: dynamic conversational CLARIFY flow (no action cards)
# PDF upload and parsing logic preserved exactly
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
from botbuilder.schema import Activity, Attachment, OAuthCard, CardAction, ActionTypes
from PyPDF2 import PdfReader
from openai import AzureOpenAI

# ──────────────── ENV & OPENAI CONFIG ────────────────
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

# ──────────────── ASSISTANT MAP ────────────────
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

# ──────────────── IN-MEMORY STATE ────────────────
thread_map = {}                 # user_id:assistant_id → thread_id
awaiting_clarification = set()  # track users awaiting clarification

# ──────────────── CLARIFY LOGIC HELPERS ────────────────


def _is_clarify(text: str) -> bool:
    """Detect explicit CLARIFY marker from Assistant instruction."""
    return bool(text and text.strip().upper().startswith("CLARIFY:"))


def _looks_like_clarify(text: str) -> bool:
    """Detect question-like messages from the Assistant."""
    if not text:
        return False
    t = text.strip().lower()
    if "?" not in t:
        return False
    return any(t.startswith(s) for s in ["what", "which", "can you", "could you", "please", "clarify", "do you"])


def _strip_clarify(text: str) -> str:
    return text[len("CLARIFY:"):].strip() if _is_clarify(text) else text


def _needs_followup_clarify(user_text: str, reply: str) -> bool:
    """Detect when a follow-up clarification is needed."""
    if not reply:
        return False
    ut, rp = user_text.lower(), reply.lower()
    ambiguous = any(k in ut for k in ("what if", "can we",
                    "options", "model", "alternative", "configuration"))
    generic = any(k in rp for k in ("may", "depends", "can be",
                  "recommended", "varies", "possible"))
    return ambiguous and generic

# ──────────────── GRAPH GROUP CHECK ────────────────


def get_user_group_level(token: str) -> str | None:
    """Fetch group membership from Graph API."""
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

# ──────────────── PDF UPLOAD ROUTE (UNCHANGED) ────────────────


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

# ──────────────── MAIN BOT LOGIC ────────────────


async def handle_activity(turn_context: TurnContext):
    a = turn_context.activity
    user_id = a.from_property.id
    user_text = (a.text or "").strip()

    # 1️⃣ On conversation join
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
            await turn_context.send_activity("🔓 Sign-in successful! You can now ask your question.")
        else:
            await turn_context.send_activity("⚠️ Sign-in failed. Please click Sign In again.")
        return

    # 3️⃣ Retrieve token
    access_token = await ensure_token(turn_context)
    if not access_token:
        return

    # 4️⃣ Resolve group → assistant
    level = get_user_group_level(access_token)
    assistant_id = ASSISTANT_MAP.get(level)
    if not assistant_id:
        await turn_context.send_activity("❌ No assistant assigned for your access level.")
        return

    key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(key)
    if not thread_id:
        thread_id = openai.beta.threads.create().id
        thread_map[key] = thread_id

    # 5️⃣ Handle clarification flow
    if user_id in awaiting_clarification:
        user_text = f"(Clarification) {user_text}"
        awaiting_clarification.discard(user_id)

    # 6️⃣ Send user message to thread
    openai.beta.threads.messages.create(
        thread_id=thread_id,
        role="user",
        content=user_text
    )

    # 7️⃣ Run the assistant
    run = openai.beta.threads.runs.create(
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
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)

    # 8️⃣ Fetch assistant reply
    msgs = openai.beta.threads.messages.list(
        thread_id=thread_id, order="desc", limit=5)
    reply = next(
        (m.content[0].text.value for m in msgs.data if m.role == "assistant"), None)

    # 9️⃣ Clarify detection
    if reply and (_is_clarify(reply) or _looks_like_clarify(reply)):
        question = _strip_clarify(reply)
        awaiting_clarification.add(user_id)
        await turn_context.send_activity(f"CLARIFY: {question}")
        return

    # 🔟 Fallback clarify if response is generic
    if reply and _needs_followup_clarify(user_text, reply):
        awaiting_clarification.add(user_id)
        await turn_context.send_activity(
            "CLARIFY: Could you please specify the exact model, configuration, or report reference?"
        )
        return

    # 11️⃣ Final fallback if no response or missing data
    if not reply or "not available" in reply.lower():
        await turn_context.send_activity(
            "We regret to inform that this information is not available in the provided documentation. "
            "You can contact us directly at contact@merford.com for further details."
        )
        return

    await turn_context.send_activity(reply)

# ─────────────────────────  FLASK ROUTES  ────────────────────────────────────


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


@app.route("/directline/token", methods=["POST"])
def directline_token():
    if not DIRECT_LINE_SECRET:
        return jsonify({"error": "DIRECT_LINE_SECRET not set"}), 500
    r = requests.post(
        "https://directline.botframework.com/v3/directline/tokens/generate",
        headers={"Authorization": f"Bearer {DIRECT_LINE_SECRET}"},
        timeout=10
    )
    if r.status_code != 200:
        return jsonify({"error": "Failed to generate token"}), 500
    return jsonify({"token": r.json().get("token")})


@app.route("/chat", methods=["GET"])
def chat():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/", methods=["GET"])
def health():
    return "Bot is running."

# ─────────────────────────  MAIN  ────────────────────────────────────────────


if __name__ == "__main__":
    logging.info("🚀 Bot starting...")
    app.run(host="0.0.0.0", port=3978)
