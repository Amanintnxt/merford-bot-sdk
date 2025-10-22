# ──────────────────────────────────────────────────────────────────────────────
# bot.py – Teams / Direct Line bridge to Azure OpenAI Assistants (SSO-first)
# Optimized version with dynamic CLARIFY logic and SSO magic code fix
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

# ──────────────── IN-MEMORY STATE ────────────────
thread_map = {}              # user_id:assistant_id → thread_id
awaiting_clarification = set()  # track users awaiting clarification reply

# ──────────────── CLARIFY LOGIC HELPERS ────────────────


def _is_clarify(text: str) -> bool:
    return bool(text and text.strip().upper().startswith("CLARIFY:"))


def _looks_like_clarify(text: str) -> bool:
    """Detect if assistant message is probably a clarifying question."""
    if not text:
        return False
    t = text.strip().lower()
    if "?" not in t:
        return False
    starts = ("what", "which", "can you", "could you",
              "please", "clarify", "do you")
    return any(t.startswith(s) for s in starts) and len(t) < 200


def _strip_clarify(text: str) -> str:
    return text[len("CLARIFY:"):].strip() if _is_clarify(text) else text


def _clarify_actions(question: str):
    """Generate context-aware quick replies."""
    lower = question.lower()
    if "model" in lower or "product" in lower:
        opts = ["Model: M-series", "Model: Unknown", "Provide model later"]
    elif "configuration" in lower:
        opts = ["Single leaf", "Double leaf", "Not sure"]
    elif "zone" in lower:
        opts = ["Zone 2 IIB T2", "Other zone", "Not sure"]
    else:
        opts = ["I'll specify", "Please repeat question", "Cancel"]
    return [CardAction(type=ActionTypes.im_back, title=o, value=o) for o in opts]


def _needs_followup_clarify(user_text: str, reply: str) -> bool:
    """Trigger extra clarification if answer looks too generic."""
    if not reply:
        return False
    ut, rp = user_text.lower(), reply.lower()
    ambiguous = any(k in ut for k in ("what if", "can we",
                    "options", "model", "alternative", "configuration"))
    generic = any(k in rp for k in ("may", "depends",
                  "options", "can be", "recommended", "varies"))
    bullets = sum(1 for l in reply.splitlines()
                  if l.strip().startswith(("-", "*", "1.")))
    return ambiguous and (generic or bullets >= 3)

# ──────────────── GRAPH LOOKUP ────────────────


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

# ──────────────── PDF UPLOAD ROUTE ────────────────


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

# ──────────────── CORE BOT HANDLER ────────────────


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

    # 2️⃣ Detect OAuth magic code
    if user_text.isdigit() and len(user_text) <= 10:
        logging.info(f"🔐 OAuth magic code detected: {user_text}")
        token = await try_get_token(turn_context, user_text)
        if token and token.token:
            await turn_context.send_activity("🔓 Sign-in successful! You can now ask your question.")
        else:
            await turn_context.send_activity("⚠️ Sign-in failed. Please click Sign In again.")
        return

    # 3️⃣ Get access token
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

    # 5️⃣ Add message to thread
    if user_text:
        if user_id in awaiting_clarification:
            user_text = f"(User clarification) {user_text}"
            awaiting_clarification.discard(user_id)
        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=user_text)

    # 6️⃣ Run assistant
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

    # 7️⃣ Get reply
    msgs = openai.beta.threads.messages.list(
        thread_id=thread_id, order="desc", limit=5)
    reply = next(
        (m.content[0].text.value for m in msgs.data if m.role == "assistant"), None)

    # 8️⃣ Clarify handling
    if reply and (_is_clarify(reply) or _looks_like_clarify(reply)):
        question = _strip_clarify(reply)
        awaiting_clarification.add(user_id)
        await turn_context.send_activity(Activity(
            type="message",
            text=question,
            suggested_actions=SuggestedActions(
                actions=_clarify_actions(question))
        ))
        return

    # 9️⃣ Post-answer ambiguity check
    if reply and _needs_followup_clarify(user_text, reply):
        question = "To provide the most accurate document-based answer, could you clarify the model or configuration?"
        awaiting_clarification.add(user_id)
        await turn_context.send_activity(Activity(
            type="message",
            text=question,
            suggested_actions=SuggestedActions(
                actions=_clarify_actions(question))
        ))
        return

    await turn_context.send_activity(reply or "❌ No response from assistant.")

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
        logging.error("Direct Line token generation failed: %s", r.text)
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
    logging.info("🚀 Bot is starting on Render...")
    logging.info("🔧 Environment check:")
    logging.info("  MicrosoftAppId: %s", "SET" if APP_ID else "MISSING")
    logging.info("  Azure OpenAI Endpoint: %s", AZURE_OPENAI_EP or "MISSING")
    logging.info("  OAuth Connection: %s", OAUTH_CONNECTION or "MISSING")
    logging.info("  Direct Line Secret: %s",
                 "SET" if DIRECT_LINE_SECRET else "MISSING")
    logging.info("  Admin Secret: %s", "SET" if ADMIN_SECRET else "NOT SET")

    app.run(host="0.0.0.0", port=3978)
