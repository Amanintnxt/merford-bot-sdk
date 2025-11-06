# ──────────────────────────────────────────────────────────────────────────────
# bot.py – Teams / Direct Line bridge to Azure OpenAI Assistants (SSO-first)
# Clarifying questions: JSON-first, text-only (no buttons), multi-turn follow-up
# with guards to avoid echoing questions or asking for “which document?”
# ──────────────────────────────────────────────────────────────────────────────
import os
import re
import json
import asyncio
import logging
import time
import requests
import openai
from dotenv import load_dotenv
from flask import Flask, request, Response, jsonify, send_from_directory, render_template
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter, TurnContext
from botbuilder.schema import Activity, Attachment, CardAction, ActionTypes, OAuthCard
from PyPDF2 import PdfReader
from openai import AzureOpenAI

# ─────────────────────────  ENV & OPENAI CONFIG  ────────────────────────────
load_dotenv()

APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_EP = os.getenv("AZURE_OPENAI_ENDPOINT")
OAUTH_CONNECTION = os.getenv("OAUTH_CONNECTION_NAME", "TeamsSSO")
DIRECT_LINE_SECRET = os.getenv("DIRECT_LINE_SECRET", "")
ADMIN_SECRET = os.getenv("ADMIN_SECRET")  # 🔐 Optional secret for /upload

openai.api_type = "azure"
openai.api_version = "2024-05-01-preview"
openai.api_key = AZURE_OPENAI_KEY
openai.azure_endpoint = AZURE_OPENAI_EP.rstrip("/")

client = AzureOpenAI(
    api_key=AZURE_OPENAI_KEY,
    azure_endpoint=AZURE_OPENAI_EP,
    api_version="2024-05-01-preview"
)

# ─────────────────────────  ASSISTANT IDS (exact!)  ──────────────────────────
ASSISTANT_MAP = {
    "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
    "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
    "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
    "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
}

# Vector store IDs
VECTOR_STORES = {
    "Level 1": "vs_ICYlowKd3PPqtSp4m4wPzD47",
    "Level 2": "vs_FeOttDiAigZaxb8fjp1rAOIF",
    "Level 3": "vs_tO6kScvWu6oBn5R8YqeDkIX1",
    "Level 4": "vs_PJIPiZ91ojScAfJmKSCHrvx2"
}

# ─────────────────────────  FLASK & BOT ADAPTER  ─────────────────────────────
app = Flask(__name__, static_folder="static", template_folder="templates")
app.logger.handlers = logging.getLogger().handlers
app.logger.setLevel(logging.INFO)
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# ─────────────────────────  IN-MEMORY STATE  ────────────────────────────────
thread_map = {}        # key = f"{user_id}:{assistant_id}" → thread_id
awaiting_clarification = set()  # user_ids awaiting a reply to a clarify question

# ─────────────────────────  CLARIFY HELPERS  ────────────────────────────────


def _is_clarify(text: str) -> bool:
    return bool(text and text.strip().upper().startswith("CLARIFY:"))


def _looks_like_clarify(text: str) -> bool:
    """Detect clarifying intent even without 'CLARIFY:' prefix."""
    if not text:
        return False
    t = text.strip().lower()
    if "?" in t and len(t) <= 260:
        return True
    cues = (
        "please specify", "please provide", "could you specify", "can you specify",
        "which model", "which configuration", "what model", "what configuration",
        "clarify", "provide model", "provide certificate", "need more information",
        "missing information", "not enough information"
    )
    return any(c in t for c in cues)


def _strip_clarify(text: str) -> str:
    return text[len("CLARIFY:"):].strip() if _is_clarify(text) else (text or "")


# JSON clarify extraction (tolerant: plain JSON, fenced, or embedded)
JSON_BLOCK_RE = re.compile(
    r"```json\s*(\{[\s\S]*?\})\s*```|^(\{[\s\S]*\})$", re.MULTILINE)
JSON_ANY_RE = re.compile(r"\{[^{}]*\"clarify\"\s*:\s*\{[\s\S]*?\}\s*\}")


def _extract_clarify_from_reply(reply: str):
    """
    Extract {"clarify":{"question":..., "options":[...]}} from assistant reply.
    Returns (question, options) or (None, None).
    """
    if not reply:
        return None, None
    txt = reply.strip()

    m = JSON_BLOCK_RE.search(txt)
    raw = (m.group(1) or m.group(2)) if m else None
    if not raw:
        m2 = JSON_ANY_RE.search(txt)
        raw = m2.group(0) if m2 else None
    if not raw:
        return None, None

    try:
        obj = json.loads(raw)
    except Exception:
        return None, None

    c = obj.get("clarify")
    if not isinstance(c, dict):
        return None, None

    q = (c.get("question") or "").strip()
    if not q:
        return None, None

    opts = c.get("options") or []
    if not isinstance(opts, list):
        opts = []
    opts = [str(o).strip() for o in opts if str(o).strip()]
    return q, opts


def _needs_followup_clarify(user_text: str, reply: str) -> bool:
    """
    Ask again if the user query is ambiguous AND the answer is generic/broad/option-heavy.
    """
    if not reply:
        return False
    ut, rp = (user_text or "").lower(), reply.lower()
    ambiguous = any(k in ut for k in (
        "what if", "is it possible", "can we", "alternative", "options",
        "non-standard", "zone", "atex", "configuration", "model",
        "rc2", "rc3", "strike plate", "el560", "top jamb", "fluid pressure", "door closer"
    ))
    generic = any(k in rp for k in (
        "options", "alternative", "may", "might", "depends", "recommended",
        "consider", "varies", "multiple", "various", "several", "it depends"
    ))
    bullets = sum(1 for l in reply.splitlines()
                  if l.strip().startswith(("-", "*", "1.", "2.", "3.")))
    return ambiguous and (generic or bullets >= 2)


# Guards: avoid asking for documents or echoing the user's question
def _is_doc_request(q: str) -> bool:
    if not q:
        return False
    t = q.lower()
    return any(p in t for p in (
        "which document", "which report", "which certificate", "reference document",
        "report id", "doc id", "source document", "which test report"
    ))


def _is_echo(user_text: str, q: str) -> bool:
    if not user_text or not q:
        return False
    ut = " ".join((user_text or "").lower().split())
    qt = " ".join((q or "").lower().rstrip("?").split())
    if qt and (qt in ut or ut in qt):
        return True
    stop = {"the", "a", "an", "of", "for", "to",
            "is", "are", "with", "on", "in", "and", "or"}
    ut_tokens = set(ut.split()) - stop
    qt_tokens = set(qt.split()) - stop
    return len(qt_tokens & ut_tokens) >= max(3, int(0.7 * (len(qt_tokens) or 1)))


def _gen_domain_clarify(user_text: str) -> str:
    """Generate a targeted clarification for common Merford domains."""
    ut = (user_text or "").lower()
    if "rc2" in ut or "rc 2" in ut or "rc3" in ut or "rc 3" in ut:
        return ("Please specify the door model/series and configuration (single or double leaf), "
                "any glazing/openings, and the required width × height. Confirm RC class (RC2/RC3).")
    if "3153" in ut or ("height" in ut and "width" not in ut):
        return ("Do you need a single or double leaf, and which RC/Fire class applies? "
                "Please provide the required width and whether glazing or extra locking is present.")
    if "el560" in ut or "strike plate" in ut:
        return ("Which door model/series and escape function (EN 179 or EN 1125) applies? "
                "Is it single or double leaf, and is there additional locking?")
    if "atex" in ut or "zone" in ut:
        return ("Which ATEX classification applies (exact zone and gas group/temperature class), "
                "and which door model/configuration?")
    if "351m" in ut or "top jamb" in ut or "electric strike" in ut:
        return ("Which door series, single or double leaf, and locking set are intended? "
                "Confirm the certification/configuration the installation must comply with.")
    if "fluid" in ut or "diesel" in ut or "pressure" in ut:
        return ("Please specify pressure head (cm), opening size (width × height), seal type, "
                "and whether the door is internal or external.")
    return ("Please specify the door model/series and configuration (single/double leaf, glazing), "
            "the required dimensions, and the certification class (RC/Fire/ATEX) if applicable.")


# ─────────────────────────  GRAPH LOOK-UP  ───────────────────────────────────
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
        name = g.get("displayName", "")
        if "Level1Access" in name:
            return "Level 1"
        if "Level2Access" in name:
            return "Level 2"
        if "Level3Access" in name:
            return "Level 3"
        if "Level4Access" in name:
            return "Level 4"
    return None

# ─────────────────────────  PDF UPLOAD HELPERS  ─────────────────────────────


def is_pdf_text_based(file_path, min_text_length=10):
    """Check if a PDF contains searchable text."""
    try:
        reader = PdfReader(file_path)
        text_content = "".join(
            [page.extract_text() or "" for page in reader.pages])
        return len(text_content.strip()) >= min_text_length
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
            return "❌ Only PDF files allowed", 400

        save_path = os.path.join("uploads", file.filename)
        os.makedirs("uploads", exist_ok=True)
        file.save(save_path)

        if not is_pdf_text_based(save_path):
            os.remove(save_path)
            return "❌ Invalid PDF (image-only, no text).", 400

        if level == "Level 1":
            targets = ["Level 1", "Level 2", "Level 3", "Level 4"]
        elif level == "Level 2":
            targets = ["Level 2", "Level 3", "Level 4"]
        elif level == "Level 3":
            targets = ["Level 3", "Level 4"]
        else:
            targets = ["Level 4"]

        try:
            with open(save_path, "rb") as f:
                new_file = client.files.create(file=f, purpose="assistants")
            for tgt in targets:
                client.vector_stores.files.create(
                    vector_store_id=VECTOR_STORES[tgt],
                    file_id=new_file.id
                )
            logging.info("Uploaded %s to %s", file.filename, targets)
            return f"✅ Uploaded {file.filename} to {', '.join(targets)}"
        except Exception as e:
            logging.exception("Upload failed")
            return f"⚠️ Upload failed: {e}", 500
        finally:
            if os.path.exists(save_path):
                os.remove(save_path)

    return render_template("upload.html")

# ─────────────────────────  TOKEN HELPERS  ───────────────────────────────────


async def try_get_token(turn_context: TurnContext, magic_code: str | None = None):
    """Try to get a token using Teams SSO (silent) or magic code if present in invoke value."""
    try:
        return await adapter.get_user_token(turn_context, OAUTH_CONNECTION, magic_code)
    except Exception as e:
        logging.info(
            "get_user_token exception (will fall back to prompt): %s", e)
        return None


async def ensure_token(turn_context: TurnContext):
    """
    1) Try silent SSO (Teams)
    2) If magic code present in value.state (Web Chat), try with code
    3) If still none, send OAuthCard for sign-in
    """
    magic = None
    if turn_context.activity.value and isinstance(turn_context.activity.value, dict):
        magic = turn_context.activity.value.get("state")

    token_resp = await try_get_token(turn_context, magic)
    if token_resp and token_resp.token:
        return token_resp.token

    # Send sign-in card (CardAction used only for OAuth; no clarify buttons anywhere)
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
    logging.info("Sent sign-in card to %s",
                 turn_context.activity.from_property.id)
    return None

# ─────────────────────────  CORE BOT HANDLER  ────────────────────────────────


async def handle_activity(turn_context: TurnContext):
    a = turn_context.activity
    user_id = (a.from_property.id or "unknown")

    # 1) Teams conversation start
    if a.type == "conversationUpdate":
        for m in a.members_added or []:
            if m.id == a.recipient.id:
                await turn_context.send_activity(
                    "✅ Connected. If you're in Teams, SSO is automatic. In Web Chat, click Sign In once."
                )
        return

    # 2) Teams SSO invokes
    if a.type == "invoke" and a.name in ("signin/verifyState", "tokenExchange"):
        token = await try_get_token(turn_context)
        if token:
            await turn_context.send_activity("🔐 You're signed in. Ask your question")
        return

    # 3) Regular messages only
    if a.type != "message":
        return

    user_text = (a.text or "").strip()
    if not user_text and not (a.value and a.value.get("state")):
        return

    # 4) Acquire token
    access_token = await ensure_token(turn_context)
    if not access_token:
        return  # waiting for sign-in

    # 5) Resolve level → assistant
    level = get_user_group_level(access_token)
    logging.info("User %s level = %s", user_id, level)
    assistant_id = ASSISTANT_MAP.get(level)
    if not assistant_id:
        await turn_context.send_activity("❌ No assistant mapped to your access. Please contact admin.")
        return

    # 6) Thread isolation
    key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(key)
    if not thread_id:
        thread_id = openai.beta.threads.create().id
        thread_map[key] = thread_id
        logging.info("Created thread %s for %s", thread_id, key)

    # 7) Add user message
    if user_text:
        content_to_send = user_text
        if user_id in awaiting_clarification:
            logging.info("↩️ Clarification received from %s: %s",
                         user_id, user_text)
            content_to_send = f"(User clarification) {user_text}"
            awaiting_clarification.discard(user_id)

        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=content_to_send
        )

    # 8) Run assistant (file_search enforced; assistant should have temperature=0)
    try:
        run = openai.beta.threads.runs.create(
            assistant_id=assistant_id,
            thread_id=thread_id,
            tool_choice={"type": "file_search"}
        )
    except Exception as e:
        logging.exception("Assistant run create failed")
        await turn_context.send_activity(f"❌ Assistant run failed: {e}")
        return

    start = time.time()
    while run.status not in ("completed", "failed", "cancelled"):
        if time.time() - start > 60:
            await turn_context.send_activity("⏳ Still working… please resend if no reply arrives.")
            break
        await asyncio.sleep(1)
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)

    # 9) Fetch newest assistant message
    try:
        msgs = openai.beta.threads.messages.list(
            thread_id=thread_id, order="desc", limit=5)
        reply = next(
            (m.content[0].text.value for m in msgs.data if m.role == "assistant"), None)
    except Exception:
        logging.exception("Fetch messages failed")
        reply = None

    # 9.1) Clarify-first handling (JSON preferred → text only)
    q_json, _opts = _extract_clarify_from_reply(reply or "")
    if q_json:
        question = q_json
    else:
        question = None
        if reply and (_is_clarify(reply) or _looks_like_clarify(reply)):
            question = _strip_clarify(reply) if _is_clarify(reply) else reply

    if question:
        # Guard: replace bad clarifications (echo/doc-request) with domain-specific one
        if _is_doc_request(question) or _is_echo(user_text, question):
            question = _gen_domain_clarify(user_text)
        logging.info("🟡 CLARIFY for %s: %s", user_id, question)
        awaiting_clarification.add(user_id)
        await turn_context.send_activity(Activity(type="message", text=question))
        return

    # 9.2) Post-answer ambiguity check (research done → still ambiguous? ask again)
    if reply and _needs_followup_clarify(user_text, reply):
        follow_q = _gen_domain_clarify(user_text)
        logging.info("🟡 POST-CLARIFY for %s: %s", user_id, follow_q)
        awaiting_clarification.add(user_id)
        await turn_context.send_activity(Activity(type="message", text=follow_q))
        return

    # 10) Send final
    await turn_context.send_activity(reply or "❌ No reply from assistant.")

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
