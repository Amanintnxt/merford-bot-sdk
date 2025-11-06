# ──────────────────────────────────────────────────────────────────────────────
# bot.py – Teams / Direct Line bridge to Azure OpenAI Assistants (SSO-first)
# Multi-turn clarification loop, reliability check, and source fallback
# PDF upload/parsing functions left unchanged
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
# thread_map: key = f"{user_id}:{assistant_id}" -> thread_id
thread_map: dict[str, str] = {}
# pending clarifications per user:
# pending_clarify[user_id] = {"thread_id":..., "assistant_id":..., "rounds": int, "original": str}
pending_clarify: dict[str, dict] = {}

# limits
MAX_CLARIFY_ROUNDS = 3
RETRY_ON_GENERIC = True  # re-query once if assistant reply looks 'generic/uncertain'

# ──────────────── LOGGING ────────────────
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("merford-bot")

# ──────────────── CLARIFY LOGIC HELPERS ────────────────


def _is_clarify(text: str) -> bool:
    return bool(text and text.strip().upper().startswith("CLARIFY:"))


def _looks_like_clarify(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    if "?" not in t:
        return False
    starts = ("what", "which", "can you", "could you",
              "please", "clarify", "do you")
    return any(t.startswith(s) for s in starts) and len(t) < 300


def _strip_clarify(text: str) -> str:
    return text[len("CLARIFY:"):].strip() if _is_clarify(text) else text


def _reply_is_generic(text: str) -> bool:
    if not text:
        return True
    t = text.lower()
    markers = ("may", "might", "could", "depends",
               "possible", "recommend", "consider")
    return any(m in t for m in markers)

# ──────────────── GRAPH GROUP CHECK ────────────────


def get_user_group_level(token: str) -> str | None:
    """Fetch group membership from Graph API using /me/memberOf."""
    url = "https://graph.microsoft.com/v1.0/me/memberOf?$select=displayName"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code != 200:
            logger.warning("Graph /me/memberOf failed: %s %s",
                           resp.status_code, resp.text)
            return None
        for g in resp.json().get("value", []):
            name = g.get("displayName", "") or ""
            logger.info("AAD group found: %s", name)
            if "Level1Access" in name:
                return "Level 1"
            if "Level2Access" in name:
                return "Level 2"
            if "Level3Access" in name:
                return "Level 3"
            if "Level4Access" in name:
                return "Level 4"
    except Exception as ex:
        logger.exception("Graph lookup error: %s", ex)
        return None
    return None

# ──────────────── PDF UPLOAD ROUTE (UNCHANGED LOGIC) ────────────────


def is_pdf_text_based(path, min_len=10):
    try:
        text = "".join([p.extract_text() or "" for p in PdfReader(path).pages])
        return len(text.strip()) > min_len
    except Exception:
        return False


@app.route("/upload", methods=["GET", "POST"])
def upload_file():
    # Keep behavior identical to your earlier upload handler
    if request.method == "POST":
        if ADMIN_SECRET and request.form.get("secret") != ADMIN_SECRET:
            return "Unauthorized", 403
        level = request.form["level"]
        file = request.files["file"]
        if not file.filename.lower().endswith(".pdf"):
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
            logger.info("Uploaded %s to %s", file.filename, targets)
            return f"✅ Uploaded {file.filename} to {', '.join(targets)}"
        except Exception as e:
            logger.exception("Upload failed")
            return f"⚠️ Upload failed: {e}", 500
        finally:
            if os.path.exists(path):
                os.remove(path)
    return render_template("upload.html")

# ──────────────── TOKEN HELPERS ────────────────


async def try_get_token(turn_context: TurnContext, magic_code=None):
    try:
        return await adapter.get_user_token(turn_context, OAUTH_CONNECTION, magic_code)
    except Exception as e:
        logger.info("get_user_token exception: %s", e)
        return None


async def ensure_token(turn_context: TurnContext):
    magic = None
    if turn_context.activity.value and isinstance(turn_context.activity.value, dict):
        magic = turn_context.activity.value.get("state")
    if not magic and turn_context.activity.text and turn_context.activity.text.strip().isdigit():
        magic = turn_context.activity.text.strip()
    token_resp = await try_get_token(turn_context, magic)
    if token_resp and token_resp.token:
        return token_resp.token

    # send oauth card
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
    logger.info("Sent sign-in card to %s",
                getattr(turn_context.activity.from_property, "id", "unknown"))
    return None

# ──────────────── CORE: multi-turn clarification loop ────────────────


async def handle_activity(turn_context: TurnContext):
    a = turn_context.activity
    user_id = (a.from_property.id or "unknown")
    user_text = (a.text or "").strip()

    # 1) Conversation join
    if a.type == "conversationUpdate":
        for m in a.members_added or []:
            if m.id == a.recipient.id:
                await turn_context.send_activity("✅ Connected. In Teams SSO is automatic; in Web Chat click Sign In.")
        return

    # only handle messages
    if a.type != "message":
        return

    # 2) OAuth magic code handling (user pastes the code)
    if user_text.isdigit() and len(user_text) <= 10:
        token = await try_get_token(turn_context, user_text)
        if token and token.token:
            await turn_context.send_activity("🔓 Sign-in successful! You can now ask your question.")
        else:
            await turn_context.send_activity("⚠️ Sign-in failed. Please click Sign In again.")
        return

    # 3) Ensure token (SSO)
    access_token = await ensure_token(turn_context)
    if not access_token:
        return  # sign-in prompt sent, wait for user

    # 4) Resolve user level and assistant
    level = get_user_group_level(access_token)
    assistant_id = ASSISTANT_MAP.get(level)
    logger.info("User %s resolved to level=%s assistant=%s",
                user_id, level, assistant_id)
    if not assistant_id:
        await turn_context.send_activity("❌ No assistant assigned for your access level. Contact admin.")
        return

    # 5) Prepare thread (per user+assistant isolation)
    key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(key)
    if not thread_id:
        try:
            thread_id = openai.beta.threads.create().id
            thread_map[key] = thread_id
            logger.info("Created new thread %s for key %s", thread_id, key)
        except Exception as e:
            logger.exception("Failed to create assistant thread")
            await turn_context.send_activity("❌ Failed to create assistant session.")
            return

    # 6) Clarification state handling:
    # If user has pending clarification, treat current message as clarification reply
    if user_id in pending_clarify:
        pc = pending_clarify[user_id]
        # ensure we are talking to the same assistant/thread
        if pc.get("assistant_id") != assistant_id:
            # different assistant -> clear pending clarify and continue as new query
            logger.info(
                "Assistant changed during pending clarify; clearing pending state for %s", user_id)
            pending_clarify.pop(user_id, None)
        else:
            # use this message as clarification input and continue the loop
            logger.info("Received clarification reply from %s (round %s): %s",
                        user_id, pc.get("rounds"), user_text)
            # send clarification to assistant and continue (prefix to show it's clarification)
            clarification_msg = f"(Clarification) {user_text}"
            openai.beta.threads.messages.create(
                thread_id=thread_id, role="user", content=clarification_msg)
            # increase round
            pc["rounds"] += 1
            # run assistant again and loop handling below
    else:
        # new question: add user message as new content
        if user_text:
            openai.beta.threads.messages.create(
                thread_id=thread_id, role="user", content=user_text)

    # 7) Run assistant (tool_choice=file_search)
    try:
        run = openai.beta.threads.runs.create(
            assistant_id=assistant_id,
            thread_id=thread_id,
            tool_choice={"type": "file_search"}
        )
    except Exception as e:
        logger.exception("Assistant run create failed")
        await turn_context.send_activity(f"❌ Assistant run failed: {e}")
        return

    # 8) Poll for run completion (with timeout)
    start = time.time()
    while run.status not in ("completed", "failed", "cancelled"):
        if time.time() - start > 60:
            await turn_context.send_activity("⏳ Still processing... please try again shortly.")
            return
        await asyncio.sleep(1)
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)

    # 9) Fetch assistant reply (most recent assistant message)
    try:
        msgs = openai.beta.threads.messages.list(
            thread_id=thread_id, order="desc", limit=8)
        reply = next(
            (m.content[0].text.value for m in msgs.data if m.role == "assistant"), None)
    except Exception:
        logger.exception("Failed to fetch assistant messages")
        reply = None

    logger.info("Assistant reply for %s: %s", user_id,
                (reply[:200] + "...") if reply and len(reply) > 200 else reply)

    # 10) Clarify detection: if assistant requests clarification, set pending state and ask user
    if reply and (_is_clarify(reply) or _looks_like_clarify(reply)):
        question = _strip_clarify(reply)
        # create or reset pending_clarify entry
        pending_clarify[user_id] = {
            "thread_id": thread_id,
            "assistant_id": assistant_id,
            "rounds": 0,
            "original": user_text
        }
        # send clarify to user (simple text, no cards)
        await turn_context.send_activity(f"CLARIFY: {question}")
        return

    # 11) If reply exists but seems generic/uncertain, optionally retry once to confirm (reliability)
    if reply and RETRY_ON_GENERIC and _reply_is_generic(reply):
        logger.info(
            "Reply looks generic; performing one reliability check for %s", user_id)
        # Add short verification prompt to assistant via user message
        verify_msg = "(Verify) Please confirm this answer strictly from the uploaded documents and include the exact source reference or say 'not available'."
        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=verify_msg)

        # run assistant verify
        try:
            run2 = openai.beta.threads.runs.create(
                assistant_id=assistant_id, thread_id=thread_id, tool_choice={"type": "file_search"})
            s2 = time.time()
            while run2.status not in ("completed", "failed", "cancelled"):
                if time.time() - s2 > 30:
                    break
                await asyncio.sleep(1)
                run2 = openai.beta.threads.runs.retrieve(
                    thread_id=thread_id, run_id=run2.id)
            msgs2 = openai.beta.threads.messages.list(
                thread_id=thread_id, order="desc", limit=6)
            verified = next(
                (m.content[0].text.value for m in msgs2.data if m.role == "assistant"), None)
            if verified:
                # prefer verified answer if it adds specificity
                reply = verified
        except Exception:
            logger.exception("Verification run failed; using original reply.")

    # 12) If we are in a pending clarification flow and assistant returned final answer (i.e., user answered a clarifying Q)
    if user_id in pending_clarify:
        # if assistant produced a non-clarify answer, clear pending state
        if reply and not (_is_clarify(reply) or _looks_like_clarify(reply)):
            logger.info(
                "Clarification resolved for %s, clearing pending state", user_id)
            pending_clarify.pop(user_id, None)
            # deliver the final reply below
        else:
            # assistant still asking for clarification -> increment rounds and check max attempts
            pc = pending_clarify[user_id]
            pc["rounds"] = pc.get("rounds", 0) + 1
            if pc["rounds"] >= MAX_CLARIFY_ROUNDS:
                pending_clarify.pop(user_id, None)
                await turn_context.send_activity("⚠️ I've asked for clarifications several times but still can't find a clear answer. Please rephrase or contact contact@merford.com.")
                return
            # if assistant still wants clarification, ask again
            if reply and (_is_clarify(reply) or _looks_like_clarify(reply)):
                question = _strip_clarify(reply)
                await turn_context.send_activity(f"CLARIFY: {question}")
                return

    # 13) Final fallback if no reply or reply indicates not available
    if not reply or "not available" in (reply or "").lower():
        await turn_context.send_activity(
            "We regret to inform that this information is not available in the provided documentation. "
            "You can contact us directly at contact@merford.com for further details."
        )
        return

    # 14) Append a short 'Source' line if assistant didn't include it.
    # We can't always reliably parse tool/file references from the thread messages here,
    # so we prefer that assistant includes Source: in its reply. As a fallback, show "Source: Not specified".
    if "source:" not in (reply or "").lower():
        reply = f"{reply}\n\nSource: Not specified in documents."

    # 15) Send final reply
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
        logger.exception("Exception in /api/messages: %s", ex)
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
        logger.error("Direct Line token generation failed: %s", r.text)
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
    logger.info("🚀 Bot starting...")
    logger.info("🔧 Environment check: MicrosoftAppId=%s, OAuth=%s",
                "SET" if APP_ID else "MISSING", OAUTH_CONNECTION)
    app.run(host="0.0.0.0", port=3978)
