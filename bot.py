# bot.py – Final merged with Teams SSO Welcome + User Name + DirectLine + Clarify + Source Logic

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

# -------------------- ENV & OPENAI --------------------
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
openai.azure_endpoint = AZURE_OPENAI_EP.rstrip(
    "/") if AZURE_OPENAI_EP else None

client = AzureOpenAI(
    api_key=AZURE_OPENAI_KEY,
    azure_endpoint=AZURE_OPENAI_EP,
    api_version="2024-05-01-preview",
)

# -------------------- ASSISTANTS & VECTOR STORES --------------------
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

# -------------------- FLASK & BOT ADAPTER --------------------
app = Flask(__name__, static_folder="static", template_folder="templates")
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# -------------------- IN-MEMORY STATE --------------------
thread_map = {}
pending_clarify = {}

MAX_CLARIFY_ROUNDS = 3
RETRY_ON_GENERIC = True  # verify run if generic answer

# -------------------- LOGGING --------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("merford-bot")

# -------------------- CLARIFY HELPERS --------------------


def _is_clarify(text: str) -> bool:
    return bool(text and text.strip().upper().startswith("CLARIFY:"))


def _looks_like_clarify(text: str) -> bool:
    if not text:
        return False
    t = text.strip().lower()
    if "?" not in t:
        return False
    starts = ("what", "which", "how", "can you",
              "could you", "do you", "please", "clarify")
    return any(t.startswith(s) for s in starts) and len(t) < 300


def _strip_clarify(text: str) -> str:
    return text[len("CLARIFY:"):].strip() if _is_clarify(text) else text.strip()


def _reply_is_generic(text: str) -> bool:
    if not text:
        return True
    t = text.lower()
    markers = ("may", "might", "could", "depends",
               "possible", "recommend", "varies")
    return any(m in t for m in markers)

# -------------------- GRAPH HELPERS --------------------


def get_user_group_level(token: str) -> str | None:
    """
    Returns Level 1-4 based on the group *IDs* stored in env.
    """

    # Load group IDs from environment
    G1 = os.getenv("LEVEL_ONE_GROUP")
    G2 = os.getenv("LEVEL_TWO_GROUP")
    G3 = os.getenv("LEVEL_THREE_GROUP")
    G4 = os.getenv("LEVEL_FOUR_GROUP")

    url = "https://graph.microsoft.com/v1.0/me/memberOf?$select=id,displayName"
    headers = {"Authorization": f"Bearer {token}"}

    try:
        resp = requests.get(url, headers=headers, timeout=10)
    except Exception as e:
        logger.warning("Graph request error: %s", e)
        return None

    if resp.status_code != 200:
        logger.warning("memberOf failed %s – %s", resp.status_code, resp.text)
        return None

    groups = resp.json().get("value", [])

    for g in groups:
        gid = g.get("id")

        if gid == G1:
            return "Level 1"
        if gid == G2:
            return "Level 2"
        if gid == G3:
            return "Level 3"
        if gid == G4:
            return "Level 4"

    return None

# -------------------- NEW: GET USER REAL NAME --------------------


def get_user_display_name(token: str) -> str | None:
    url = "https://graph.microsoft.com/v1.0/me?$select=displayName"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code == 200:
            return resp.json().get("displayName")
        else:
            logger.warning("Graph /me failed %s – %s",
                           resp.status_code, resp.text)
            return None
    except Exception as e:
        logger.warning("Graph /me error: %s", e)
        return None

# -------------------- PDF UPLOAD (unchanged) --------------------


def is_pdf_text_based(path, min_len=10):
    try:
        text = "".join([p.extract_text() or "" for p in PdfReader(path).pages])
        return len(text.strip()) >= min_len
    except Exception:
        return False


@app.route("/upload", methods=["GET", "POST"])
def upload_file():
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

# -------------------- TOKEN HELPERS --------------------


async def try_get_token(turn_context: TurnContext, magic_code=None):
    try:
        resp = await adapter.get_user_token(turn_context, OAUTH_CONNECTION, magic_code)
        if resp and resp.token:
            logger.info("get_user_token OK (len=%s) magic_code=%s",
                        len(resp.token), "yes" if magic_code else "no")
        else:
            logger.info("get_user_token returned NONE magic_code=%s",
                        "yes" if magic_code else "no")
        return resp
    except Exception as e:
        logger.info("get_user_token exception: %s", e)
        return None


async def ensure_token(turn_context: TurnContext):
    magic = None

    # magic code from OAuth flow
    if turn_context.activity.value and isinstance(turn_context.activity.value, dict):
        magic = turn_context.activity.value.get("state")

    # 6–8 digit codes only
    if not magic and turn_context.activity.text:
        _txt = turn_context.activity.text.strip()
        if _txt.isdigit() and len(_txt) in (6, 7, 8):
            magic = _txt

    token_resp = await try_get_token(turn_context, magic)
    if token_resp and token_resp.token:
        return token_resp.token

    # DIRECT LINE ONLY: Show OAuth card
    if turn_context.activity.channel_id != "msteams":
        url = await adapter.get_oauth_sign_in_link(turn_context, OAUTH_CONNECTION)
        card = OAuthCard(
            text="Please sign in to continue.",
            connection_name=OAUTH_CONNECTION,
            buttons=[CardAction(type=ActionTypes.signin,
                                title="Sign In", value=url)],
        )
        await turn_context.send_activity(Activity(
            attachments=[Attachment(
                content_type="application/vnd.microsoft.card.oauth",
                content=card
            )]
        ))
        logger.info("Sent sign-in card to %s",
                    getattr(turn_context.activity.from_property, "id", "unknown"))
    else:
        # In Teams, do NOT send OAuthCard
        logger.info("Teams SSO waiting → no OAuthCard sent.")

    return None


# -------------------- CORE BOT HANDLER --------------------

async def handle_activity(turn_context: TurnContext):
    a = turn_context.activity
    user_id = (a.from_property.id or "unknown")
    user_text = (a.text or "").strip()

    # request trace for every activity
    logger.info(
        "RX activity: type=%s name=%s channel=%s convId=%s svc=%s",
        a.type, getattr(a, "name", None), getattr(a, "channel_id", None),
        getattr(getattr(a, "conversation", None), "id", None),
        getattr(a, "service_url", None),
    )

    # -------------------- 1. Teams Greeting with User Name --------------------
    if a.type == "conversationUpdate":
        for m in a.members_added or []:
            if m.id == a.recipient.id:
                # Try silent SSO token
                token_resp = await try_get_token(turn_context)
                user_name = None

                if token_resp and token_resp.token:
                    user_name = get_user_display_name(token_resp.token)

                if user_name:
                    await turn_context.send_activity(f"👋 Hi **{user_name}**, you are already logged in.")
                else:
                    await turn_context.send_activity("👋 Hi! Please sign in to continue.")

        return

    # -------------------- 1.5 Invoke events (Teams SSO) --------------------
    if a.type == "invoke" and a.name in ("signin/verifyState", "tokenExchange"):
        logger.info("Invoke received: name=%s from=%s", a.name,
                    getattr(a.from_property, "id", None))
        token = await try_get_token(turn_context)
        if token and token.token:
            logger.info("Invoke token acquired (len=%s)", len(token.token))

            # Fetch user name
            user_name = get_user_display_name(token.token)
            if user_name:
                await turn_context.send_activity(f"🔓 Welcome **{user_name}**! You're signed in.")
            else:
                await turn_context.send_activity("🔓 You're signed in.")

        else:
            logger.info("Invoke token acquisition failed")
        return

    # -------------------- 2. MESSAGE --------------------
    if a.type != "message":
        return

    # 2.1 Magic code
    if user_text.isdigit() and len(user_text) in (6, 7, 8):
        logger.info("🔐 OAuth magic code detected: %s", user_text)
        token = await try_get_token(turn_context, user_text)
        if token and token.token:
            user_name = get_user_display_name(token.token)
            if user_name:
                await turn_context.send_activity(f"🔓 Welcome **{user_name}**! You're signed in.")
            else:
                await turn_context.send_activity("🔓 Sign-in successful! You can now ask your question.")
        else:
            await turn_context.send_activity("⚠️ Sign-in failed. Please click Sign In again.")
        return

    # -------------------- 3. Ensure token exists (Teams auto / DirectLine magic) --------------------
    access_token = await ensure_token(turn_context)
    if not access_token:
        return

    # 3.1 Load display name for usage later
    user_name = get_user_display_name(access_token)
    if user_name:
        logger.info("User name resolved: %s", user_name)

    # -------------------- 4. Resolve user level → assistant --------------------
    level = get_user_group_level(access_token)
    assistant_id = ASSISTANT_MAP.get(level)
    logger.info("User %s resolved: level=%s assistant=%s",
                user_id, level, assistant_id)

    if not assistant_id:
        await turn_context.send_activity("❌ No assistant assigned for your access level. Contact admin.")
        return

    # -------------------- 5. Thread isolation --------------------
    key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(key)

    if not thread_id:
        try:
            thread_id = openai.beta.threads.create().id
            thread_map[key] = thread_id
            logger.info("Created thread %s for %s", thread_id, key)
        except Exception as e:
            logger.exception("Failed to create assistant thread")
            await turn_context.send_activity("❌ Failed to create assistant session.")
            return

    # -------------------- 6. Clarification flow --------------------
    if user_id in pending_clarify:
        pc = pending_clarify[user_id]
        if pc.get("assistant_id") != assistant_id:
            pending_clarify.pop(user_id, None)
        else:
            logger.info("Received clarification from %s (round %s): %s",
                        user_id, pc.get("rounds"), user_text)
            clarification_msg = f"(Clarification) {user_text}"
            openai.beta.threads.messages.create(
                thread_id=thread_id, role="user", content=clarification_msg)
            pc["rounds"] += 1
    else:
        if user_text:
            openai.beta.threads.messages.create(
                thread_id=thread_id, role="user", content=user_text)

    # -------------------- 7. Create assistant run --------------------
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

    # -------------------- 8. Poll --------------------
    start = time.time()
    while run.status not in ("completed", "failed", "cancelled"):
        if time.time() - start > 60:
            await turn_context.send_activity("⏳ Still processing... please try again shortly.")
            return
        await asyncio.sleep(1)
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)

    # -------------------- 9. Fetch assistant reply --------------------
    try:
        msgs = openai.beta.threads.messages.list(
            thread_id=thread_id, order="desc", limit=8
        )
        reply = next(
            (m.content[0].text.value for m in msgs.data if m.role == "assistant"),
            None
        )
    except Exception:
        logger.exception("Failed to fetch assistant messages")
        reply = None

    logger.info("Assistant reply (truncated): %s",
                (reply[:200] + "...") if reply and len(reply) > 200 else reply)

    # -------------------- 10. Clarify logic --------------------
    if reply and (_is_clarify(reply) or _looks_like_clarify(reply)):
        question = _strip_clarify(reply)
        pending_clarify[user_id] = {
            "thread_id": thread_id, "assistant_id": assistant_id,
            "rounds": 0, "original": user_text
        }
        await turn_context.send_activity(f"CLARIFY: {question}")
        return

    # -------------------- 11. Generic reply → verify --------------------
    if reply and RETRY_ON_GENERIC and _reply_is_generic(reply):
        logger.info("Generic reply detected → verification run")
        verify_msg = "(Verify) Please confirm this answer strictly from the uploaded documents and include the exact source reference or say 'not available'."
        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=verify_msg
        )

        try:
            run2 = openai.beta.threads.runs.create(
                assistant_id=assistant_id,
                thread_id=thread_id,
                tool_choice={"type": "file_search"},
            )
            s2 = time.time()
            while run2.status not in ("completed", "failed", "cancelled"):
                if time.time() - s2 > 30:
                    break
                await asyncio.sleep(1)
                run2 = openai.beta.threads.runs.retrieve(
                    thread_id=thread_id, run_id=run2.id
                )

            msgs2 = openai.beta.threads.messages.list(
                thread_id=thread_id, order="desc", limit=6
            )
            verified = next(
                (m.content[0].text.value for m in msgs2.data if m.role == "assistant"),
                None
            )
            if verified:
                reply = verified
        except Exception:
            logger.exception("Verification run failed; using original reply.")

    # -------------------- 12. Clarify exit logic --------------------
    if user_id in pending_clarify:
        pc = pending_clarify[user_id]

        if reply and not (_is_clarify(reply) or _looks_like_clarify(reply)):
            pending_clarify.pop(user_id, None)
        else:
            pc["rounds"] += 1
            if pc["rounds"] >= MAX_CLARIFY_ROUNDS:
                pending_clarify.pop(user_id, None)
                await turn_context.send_activity(
                    "⚠️ I've asked clarifying questions several times but couldn't resolve this. Please rephrase or contact support."
                )
                return

            if reply and (_is_clarify(reply) or _looks_like_clarify(reply)):
                await turn_context.send_activity(f"CLARIFY: {_strip_clarify(reply)}")
                return

    # -------------------- 13. Fallback --------------------
    if not reply or "not available" in (reply or "").lower():
        await turn_context.send_activity(
            "We regret to inform that this information is not available in the provided documentation. You can contact support for further details."
        )
        return

    # -------------------- 15. SEND FINAL REPLY --------------------
    await turn_context.send_activity(reply)

# -------------------- FLASK ROUTES --------------------


@app.route("/api/messages", methods=["POST"])
def messages():
    try:
        if "application/json" not in request.headers.get("Content-Type", ""):
            return Response("Unsupported Media Type", 415)

        activity = Activity().deserialize(request.json)
        auth_hdr = request.headers.get("Authorization", "")

        logger.info("HTTP POST /api/messages (channelId=%s)",
                    getattr(activity, "channel_id", None))

        async def _proc():
            return await adapter.process_activity(activity, auth_hdr, handle_activity)

        asyncio.run(_proc())
        return Response(status=200)

    except Exception as ex:
        logger.exception("Exception in /api/messages: %s", ex)
        return Response("Internal Server Error", 500)


@app.route("/directline/token", methods=["POST"])
def directline_token():
    """
    ONLY for DirectLine WebChat client.
    Teams never uses this.
    """
    if not DIRECT_LINE_SECRET:
        return jsonify({"error": "DIRECT_LINE_SECRET not set"}), 500

    r = requests.post(
        "https://directline.botframework.com/v3/directline/tokens/generate",
        headers={"Authorization": f"Bearer {DIRECT_LINE_SECRET}"},
        timeout=10
    )

    if r.status_code != 200:
        logger.error("Direct Line token generation failed: %s", r.text)
        return jsonify({"error": "Failed to generate token", "details": r.text}), 500

    return jsonify({"token": r.json().get("token")})


@app.route("/chat", methods=["GET"])
def chat():
    """
    Serves your WebChat index.html (if using Direct Line).
    """
    return send_from_directory(app.static_folder, "index.html")


@app.route("/", methods=["GET"])
def health():
    return "Teams Bot is running."


# -------------------- MAIN --------------------
if __name__ == "__main__":
    logger.info("🚀 Bot starting...")
    logger.info("🔧 Environment check: MicrosoftAppId=%s, OAuth=%s, DirectLine=%s",
                "SET" if APP_ID else "MISSING",
                OAUTH_CONNECTION,
                "SET" if DIRECT_LINE_SECRET else "MISSING")

    app.run(host="0.0.0.0", port=3978)
