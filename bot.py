# ──────────────────────────────────────────────────────────────────────────────
# bot.py – Teams / Direct Line bridge to Azure OpenAI Assistants (SSO-first)
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

# ─────────────────────────  LOGGING  ─────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-7s %(message)s")

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
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# ─────────────────────────  IN-MEMORY STATE  ────────────────────────────────
thread_map = {}        # key = f"{user_id}:{assistant_id}" → thread_id

# ─────────────────────────  GRAPH LOOK-UP  ───────────────────────────────────


def get_user_group_level(token: str) -> str | None:
    url = "https://graph.microsoft.com/v1.0/me/memberOf?$select=displayName"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
    except requests.RequestException as e:
        logging.warning("Graph request error: %s", e)
        return None

    if resp.status_code != 200:
        logging.warning("Graph /me/memberOf failed %s – %s",
                        resp.status_code, resp.text)
        return None

    for g in resp.json().get("value", []):
        name = g.get("displayName")
        if not name:
            continue
        logging.info("AAD group found: %s", name)
        if name == "Level1Access":
            return "Level 1"
        if name == "Level2Access":
            return "Level 2"
        if name == "Level3Access":
            return "Level 3"
        if name == "Level4Access":
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
        # 🔐 Optional admin protection
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

        # Decide targets
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
            # cleanup
            if os.path.exists(save_path):
                os.remove(save_path)

    return render_template("upload.html")

# ─────────────────────────  TOKEN HELPERS  ───────────────────────────────────


async def try_get_token(turn_context: TurnContext, magic_code: str | None = None):
    """Try to get a token using Teams SSO (silent), or magic code if present."""
    try:
        return await adapter.get_user_token(turn_context, OAUTH_CONNECTION, magic_code)
    except Exception as e:
        logging.info(
            "get_user_token exception (will fall back to prompt): %s", e)
        return None


async def ensure_token(turn_context: TurnContext):
    """
    1) Try silent SSO (Teams)
    2) If user typed/passed a magic code (Web Chat), try with code
    3) If still none, send OAuthCard for sign-in
    """
    magic = None
    if turn_context.activity.value and isinstance(turn_context.activity.value, dict):
        magic = turn_context.activity.value.get("state")
    if not magic and turn_context.activity.text and turn_context.activity.text.strip().isdigit():
        magic = turn_context.activity.text.strip()

    token_resp = await try_get_token(turn_context, magic)
    if token_resp and token_resp.token:
        return token_resp.token

    # Send sign-in card
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
                    "✅ Connected. If you're in Teams, SSO is automatic. In Web Chat, click **Sign In** once."
                )
        return

    # 2) Teams SSO invokes
    if a.type == "invoke" and a.name in ("signin/verifyState", "tokenExchange"):
        logging.info("Received invoke: %s", a.name)
        token = await try_get_token(turn_context)
        if token:
            await turn_context.send_activity("🔐 You're signed in. Ask your question")
        return  # ⚠️ No Response object here

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
    if user_text and not user_text.isdigit():
        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=user_text
        )

    # 8) Run assistant
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
            await turn_context.send_activity("⏳ Still working… please send again if no reply arrives.")
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
    app.run(host="0.0.0.0", port=3978)
