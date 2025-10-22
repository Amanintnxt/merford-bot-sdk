# ──────────────────────────────────────────────────────────────────────────────
# bot.py – Teams / Direct Line bridge to Azure OpenAI Assistants (Clarify 2.0)
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

# ─────────────────────────  ENV & OPENAI CONFIG  ────────────────────────────
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
                    format="%(asctime)s  %(levelname)-7s %(message)s")

# ─────────────────────────  ASSISTANT IDS  ──────────────────────────
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

thread_map = {}
awaiting_clarification = set()

# ─────────────────────────  HELPER FUNCTIONS  ────────────────────────────────


def _is_clarify(text: str) -> bool:
    return bool(text and text.strip().upper().startswith("CLARIFY:"))


def _strip_clarify(text: str) -> str:
    return text[len("CLARIFY:"):].strip() if _is_clarify(text) else (text or "")

# ─────────────────────────  DYNAMIC CLARIFY GENERATOR  ───────────────────────


def _clarify_actions(question: str, assistant_id: str = None, vector_store_id: str = None, thread_id: str = None):
    """
    Dynamically generate clarifying options based on the SAME assistant and vector store.
    """
    opts = []
    default_opts = ["I'll specify", "Please repeat question", "Cancel"]

    try:
        if assistant_id and thread_id:
            logging.info(
                f"🧠 Generating clarify options via assistant {assistant_id}")

            run = openai.beta.threads.runs.create(
                assistant_id=assistant_id,
                thread_id=thread_id,
                tool_choice={"type": "file_search"},
                instructions=(
                    f"You are a clarification assistant for Merford’s level-specific knowledge base. "
                    f"Refer ONLY to your attached documents (vector store: {vector_store_id}). "
                    "Given the user's query:\n"
                    f"'{question}'\n"
                    "Generate ONE short clarifying question and 2–3 multiple-choice options "
                    "strictly based on distinctions present in the documents. "
                    "Format: First line = Clarifying question. Next lines = possible answer options."
                )
            )

            start = time.time()
            while run.status not in ("completed", "failed", "cancelled"):
                if time.time() - start > 15:
                    break
                time.sleep(0.5)
                run = openai.beta.threads.runs.retrieve(
                    thread_id=thread_id, run_id=run.id)

            msgs = openai.beta.threads.messages.list(
                thread_id=thread_id, order="desc", limit=3)
            answer = next(
                (m.content[0].text.value for m in msgs.data if m.role == "assistant"), "")

            if answer:
                lines = [l.strip(" -*•")
                         for l in answer.splitlines() if len(l.strip()) > 2]
                if len(lines) > 1:
                    opts = lines[1:4]
                    question = lines[0]
                else:
                    question = answer.strip()

    except Exception as e:
        logging.warning("Clarify generation failed: %s", e)
        question = question or "Can you clarify your question?"

    if not opts:
        opts = default_opts

    return question, [CardAction(type=ActionTypes.im_back, title=o, value=o) for o in opts]

# ─────────────────────────  GRAPH USER LEVEL  ────────────────────────────────


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
        if name == "Level1Access":
            return "Level 1"
        if name == "Level2Access":
            return "Level 2"
        if name == "Level3Access":
            return "Level 3"
        if name == "Level4Access":
            return "Level 4"
    return None

# ─────────────────────────  TOKEN HELPERS  ────────────────────────────────


async def try_get_token(turn_context: TurnContext, magic_code: str | None = None):
    try:
        return await adapter.get_user_token(turn_context, OAUTH_CONNECTION, magic_code)
    except Exception as e:
        logging.warning("get_user_token exception: %s", e)
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

# ─────────────────────────  MAIN BOT LOGIC  ────────────────────────────────


async def handle_activity(turn_context: TurnContext):
    a = turn_context.activity
    user_id = a.from_property.id or "unknown"

    if a.type == "conversationUpdate":
        for m in a.members_added or []:
            if m.id == a.recipient.id:
                await turn_context.send_activity("✅ Connected. Sign in to continue.")
        return

    if a.type == "invoke" and a.name in ("signin/verifyState", "tokenExchange"):
        token = await try_get_token(turn_context)
        if token:
            await turn_context.send_activity("🔐 You're signed in. Ask your question.")
        return

    if a.type != "message":
        return
    user_text = (a.text or "").strip()
    if not user_text and not (a.value and a.value.get("state")):
        return

    access_token = await ensure_token(turn_context)
    if not access_token:
        return

    level = get_user_group_level(access_token)
    logging.info(f"User {user_id} level = {level}")
    assistant_id = ASSISTANT_MAP.get(level)
    vector_store_id = VECTOR_STORES.get(level)
    if not assistant_id:
        await turn_context.send_activity("❌ No assistant mapped to your access.")
        return

    key = f"{user_id}:{assistant_id}"
    thread_id = thread_map.get(key)
    if not thread_id:
        thread_id = openai.beta.threads.create().id
        thread_map[key] = thread_id

    # Add user message
    if user_text and not user_text.isdigit():
        content_to_send = f"(User clarification) {user_text}" if user_id in awaiting_clarification else user_text
        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=content_to_send)
        awaiting_clarification.discard(user_id)

    # Run assistant
    run = openai.beta.threads.runs.create(
        assistant_id=assistant_id, thread_id=thread_id, tool_choice={"type": "file_search"})
    start = time.time()
    while run.status not in ("completed", "failed", "cancelled"):
        if time.time() - start > 60:
            break
        await asyncio.sleep(1)
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id, run_id=run.id)

    # Get response
    msgs = openai.beta.threads.messages.list(
        thread_id=thread_id, order="desc", limit=5)
    reply = next(
        (m.content[0].text.value for m in msgs.data if m.role == "assistant"), None)

    # Detect CLARIFY
    if reply and _is_clarify(reply):
        question_text = _strip_clarify(reply)
        question, actions = _clarify_actions(
            question_text, assistant_id, vector_store_id, thread_id)
        awaiting_clarification.add(user_id)
        await turn_context.send_activity(Activity(type="message", text=question, suggested_actions=SuggestedActions(actions=actions)))
        return

    await turn_context.send_activity(reply or "❌ No reply from assistant.")

# ─────────────────────────  FLASK ROUTES  ────────────────────────────────────


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


@app.route("/directline/token", methods=["POST"])
def directline_token():
    if not DIRECT_LINE_SECRET:
        return jsonify({"error": "DIRECT_LINE_SECRET not set"}), 500
    r = requests.post(
        "https://directline.botframework.com/v3/directline/tokens/generate",
        headers={"Authorization": f"Bearer {DIRECT_LINE_SECRET}"}, timeout=10)
    if r.status_code != 200:
        return jsonify({"error": "Failed to generate token"}), 500
    return jsonify({"token": r.json().get("token")})


@app.route("/", methods=["GET"])
def health():
    return "Bot is running."


# ─────────────────────────  MAIN  ───────────────────────────────
if __name__ == "__main__":
    logging.info("🚀 Bot running with dynamic clarify logic")
    app.run(host="0.0.0.0", port=3978)
