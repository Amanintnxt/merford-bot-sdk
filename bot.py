import os
import asyncio
from flask import Flask, request, jsonify, send_from_directory
from botbuilder.core import (
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    ConversationState,
    MemoryStorage,
    TurnContext,
    UserState,
)
from botbuilder.schema import Activity
from botbuilder.dialogs import (
    DialogSet,
    DialogTurnStatus,
    OAuthPrompt,
    OAuthPromptSettings,
    WaterfallDialog,
    WaterfallStepContext,
)
from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from openai import AzureOpenAI
import requests
from dotenv import load_dotenv

load_dotenv()

# ENV VARIABLES
APP_ID = os.environ["MicrosoftAppId"]
APP_PASSWORD = os.environ["MicrosoftAppPassword"]
CONNECTION_NAME = os.environ["OAUTH_CONNECTION_NAME"]
ASSISTANT_ID_DEFAULT = os.environ["ASSISTANT_ID"]
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
AZURE_OPENAI_API_KEY = os.environ["AZURE_OPENAI_API_KEY"]
AZURE_OPENAI_ENDPOINT = os.environ["AZURE_OPENAI_ENDPOINT"]
DIRECT_LINE_SECRET = os.environ["DIRECT_LINE_SECRET"]

# ====== Azure OpenAI Setup ======
token_provider = get_bearer_token_provider(
    DefaultAzureCredential(),
    "https://cognitiveservices.azure.com/.default"
)

openai = AzureOpenAI(
    azure_ad_token_provider=token_provider,
    api_key=AZURE_OPENAI_API_KEY,
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
    api_version="2024-05-01-preview"
)

# ====== Bot Setup ======
app = Flask(__name__)
loop = asyncio.get_event_loop()
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

memory = MemoryStorage()
user_state = UserState(memory)
conversation_state = ConversationState(memory)
dialog_state = conversation_state.create_property("DialogState")
user_profile_accessor = user_state.create_property("UserProfile")

dialogs = DialogSet(dialog_state)

oauth_prompt = OAuthPrompt(
    OAuthPrompt.__name__,
    OAuthPromptSettings(
        connection_name=CONNECTION_NAME,
        text="Please sign in to continue.",
        title="Sign In",
        timeout=300000,
    ),
)
dialogs.add(oauth_prompt)

# ====== Group → Assistant Map ======
assistant_map = {
    "Level1Access": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
    "Level2Access": "asst_BIOAPR48tzth4k79U4h0cPtu",
    "Level3Access": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
    "Level4Access": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
}

# Waterfall Dialog


async def prompt_step(step: WaterfallStepContext):
    return await step.begin_dialog(OAuthPrompt.__name__)


async def token_step(step: WaterfallStepContext):
    token_response = step.result
    if not token_response:
        await step.context.send_activity("Login was not successful.")
        return await step.end_dialog()

    access_token = token_response.token
    headers = {"Authorization": f"Bearer {access_token}"}
    graph_url = "https://graph.microsoft.com/v1.0/me/memberOf"
    res = requests.get(graph_url, headers=headers)

    if res.status_code != 200:
        await step.context.send_activity("Failed to retrieve group info.")
        return await step.end_dialog()

    groups = res.json().get("value", [])
    assistant_id = ASSISTANT_ID_DEFAULT

    for g in groups:
        group_name = g.get("displayName")
        if group_name in assistant_map:
            assistant_id = assistant_map[group_name]
            break

    user_profile = await user_profile_accessor.get(step.context, {})
    user_profile["assistant_id"] = assistant_id

    if "thread_id" not in user_profile:
        thread = openai.beta.threads.create()
        user_profile["thread_id"] = thread.id

    await user_profile_accessor.set(step.context, user_profile)
    await user_state.save_changes(step.context)

    await step.context.send_activity("✅ Signed in successfully!")
    await step.context.send_activity("Hello! How can I assist you today?")
    return await step.end_dialog()

dialogs.add(WaterfallDialog("main_dialog", [prompt_step, token_step]))

# ====== Message Handler ======


async def handle_message(context: TurnContext):
    dialog_ctx = await dialogs.create_context(context)
    results = await dialog_ctx.continue_dialog()

    if results.status == DialogTurnStatus.Empty:
        await dialog_ctx.begin_dialog("main_dialog")
        return

    elif results.status == DialogTurnStatus.Complete:
        user_profile = await user_profile_accessor.get(context, {})
        assistant_id = user_profile.get("assistant_id", ASSISTANT_ID_DEFAULT)
        thread_id = user_profile.get("thread_id")

        message = context.activity.text
        openai.beta.threads.messages.create(
            thread_id, role="user", content=message
        )

        run = openai.beta.threads.runs.create(
            thread_id, assistant_id=assistant_id
        )

        while True:
            run_status = openai.beta.threads.runs.retrieve(thread_id, run.id)
            if run_status.status == "completed":
                break
            await asyncio.sleep(1)

        messages = openai.beta.threads.messages.list(thread_id)
        last_msg = messages.data[0].content[0].text.value
        await context.send_activity(last_msg)

# ====== Flask Routes ======


@app.route("/api/messages", methods=["POST"])
def messages():
    if "application/json" not in request.headers["Content-Type"]:
        return jsonify({"error": "Unsupported Media Type"}), 415

    activity = Activity().deserialize(request.json)
    auth_header = request.headers.get("Authorization", "")

    async def call_bot():
        await adapter.process_activity(activity, auth_header, handle_message)

    loop.run_until_complete(call_bot())
    return "", 202


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/directline/token", methods=["GET"])
def direct_line_token():
    url = "https://directline.botframework.com/v3/directline/tokens/generate"
    headers = {
        "Authorization": f"Bearer {DIRECT_LINE_SECRET}",
        "Content-Type": "application/json"
    }
    res = requests.post(url, headers=headers)
    return jsonify(res.json())


# ====== Run Server ======
if __name__ == "__main__":
    app.run(port=3978, debug=True)
