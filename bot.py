import os
import json
import asyncio
import requests
from flask import Flask, request, send_from_directory, jsonify
from dotenv import load_dotenv
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings, TurnContext, ConversationState, MemoryStorage
from botbuilder.schema import Activity
from botbuilder.dialogs import DialogSet, OAuthPrompt, OAuthPromptSettings, DialogTurnStatus
from openai import AzureOpenAI
from azure.identity import DefaultAzureCredential, get_bearer_token_provider

load_dotenv()

app = Flask(__name__)

# === Load environment variables ===
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_ENTRA_TOKEN = os.getenv("AZURE_ENTRA_TOKEN")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
OAUTH_CONNECTION_NAME = os.getenv("OAUTH_CONNECTION_NAME")
MicrosoftAppId = os.getenv("MicrosoftAppId")
MicrosoftAppPassword = os.getenv("MicrosoftAppPassword")
ASSISTANT_ID = os.getenv("ASSISTANT_ID")
DIRECT_LINE_SECRET = os.getenv("DIRECT_LINE_SECRET")

# === Azure OpenAI client using Azure AD Token ===
token_provider = get_bearer_token_provider(
    DefaultAzureCredential(),
    "https://cognitiveservices.azure.com/.default"
)
openai = AzureOpenAI(
    azure_ad_token_provider=token_provider,
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
    api_key=AZURE_OPENAI_API_KEY,
    api_version="2023-12-01-preview"  # required param
)

# === Bot Adapter ===
adapter_settings = BotFrameworkAdapterSettings(
    MicrosoftAppId, MicrosoftAppPassword)
adapter = BotFrameworkAdapter(adapter_settings)

# === Memory Storage & State ===
memory = MemoryStorage()
conversation_state = ConversationState(memory)
dialog_state = conversation_state.create_property("DialogState")
dialogs = DialogSet(dialog_state)

# === OAuth Prompt ===
oauth_prompt_settings = OAuthPromptSettings(
    connection_name=OAUTH_CONNECTION_NAME,
    text="Please sign in to continue",
    title="Sign In",
    timeout=300000  # 5 minutes
)
dialogs.add(OAuthPrompt("OAuthPrompt", oauth_prompt_settings))

# === HTML frontend ===


@app.route("/")
def index():
    return send_from_directory("static", "index.html")

# Bot endpoint


@app.route("/api/messages", methods=["POST"])
async def messages():
    if "application/json" in request.headers["Content-Type"]:
        body = request.json
    else:
        return jsonify({"error": "Unsupported Media Type"}), 415

    activity = Activity().deserialize(body)
    auth_header = request.headers.get("Authorization", "")

    async def call_bot(turn_context: TurnContext):
        dc = await dialogs.create_context(turn_context)
        results = await dc.continue_dialog()

        if results.status == DialogTurnStatus.Empty:
            await dc.begin_dialog("OAuthPrompt")
        elif results.status == DialogTurnStatus.Complete:
            if results.result:
                token = results.result.token

                # === Replace this with Assistant logic ===
                user_input = turn_context.activity.text
                thread = openai.beta.threads.create()
                openai.beta.threads.messages.create(
                    thread_id=thread.id,
                    role="user",
                    content=user_input
                )
                run = openai.beta.threads.runs.create(
                    thread_id=thread.id,
                    assistant_id=ASSISTANT_ID
                )
                while True:
                    status = openai.beta.threads.runs.retrieve(
                        thread_id=thread.id, run_id=run.id)
                    if status.status == "completed":
                        break
                    await asyncio.sleep(1)
                messages = openai.beta.threads.messages.list(
                    thread_id=thread.id)
                reply = messages.data[0].content[0].text.value
                await turn_context.send_activity(reply)
            else:
                await turn_context.send_activity("You must sign in to proceed.")

        await conversation_state.save_changes(turn_context)

    await adapter.process_activity(activity, auth_header, call_bot)
    return ("", 200)


@app.route("/directline/token", methods=["GET"])
def direct_line_token():
    url = "https://directline.botframework.com/v3/directline/tokens/generate"
    headers = {
        "Authorization": f"Bearer {DIRECT_LINE_SECRET}",
        "Content-Type": "application/json"
    }
    res = requests.post(url, headers=headers)
    return jsonify(res.json())


# === Run server ===
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3978))
    app.run(host="0.0.0.0", port=port, debug=True)
