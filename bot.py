import os
import json
import time
import requests
from flask import Flask, request, jsonify
from botbuilder.core import (
    BotFrameworkAdapterSettings,
    TurnContext,
    ConversationState,
    MemoryStorage,
    BotFrameworkAdapter,
)
from botbuilder.schema import Activity
from botbuilder.dialogs import DialogSet, DialogTurnStatus, OAuthPrompt, OAuthPromptSettings, PromptOptions
from openai import AzureOpenAI

# === Environment Variables ===
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
OAUTH_CONNECTION_NAME = os.getenv("OAUTH_CONNECTION_NAME")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_ENTRA_TOKEN = os.getenv("AZURE_ENTRA_TOKEN")

ASSISTANT_IDS = {
    "Level1Access": os.getenv("ASSISTANT_ID_LEVEL1"),
    "Level2Access": os.getenv("ASSISTANT_ID_LEVEL2"),
    "Level3Access": os.getenv("ASSISTANT_ID_LEVEL3"),
    "Level4Access": os.getenv("ASSISTANT_ID_LEVEL4"),
}

# === Flask Setup ===
app = Flask(__name__)

# === Bot Framework Adapter ===
adapter_settings = BotFrameworkAdapterSettings(
    app_id=os.getenv("MicrosoftAppId"),
    app_password=os.getenv("MicrosoftAppPassword"),
)
adapter = BotFrameworkAdapter(adapter_settings)

# === Conversation State ===
memory = MemoryStorage()
conversation_state = ConversationState(memory)
dialog_state = conversation_state.create_property("DialogState")
dialogs = DialogSet(dialog_state)

oauth_settings = OAuthPromptSettings(
    connection_name=OAUTH_CONNECTION_NAME,
    text="Please sign in",
    title="Sign In",
    timeout=300000,
)

dialogs.add(OAuthPrompt("OAuthPrompt", oauth_settings))

# === Azure OpenAI Client ===
openai_client = AzureOpenAI(
    api_key=AZURE_OPENAI_API_KEY,
    api_version=os.getenv("OPENAI_API_VERSION"),
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
)

# === Helper Functions ===


def get_user_groups(token):
    graph_url = "https://graph.microsoft.com/v1.0/me/memberOf"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(graph_url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return [group["displayName"] for group in data.get("value", [])]
    return []


def get_assistant_id_from_groups(groups):
    for group in groups:
        if group in ASSISTANT_IDS:
            return ASSISTANT_IDS[group]
    raise Exception("❌ No assistant mapped for user group.")


def send_to_openai_assistant(assistant_id, user_input):
    thread = openai_client.beta.threads.create()
    openai_client.beta.threads.messages.create(
        thread_id=thread.id,
        role="user",
        content=user_input,
    )
    run = openai_client.beta.threads.runs.create(
        thread_id=thread.id,
        assistant_id=assistant_id,
    )
    while True:
        run_status = openai_client.beta.threads.runs.retrieve(
            thread_id=thread.id, run_id=run.id
        )
        if run_status.status == "completed":
            break
        time.sleep(1)
    messages = openai_client.beta.threads.messages.list(thread_id=thread.id)
    return messages.data[0].content[0].text.value

# === Flask Bot Route ===


@app.route("/api/messages", methods=["POST"])
def messages():
    activity = Activity().deserialize(request.json)

    if activity.type != "message":
        return jsonify({"status": "ignored"})

    async def call_bot(turn_context: TurnContext):
        dc = await dialogs.create_context(turn_context)
        results = await dc.continue_dialog()

        if results.status == DialogTurnStatus.Empty:
            await dc.begin_dialog("OAuthPrompt")
        elif results.status == DialogTurnStatus.Complete:
            token_response = results.result
            if token_response:
                token = token_response.token
                user_groups = get_user_groups(token)
                assistant_id = get_assistant_id_from_groups(user_groups)

                response = send_to_openai_assistant(
                    assistant_id=assistant_id,
                    user_input=turn_context.activity.text,
                )
                await turn_context.send_activity(response)
            else:
                await turn_context.send_activity("❌ Login failed.")

        await conversation_state.save_changes(turn_context)

    task = adapter.process_activity(activity, "", call_bot)
    return jsonify({"status": "ok"})


# === App Entry Point ===
if __name__ == "__main__":
    try:
        print("🚀 Starting bot on Render...")
        app.run(host="0.0.0.0", port=10000)
    except Exception as e:
        print("❌ Bot failed to start:", str(e))
