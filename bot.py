import os
import time
import openai
import asyncio
import logging
import requests
from dotenv import load_dotenv
from flask import Flask, request, Response

from botbuilder.core import (
    BotFrameworkAdapterSettings,
    BotFrameworkAdapter,
    TurnContext,
    MemoryStorage,
    ConversationState,
    UserState
)
from botbuilder.schema import Activity
from botbuilder.dialogs import DialogSet, DialogTurnStatus, WaterfallDialog, WaterfallStepContext, OAuthPrompt, OAuthPromptSettings

# Load environment variables
load_dotenv()
# Environment variables
APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
# OAuth Connection in Azure
CONNECTION_NAME = os.getenv("OAUTH_CONNECTION_NAME")
ASSISTANT_ID = os.getenv("ASSISTANT_ID")


# Configure OpenAI
openai.api_type = "azure"
openai.api_version = "2024-05-01-preview"
openai.api_key = AZURE_OPENAI_API_KEY
openai.azure_endpoint = AZURE_OPENAI_ENDPOINT.rstrip("/")

# Flask app and adapter
app = Flask(__name__)
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# Bot state
memory = MemoryStorage()
conversation_state = ConversationState(memory)
user_state = UserState(memory)
dialogs = DialogSet(conversation_state.create_property("DialogState"))

# OAuth Prompt setup
dialogs.add(OAuthPrompt(
    OAuthPrompt.__name__,
    OAuthPromptSettings(
        connection_name=CONNECTION_NAME,
        text="Please sign in to access your profile.",
        title="Sign In",
        timeout=300000  # 5 minutes
    )
))

dialogs.add(WaterfallDialog(
    "MainDialog",
    [
        lambda step: step.begin_dialog(OAuthPrompt.__name__),
        lambda step: handle_token(step)
    ]
))

# Map thread to user
thread_map = {}

# Handle token and continue conversation


async def handle_token(step: WaterfallStepContext):
    token_response = step.result
    if not token_response:
        await step.context.send_activity("Sorry, I couldn't log you in.")
        return await step.end_dialog()

    user_id = step.context.activity.from_property.aad_object_id
    access_token = token_response.token

    # 🔐 Use delegated token to call Microsoft Graph (e.g., group check)
    graph_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/memberOf?$select=id,displayName"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(graph_url, headers=headers)
    groups = response.json().get("value", [])

    level = None
    for g in groups:
        name = g.get("displayName")
        if name == "Level1Access":
            level = "Level 1"
        elif name == "Level2Access":
            level = "Level 2"
        elif name == "Level3Access":
            level = "Level 3"
        elif name == "Level4Access":
            level = "Level 4"

    # Choose assistant
    assistant_map = {
        "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
        "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
        "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
        "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
    }
    assistant_id = assistant_map.get(level, ASSISTANT_ID)

    # Thread setup
    if user_id not in thread_map:
        thread = openai.beta.threads.create()
        thread_map[user_id] = thread.id
    thread_id = thread_map[user_id]

    # Add greeting
    openai.beta.threads.messages.create(
        thread_id=thread_id,
        role="user",
        content="Hi"
    )

    run = openai.beta.threads.runs.create(
        assistant_id=assistant_id,
        thread_id=thread_id
    )

    while run.status not in ["completed", "failed", "cancelled"]:
        time.sleep(1)
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id,
            run_id=run.id
        )

    messages = openai.beta.threads.messages.list(thread_id=thread_id)
    assistant_reply = None
    for msg in messages.data:
        if msg.role == "assistant":
            assistant_reply = msg.content[0].text.value
            break

    if not assistant_reply:
        assistant_reply = "I couldn't get a reply from the assistant."

    await step.context.send_activity(assistant_reply)
    return await step.end_dialog()

# Handle every message


async def handle_message(turn_context: TurnContext):
    if turn_context.activity.type == "message":
        dc = await dialogs.create_context(turn_context)
        result = await dc.continue_dialog()
        if result.status == DialogTurnStatus.Empty:
            await dc.begin_dialog("MainDialog")
    elif turn_context.activity.type == "conversationUpdate":
        for member in turn_context.activity.members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity("Hello! You may be prompted to sign in.")

# Bot endpoint


@app.route("/api/messages", methods=["POST"])
def messages():
    activity = Activity().deserialize(request.json)
    auth_header = request.headers.get("Authorization", "")

    async def aux():
        await adapter.process_activity(activity, auth_header, handle_message)
    try:
        asyncio.run(aux())
        return Response(status=200)
    except Exception as e:
        logging.error(f"Exception: {e}")
        return Response("Internal Server Error", status=500)


@app.route("/", methods=["GET"])
def health():
    return "Bot running with SSO!"


@app.route("/teamsso/callback", methods=["GET"])
def teams_callback():
    # Print the entire query parameters
    print("Query Params:", request.args)
    return "Bot running with SSO! callback"


# Run the app
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app.run("0.0.0.0", port=3978)
