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
from botbuilder.dialogs import DialogSet, DialogTurnStatus, WaterfallDialog, WaterfallStepContext

# Load .env values
load_dotenv()

# Environment variables
APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
ASSISTANT_ID = os.getenv("ASSISTANT_ID")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# OpenAI config
openai.api_type = "azure"
openai.api_version = "2024-05-01-preview"
openai.api_key = AZURE_OPENAI_API_KEY
openai.azure_endpoint = AZURE_OPENAI_ENDPOINT.rstrip("/")

# Flask + Bot Framework setup
app = Flask(__name__)
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)
memory = MemoryStorage()
conversation_state = ConversationState(memory)
user_state = UserState(memory)
dialogs = DialogSet(conversation_state.create_property("DialogState"))

# Session mapping
thread_map = {}
access_token_map = {}

# Bot dialog
dialogs.add(WaterfallDialog("MainDialog", [
    lambda step: send_signin_link(step),
    lambda step: complete_signin(step)
]))


async def send_signin_link(step: WaterfallStepContext):
    user_id = step.context.activity.from_property.aad_object_id
    state = user_id  # Can encrypt or JWT for security

    login_url = (
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
        f"?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri=https://merford-bot-sdk.onrender.com/teamsso/callback"
        f"&response_mode=query"
        f"&scope=openid profile email offline_access https://graph.microsoft.com/.default"
        f"&state={state}"
    )

    await step.context.send_activity(f"Please [sign in]({login_url}) to continue.")
    return await step.end_dialog()


async def complete_signin(step: WaterfallStepContext):
    await step.context.send_activity("After signing in, please send any message to continue.")
    return await step.end_dialog()

# Message handler


async def handle_message(turn_context: TurnContext):
    if turn_context.activity.type == "message":
        user_id = turn_context.activity.from_property.aad_object_id
        dc = await dialogs.create_context(turn_context)
        print(f'access_token = {access_token_map}')
        # Check if token is available for user
        token = access_token_map.get(user_id)
        if not token:
            result = await dc.continue_dialog()
            if result.status == DialogTurnStatus.Empty:
                await dc.begin_dialog("MainDialog")
            return

        # Fetch user groups
        headers = {"Authorization": f"Bearer {token}"}
        graph_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/memberOf?$select=id,displayName"
        response = requests.get(graph_url, headers=headers)

        if response.status_code != 200:
            await turn_context.send_activity("Failed to fetch group info.")
            return

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

        assistant_map = {
            "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
            "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
            "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
            "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
        }
        assistant_id = assistant_map.get(level, ASSISTANT_ID)

        # Set up thread for assistant
        if user_id not in thread_map:
            thread = openai.beta.threads.create()
            thread_map[user_id] = thread.id
        thread_id = thread_map[user_id]

        # Add user message and get assistant reply
        openai.beta.threads.messages.create(
            thread_id=thread_id, role="user", content=turn_context.activity.text)
        run = openai.beta.threads.runs.create(
            assistant_id=assistant_id, thread_id=thread_id)

        while run.status not in ["completed", "failed", "cancelled"]:
            time.sleep(1)
            run = openai.beta.threads.runs.retrieve(
                thread_id=thread_id, run_id=run.id)

        messages = openai.beta.threads.messages.list(thread_id=thread_id)
        assistant_reply = next((m.content[0].text.value for m in messages.data if m.role ==
                               "assistant"), "I couldn't get a reply from the assistant.")
        await turn_context.send_activity(assistant_reply)

    elif turn_context.activity.type == "conversationUpdate":
        for member in turn_context.activity.members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity("Welcome! Type anything to begin.")

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

# Custom OAuth callback


@app.route("/teamsso/callback", methods=["GET"])
def teams_callback():
    code = request.args.get("code")
    state = request.args.get("state")  # user_id
    error = request.args.get("error")

    if error or not code:
        return f"Login failed: {error or 'No code provided'}", 400

    token_data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": "https://merford-bot-sdk.onrender.com/teamsso/callback",
        "scope": "openid profile email offline_access https://graph.microsoft.com/.default"
    }

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    token_response = requests.post(token_url, data=token_data)

    if token_response.status_code != 200:
        logging.error("Token exchange failed: %s", token_response.text)
        return "Token exchange failed", 500

    tokens = token_response.json()
    access_token = tokens.get("access_token")

    if not access_token:
        return "No access token in response", 500

    # Store access token in memory (use Redis or DB in production)
    access_token_map[state] = access_token

    return "✅ Logged in! Return to chat and type anything to continue."

# Health check


@app.route("/", methods=["GET"])
def health():
    return "Bot running with SSO!"


# Start the Flask app
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app.run("0.0.0.0", port=3978)
