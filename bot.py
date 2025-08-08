import os
import time
import openai
import asyncio
import logging
import requests
from dotenv import load_dotenv
from flask import Flask, request, Response
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter, TurnContext
from botbuilder.schema import Activity
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.core.oauth import UserTokenClient

# Load environment variables
load_dotenv()

# Credentials
APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")

# OAuth Connection Name from Azure Bot Service
OAUTH_CONNECTION_NAME = os.getenv("OAUTH_CONNECTION_NAME", "TeamsSSO")

# Configure OpenAI Azure API
openai.api_type = "azure"
openai.api_version = "2024-05-01-preview"
openai.api_key = AZURE_OPENAI_API_KEY
openai.azure_endpoint = AZURE_OPENAI_ENDPOINT.rstrip("/")

# Flask & Bot setup
app = Flask(__name__)
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# Memory store
thread_map = {}


def get_user_group_level(user_id, access_token):
    """Get the user's group level using a live Graph API token."""
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/memberOf?$select=id,displayName"
    headers = {"Authorization": f"Bearer {access_token}"}

    logging.info(f"Fetching group membership for user {user_id}")
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        logging.warning(
            f"Group lookup failed: {response.status_code} - {response.text}")
        return None

    groups = response.json().get("value", [])
    for group in groups:
        name = group.get("displayName")
        if name == "Level1Access":
            return "Level 1"
        elif name == "Level2Access":
            return "Level 2"
        elif name == "Level3Access":
            return "Level 3"
        elif name == "Level4Access":
            return "Level 4"

    return None  # No matching group


async def handle_message(turn_context: TurnContext):
    """Main bot message handler."""
    if turn_context.activity.type == "conversationUpdate":
        members_added = turn_context.activity.members_added
        if members_added:
            for member in members_added:
                if member.id == turn_context.activity.recipient.id:
                    await turn_context.send_activity("Hello! How can I assist you today?")
        return

    if turn_context.activity.type != "message" or not turn_context.activity.text.strip():
        return

    # 1️⃣ Retrieve user token from Teams SSO
    token_client = await adapter.create_user_token_client(turn_context)
    token_response = await token_client.get_user_token(
        turn_context.activity.from_property.id,
        OAUTH_CONNECTION_NAME,
        turn_context.activity.channel_id,
        turn_context.activity.from_property.aad_object_id
    )

    if not token_response or not token_response.token:
        await turn_context.send_activity("You need to sign in to use this bot.")
        return

    access_token = token_response.token
    user_id = turn_context.activity.from_property.aad_object_id
    user_input = turn_context.activity.text

    try:
        await turn_context.send_activity(Activity(type="typing"))

        # 2️⃣ Get the user's group level using live token
        level = get_user_group_level(user_id, access_token)
        logging.info(f"User {user_id} has level: {level}")

        if not level:
            await turn_context.send_activity("You do not have permission to access this bot.")
            return

        # 3️⃣ Map assistant based on level
        assistant_map = {
            "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
            "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
            "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
            "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
        }
        assistant_id = assistant_map.get(level)
        logging.info(f"Using assistant: {assistant_id} for user {user_id}")

        # 4️⃣ Get or create conversation thread
        thread_id = thread_map.get(user_id)
        if not thread_id:
            thread = openai.beta.threads.create()
            thread_id = thread.id
            thread_map[user_id] = thread_id

        # Add user message to thread
        openai.beta.threads.messages.create(
            thread_id=thread_id,
            role="user",
            content=user_input
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

        # Fetch assistant reply message
        messages = openai.beta.threads.messages.list(thread_id=thread_id)
        assistant_reply = None
        for msg in messages.data:
            if msg.role == "assistant" and msg.content:
                assistant_reply = msg.content[0].text.value
                break

        if not assistant_reply:
            assistant_reply = "Sorry, I didn't get a reply from the assistant."

    except Exception as e:
        logging.error(f"Error handling message: {e}")
        assistant_reply = "Something went wrong."

    # 5️⃣ Send reply
    await turn_context.send_activity(Activity(
        type="message",
        text=assistant_reply,
        recipient=turn_context.activity.from_property,
        from_property=turn_context.activity.recipient,
        conversation=turn_context.activity.conversation,
        channel_id=turn_context.activity.channel_id,
        service_url=turn_context.activity.service_url
    ))


@app.route("/api/messages", methods=["POST"])
def messages():
    try:
        if "application/json" not in request.headers.get("Content-Type", ""):
            return Response("Unsupported Media Type", status=415)

        activity = Activity().deserialize(request.json)
        auth_header = request.headers.get("Authorization", "")

        async def process():
            return await adapter.process_activity(activity, auth_header, handle_message)

        asyncio.run(process())
        return Response(status=200)

    except Exception as e:
        logging.error(f"Exception in /api/messages: {e}")
        return Response("Internal Server Error", status=500)


@app.route("/", methods=["GET"])
def health_check():
    return "Teams Bot is running."


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app.run(host="0.0.0.0", port=3978)
