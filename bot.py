import os
import time
import openai
import asyncio
import logging
import requests
from dotenv import load_dotenv
from flask import Flask, request, Response
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter, TurnContext
from botbuilder.schema import Activity, Attachment, CardAction, ActionTypes, OAuthCard

# Load environment variables
load_dotenv()

# Credentials
APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
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

# Store conversation threads
thread_map = {}

# ---------------------------
# Helper: Get user group level from Microsoft Graph using /me
# ---------------------------


def get_user_group_level(access_token):
    """Get the user's group level using /me/memberOf."""
    url = "https://graph.microsoft.com/v1.0/me/memberOf?$select=id,displayName"
    headers = {"Authorization": f"Bearer {access_token}"}
    logging.info("Fetching group membership for signed-in user via /me")

    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        logging.warning(
            f"Group lookup failed: {response.status_code} - {response.text}")
        return None

    groups = response.json().get("value", [])
    logging.info(f"Found {len(groups)} groups")

    for group in groups:
        name = group.get("displayName")
        logging.info(f" - Group: {name}")
        if name == "Level1Access":
            return "Level 1"
        elif name == "Level2Access":
            return "Level 2"
        elif name == "Level3Access":
            return "Level 3"
        elif name == "Level4Access":
            return "Level 4"
    return None


# ---------------------------
# Main Bot Logic
# ---------------------------
async def handle_message(turn_context: TurnContext):
    # Send greeting when conversation starts
    if turn_context.activity.type == "conversationUpdate":
        members_added = turn_context.activity.members_added
        if members_added:
            for member in members_added:
                if member.id == turn_context.activity.recipient.id:
                    await turn_context.send_activity("Hello! How can I assist you today?")
        return

    # Ignore empty messages
    if turn_context.activity.type != "message" or not turn_context.activity.text.strip():
        return

    # Detect magic code (from OAuth flow)
    magic_code = None
    if turn_context.activity.value and "state" in turn_context.activity.value:
        magic_code = turn_context.activity.value["state"]
    elif turn_context.activity.text and turn_context.activity.text.strip().isdigit():
        magic_code = turn_context.activity.text.strip()

    # Try to get token
    token_response = await adapter.get_user_token(
        turn_context,
        OAUTH_CONNECTION_NAME,
        magic_code
    )

    if not token_response or not token_response.token:
        # Ask user to sign in
        sign_in_url = await adapter.get_oauth_sign_in_link(turn_context, OAUTH_CONNECTION_NAME)
        oauth_card = OAuthCard(
            text="Please sign in to continue.",
            connection_name=OAUTH_CONNECTION_NAME,
            buttons=[
                CardAction(
                    type=ActionTypes.signin,
                    title="Sign In",
                    value=sign_in_url
                )
            ]
        )
        attachment = Attachment(
            content_type="application/vnd.microsoft.card.oauth",
            content=oauth_card
        )
        await turn_context.send_activity(Activity(attachments=[attachment]))
        return

    # We have a valid token
    access_token = token_response.token
    level = get_user_group_level(access_token)

    logging.info(f"User is at: {level}")

    if not level:
        await turn_context.send_activity("You do not have permission to access this bot.")
        return

    # Assistant mapping
    assistant_map = {
        "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
        "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
        "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
        "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
    }
    assistant_id = assistant_map.get(level)

    # Create or get thread for user
    user_id = turn_context.activity.from_property.id
    thread_id = thread_map.get(user_id)
    if not thread_id:
        thread = openai.beta.threads.create()
        thread_id = thread.id
        thread_map[user_id] = thread_id

    # Add user message to thread
    openai.beta.threads.messages.create(
        thread_id=thread_id,
        role="user",
        content=turn_context.activity.text
    )

    try:
        run = openai.beta.threads.runs.create(
            assistant_id=assistant_id,
            thread_id=thread_id
        )
    except Exception as e:
        logging.error(f"Failed to create run: {e}")
        await turn_context.send_activity("Something went wrong while connecting to the assistant.")
        return

    # Wait for completion
    while run.status not in ["completed", "failed", "cancelled"]:
        time.sleep(1)
        run = openai.beta.threads.runs.retrieve(
            thread_id=thread_id,
            run_id=run.id
        )

    # Fetch assistant reply
    messages = openai.beta.threads.messages.list(thread_id=thread_id)
    assistant_reply = None
    for msg in messages.data:
        if msg.role == "assistant" and msg.content:
            assistant_reply = msg.content[0].text.value
            break

    if not assistant_reply:
        assistant_reply = "Sorry, I didn't get a reply from the assistant."

    await turn_context.send_activity(Activity(
        type="message",
        text=assistant_reply
    ))


# ---------------------------
# Flask Endpoints
# ---------------------------
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
    return "Bot is running."


# ---------------------------
# Start Flask app
# ---------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app.run(host="0.0.0.0", port=3978)
