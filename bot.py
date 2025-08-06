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

# Load environment variables
load_dotenv()

# Credentials
APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

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
access_token_cache = {"token": None, "expiry": 0}


def get_graph_api_token():
    now = time.time()
    if access_token_cache["token"] and now < access_token_cache["expiry"]:
        return access_token_cache["token"]

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }

    response = requests.post(url, headers=headers, data=data)
    if response.status_code != 200:
        raise Exception(
            f"Failed to get token: {response.status_code}, {response.text}")

    token_data = response.json()
    token = token_data["access_token"]
    expires_in = token_data.get("expires_in", 1800)
    access_token_cache["token"] = token
    access_token_cache["expiry"] = now + \
        expires_in - 60  # renew 1 min before expiry

    return token

# Group Lookup Function with displayName selection


def get_user_group_level(user_id):
    print(f"=== GROUP LOOKUP DEBUG ===")
    print(f"Looking up groups for user: {user_id}")

    access_token = os.getenv("AZURE_ENTRA_TOKEN")
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/memberOf?$select=id,displayName"
    headers = {"Authorization": f"Bearer {access_token}"}

    print(f"Making request to: {url}")
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        logging.warning(
            f"Group lookup failed: {response.status_code} - {response.text}")
        print(f"ERROR: HTTP {response.status_code} - {response.text}")
        return None

    groups = response.json().get("value", [])
    print(f"Found {len(groups)} groups for user {user_id}")

    for group in groups:
        name = group.get("displayName")
        print(f"  - Group: {name}")
        if name == "Level1Access":
            print(f"    -> Returning Level 1")
            return "Level 1"
        elif name == "Level2Access":
            print(f"    -> Returning Level 2")
            return "Level 2"
        elif name == "Level3Access":
            print(f"    -> Returning Level 3")
            return "Level 3"
        elif name == "Level4Access":
            print(f"    -> Returning Level 4")
            return "Level 4"

    print(f"No matching level groups found for user {user_id}")
    print(f"==========================")
    return None  # Not found

# Main Bot Handler


async def handle_message(turn_context: TurnContext):
    # Handle conversationUpdate event to send greeting once
    if turn_context.activity.type == "conversationUpdate":
        members_added = turn_context.activity.members_added
        if members_added:
            for member in members_added:
                if member.id == turn_context.activity.recipient.id:
                    await turn_context.send_activity("Hello! How can I assist you today?")
        return

    # Only handle non-empty 'message' activities
    if turn_context.activity.type != "message" or not turn_context.activity.text or not turn_context.activity.text.strip():
        return  # Ignore empty or whitespace messages

    user_id = turn_context.activity.from_property.aad_object_id or turn_context.activity.from_property.id
    user_input = turn_context.activity.text

    try:
        await turn_context.send_activity(Activity(type="typing"))

        level = get_user_group_level(user_id)
        print(f"User {user_id} has level: {level}")

        assistant_map = {
            "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
            "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
            "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
            "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
        }
        assistant_id = assistant_map.get(level, os.getenv("ASSISTANT_ID"))
        print(f"Using assistant: {assistant_id} for user {user_id}")
        # Get or create conversation thread
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

        # Run assistant
        print(f"=== ASSISTANT RUN DEBUG ===")
        print(f"Creating run with assistant_id: {assistant_id}")
        print(f"Thread ID: {thread_id}")

        try:
            run = openai.beta.threads.runs.create(
                assistant_id=assistant_id,
                thread_id=thread_id
            )
            print(f"Run created successfully with ID: {run.id}")
        except Exception as e:
            print(f"ERROR creating run: {e}")
            logging.error(
                f"Failed to create run with assistant {assistant_id}: {e}")
            # Try with Level 1 assistant as fallback
            fallback_assistant = "asst_r6q2Ve7DDwrzh0m3n3sbOote"
            print(f"Trying fallback assistant: {fallback_assistant}")
            run = openai.beta.threads.runs.create(
                assistant_id=fallback_assistant,
                thread_id=thread_id
            )
            print(f"Fallback run created with ID: {run.id}")

        print(f"============================")

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

    # Send the assistant reply message back to the user
    await turn_context.send_activity(Activity(
        type="message",
        text=assistant_reply,
        recipient=turn_context.activity.from_property,
        from_property=turn_context.activity.recipient,
        conversation=turn_context.activity.conversation,
        channel_id=turn_context.activity.channel_id,
        service_url=turn_context.activity.service_url
    ))

# Flask Endpoints


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


@app.route("/debug", methods=["GET"])
def debug_info():
    try:
        # Check if we can get a token
        token = get_graph_api_token()
        token_status = "Token obtained successfully"
    except Exception as e:
        token_status = f"Token error: {str(e)}"

    return {
        "status": "Bot is running",
        "token_status": token_status,
        "environment_vars": {
            "TENANT_ID": "SET" if os.getenv("TENANT_ID") else "NOT SET",
            "CLIENT_ID": "SET" if os.getenv("CLIENT_ID") else "NOT SET",
            "ASSISTANT_ID": os.getenv("ASSISTANT_ID", "NOT SET"),
            "AZURE_OPENAI_ENDPOINT": "SET" if os.getenv("AZURE_OPENAI_ENDPOINT") else "NOT SET"
        }
    }


# Run Flask Server
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    app.run(host="0.0.0.0", port=3978)
