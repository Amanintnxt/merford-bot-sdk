import os
import time
import openai
import asyncio
import logging
import requests
import msal
from dotenv import load_dotenv
from flask import Flask, request, Response
from botbuilder.core import BotFrameworkAdapterSettings, BotFrameworkAdapter, TurnContext
from botbuilder.schema import Activity

# Load environment variables
load_dotenv()

# Credentials and config
APP_ID = os.getenv("MicrosoftAppId", "")
APP_PASSWORD = os.getenv("MicrosoftAppPassword", "")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")

# MSAL setup for device code flow (public client, no client secret needed)
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Directory.Read.All",
          "Group.Read.All", "GroupMember.Read.All"]
msal_app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# OpenAI config
openai.api_type = "azure"
openai.api_version = "2024-05-01-preview"
openai.api_key = AZURE_OPENAI_API_KEY
openai.azure_endpoint = AZURE_OPENAI_ENDPOINT.rstrip("/")

# Flask & Bot Framework setup
app = Flask(__name__)
adapter_settings = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
adapter = BotFrameworkAdapter(adapter_settings)

# Memory stores
thread_map = {}
token_cache = {"token": None, "expiry": 0}


def get_token_device_code_flow():
    now = time.time()
    if token_cache["token"] and now < token_cache["expiry"]:
        return token_cache["token"]

    flow = msal_app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        raise Exception("Failed to initiate device flow: " + str(flow))

    # Instruct user to sign in using the printed code and URL
    print(flow["message"])

    result = msal_app.acquire_token_by_device_flow(
        flow)  # This blocks until sign-in completed

    if "access_token" in result:
        token_cache["token"] = result["access_token"]
        token_cache["expiry"] = now + int(result.get("expires_in", 3600)) - 60
        return token_cache["token"]
    else:
        raise Exception("Failed to acquire token: " +
                        str(result.get("error_description", result)))


def get_user_group_level(user_id):
    access_token = get_token_device_code_flow()
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/memberOf?$select=id,displayName"
    headers = {"Authorization": f"Bearer {access_token}"}

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
    return None


async def handle_message(turn_context: TurnContext):
    if turn_context.activity.type == "conversationUpdate":
        members_added = turn_context.activity.members_added
        if members_added:
            for member in members_added:
                if member.id == turn_context.activity.recipient.id:
                    await turn_context.send_activity("Hello! How can I assist you today?")
        return

    if turn_context.activity.type != "message" or not turn_context.activity.text or not turn_context.activity.text.strip():
        return

    user_id = turn_context.activity.from_property.aad_object_id or turn_context.activity.from_property.id
    user_input = turn_context.activity.text

    try:
        await turn_context.send_activity(Activity(type="typing"))

        # Optionally, override user_id for testing:
        # user_id = "a62bf818-86a9-4a27-80d3-b087ea19e3f8"

        level = get_user_group_level(user_id)
        assistant_map = {
            "Level 1": "asst_r6q2Ve7DDwrzh0m3n3sbOote",
            "Level 2": "asst_BIOAPR48tzth4k79U4h0cPtu",
            "Level 3": "asst_SLWGUNXMQrmzpJIN1trU0zSX",
            "Level 4": "asst_s1OefDDIgDVpqOgfp5pfCpV1"
        }
        assistant_id = assistant_map.get(level, os.getenv("ASSISTANT_ID"))
        print(f"Using assistant: {assistant_id} for user {user_id}")

        thread_id = thread_map.get(user_id)
        if not thread_id:
            thread = openai.beta.threads.create()
            thread_id = thread.id
            thread_map[user_id] = thread_id

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
