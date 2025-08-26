import os
import logging
import openai
from flask import Flask, request, Response
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings, TurnContext, ConversationState, MemoryStorage
from botbuilder.schema import Activity
import requests

# Flask app
app = Flask(__name__)

# Logging
logging.basicConfig(level=logging.INFO)

# Bot Framework adapter
SETTINGS = BotFrameworkAdapterSettings(os.environ.get(
    "MicrosoftAppId"), os.environ.get("MicrosoftAppPassword"))
ADAPTER = BotFrameworkAdapter(SETTINGS)

# Memory storage
MEMORY = MemoryStorage()
CONVERSATION_STATE = ConversationState(MEMORY)

# OpenAI config
openai.api_key = os.environ.get("AZURE_OPENAI_API_KEY")
openai.api_base = os.environ.get("AZURE_OPENAI_ENDPOINT")
openai.api_version = os.environ.get("OPENAI_API_VERSION")
openai.api_type = "azure"

# Azure Entra credentials
TENANT_ID = os.environ.get("TENANT_ID")
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")

# Assistant mapping (NO FALLBACK)
ASSISTANT_MAP = {
    "Level 1": os.environ.get("ASSISTANT_ID_LEVEL1"),
    "Level 2": os.environ.get("ASSISTANT_ID_LEVEL2"),
    "Level 3": os.environ.get("ASSISTANT_ID_LEVEL3"),
    "Level 4": os.environ.get("ASSISTANT_ID_LEVEL4"),
}

# Get Graph API Token


def get_graph_api_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default"
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    resp = requests.post(url, data=payload, headers=headers)
    resp.raise_for_status()
    return resp.json().get("access_token")

# Get user group level


def get_user_group_level(user_id, token):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/memberOf"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()

    groups = resp.json().get("value", [])
    for g in groups:
        name = g.get("displayName")
        if name in ASSISTANT_MAP:
            return name
    return None

# Handle incoming messages


async def process_message(turn_context: TurnContext):
    if turn_context.activity.type == "message":
        user_id = turn_context.activity.from_property.id
        logging.info(f"User ID: {user_id}")

        try:
            token = get_graph_api_token()
            level = get_user_group_level(user_id, token)
            logging.info(f"User level: {level}")

            if not level:
                await turn_context.send_activity("❌ You do not have permission to access this bot.")
                return

            assistant_id = ASSISTANT_MAP.get(level)
            if not assistant_id:
                await turn_context.send_activity("❌ Error: No assistant mapping found for your access level.")
                return

            logging.info(f"Assigned assistant: {assistant_id}")

            # Create OpenAI thread
            thread = openai.beta.threads.create()
            openai.beta.threads.messages.create(
                thread_id=thread.id,
                role="user",
                content=turn_context.activity.text
            )

            # Run assistant
            run = openai.beta.threads.runs.create_and_poll(
                thread_id=thread.id,
                assistant_id=assistant_id,
            )

            if run.status == "completed":
                messages = openai.beta.threads.messages.list(
                    thread_id=thread.id)
                reply = messages.data[0].content[0].text.value
                await turn_context.send_activity(reply)
            else:
                await turn_context.send_activity("⚠️ Assistant could not complete your request.")

        except Exception as e:
            logging.error(f"Error: {str(e)}")
            await turn_context.send_activity("❌ An error occurred while processing your request.")

# Bot messages endpoint


@app.route("/api/messages", methods=["POST"])
def messages():
    if "application/json" in request.headers["Content-Type"]:
        body = request.json
    else:
        return Response(status=415)

    activity = Activity().deserialize(body)

    async def aux_func(turn_context):
        await process_message(turn_context)

    task = ADAPTER.process_activity(
        activity, body.get("authHeader", ""), aux_func)
    return Response(status=201)


if __name__ == "__main__":
    app.run(debug=True, port=3978)
