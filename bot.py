import os
import json
import requests
from flask import Flask, request
from botbuilder.core import (
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    MemoryStorage,
    ConversationState,
    TurnContext
)
from botbuilder.schema import Activity
from botbuilder.dialogs import DialogSet, DialogTurnStatus, OAuthPrompt, OAuthPromptSettings, PromptOptions

app = Flask(__name__)

# Adapter setup
settings = BotFrameworkAdapterSettings(
    app_id=os.getenv("MicrosoftAppId"),
    app_password=os.getenv("MicrosoftAppPassword")
)
adapter = BotFrameworkAdapter(settings)

# Memory & conversation state
memory = MemoryStorage()
conversation_state = ConversationState(memory)
dialogs = DialogSet(conversation_state.create_property("DialogState"))

# OAuthPrompt setup
OAUTH_CONNECTION_NAME = os.getenv("OAUTH_CONNECTION_NAME")
dialogs.add(OAuthPrompt(
    "OAuthPrompt",
    OAuthPromptSettings(
        connection_name=OAUTH_CONNECTION_NAME,
        text="Please sign in to continue",
        title="Sign In",
        timeout=300000
    )
))

# Group to Assistant mapping
group_to_assistant = {
    "Level1Access": os.getenv("ASSISTANT_ID_LEVEL1"),
    "Level2Access": os.getenv("ASSISTANT_ID_LEVEL2"),
    "Level3Access": os.getenv("ASSISTANT_ID_LEVEL3"),
    "Level4Access": os.getenv("ASSISTANT_ID_LEVEL4"),
}

# Azure AD & OpenAI
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
OPENAI_API_VERSION = os.getenv("OPENAI_API_VERSION")

# Util: get user groups


def get_user_groups(token):
    headers = {
        "Authorization": f"Bearer {token}"
    }
    url = "https://graph.microsoft.com/v1.0/me/memberOf"
    res = requests.get(url, headers=headers)
    print(f"[Graph API] Response: {res.status_code}")
    print(res.text)
    if res.status_code == 200:
        data = res.json()
        return [group["displayName"] for group in data["value"]]
    return []

# Util: call Azure OpenAI Assistant


def call_assistant(assistant_id, user_input):
    url = f"{AZURE_OPENAI_ENDPOINT}/openai/assistants/{assistant_id}/threads"
    headers = {
        "api-key": AZURE_OPENAI_API_KEY,
        "Content-Type": "application/json"
    }
    body = {
        "messages": [{"role": "user", "content": user_input}]
    }
    res = requests.post(url, headers=headers, json=body)
    if res.status_code == 200:
        return res.json()["choices"][0]["message"]["content"]
    else:
        print(f"[OpenAI Error]: {res.status_code} {res.text}")
        return "Sorry, I couldn't get a response from the assistant."


@app.route("/api/messages", methods=["POST"])
async def messages():
    activity = Activity().deserialize(request.json)

    async def call_bot(turn_context: TurnContext):
        dialog_context = await dialogs.create_context(turn_context)

        if activity.type != "message":
            return

        if dialog_context.active_dialog is None:
            prompt_options = PromptOptions(prompt=Activity(
                type="message", text="Please sign in"))
            await dialog_context.prompt("OAuthPrompt", prompt_options)
        else:
            result = await dialog_context.continue_dialog()
            if result.status == DialogTurnStatus.Complete:
                token_response = result.result
                access_token = token_response.token
                print(f"[Access Token] {access_token}")

                # Log Graph ID token if available
                print(
                    f"[SSO] OAuth Token Response: {json.dumps(token_response.additional_properties, indent=2)}")

                user_groups = get_user_groups(access_token)
                print(f"[User Groups] {user_groups}")

                assistant_id = None
                for group in user_groups:
                    if group in group_to_assistant:
                        assistant_id = group_to_assistant[group]
                        break

                if assistant_id is None:
                    await turn_context.send_activity("You are not assigned to any group assistant.")
                else:
                    user_input = activity.text
                    response = call_assistant(assistant_id, user_input)
                    await turn_context.send_activity(response)

    await adapter.process_activity(activity, "", call_bot)
    return "", 200
