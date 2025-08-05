# Merford Bot SDK

A Teams bot that provides different levels of access based on Azure AD group membership.

## Environment Variables Required

Create a `.env` file in the root directory with the following variables:

```env
# Bot Framework Configuration
MicrosoftAppId=your_bot_app_id
MicrosoftAppPassword=your_bot_app_password

# Azure OpenAI Configuration
AZURE_OPENAI_API_KEY=your_azure_openai_api_key
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/

# Azure AD Configuration for Group Lookup
TENANT_ID=your_tenant_id
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret

# Assistant IDs (Optional - will default to Level 1 if not set)
ASSISTANT_ID=asst_r6q2Ve7DDwrzh0m3n3sbOote
```

## Assistant IDs

The bot uses different assistants based on user group membership:

- **Level 1**: `asst_r6q2Ve7DDwrzh0m3n3sbOote`
- **Level 2**: `asst_BIOAPR48tzth4k79U4h0cPtu`
- **Level 3**: `asst_SLWGUNXMQrmzpJIN1trU0zSX`
- **Level 4**: `asst_s1OefDDIgDVpqOgfp5pfCpV1`

## Azure AD Groups

The bot looks for users in these groups:
- `Level1Access`
- `Level2Access`
- `Level3Access`
- `Level4Access`

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Set up environment variables in `.env` file

3. Run the bot:
   ```bash
   python bot.py
   ```

## Troubleshooting

- If you get "No assistant found" errors, ensure the assistant IDs are correct and exist in your Azure OpenAI resource
- If you get authorization errors, check your Azure AD app registration permissions
- The bot will fallback to Level 1 assistant if no group is found or if the primary assistant fails