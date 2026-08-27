# Basic
import os
import requests
from dotenv import load_dotenv

# MSAL for On-Behalf-Of token exchange
import msal

# FastMCP
from fastmcp import FastMCP

# Required for health check
from starlette.requests import Request
from starlette.responses import PlainTextResponse

# Load environmental variables
load_dotenv(override=True)

# Entra app registration config
CLIENT_ID = os.getenv("ENTRA_CLIENT_ID")
CLIENT_SECRET = os.getenv("ENTRA_CLIENT_SECRET")
TENANT_ID = os.getenv("ENTRA_TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_SCOPE = ["https://graph.microsoft.com/Mail.Read"]

# Build the MSAL confidential client once at startup
_msal_app = msal.ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=AUTHORITY,
)

# Setup MCP server with tool instructions
mcp = FastMCP(
    name="GraphOBOServer",
    instructions="""
        This server provides one tool: get_last_email.
        It returns the last email for the signed-in user.
        The caller must supply a valid user assertion token from the upstream app.
        Response includes `from`, `to`, `subject`, `date`, and `body` fields.
    """
)

# Add a health check endpoint
@mcp.custom_route("/healthz", methods=["GET"])
async def healthz(request: Request) -> PlainTextResponse:
    return PlainTextResponse("healthy", status_code=200)


def _exchange_obo_token(user_assertion: str) -> str:
    """Exchange the upstream user assertion for a Graph API access token via OBO flow."""
    result = _msal_app.acquire_token_on_behalf_of(
        user_assertion=user_assertion,
        scopes=GRAPH_SCOPE,
    )
    if "access_token" in result:
        return result["access_token"]

    error_desc = result.get("error_description", result.get("error", "Unknown OBO error"))
    raise RuntimeError(f"OBO token exchange failed: {error_desc}")


# Create the get_last_email tool
@mcp.tool()
def get_last_email(user_assertion: str) -> dict:
    """Get the last email for the signed-in user.

    Args:
        user_assertion: The access token from the upstream application to exchange via OBO flow.

    Returns:
        A JSON object containing:
        - from: sender of the email
        - to: recipients of the email
        - subject: subject of the email
        - date: date the email was received
        - body: body content of the email
    """

    # Exchange the upstream token for a Graph token via OBO
    try:
        graph_token = _exchange_obo_token(user_assertion)
    except RuntimeError as e:
        print(f"OBO token exchange error: {e}")
        return {
            "error": {
                "type": "tool_error",
                "code": "OBO_TOKEN_ERROR",
                "message": str(e),
            }
        }

    # Call Microsoft Graph to get the last email
    try:
        result = requests.get(
            url="https://graph.microsoft.com/v1.0/me/messages",
            headers={"Authorization": f"Bearer {graph_token}"},
            params={
                "$top": 1,
                "$orderby": "receivedDateTime desc",
                "$select": "from,toRecipients,subject,receivedDateTime,body",
            },
            timeout=30,
        )
    except requests.RequestException as e:
        print(f"Network error fetching email: {e}")
        return {
            "error": {
                "type": "tool_error",
                "code": "GRAPH_NETWORK_ERROR",
                "message": "Unable to reach Microsoft Graph",
            }
        }

    if result.status_code != 200:
        print(f"Graph API error: {result.status_code} - {result.text}")
        return {
            "error": {
                "type": "tool_error",
                "code": "GRAPH_API_ERROR",
                "message": f"Graph API returned {result.status_code}",
            }
        }

    data = result.json()
    messages = data.get("value", [])
    if not messages:
        return {"message": "No emails found"}

    msg = messages[0]
    return {
        "from": msg.get("from", {}).get("emailAddress", {}).get("address", ""),
        "to": [
            r.get("emailAddress", {}).get("address", "")
            for r in msg.get("toRecipients", [])
        ],
        "subject": msg.get("subject", ""),
        "date": msg.get("receivedDateTime", ""),
        "body": msg.get("body", {}).get("content", ""),
    }

if __name__ == "__main__":
    mcp.run("streamable-http", host="0.0.0.0", port=8080, show_banner="My OBO MCP Server")