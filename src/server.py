# Basic
import sys
import os
import logging
import contextvars
import httpx
from dotenv import load_dotenv

import msal
from fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import PlainTextResponse
from starlette.types import ASGIApp, Receive, Scope, Send

## Load environment variables
load_dotenv(override=True)

## This variable stores the user's token from the incoming HTTP request
## so the tool function can access it later. Think of it like a per-request global.
incoming_token: contextvars.ContextVar[str | None] = contextvars.ContextVar(
    "incoming_token", default=None
)

## Entra app registration config
CLIENT_ID = os.getenv("ENTRA_CLIENT_ID")
CLIENT_SECRET = os.getenv("ENTRA_CLIENT_SECRET")
TENANT_ID = os.getenv("ENTRA_TENANT_ID")

## MSAL client — handles the OBO token exchange with Entra
msal_app = msal.ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
)


## Configure logging to stdout so it shows up in the Uvicorn terminal
def configure_logging(level="ERROR"):
    try:
        logging.basicConfig(
            level=getattr(logging, level.upper()),
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler(sys.stdout)]
        )
    except Exception as e:
        print(f"Failed to set up logging: {e}", file=sys.stderr)
        sys.exit(1)


## This grabs the Authorization header from every incoming HTTP request
## and saves the Bearer token so the tool can use it for OBO.
## MCP tools don't have access to HTTP headers like FastAPI routes do,
## so this is the workaround.
class AuthHeaderMiddleware:
    def __init__(self, app: ASGIApp):
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send):
        if scope["type"] == "http":
            headers = dict(scope.get("headers", []))
            auth = headers.get(b"authorization", b"").decode()
            if auth.startswith("Bearer "):
                incoming_token.set(auth.removeprefix("Bearer "))
        await self.app(scope, receive, send)


## Setup MCP server
mcp = FastMCP(
    name="GraphOBOServer",
    instructions="This server provides one tool: get_last_email. It returns the last email for the signed-in user."
)


@mcp.custom_route("/healthz", methods=["GET"])
async def healthz(request: Request) -> PlainTextResponse:
    return PlainTextResponse("healthy", status_code=200)


@mcp.tool()
def get_last_email() -> dict:
    """Get the last email for the signed-in user."""

    ## Get the token that the upstream app sent in the Authorization header
    user_token = incoming_token.get()
    if not user_token:
        logging.error("No Bearer token in Authorization header")
        return {"error": "No Bearer token in Authorization header"}

    ## Exchange it for a Graph API token using OBO
    result = msal_app.acquire_token_on_behalf_of(
        user_assertion=user_token,
        scopes=["https://graph.microsoft.com/Mail.Read"],
    )
    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        logging.error("OBO token exchange failed: %s", error)
        return {"error": f"OBO token exchange failed: {error}"}

    graph_token = result["access_token"]

    ## Call Graph API with the OBO token to get the user's last email
    try:
        response = httpx.get(
            "https://graph.microsoft.com/v1.0/me/messages",
            headers={
                "Authorization": f"Bearer {graph_token}",
                "Prefer": 'outlook.body-content-type="text"',
            },
            params={
                "$top": 1,
                "$select": "subject,from,receivedDateTime,bodyPreview",
            },
            timeout=30,
        )
    except httpx.RequestError as e:
        logging.error("Network error calling Graph: %s", e)
        return {"error": "Unable to reach Microsoft Graph"}

    if response.status_code != 200:
        logging.error("Graph API error: %s", response.status_code)
        return {"error": f"Graph API returned {response.status_code}"}

    messages = response.json().get("value", [])
    if not messages:
        return {"message": "No emails found"}

    msg = messages[0]
    return {
        "from": msg.get("from", {}).get("emailAddress", {}).get("address", ""),
        "subject": msg.get("subject", ""),
        "date": msg.get("receivedDateTime", ""),
        "bodyPreview": msg.get("bodyPreview", ""),
    }


if __name__ == "__main__":
    import uvicorn
    configure_logging(level="INFO")
    app = mcp.http_app()
    app = AuthHeaderMiddleware(app)
    uvicorn.run(app, host="0.0.0.0", port=80)