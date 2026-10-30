import os
import sys
import webbrowser
import subprocess
import requests
import msal
from dotenv import load_dotenv

# Set the preferred browser environment variable
os.environ["BROWSER"] = "firefox"

# 1. Load configuration from .env
load_dotenv()
CLIENT_ID = os.getenv("APPID", "").strip()
TENANT_ID = os.getenv("TENANTID", "").strip()
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Mail.Read", "Device.Read"]

CACHE_FILE = os.path.join(os.path.dirname(__file__), ".token_cache.bin")
AUTH_PORT = 8400  # Fixed port for PKCE redirect (http://localhost:8400)


class InteractiveBrowserLauncher(webbrowser.BaseBrowser):
    """Custom browser launcher that prints the URL prominently and attempts to open the browser."""
    def open(self, url, new=0, autoraise=True):
        print("\n" + "=" * 78, flush=True)
        print(" MICROSOFT 365 AUTHENTICATION (PKCE) ", flush=True)
        print(" Open this link in your browser to sign in:", flush=True)
        print(f"\n {url}\n", flush=True)
        print("=" * 78 + "\n", flush=True)

        # Attempt to launch the configured BROWSER (firefox)
        browser = os.environ.get("BROWSER", "firefox")
        try:
            subprocess.Popen([browser, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except Exception:
            pass

        # Fallback to system xdg-open if available
        try:
            subprocess.Popen(["xdg-open", url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except Exception:
            pass

        return True


# Register our launcher so MSAL uses it and always displays the auth URL in the terminal
webbrowser.register("interactive_launcher", None, InteractiveBrowserLauncher(), preferred=True)


def build_msal_app():
    """Build MSAL PublicClientApplication backed by a persistent token cache file."""
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    return app, cache


def save_cache(cache):
    """Save token cache to disk if changes occurred."""
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def get_access_token():
    """Acquire token silently from persistent cache."""
    app, cache = build_msal_app()
    accounts = app.get_accounts()
    result = None

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        save_cache(cache)

    if result and "access_token" in result:
        return result["access_token"]

    raise Exception(
        "Exchange MCP Server is not authenticated. Please run:\n"
        "  uv run Exchnage-MCP-Demo.py --login\n"
        "in your terminal to complete the one-time PKCE sign-in."
    )


def login():
    """One-time interactive PKCE login via browser."""
    app, cache = build_msal_app()
    print(f"Initiating login on http://localhost:{AUTH_PORT} (BROWSER={os.environ.get('BROWSER')})...")
    result = app.acquire_token_interactive(
        scopes=SCOPES,
        port=AUTH_PORT,
        prompt="select_account"
    )

    if result and "access_token" in result:
        save_cache(cache)
        print("\nAuthentication successful! Token saved to cache.")
    else:
        err = result.get("error_description") if result else "Unknown error"
        print(f"\nAuthentication failed: {err}")


### MCP Server Section
from mcp.server.mcpserver import MCPServer

# Initialize our MCP Server
mcp = MCPServer("exchange-device-mcp-demos")


@mcp.tool()
def get_recent_emails(top: int = 5) -> str:
    """Fetch recent emails from the signed-in user's Exchange Online inbox."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    params = {"$top": top, "$select": "subject,from,receivedDateTime,bodyPreview"}
    res = requests.get("https://graph.microsoft.com/v1.0/me/messages", headers=headers, params=params)
    return str(res.json().get("value", []))


@mcp.tool()
def get_user_devices() -> str:
    """List registered devices for the signed-in user."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get("https://graph.microsoft.com/v1.0/me/registeredDevices", headers=headers)
    return str(res.json().get("value", []))


if __name__ == "__main__":
    # If explicitly passed --login, run login
    if len(sys.argv) > 1 and sys.argv[1] in ("--login", "-l", "login"):
        login()
    # If run interactively in a human terminal without arguments, check auth or prompt login
    elif sys.stdin.isatty():
        try:
            get_access_token()
            print("Exchange MCP Server: Authenticated. Running stdio server (waiting for MCP host input)...")
            mcp.run()
        except Exception:
            print("No active login found. Launching authentication...\n")
            login()
    # Otherwise run as stdio server for MCP agent/IDE
    else:
        mcp.run()
