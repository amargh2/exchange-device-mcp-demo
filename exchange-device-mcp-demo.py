import os
import sys
import json
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

DEFAULT_SCOPES = ["User.Read", "Mail.Read", "Device.Read"]
ADMIN_SCOPES = ["User.Read.All", "DeviceManagementManagedDevices.Read.All"]
ALL_SCOPES = DEFAULT_SCOPES + ADMIN_SCOPES

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


def get_access_token(scopes=None):
    """Acquire token silently from persistent cache."""
    target_scopes = scopes or DEFAULT_SCOPES
    app, cache = build_msal_app()
    accounts = app.get_accounts()
    result = None

    if accounts:
        result = app.acquire_token_silent(target_scopes, account=accounts[0])
        save_cache(cache)

    if result and "access_token" in result:
        return result["access_token"]

    raise Exception(
        f"Exchange MCP Server requires authentication for {target_scopes}. Please run:\n"
        "  uv run exchange-device-mcp-demo.py --login\n"
        "in your terminal to complete sign-in."
    )


def login():
    """One-time interactive PKCE login via browser with all scopes requested."""
    app, cache = build_msal_app()
    print(f"Initiating login on http://localhost:{AUTH_PORT} (BROWSER={os.environ.get('BROWSER')})...")
    result = app.acquire_token_interactive(
        scopes=ALL_SCOPES,
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


@mcp.tool()
def audit_domain_migration(target_domain: str = "", old_domain: str = "") -> str:
    """Audit user UPNs, primary emails, and proxyAddresses for email domain migration and UPN cutover discrepancies."""
    try:
        token = get_access_token(scopes=ALL_SCOPES)
    except Exception:
        token = get_access_token(scopes=DEFAULT_SCOPES)

    headers = {"Authorization": f"Bearer {token}"}
    url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail,proxyAddresses,accountEnabled&$top=999"
    res = requests.get(url, headers=headers)

    if res.status_code == 403:
        return (
            "Access Denied (403 Forbidden). To run this audit across all tenant users, "
            "add the 'User.Read.All' delegated permission to your Azure App Registration "
            "and have an administrator grant consent."
        )
    if not res.ok:
        return f"Error from Microsoft Graph: {res.status_code} - {res.text}"

    users = res.json().get("value", [])
    total_users = len(users)
    mismatches = []
    missing_old_alias = []
    target_upn_count = 0
    old_upn_count = 0

    for u in users:
        upn = (u.get("userPrincipalName") or "").lower()
        mail = (u.get("mail") or "").lower()
        proxies = [p.lower() for p in (u.get("proxyAddresses") or [])]

        if target_domain and upn.endswith(f"@{target_domain.lower()}"):
            target_upn_count += 1
        elif old_domain and upn.endswith(f"@{old_domain.lower()}"):
            old_upn_count += 1

        # Check UPN vs Mail mismatch
        if mail and upn != mail:
            mismatches.append({
                "displayName": u.get("displayName"),
                "userPrincipalName": u.get("userPrincipalName"),
                "primaryMail": u.get("mail"),
                "issue": "UPN does not match primary mail"
            })

        # Check if migrated to target_domain but missing old_domain fallback alias in proxyAddresses
        if target_domain and old_domain and (upn.endswith(f"@{target_domain.lower()}") or mail.endswith(f"@{target_domain.lower()}")):
            has_old_alias = any(p.endswith(f"@{old_domain.lower()}") for p in proxies)
            if not has_old_alias:
                missing_old_alias.append({
                    "displayName": u.get("displayName"),
                    "userPrincipalName": u.get("userPrincipalName"),
                    "primaryMail": u.get("mail"),
                    "issue": f"Missing secondary alias (@{old_domain}) in proxyAddresses - incoming mail may bounce!"
                })

    report = {
        "summary": {
            "total_users_scanned": total_users,
            f"users_with_target_domain_{target_domain}": target_upn_count if target_domain else "N/A",
            f"users_with_old_domain_{old_domain}": old_upn_count if old_domain else "N/A",
            "upn_and_mail_mismatches_count": len(mismatches),
            "missing_old_alias_count": len(missing_old_alias)
        },
        "upn_vs_mail_mismatches": mismatches[:25],
        "at_risk_missing_old_alias": missing_old_alias[:25]
    }
    return json.dumps(report, indent=2)


@mcp.tool()
def get_noncompliant_devices() -> str:
    """Find noncompliant devices in Microsoft Intune and list the specific compliance policy failure reasons."""
    try:
        token = get_access_token(scopes=ALL_SCOPES)
    except Exception:
        token = get_access_token(scopes=DEFAULT_SCOPES)

    headers = {"Authorization": f"Bearer {token}"}
    url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$filter=complianceState eq 'noncompliant'&$select=id,deviceName,userPrincipalName,operatingSystem,osVersion,complianceState,lastSyncDateTime"
    res = requests.get(url, headers=headers)

    if res.status_code == 403:
        return (
            "Access Denied (403 Forbidden). To query Intune managed devices, "
            "add the 'DeviceManagementManagedDevices.Read.All' delegated permission to your Azure App Registration "
            "and have an administrator grant consent."
        )
    if not res.ok:
        return f"Error from Microsoft Graph: {res.status_code} - {res.text}"

    devices = res.json().get("value", [])
    if not devices:
        return "No noncompliant devices found in Intune! All managed devices are currently compliant."

    results = []
    for d in devices[:20]:
        dev_id = d.get("id")
        policy_url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{dev_id}/deviceCompliancePolicyStates"
        p_res = requests.get(policy_url, headers=headers)
        failing_rules = []

        if p_res.ok:
            for policy in p_res.json().get("value", []):
                if policy.get("state") in ("nonCompliant", "error", "conflict"):
                    policy_name = policy.get("displayName", "Unknown Policy")
                    for setting in policy.get("settingStates", []):
                        if setting.get("state") in ("nonCompliant", "error", "conflict"):
                            failing_rules.append({
                                "policy": policy_name,
                                "setting": setting.get("settingName"),
                                "state": setting.get("state")
                            })

        results.append({
            "deviceName": d.get("deviceName"),
            "userPrincipalName": d.get("userPrincipalName"),
            "operatingSystem": f"{d.get('operatingSystem')} {d.get('osVersion')}",
            "lastSyncDateTime": d.get("lastSyncDateTime"),
            "failing_compliance_rules": failing_rules or ["Policy flagged noncompliant, setting details not available"]
        })

    return json.dumps(results, indent=2)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("--login", "-l", "login"):
        login()
    elif sys.stdin.isatty():
        try:
            get_access_token()
            print("Exchange MCP Server: Authenticated. Running stdio server (waiting for MCP host input)...")
            mcp.run()
        except Exception:
            print("No active login found. Launching authentication...\n")
            login()
    else:
        mcp.run()
