"""
================================================================================
EXCHANGE & INTUNE AI ASSISTANT (MCP SERVER)
================================================================================
What is this file?
  This script connects an AI Assistant (via Model Context Protocol / MCP) to
  Microsoft 365. It lets the agent answer questions about your tenant in real time.

How is this different from PowerShell or Admin Portals?
  - Admin Portals: Require clicking through 10+ blade menus in Entra / Intune / Exchange.
  - PowerShell: Requires writing complex loops with Get-MgUser or Get-MgDevice.
  - This Tool: Connects to the exact same Microsoft Graph APIs, but packages the
    answers into conversational tools for an AI agent.

Security & Permissions:
  - 100% standard Microsoft Entra ID OAuth 2.0 (PKCE interactive login).
  - No secrets or passwords stored anywhere in code.
  - Runs with DELEGATED permissions: The agent can only see what the signed-in
    administrator is authorized to see in the tenant.
================================================================================
"""

import os
import sys
import json
import webbrowser
import subprocess
import requests
import msal
from dotenv import load_dotenv

# Set preferred browser for the sign-in prompt (can be changed to chrome, edge, etc.)
os.environ["BROWSER"] = "firefox"

# ==============================================================================
# SECTION 1: CONFIGURATION & PERMISSIONS
# ==============================================================================
# Load tenant and app credentials from the local .env file (never hardcoded in git)
load_dotenv()
CLIENT_ID = os.getenv("APPID", "").strip()     # Your Azure App Registration ID
TENANT_ID = os.getenv("TENANTID", "").strip() # Your Microsoft 365 Tenant ID
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Scopes explain: These are the exact Microsoft Graph permissions we request:
# - User.Read, Mail.Read, Device.Read: Basic permissions to read own mailbox and devices.
# - User.Read.All: Equivalent to reading the directory (like Get-MgUser).
# - DeviceManagementManagedDevices.Read.All: Reading Intune device compliance states.
DEFAULT_SCOPES = ["User.Read", "Mail.Read", "Device.Read"]
ADMIN_SCOPES = ["User.Read.All", "DeviceManagementManagedDevices.Read.All"]
ALL_SCOPES = DEFAULT_SCOPES + ADMIN_SCOPES

# Where to save the token on disk so you don't have to sign in every time
CACHE_FILE = os.path.join(os.path.dirname(__file__), ".token_cache.bin")

# Fixed localhost port for the Microsoft browser redirect
AUTH_PORT = 8400


# ==============================================================================
# SECTION 2: BROWSER LOGIN HELPER
# ==============================================================================
class InteractiveBrowserLauncher(webbrowser.BaseBrowser):
    """
    Ensures that when a sign-in is required:
    1. The Microsoft login link is printed clearly in the terminal (so you can click it).
    2. It attempts to launch your system browser automatically in the background.
    """
    def open(self, url, new=0, autoraise=True):
        print("\n" + "=" * 78, flush=True)
        print(" MICROSOFT 365 SIGN-IN REQUIRED (PKCE) ", flush=True)
        print(" Open this link in your browser to sign in:", flush=True)
        print(f"\n {url}\n", flush=True)
        print("=" * 78 + "\n", flush=True)

        # Try launching the configured browser (Firefox)
        browser = os.environ.get("BROWSER", "firefox")
        try:
            subprocess.Popen([browser, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except Exception:
            pass

        # Fallback to system default opener (xdg-open on Linux)
        try:
            subprocess.Popen(["xdg-open", url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except Exception:
            pass

        return True


# Register our launcher so MSAL uses it whenever browser authentication runs
webbrowser.register("interactive_launcher", None, InteractiveBrowserLauncher(), preferred=True)


# ==============================================================================
# SECTION 3: TOKEN CACHING & AUTHENTICATION (MSAL)
# ==============================================================================
def build_msal_app():
    """
    Creates the Microsoft Authentication Library (MSAL) client.
    Loads the persistent cache file (.token_cache.bin) so previous logins are remembered.
    """
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    return app, cache


def save_cache(cache):
    """Saves updated access/refresh tokens to .token_cache.bin."""
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def get_access_token(scopes=None):
    """
    Silently retrieves an access token from the cache.
    If the token has expired, MSAL automatically uses the refresh token in the
    background without prompting the user.
    """
    target_scopes = scopes or DEFAULT_SCOPES
    app, cache = build_msal_app()
    accounts = app.get_accounts()
    result = None

    if accounts:
        # Silent token renewal (no popups or prompts)
        result = app.acquire_token_silent(target_scopes, account=accounts[0])
        save_cache(cache)

    if result and "access_token" in result:
        return result["access_token"]

    raise Exception(
        f"Authentication required for {target_scopes}. Please run:\n"
        "  uv run exchange-device-mcp-demo.py --login\n"
        "in your terminal to sign in."
    )


def login():
    """
    Interactive sign-in: opens the browser on localhost:8400, completes the
    OAuth PKCE exchange, and writes the tokens to .token_cache.bin.
    """
    app, cache = build_msal_app()
    print(f"Starting Microsoft login on http://localhost:{AUTH_PORT}...")
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


# ==============================================================================
# SECTION 4: MCP TOOLS (WHAT THE AGENT CAN CALL)
# ==============================================================================
from mcp.server.mcpserver import MCPServer

# Initialize our MCP Server with the project name
mcp = MCPServer("exchange-device-mcp-demos")


# ------------------------------------------------------------------------------
# TOOL 1: Read Recent Emails
# ------------------------------------------------------------------------------
@mcp.tool()
def get_recent_emails(top: int = 5) -> str:
    """
    Fetch recent emails from the signed-in user's Exchange Online inbox.
    Returns: subject, sender, received date/time, and a short body preview.
    """
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    params = {"$top": top, "$select": "subject,from,receivedDateTime,bodyPreview"}
    res = requests.get("https://graph.microsoft.com/v1.0/me/messages", headers=headers, params=params)
    return str(res.json().get("value", []))


# ------------------------------------------------------------------------------
# TOOL 2: Read Personal Registered Devices
# ------------------------------------------------------------------------------
@mcp.tool()
def get_user_devices() -> str:
    """
    List registered hardware/phones for the signed-in user from Microsoft Entra ID.
    Equivalent to viewing 'My Devices' in https://myaccount.microsoft.com.
    """
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get("https://graph.microsoft.com/v1.0/me/registeredDevices", headers=headers)
    return str(res.json().get("value", []))


# ------------------------------------------------------------------------------
# TOOL 3: Domain Migration & UPN Cutover Auditor
# ------------------------------------------------------------------------------
@mcp.tool()
def audit_domain_migration(target_domain: str = "", old_domain: str = "") -> str:
    """
    Audit user UPNs, primary emails, and proxyAddresses for email domain migrations.
    
    Why this matters during an email cutover:
      1. UPN vs Mail Mismatch: A user's sign-in name changed, but their primary email didn't.
      2. Missing Alias Risk: If an account is moved to @newdomain.com, but its old address
         (@olddomain.com) wasn't preserved as an alias in 'proxyAddresses', incoming emails
         from external clients or vendors sent to the old address will bounce (NDR 550)!
    """
    try:
        token = get_access_token(scopes=ALL_SCOPES)
    except Exception:
        token = get_access_token(scopes=DEFAULT_SCOPES)

    headers = {"Authorization": f"Bearer {token}"}
    # Microsoft Graph endpoint equivalent to running:
    # Get-MgUser -Select id,displayName,userPrincipalName,mail,proxyAddresses,accountEnabled -All
    url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail,proxyAddresses,accountEnabled&$top=999"
    res = requests.get(url, headers=headers)

    if res.status_code == 403:
        return (
            "Access Denied (403 Forbidden). To audit users across the tenant, "
            "add the 'User.Read.All' delegated permission to your Azure App Registration "
            "and grant admin consent in Azure Portal."
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
        # proxyAddresses holds all Exchange email aliases (e.g. SMTP:primary, smtp:alias)
        proxies = [p.lower() for p in (u.get("proxyAddresses") or [])]

        # 1. Count progress towards target domain
        if target_domain and upn.endswith(f"@{target_domain.lower()}"):
            target_upn_count += 1
        elif old_domain and upn.endswith(f"@{old_domain.lower()}"):
            old_upn_count += 1

        # 2. Check for UPN and Primary Email discrepancies
        if mail and upn != mail:
            mismatches.append({
                "displayName": u.get("displayName"),
                "userPrincipalName": u.get("userPrincipalName"),
                "primaryMail": u.get("mail"),
                "issue": "UPN (sign-in) does not match primary mail address"
            })

        # 3. Check for Bounce Risk (missing old domain fallback alias)
        if target_domain and old_domain and (upn.endswith(f"@{target_domain.lower()}") or mail.endswith(f"@{target_domain.lower()}")):
            has_old_alias = any(p.endswith(f"@{old_domain.lower()}") for p in proxies)
            if not has_old_alias:
                missing_old_alias.append({
                    "displayName": u.get("displayName"),
                    "userPrincipalName": u.get("userPrincipalName"),
                    "primaryMail": u.get("mail"),
                    "issue": f"Missing @{old_domain} fallback in proxyAddresses. Senders emailing the old address will get bounce errors!"
                })

    report = {
        "summary": {
            "total_users_scanned": total_users,
            f"users_on_new_domain_{target_domain}": target_upn_count if target_domain else "N/A",
            f"users_still_on_old_domain_{old_domain}": old_upn_count if old_domain else "N/A",
            "upn_and_mail_mismatches_count": len(mismatches),
            "missing_old_alias_bounce_risk_count": len(missing_old_alias)
        },
        "upn_vs_mail_mismatches": mismatches[:25],
        "at_risk_missing_old_alias": missing_old_alias[:25]
    }
    return json.dumps(report, indent=2)


# ------------------------------------------------------------------------------
# TOOL 4: Intune Noncompliance & Root-Cause Inspector
# ------------------------------------------------------------------------------
@mcp.tool()
def get_noncompliant_devices() -> str:
    """
    Find noncompliant devices in Microsoft Intune and report the EXACT failure reasons.
    
    Why this is faster than the Intune Admin Center:
      In the web portal, finding why a machine is noncompliant takes 5+ clicks per device.
      This tool pulls all noncompliant devices across the company in one query and
      extracts the specific failing policy setting (e.g. BitLocker missing, OS outdated).
    """
    try:
        token = get_access_token(scopes=ALL_SCOPES)
    except Exception:
        token = get_access_token(scopes=DEFAULT_SCOPES)

    headers = {"Authorization": f"Bearer {token}"}
    # Microsoft Graph endpoint equivalent to running:
    # Get-MgDeviceManagementManagedDevice -Filter "complianceState eq 'noncompliant'"
    url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$filter=complianceState eq 'noncompliant'&$select=id,deviceName,userPrincipalName,operatingSystem,osVersion,complianceState,lastSyncDateTime"
    res = requests.get(url, headers=headers)

    if res.status_code == 403:
        return (
            "Access Denied (403 Forbidden). To query Intune managed devices, "
            "add the 'DeviceManagementManagedDevices.Read.All' delegated permission to your Azure App Registration "
            "and grant admin consent in Azure Portal."
        )
    if not res.ok:
        return f"Error from Microsoft Graph: {res.status_code} - {res.text}"

    devices = res.json().get("value", [])
    if not devices:
        return "All clear! No noncompliant devices found in Intune. All managed machines meet policy requirements."

    results = []
    # Inspect each noncompliant device to find which specific setting triggered the failure
    for d in devices[:20]:
        dev_id = d.get("id")
        policy_url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{dev_id}/deviceCompliancePolicyStates"
        p_res = requests.get(policy_url, headers=headers)
        failing_rules = []

        if p_res.ok:
            for policy in p_res.json().get("value", []):
                if policy.get("state") in ("nonCompliant", "error", "conflict"):
                    policy_name = policy.get("displayName", "Unknown Policy")
                    # Inspect settingStates to find the exact rule that broke
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


# ==============================================================================
# SECTION 5: SERVER EXECUTION & ENTRYPOINT
# ==============================================================================
if __name__ == "__main__":
    # Case 1: Human explicitly ran `--login`
    if len(sys.argv) > 1 and sys.argv[1] in ("--login", "-l", "login"):
        login()

    # Case 2: Human ran this in a terminal without arguments
    # Check if we're already logged in; if not, kick off sign-in automatically
    elif sys.stdin.isatty():
        try:
            get_access_token()
            print("Exchange MCP Server: Authenticated. Running stdio server (waiting for AI agent commands)...")
            mcp.run()
        except Exception:
            print("No active login found. Launching authentication...\n")
            login()

    # Case 3: Spawned in the background by an AI agent or IDE (stdio mode)
    else:
        mcp.run()
