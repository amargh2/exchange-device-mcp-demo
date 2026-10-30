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

# ==============================================================================
# SECTION 1: CONFIGURATION & PERMISSIONS
# ==============================================================================
# Load tenant and app credentials from the local .env file (never hardcoded in git)
load_dotenv()
CLIENT_ID = os.getenv("APPID", "").strip()     # Your Azure App Registration ID
TENANT_ID = os.getenv("TENANTID", "").strip() # Your Microsoft 365 Tenant ID
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Scopes explain: These are the exact Microsoft Graph permissions we request:
# Notice: WE DO NOT REQUEST 'Mail.Read' (we cannot read user email messages!)
# - User.Read, Device.Read: Read own basic profile and registered personal hardware.
# - MailboxSettings.Read: Read Out-of-Office auto-replies, regional time zone, and language.
# - Calendars.Read: Read calendar folder sharing permissions (ACLs).
# - User.Read.All: Directory-wide read for migration audits (like Get-MgUser).
# - DeviceManagementManagedDevices.Read.All: Reading Intune device compliance states.
DEFAULT_SCOPES = [
    "User.Read",
    "Device.Read",
    "MailboxSettings.Read",
    "Calendars.Read"
]
ADMIN_SCOPES = [
    "User.Read.All",
    "DeviceManagementManagedDevices.Read.All"
]
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
    2. It attempts to launch your system browser automatically in the background (works on Windows, Mac, and Linux).
    """
    def open(self, url, new=0, autoraise=True):
        print("\n" + "=" * 78, flush=True)
        print(" MICROSOFT 365 SIGN-IN REQUIRED (PKCE) ", flush=True)
        print(" Open this link in your browser to sign in:", flush=True)
        print(f"\n {url}\n", flush=True)
        print("=" * 78 + "\n", flush=True)

        # On Windows, os.startfile opens the default system browser natively (Edge, Chrome, etc.)
        if sys.platform == "win32":
            try:
                os.startfile(url)
                return True
            except Exception:
                pass

        # Try launching the configured browser (Firefox / Chrome / etc.)
        browser = os.environ.get("BROWSER")
        if browser:
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
# TOOL 1: Calendar Folder Sharing & Permissions
# ------------------------------------------------------------------------------
@mcp.tool()
def get_calendar_permissions(user_email: str = "") -> str:
    """
    Check who has permission to view or edit a user's calendar.
    PowerShell equivalent: Get-MailboxFolderPermission -Identity user:\\Calendar
    
    Parameters:
      user_email (optional): Target user UPN or email (e.g. 'john@company.com').
                             If left blank, checks your own calendar.
    """
    access_token = get_access_token()
    http_headers = {"Authorization": f"Bearer {access_token}"}

    # If a specific email is provided, query that mailbox; otherwise query yourself (/me)
    if user_email.strip():
        graph_endpoint = f"https://graph.microsoft.com/v1.0/users/{user_email.strip()}/calendar/calendarPermissions"
    else:
        graph_endpoint = "https://graph.microsoft.com/v1.0/me/calendar/calendarPermissions"

    graph_response = requests.get(graph_endpoint, headers=http_headers)

    if graph_response.status_code == 404:
        return "Mailbox not found or not provisioned in Exchange Online (user may not have an active Exchange Online license)."
    if not graph_response.ok:
        return f"Error from Microsoft Graph: {graph_response.status_code} - {graph_response.text}"

    permissions_list = graph_response.json().get("value", [])
    if not permissions_list:
        return "No explicit calendar sharing permissions found."

    clean_permissions = []
    for permission in permissions_list:
        user_info = permission.get("emailAddress", {})
        shared_with_name = user_info.get("name", "Default / Anonymous")
        shared_with_email = user_info.get("address", "N/A")
        access_role = permission.get("role", "custom")
        allowed_roles = permission.get("allowedRoles", [])

        clean_permissions.append({
            "sharedWith": shared_with_name,
            "email": shared_with_email,
            "accessRole": access_role,
            "allowedCapabilities": allowed_roles
        })

    return json.dumps(clean_permissions, indent=2)


# ------------------------------------------------------------------------------
# TOOL 2: Mailbox Settings & Out-of-Office (Auto-Replies)
# ------------------------------------------------------------------------------
@mcp.tool()
def get_mailbox_settings(user_email: str = "") -> str:
    """
    Check mailbox configuration, Out-of-Office (automatic replies), and regional time zone.
    PowerShell equivalent: Get-MailboxAutoReplyConfiguration & Get-MailboxRegionalConfiguration
    
    Parameters:
      user_email (optional): Target user UPN or email (e.g. 'john@company.com').
                             If left blank, checks your own mailbox.
    """
    access_token = get_access_token()
    http_headers = {"Authorization": f"Bearer {access_token}"}

    if user_email.strip():
        graph_endpoint = f"https://graph.microsoft.com/v1.0/users/{user_email.strip()}/mailboxSettings"
    else:
        graph_endpoint = "https://graph.microsoft.com/v1.0/me/mailboxSettings"

    graph_response = requests.get(graph_endpoint, headers=http_headers)

    if graph_response.status_code == 404:
        return "Mailbox not found or not provisioned in Exchange Online (user may not have an active Exchange Online license)."
    if not graph_response.ok:
        return f"Error from Microsoft Graph: {graph_response.status_code} - {graph_response.text}"

    settings_data = graph_response.json()
    auto_reply = settings_data.get("automaticRepliesSetting", {})

    clean_summary = {
        "autoRepliesStatus": auto_reply.get("status", "unknown"),
        "autoReplyMessageInternal": auto_reply.get("internalReplyMessage", ""),
        "autoReplyMessageExternal": auto_reply.get("externalReplyMessage", ""),
        "scheduledStartTime": auto_reply.get("scheduledStartDateTime", {}).get("dateTime", "N/A"),
        "scheduledEndTime": auto_reply.get("scheduledEndDateTime", {}).get("dateTime", "N/A"),
        "mailboxTimeZone": settings_data.get("timeZone", "N/A"),
        "mailboxLanguage": settings_data.get("language", {}).get("locale", "N/A"),
        "mailboxType": settings_data.get("userPurpose", "user")
    }

    return json.dumps(clean_summary, indent=2)


# ------------------------------------------------------------------------------
# TOOL 2: Read Personal Registered Devices
# ------------------------------------------------------------------------------
@mcp.tool()
def get_user_devices() -> str:
    """
    List registered hardware/phones for the signed-in user from Microsoft Entra ID.
    Equivalent to viewing 'My Devices' in https://myaccount.microsoft.com.
    """
    access_token = get_access_token()
    http_headers = {"Authorization": f"Bearer {access_token}"}
    graph_response = requests.get(
        "https://graph.microsoft.com/v1.0/me/registeredDevices",
        headers=http_headers
    )
    user_devices = graph_response.json().get("value", [])
    return str(user_devices)


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
        access_token = get_access_token(scopes=ALL_SCOPES)
    except Exception:
        access_token = get_access_token(scopes=DEFAULT_SCOPES)

    http_headers = {"Authorization": f"Bearer {access_token}"}

    # Microsoft Graph endpoint equivalent to running:
    # Get-MgUser -Select id,displayName,userPrincipalName,mail,proxyAddresses,accountEnabled -All
    graph_endpoint = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail,proxyAddresses,accountEnabled&$top=999"
    graph_response = requests.get(graph_endpoint, headers=http_headers)

    if graph_response.status_code == 403:
        return (
            "Access Denied (403 Forbidden). To audit users across the tenant, "
            "add the 'User.Read.All' delegated permission to your Azure App Registration "
            "and grant admin consent in Azure Portal."
        )
    if not graph_response.ok:
        return f"Error from Microsoft Graph: {graph_response.status_code} - {graph_response.text}"

    # Extract the user list from the Microsoft Graph response
    all_users = graph_response.json().get("value", [])
    total_users_count = len(all_users)

    # Tracking lists and counts for our audit report
    mismatched_accounts = []
    missing_old_alias_accounts = []
    target_domain_user_count = 0
    old_domain_user_count = 0

    # Step through every user account in the tenant
    for user_account in all_users:
        user_display_name = user_account.get("displayName")
        user_sign_in_name = (user_account.get("userPrincipalName") or "").lower()
        user_primary_email = (user_account.get("mail") or "").lower()

        # proxyAddresses contains all Exchange routing addresses (e.g., SMTP:primary, smtp:alias)
        raw_aliases = user_account.get("proxyAddresses") or []
        email_aliases = []
        for alias in raw_aliases:
            email_aliases.append(alias.lower())

        # Check 1: Count how many users have the target domain vs old domain
        if target_domain and user_sign_in_name.endswith(f"@{target_domain.lower()}"):
            target_domain_user_count += 1
        elif old_domain and user_sign_in_name.endswith(f"@{old_domain.lower()}"):
            old_domain_user_count += 1

        # Check 2: Check for UPN vs Primary Email mismatches
        # In a healthy migration, sign-in name (UPN) and primary email should generally match
        if user_primary_email and (user_sign_in_name != user_primary_email):
            mismatched_accounts.append({
                "displayName": user_display_name,
                "userPrincipalName": user_account.get("userPrincipalName"),
                "primaryMail": user_account.get("mail"),
                "issue": "UPN (sign-in) does not match primary mail address"
            })

        # Check 3: Check for Bounce Risk (missing old domain fallback alias)
        # If a user is on the new domain, do they still have their old address as an alias?
        is_on_new_domain = False
        if target_domain:
            if user_sign_in_name.endswith(f"@{target_domain.lower()}") or user_primary_email.endswith(f"@{target_domain.lower()}"):
                is_on_new_domain = True

        if is_on_new_domain and old_domain:
            has_old_domain_alias = False
            for alias in email_aliases:
                if alias.endswith(f"@{old_domain.lower()}"):
                    has_old_domain_alias = True
                    break

            # If the old alias is missing, incoming emails to their old address will bounce!
            if not has_old_domain_alias:
                missing_old_alias_accounts.append({
                    "displayName": user_display_name,
                    "userPrincipalName": user_account.get("userPrincipalName"),
                    "primaryMail": user_account.get("mail"),
                    "issue": f"Missing @{old_domain} fallback in proxyAddresses. Senders emailing the old address will get bounce errors!"
                })

    # Build the clean, easy-to-read summary report
    audit_report = {
        "summary": {
            "total_users_scanned": total_users_count,
            f"users_on_new_domain_{target_domain}": target_domain_user_count if target_domain else "N/A",
            f"users_still_on_old_domain_{old_domain}": old_domain_user_count if old_domain else "N/A",
            "upn_and_mail_mismatches_count": len(mismatched_accounts),
            "missing_old_alias_bounce_risk_count": len(missing_old_alias_accounts)
        },
        "upn_vs_mail_mismatches": mismatched_accounts[:25],
        "at_risk_missing_old_alias": missing_old_alias_accounts[:25]
    }
    return json.dumps(audit_report, indent=2)


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
        access_token = get_access_token(scopes=ALL_SCOPES)
    except Exception:
        access_token = get_access_token(scopes=DEFAULT_SCOPES)

    http_headers = {"Authorization": f"Bearer {access_token}"}

    # Microsoft Graph endpoint equivalent to running:
    # Get-MgDeviceManagementManagedDevice -Filter "complianceState eq 'noncompliant'"
    intune_devices_url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$filter=complianceState eq 'noncompliant'&$select=id,deviceName,userPrincipalName,operatingSystem,osVersion,complianceState,lastSyncDateTime"
    graph_response = requests.get(intune_devices_url, headers=http_headers)

    if graph_response.status_code == 403:
        return (
            "Access Denied (403 Forbidden). To query Intune managed devices, "
            "add the 'DeviceManagementManagedDevices.Read.All' delegated permission to your Azure App Registration "
            "and grant admin consent in Azure Portal."
        )
    if not graph_response.ok:
        return f"Error from Microsoft Graph: {graph_response.status_code} - {graph_response.text}"

    noncompliant_devices = graph_response.json().get("value", [])
    if not noncompliant_devices:
        return "All clear! No noncompliant devices found in Intune. All managed machines meet policy requirements."

    device_investigation_results = []

    # Step through each noncompliant device and find out WHY it failed compliance
    for device in noncompliant_devices[:20]:
        device_id = device.get("id")
        device_name = device.get("deviceName")
        device_owner = device.get("userPrincipalName")
        device_os = f"{device.get('operatingSystem')} {device.get('osVersion')}"
        device_last_sync = device.get("lastSyncDateTime")

        # Query the device's compliance policy states to drill down into the specific setting
        policy_url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{device_id}/deviceCompliancePolicyStates"
        policy_response = requests.get(policy_url, headers=http_headers)
        failing_rules = []

        if policy_response.ok:
            compliance_policies = policy_response.json().get("value", [])
            for policy in compliance_policies:
                policy_state = policy.get("state")

                # If the policy state is failed or error, check which setting broke
                if policy_state in ("nonCompliant", "error", "conflict"):
                    policy_name = policy.get("displayName", "Unknown Policy")
                    setting_states = policy.get("settingStates", [])

                    for setting in setting_states:
                        setting_state = setting.get("state")
                        if setting_state in ("nonCompliant", "error", "conflict"):
                            failing_rules.append({
                                "policy": policy_name,
                                "setting": setting.get("settingName"),
                                "state": setting_state
                            })

        device_investigation_results.append({
            "deviceName": device_name,
            "userPrincipalName": device_owner,
            "operatingSystem": device_os,
            "lastSyncDateTime": device_last_sync,
            "failing_compliance_rules": failing_rules or ["Policy flagged noncompliant, setting details not available"]
        })

    return json.dumps(device_investigation_results, indent=2)


# ------------------------------------------------------------------------------
# TOOL 5: Intune Installed Software & Application Inventory Search
# ------------------------------------------------------------------------------
@mcp.tool()
def search_installed_apps(app_name: str = "") -> str:
    """
    Search software detected by Microsoft Intune across company-managed devices.
    Used to reconcile security/cyber audit requests (e.g. 'Who has app X installed, and what version?').
    
    PowerShell equivalent: Get-MgDeviceManagementDetectedApp

    Parameters:
      app_name (optional): Software name to search for (e.g. 'Chrome', 'Zoom', 'VLC', '7-Zip').
                           If left blank, returns the top discovered applications across the fleet.
    """
    try:
        access_token = get_access_token(scopes=ALL_SCOPES)
    except Exception:
        access_token = get_access_token(scopes=DEFAULT_SCOPES)

    http_headers = {"Authorization": f"Bearer {access_token}"}

    # Query Microsoft Graph for detected software inventory
    if app_name.strip():
        search_query = app_name.strip()
        graph_endpoint = f"https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$filter=contains(displayName, '{search_query}')&$top=20"
    else:
        graph_endpoint = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$top=25"

    graph_response = requests.get(graph_endpoint, headers=http_headers)

    if graph_response.status_code == 403:
        return (
            "Access Denied (403 Forbidden). To query Intune detected software, "
            "ensure the 'DeviceManagementManagedDevices.Read.All' delegated permission is consented in Azure Portal."
        )
    if not graph_response.ok:
        return f"Error from Microsoft Graph: {graph_response.status_code} - {graph_response.text}"

    detected_apps = graph_response.json().get("value", [])
    if not detected_apps:
        if app_name.strip():
            return f"No installed software matching '{app_name}' was found across Intune managed devices."
        return "No detected applications found in Intune."

    app_inventory_results = []

    for app in detected_apps[:15]:
        app_id = app.get("id")
        app_display_name = app.get("displayName")
        app_version = app.get("version", "N/A")
        installed_device_count = app.get("deviceCount", 0)
        app_platform = app.get("platform", "Unknown")

        app_entry = {
            "applicationName": app_display_name,
            "version": app_version,
            "platform": app_platform,
            "installedDeviceCount": installed_device_count,
            "devicesInstalledOn": []
        }

        # If searching for a specific software, drill into which exact machines have it installed
        if app_name.strip() and installed_device_count > 0:
            devices_endpoint = f"https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/{app_id}/managedDevices?$select=deviceName,userPrincipalName,operatingSystem"
            dev_response = requests.get(devices_endpoint, headers=http_headers)
            if dev_response.ok:
                for dev in dev_response.json().get("value", [])[:10]:
                    app_entry["devicesInstalledOn"].append({
                        "deviceName": dev.get("deviceName"),
                        "userPrincipalName": dev.get("userPrincipalName"),
                        "operatingSystem": dev.get("operatingSystem")
                    })

        app_inventory_results.append(app_entry)

    return json.dumps(app_inventory_results, indent=2)


# ==============================================================================
# SECTION 5: SERVER EXECUTION & ENTRYPOINT
# ==============================================================================
# In Python, `if __name__ == "__main__":` means:
# "If this file was directly run by a user in a terminal (not imported by another script), run this block."
if __name__ == "__main__":

    # `sys.argv` is the list of command-line arguments passed to the script.
    # Exactly like `$args` in PowerShell or `%1, %2` in batch files.
    # If the user typed `python exchange-device-mcp-demo.py --login`, run the login function.
    if len(sys.argv) > 1 and sys.argv[1] in ("--login", "-l", "login"):
        login()

    # `sys.stdin.isatty()` asks: "Is a human sitting at a real terminal keyboard?"
    # If a person ran `uv run exchange-device-mcp-demo.py` in their terminal without arguments,
    # we check if they are logged in. If not, we automatically trigger the login prompt for them!
    elif sys.stdin.isatty():
        try:
            get_access_token()
            print("Exchange MCP Server: Authenticated. Running stdio server (waiting for AI agent commands)...")
            mcp.run()
        except Exception:
            print("No active login found. Launching authentication...\n")
            login()

    # If `sys.stdin.isatty()` is False, it means an automated program (like VS Code or an AI agent)
    # launched this script in the background to send and receive JSON-RPC commands.
    else:
        mcp.run()
