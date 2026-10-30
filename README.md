# Exchange & Intune Management MCP Server (`exchange-device-mcp-demos`)

A Model Context Protocol (MCP) server that connects AI assistants and agents (Antigravity, VS Code, Claude Desktop, Cursor) directly to Microsoft 365. It allows administrators to query Microsoft Entra ID, Exchange Online, and Intune using natural language instead of hopping through admin portal blades or writing complex PowerShell loops.

---

## 🚀 Available Tools

| Tool | PowerShell Equivalent | What It Does |
| :--- | :--- | :--- |
| **`audit_domain_migration`** | `Get-MgUser` + UPN/Proxy audit | Audits user sign-in names (UPN), primary emails, and `proxyAddresses` during an email cutover. Catches UPN mismatches and flags users missing fallback aliases to prevent incoming mail bounces (NDR 550). |
| **`get_noncompliant_devices`** | `Get-MgDeviceManagementManagedDevice` | Queries Microsoft Intune for noncompliant hardware and extracts the exact failing policy setting (e.g., BitLocker missing, outdated OS, PIN complexity). |
| **`get_calendar_permissions`** | `Get-MailboxFolderPermission -Identity user:\Calendar` | Checks calendar sharing and delegate Access Control Lists (ACLs) to see who has Reviewer, Editor, or Delegate rights. |
| **`get_mailbox_settings`** | `Get-MailboxAutoReplyConfiguration` | Inspects Out-of-Office (automatic reply) status, internal/external reply messages, and regional mailbox time zones. |
| **`get_user_devices`** | `Get-MgUserRegisteredDevice` | Lists registered mobile phones and personal hardware associated with a user in Microsoft Entra ID. |

---

## 🔒 Security Architecture

- **100% Delegated Permissions**: The server operates strictly under the logged-in administrator's RBAC rights. It can never see or do anything the administrator is not authorized to do.
- **Enterprise OAuth 2.0 with PKCE**: Tokens are requested directly from `login.microsoftonline.com`, fully respecting your tenant's Conditional Access, Multi-Factor Authentication (MFA), and compliance policies.
- **Persistent Local Token Cache**: Sign-in tokens are saved locally in `.token_cache.bin` (which is excluded from Git). Once authenticated, the server refreshes tokens silently in the background with zero popups.
- **Auditable**: Every single query is logged in your Entra ID Sign-In & Audit logs under your registered App ID.

---

## 📋 Prerequisites

- [Python 3.12+](https://www.python.org/)
- [`uv`](https://docs.astral.sh/uv/) (Fast Python package manager)
- An Azure / Microsoft Entra ID App Registration with the following delegated permissions:
  - `User.Read`
  - `Device.Read`
  - `MailboxSettings.Read`
  - `Calendars.Read`
  - `User.Read.All`
  - `DeviceManagementManagedDevices.Read.All`
  *(Note: `Mail.Read` is NOT required. This tool physically cannot read email messages).*

---

## ⚙️ Setup & Installation

### 1. Clone & Configure Environment

```bash
git clone git@github.com:amargh2/exchange-device-mcp-demo.git
cd exchange-device-mcp-demo

# Copy the example environment file and fill in your Azure App details
cp .env.example .env
```

Edit `.env`:
```ini
APPID=your_azure_client_app_id
TENANTID=your_microsoft_365_tenant_id
```

### 2. One-Time Sign-In (PKCE)

Run the login command in your terminal to complete interactive Microsoft sign-in:

```bash
uv run exchange-device-mcp-demo.py --login
```

A browser window will launch (and the sign-in URL is printed in the terminal). Sign in with your administrator account and accept the requested permissions. Tokens are saved to `.token_cache.bin`.

---

## 🤖 Running with AI Clients

### In VS Code / Antigravity / Cursor

Add the server to your `.vscode/mcp.json`:

```json
{
  "servers": {
    "exchange-device-mcp-demos": {
      "type": "stdio",
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "${workspaceFolder}",
        "exchange-device-mcp-demo.py"
      ],
      "env": {
        "BROWSER": "firefox",
        "DISPLAY": ":0"
      }
    }
  }
}
```

### Example Prompts for the AI Agent:

- *"Run a domain migration audit for target domain `newcompany.com` and old domain `oldcompany.com`."*
- *"Check Intune for noncompliant devices and tell me why they are failing."*
- *"Who has permission to view or edit `ceo@company.com`'s calendar?"*
- *"Does `john@company.com` have an Out of Office message turned on?"*
- *"What personal devices are registered to my account in Entra ID?"*

