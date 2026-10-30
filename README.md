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

## 🔒 Security & Governance

- **100% Read-Only**: There are zero write, modify, or delete capabilities. The agent cannot alter mailboxes, reset passwords, or wipe devices.
- **No `Mail.Read` Scope**: The tool physically lacks the permission to read user email messages. It only reads calendar ACLs and mailbox Out-of-Office settings.
- **Delegated RBAC Permissions**: Operates under the signed-in administrator's exact Entra ID permissions.
- **Enterprise OAuth 2.0 with PKCE**: Tokens are requested directly from `login.microsoftonline.com`, fully respecting your tenant's Conditional Access, Multi-Factor Authentication (MFA), and compliance policies.
- **Local Token Isolation**: Tokens are stored locally in `.token_cache.bin` on the admin's machine (never committed to Git).
- **Auditable**: Every API call is logged in your Entra ID Sign-In & Audit logs under your registered App ID.

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
  *(Note: `Mail.Read` is NOT required).*

---

## ⚙️ Setup & Installation

### Step 1: Install `uv`

#### Windows (PowerShell):
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

#### macOS / Linux:
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

---

### Step 2: Clone & Configure Environment

#### Windows (PowerShell):
```powershell
git clone git@github.com:amargh2/exchange-device-mcp-demo.git
cd exchange-device-mcp-demo

# Copy the example environment file
Copy-Item .env.example .env
```

#### macOS / Linux (Bash):
```bash
git clone git@github.com:amargh2/exchange-device-mcp-demo.git
cd exchange-device-mcp-demo
cp .env.example .env
```

Edit `.env` with your Azure App Registration details:
```ini
APPID=your_azure_client_app_id
TENANTID=your_microsoft_365_tenant_id
```

---

### Step 3: One-Time Sign-In (Interactive PKCE)

Run the login command in PowerShell or Bash:

```powershell
uv run exchange-device-mcp-demo.py --login
```

Your system browser will automatically launch with the Microsoft sign-in prompt (the link is also printed in your terminal as a backup). Sign in with your administrator account and accept the requested permissions. Tokens are saved to your local `.token_cache.bin`.

---

## 🤖 Running with AI Clients

### In VS Code / Antigravity / Cursor

Add the server to your project's `.vscode/mcp.json`:

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
      ]
    }
  }
}
```

> [!NOTE]
> **What is `${workspaceFolder}`?**
> In VS Code, `${workspaceFolder}` is a built-in variable that dynamically resolves to the root folder of the project currently open in your editor (similar to `$PSScriptRoot` in PowerShell). This ensures the configuration works seamlessly on any machine without hardcoding user directory paths (e.g., `C:\Users\John\...`).

---

## 💬 Example Prompts for the AI Agent

Once connected, you can ask your AI agent questions like:

- *"Run a domain migration audit for target domain `newcompany.com` and old domain `oldcompany.com`."*
- *"Check Intune for noncompliant devices and tell me why they are failing."*
- *"Who has permission to view or edit `ceo@company.com`'s calendar?"*
- *"Does `john@company.com` have an Out of Office message turned on?"*
- *"What personal devices are registered to my account in Entra ID?"*
