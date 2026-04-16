# MSGraphMCP

A **Model Context Protocol (MCP) server** that exposes Microsoft Graph API
capabilities so that AI assistants (Claude, Copilot, etc.) can interact with
**Microsoft Teams**, **Outlook Mail**, **Calendar**, **OneDrive/SharePoint**,
and the **user directory** on behalf of your organisation.

---

## Features

| Area | Tools |
|------|-------|
| **Teams** | `list_teams`, `get_team`, `list_channels`, `get_channel_messages`, `send_channel_message`, `create_channel`, `list_team_members`, `list_chats`, `get_chat_messages`, `send_chat_message` |
| **Calendar** | `list_calendar_events`, `get_calendar_event`, `create_calendar_event`, `update_calendar_event` |
| **Mail** | `list_emails`, `get_email`, `send_email`, `reply_to_email`, `move_email` |
| **Users** | `get_current_user`, `get_user`, `search_users`, `get_user_presence`, `get_my_presence`, `list_my_manager_and_reports` |
| **Files** | `list_drive_items`, `get_drive_item`, `search_drive_items`, `list_sharepoint_sites`, `list_sharepoint_site_files` |

---

## Prerequisites

* Python 3.10+
* An **Azure AD app registration** with the permissions listed below

### Required Azure AD Permissions

**Application permissions** (for `client_credentials` / app-only flow):

| Permission | Purpose |
|------------|---------|
| `Team.ReadBasic.All` | List teams |
| `Channel.ReadBasic.All` | List channels |
| `ChannelMessage.Read.All` | Read channel messages |
| `ChannelMessage.Send` | Send channel messages |
| `Chat.Read.All` | Read chats |
| `Chat.ReadWrite.All` | Send chat messages |
| `Calendars.ReadWrite` | Read/write calendar events |
| `Mail.ReadWrite` | Read/write emails |
| `Mail.Send` | Send emails |
| `Files.Read.All` | Read OneDrive/SharePoint files |
| `Presence.Read.All` | Read user presence |
| `User.Read.All` | Read user profiles |
| `Sites.Read.All` | List SharePoint sites |

> **Tip:** For personal/delegated access (your own data only), use the
> `device_code` flow and grant the corresponding **Delegated** permissions.

---

## Installation

```bash
pip install msgraph-mcp
```

Or install from source:

```bash
git clone https://github.com/Luis-Higueros/MSGraphMCP.git
cd MSGraphMCP
pip install -e .
```

---

## Configuration

Copy `.env.example` to `.env` and fill in your Azure AD credentials:

```env
TENANT_ID=your-azure-ad-tenant-id
CLIENT_ID=your-app-registration-client-id
CLIENT_SECRET=your-client-secret   # only for client_credentials flow
AUTH_FLOW=client_credentials       # or 'device_code'
```

| Variable | Required | Description |
|----------|----------|-------------|
| `TENANT_ID` | ✅ | Your Azure AD tenant ID |
| `CLIENT_ID` | ✅ | App registration client ID |
| `CLIENT_SECRET` | ✅ (app-only) | Client secret (not needed for `device_code`) |
| `AUTH_FLOW` | optional | `client_credentials` (default) or `device_code` |

---

## Usage

### stdio transport (default – for Claude Desktop)

```bash
msgraph-mcp
# or
python -m msgraph_mcp.server
```

### SSE transport (for web-based MCP clients)

```bash
msgraph-mcp --sse --port 8000
# or
python -m msgraph_mcp.server --sse --port 8000
```

### Claude Desktop configuration

Add the following to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "msgraph": {
      "command": "msgraph-mcp",
      "env": {
        "TENANT_ID": "your-tenant-id",
        "CLIENT_ID": "your-client-id",
        "CLIENT_SECRET": "your-client-secret",
        "AUTH_FLOW": "client_credentials"
      }
    }
  }
}
```

---

## Example interactions

Once connected to an AI assistant you can ask questions like:

* *"List all my Teams and their channels."*
* *"Send a message 'Sprint planning at 2 PM' to the #general channel in the Engineering team."*
* *"What meetings do I have tomorrow?"*
* *"Create a calendar event called 'Design Review' for next Monday 3–4 PM and invite alice@example.com."*
* *"Show me my unread emails from this week."*
* *"What files are in my Documents folder on OneDrive?"*
* *"Is John Smith currently available?"*

---

## Development

```bash
# Install with dev dependencies
pip install -e ".[dev]"

# Run tests
pytest tests/ -v
```

---

## Project structure

```
src/msgraph_mcp/
├── __init__.py
├── auth.py           # MSAL authentication (client_credentials / device_code)
├── graph_client.py   # Async HTTP client for Microsoft Graph REST API
├── server.py         # FastMCP server – registers all tools and starts transport
└── tools/
    ├── calendar.py   # Calendar tools
    ├── files.py      # OneDrive / SharePoint tools
    ├── mail.py       # Outlook Mail tools
    ├── teams.py      # Microsoft Teams tools
    └── users.py      # User directory & presence tools
tests/
├── conftest.py
├── test_auth.py
├── test_graph_client.py
├── test_tools_calendar.py
├── test_tools_files.py
├── test_tools_mail.py
├── test_tools_teams.py
└── test_tools_users.py
```

---

## License

MIT
