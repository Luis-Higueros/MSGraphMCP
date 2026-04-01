# MSGraphMCP — Microsoft 365 MCP Server

An MCP (Model Context Protocol) server exposing Microsoft 365 capabilities — Mail, Calendar, Teams, Files, OneNote, and Planner — via delegated OAuth2 authentication. Built with C# / ASP.NET Core, designed to run on Azure Container Instances and integrate with [Relevance AI](https://relevanceai.com).

---

## Features

- **One-time login** — device code flow on first use; MSAL refresh tokens persisted to Azure Blob Storage survive container restarts indefinitely
- **Proactive token refresh** — background timer silently renews access tokens 5 minutes before expiry
- **Session management** — stateful sessions keyed by `sessionId`; Relevance AI passes it as a header on every call
- **Intelligent tools** — free-slot finder with timezone awareness, multi-month email search, meeting time suggestions via Graph's native `findMeetingTimes` API, and more
- **Azure-ready** — Dockerfile + deployment scripts (Bash + PowerShell) for ACI; GitHub Actions CI/CD included

---

## Tool Reference

### Auth
| Tool | Description |
|---|---|
| `graph_initiate_login` | Start login. Returns `sessionId` immediately. Silent if user has logged in before. |
| `graph_check_login_status` | Poll until `status == "authenticated"`. |
| `graph_who_am_i` | Returns display name and email for the active session. |
| `graph_logout` | Ends session. Optionally revokes blob-cached token (forces re-login next time). |

### Mail
| Tool | Description |
|---|---|
| `mail_search` | Search by keywords, sender, date range, attachment flag. |
| `mail_summarize` | Search + return structured payload for LLM summarization. |
| `mail_get_thread` | Full conversation thread by `conversationId`. |
| `mail_get_by_id` | Full body of a single email. |
| `mail_send` | Send an email (plain text or HTML). |
| `mail_draft_reply` | Save a draft reply (reply or reply-all). |
| `mail_send_draft` | Send a previously saved draft. |

### Calendar
| Tool | Description |
|---|---|
| `calendar_find_free_slots` | Find gaps ≥ N minutes in the user's calendar, respecting their timezone and working hours. |
| `calendar_suggest_meeting_times` | Use Graph's `findMeetingTimes` API to find slots that work for all attendees. |
| `calendar_get_agenda` | Day-by-day agenda for a date range. |
| `calendar_create_event` | Create an event with optional attendees. Times interpreted in user's timezone. |

### Teams
| Tool | Description |
|---|---|
| `teams_list_my_teams` | All teams the user belongs to. |
| `teams_list_channels` | Channels in a team. |
| `teams_get_channel_messages` | Recent messages from a channel. |
| `teams_send_channel_message` | Post a message to a channel. |
| `teams_list_chats` | 1:1 and group chats. |
| `teams_get_chat_messages` | Recent messages from a chat. |

### Files (OneDrive / SharePoint)
| Tool | Description |
|---|---|
| `files_list_items` | List files/folders at a path. |
| `files_get_content` | Download text content of a file. |
| `files_upload_text` | Create or overwrite a text file. |
| `files_search` | Search files by name or content. |
| `files_create_share_link` | Generate a shareable link. |

### OneNote
| Tool | Description |
|---|---|
| `one_note_list_notebooks` | All notebooks. |
| `one_note_list_sections` | Sections in a notebook. |
| `one_note_list_pages` | Pages in a section. |
| `one_note_get_page_content` | Full plain-text content of a page. |
| `one_note_search_pages` | Full-text search across pages. |
| `one_note_create_page` | Create a new page in a section. |

### Planner
| Tool | Description |
|---|---|
| `planner_list_plans` | All plans accessible via the user's groups. |
| `planner_list_tasks` | Tasks in a plan, filterable by bucket, assignee, completion. |
| `planner_create_task` | Create a task with priority, due date, and notes. |
| `planner_update_task` | Update completion %, priority, due date, or title. |

---

## Quick Start

### Prerequisites

- [.NET 9 SDK](https://dotnet.microsoft.com/download)
- [Docker](https://www.docker.com/products/docker-desktop)
- [Azure CLI](https://learn.microsoft.com/en-us/cli/azure/install-azure-cli) (`az login`)
- An Azure subscription
- An Azure AD App Registration (see below)

### 1. Azure AD App Registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory → App Registrations → New Registration**
2. Name: `MSGraphMCP`
3. Supported account types: **Accounts in this organizational directory only**
4. Redirect URI: leave blank (device code flow doesn't need one)
5. Under **API Permissions → Add a permission → Microsoft Graph → Delegated**:

```
User.Read
Mail.Read
Mail.Send
MailboxSettings.Read
Calendars.ReadWrite
Channel.ReadBasic.All
ChannelMessage.Read.All
Team.ReadBasic.All
Files.ReadWrite.All
Notes.ReadWrite.All
Tasks.ReadWrite
offline_access
```

6. Click **Grant admin consent**
7. Note your **Tenant ID** and **Client ID** from the Overview page

### 2. Local Development

```bash
# Clone
git clone https://github.com/your-org/MSGraphMCP.git
cd MSGraphMCP

# Set configuration
cp src/MSGraphMCP/appsettings.json src/MSGraphMCP/appsettings.Local.json
# Edit appsettings.Local.json with your TenantId, ClientId
# For TokenCache:StorageConnectionString use "UseDevelopmentStorage=true" + Azurite

# Run
dotnet run --project src/MSGraphMCP

# Server available at http://localhost:8080
# Health: http://localhost:8080/health
# MCP:    http://localhost:8080/mcp
```

> **Local token cache**: Install [Azurite](https://learn.microsoft.com/en-us/azure/storage/common/storage-use-azurite) for local blob emulation, or point at a real Azure Storage account.

#### One-command local startup

Use the helper scripts to start Azurite (if needed) and run the app with local token cache wiring:

```bash
# macOS / Linux
./deploy/run-local.sh

# Windows PowerShell
.\deploy\run-local.ps1

# Windows PowerShell, force restart if already running
.\deploy\run-local.ps1 -Restart
```

### 3. Deploy to Azure

```bash
# Bash (Linux / macOS)
export AAD_TENANT_ID="your-tenant-id"
export AAD_CLIENT_ID="your-client-id"
chmod +x deploy/deploy-aci.sh
./deploy/deploy-aci.sh

# PowerShell (Windows)
$env:AAD_TENANT_ID = "your-tenant-id"
$env:AAD_CLIENT_ID = "your-client-id"
.\deploy\deploy-aci.ps1
```

The script will:
- Create a Resource Group, ACR, Storage Account, and ACI
- Build and push the Docker image
- Deploy the container with environment variables injected securely

---

## How One-Time Login Works

```
First call ever
  → graph_initiate_login(userHint: "alice@company.com")
  → Checks Azure Blob: cache MISS
  → Returns { status: "pending", verificationUrl, userCode }
  → User visits URL, enters code (one time, ~30 seconds)
  → MSAL saves refresh token to Azure Blob Storage

All future calls (even after container restarts)
  → graph_initiate_login(userHint: "alice@company.com")
  → Checks Azure Blob: cache HIT
  → Silent token acquisition (no device code)
  → Returns { status: "authenticated", sessionId }

Background (every ~55 minutes)
  → Proactive silent refresh keeps access token valid
  → Refresh tokens valid 90 days sliding — resets on each use
```

---

## Relevance AI Integration

### Workflow Pattern

```
Step 1:  graph_initiate_login { userHint: "{{user_email}}" }
           → if status == "authenticated": jump to Step 4
           → if status == "pending": continue to Step 2

Step 2:  [UI Block] Show user: verificationUrl + userCode

Step 3:  [Loop] graph_check_login_status { sessionId }
           → poll every 5 seconds until status == "authenticated"

Step 4:  Store sessionId as workflow variable

Step 5+: All tool calls include sessionId
```

### HTTP Configuration

```
Method:  POST
URL:     http://<your-aci-fqdn>:8080/mcp
Headers:
  Content-Type: application/json
  X-Session-Id: {{session_id}}
```

### Scope Smoke Test Endpoint

Use this endpoint to run a quick pass/fail sweep of the core Graph capabilities for an existing authenticated session.

```
POST /test/scope-smoke
Content-Type: application/json

{
  "sessionId": "<session-id>",
  "teamId": "<optional-team-id>",
  "channelId": "<optional-channel-id>",
  "fromUtc": "2026-04-01T00:00:00Z",
  "toUtc": "2026-04-03T00:00:00Z"
}
```

The response includes a summary and per-check status for:
- `GraphWhoAmI`
- `MailSearch`
- `CalendarGetAgenda`
- `FilesListItems`
- `TeamsListMyTeams`
- `TeamsListChannels`
- `TeamsGetChannelMessages`
- `OneNoteListNotebooks`
- `PlannerListPlans`

---

## GitHub Actions CI/CD

Set these secrets in your GitHub repository (**Settings → Secrets → Actions**):

| Secret | Value |
|---|---|
| `AZURE_CREDENTIALS` | Output of `az ad sp create-for-rbac --sdk-auth` |

The workflow will:
- Build and test on every push and PR
- Build + push Docker image to ACR on merges to `main`
- Restart the ACI and verify `/health` passes

---

## Configuration Reference

All settings can be overridden via environment variables (use `__` as separator):

| `appsettings.json` key | Environment variable | Description |
|---|---|---|
| `AzureAd:TenantId` | `AzureAd__TenantId` | Azure AD tenant ID |
| `AzureAd:ClientId` | `AzureAd__ClientId` | App registration client ID |
| `TokenCache:StorageConnectionString` | `TokenCache__StorageConnectionString` | Azure Storage connection string |
| `TokenCache:ContainerName` | `TokenCache__ContainerName` | Blob container name (default: `msal-token-cache`) |
| `Session:TimeoutHours` | `Session__TimeoutHours` | Idle session timeout in hours (default: `8`) |

---

## Project Structure

```
MSGraphMCP/
├── .github/workflows/deploy.yml    # CI/CD pipeline
├── deploy/
│   ├── deploy-aci.sh               # Bash deployment script
│   └── deploy-aci.ps1              # PowerShell deployment script
├── src/MSGraphMCP/
│   ├── Auth/
│   │   ├── BlobTokenCache.cs       # MSAL token persistence to Azure Blob
│   │   └── GraphAuthProvider.cs    # Device code flow + silent refresh
│   ├── Session/
│   │   ├── SessionContext.cs       # Per-session state (GraphClient, timers)
│   │   └── SessionStore.cs         # Thread-safe in-memory session registry
│   ├── Tools/
│   │   ├── AuthTools.cs            # Login, status, logout, whoami
│   │   ├── MailTools.cs            # Search, summarize, thread, send, draft
│   │   ├── CalendarTools.cs        # Free slots, meeting suggestions, agenda, create
│   │   ├── TeamsTools.cs           # Teams, channels, chats, messages
│   │   ├── FilesTools.cs           # OneDrive list, get, upload, search, share
│   │   ├── OneNoteTools.cs         # Notebooks, sections, pages, search, create
│   │   └── PlannerTools.cs         # Plans, tasks, create, update
│   ├── Program.cs                  # ASP.NET Core + MCP HTTP/SSE wiring
│   ├── appsettings.json            # Configuration template
│   └── MSGraphMCP.csproj           # Project file + NuGet dependencies
├── Dockerfile                      # Multi-stage build, non-root user
├── .dockerignore
├── .gitignore
└── MSGraphMCP.sln
```

---

## Security Notes

- **Never commit** `appsettings.Production.json` or any file containing secrets — it's in `.gitignore`
- The deployment scripts inject secrets as **secure environment variables** (not visible in ACI portal)
- The MSAL token cache blob is private (no public access) by default
- Consider adding Azure API Management in front of ACI for HTTPS termination and IP allowlisting before exposing to Relevance AI

---

## License

MIT
