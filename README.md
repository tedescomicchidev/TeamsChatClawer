# TeamsChatCrawler

A .NET 8 CLI tool that exports Microsoft Teams chat messages, images, and file
attachments from the last 5 days prior to a reference date, using the
**Microsoft Graph REST API**.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| .NET 8 SDK | https://dotnet.microsoft.com/download |
| Azure AD App Registration | See setup below |

---

## Azure AD App Registration

1. Go to **Azure Portal → Azure Active Directory → App registrations → New registration**.
2. Set *Supported account types* to match your tenant.
3. Under **Authentication**, add a **Mobile and desktop application** platform with
   redirect URI `https://login.microsoftonline.com/common/oauth2/nativeclient`.
4. Enable **Allow public client flows** (required for device-code flow).
5. Under **API permissions**, add **Microsoft Graph Delegated** permissions:
   - `Chat.Read`
   - `Chat.ReadBasic`
   - `Files.Read.All`
   - `User.Read`
6. Grant admin consent if required by your tenant.
7. Note the **Application (client) ID** and **Directory (tenant) ID**.

---

## Installation

```bash
git clone <repo-url>
cd TeamsChatClawer/src
dotnet build -c Release
```

Or publish a self-contained binary:

```bash
dotnet publish -c Release -r linux-x64 --self-contained true -o ../publish
```

---

## Usage

```
USAGE:
  TeamsChatCrawler [OPTIONS]

OPTIONS:
  -d, --date <date>           Reference date yyyy-MM-dd (default: today)
  -c, --chats <name> ...      Chat topic / member name filter(s) (default: all chats)
  -o, --output <dir>          Output directory (default: ./teams-export)
  -t, --tenant-id <id>        Azure AD tenant ID  [env: TEAMS_TENANT_ID]
  -i, --client-id <id>        Azure AD client ID  [env: TEAMS_CLIENT_ID]
  -?, -h, --help              Show help
```

### Examples

**Crawl all chats for the 5 days before 2026-02-27:**
```bash
dotnet run -- --date 2026-02-27 \
  --tenant-id <your-tenant-id> \
  --client-id <your-client-id>
```

**Crawl only chats matching "Project Alpha" or "John":**
```bash
dotnet run -- --chats "Project Alpha" "John" \
  --tenant-id <your-tenant-id> \
  --client-id <your-client-id>
```

**Use environment variables to avoid repeating credentials:**
```bash
export TEAMS_TENANT_ID=<your-tenant-id>
export TEAMS_CLIENT_ID=<your-client-id>

dotnet run -- --date 2026-02-20 --chats "Finance" --output ./my-export
```

---

## Authentication

The tool uses the **device code flow** (no browser pop-up, works in headless
terminals).  On first run you will see:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXXXXXX to authenticate.
```

The acquired token is cached in memory for the duration of the run.

---

## Output Structure

```
teams-export/
├── summary.txt                         ← full message transcript
└── <Chat_Name>/
    └── <yyyy-MM-dd>/
        └── attachments/
            ├── photo.png
            └── document.docx
```

`summary.txt` contains every message in chronological order with sender,
timestamp, body text, and local paths to any downloaded attachments.

---

## Architecture

```
src/
├── Program.cs                    CLI entry point (System.CommandLine)
├── Models/
│   ├── CrawlOptions.cs           Date window + filter parameters
│   └── ChatMessage.cs            Domain model for messages and attachments
└── Services/
    ├── AuthService.cs            DeviceCodeCredential → GraphServiceClient
    ├── ChatService.cs            Graph /me/chats with paging + filtering
    ├── MessageService.cs         Graph /me/chats/{id}/messages with $filter
    ├── DownloadService.cs        Hosted content + SharePoint file downloads
    └── CrawlerOrchestrator.cs   Coordinates the full crawl pipeline
```

### Key Graph API endpoints used

| Purpose | Endpoint |
|---|---|
| List chats | `GET /me/chats?$expand=members` |
| List messages (filtered) | `GET /me/chats/{id}/messages?$filter=createdDateTime ge ... and ...` |
| Download hosted content | `GET /me/chats/{chatId}/messages/{msgId}/hostedContents/{id}/$value` |
| Download SharePoint file | Direct URL from `attachment.contentUrl` |
