# TeamsChatCrawler

A .NET CLI utility that exports Microsoft Teams chat messages and downloads message attachments (pictures/documents) for the 5-day window prior to a given date.

## What this tool does

- Accepts an optional anchor date (`--date yyyy-MM-dd`).
  - If not provided, the tool defaults to today (UTC).
- Accepts optional filters:
  - `--chat` for chat titles/topics.
  - `--group` for named chat groupings from a JSON mapping file (`--group-config`).
- Calls Microsoft Graph Teams chat REST APIs to:
  - list your chats,
  - fetch messages per chat,
  - download eligible attachments.

## Microsoft Teams/Graph REST API endpoints used

The implementation is aligned to Graph v1.0 chat endpoints:

- `GET /me/chats`
- `GET /chats/{chat-id}/messages`
- File download from `chatMessageAttachment.contentUrl`

Reference docs:

- https://learn.microsoft.com/graph/api/chat-list
- https://learn.microsoft.com/graph/api/chat-list-messages
- https://learn.microsoft.com/graph/api/resources/chatmessageattachment

## Prerequisites

- .NET SDK 8.0+
- A delegated Microsoft Graph access token in `TEAMS_ACCESS_TOKEN`
  - Minimum practical scopes: `Chat.Read`, `Files.Read`.

## Usage

```bash
dotnet run -- \
  --date 2026-02-20 \
  --chat "Project Falcon,Leadership" \
  --group "Finance,Delivery" \
  --group-config ./chat-groups.json \
  --output ./exports
```

### Defaults

```bash
dotnet run --
```

- Date defaults to today (UTC).
- No chat/group filters means **all chats** are processed.
- Output defaults to `./output`.

## Group configuration file format

`chat-groups.json` example:

```json
{
  "groups": [
    {
      "name": "Finance",
      "chatIds": ["19:123abc@thread.v2"],
      "chatTitles": ["finance", "budget"]
    },
    {
      "name": "Delivery",
      "chatIds": [],
      "chatTitles": ["delivery standup", "release"]
    }
  ]
}
```

## Output structure

```text
output/
  <chat-topic-or-id>/
    messages.json
    attachments/
      <messageId>_<fileName>
```

## Notes

- The date window is inclusive and spans from `anchorDate - 5 days` to `anchorDate` in UTC.
- Message body HTML/text is preserved in `messages.json`.
- Attachment downloads are attempted for attachment types containing `image`, `application/*`, or `reference` content types.
