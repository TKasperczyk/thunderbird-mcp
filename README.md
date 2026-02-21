# Thunderbird MCP

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Thunderbird](https://img.shields.io/badge/Thunderbird-102%2B-0a84ff.svg)](https://www.thunderbird.net/)
[![MCP](https://img.shields.io/badge/MCP-compatible-green.svg)](https://modelcontextprotocol.io/)

> Inspired by [bb1/thunderbird-mcp](https://github.com/bb1/thunderbird-mcp). Rewritten from scratch with a bundled HTTP server, proper MIME decoding, and UTF-8 handling throughout.

An MCP server that lets AI assistants read email, manage calendars, compose messages, and control message filters in Thunderbird.

## How it works

```
MCP Client <--stdio--> mcp-bridge.cjs <--HTTP--> Thunderbird Extension
```

The Thunderbird extension runs a local HTTP server on port 8765. The Node.js bridge translates between MCP's stdio protocol and HTTP so MCP clients can talk to it.

## Setup

**1. Install the extension**

```bash
./scripts/build.sh
./scripts/install.sh
```

Restart Thunderbird.

**2. Configure your MCP client**

Example for `~/.claude.json`:

```json
{
  "mcpServers": {
    "thunderbird-mail": {
      "command": "node",
      "args": ["/path/to/thunderbird-mcp/mcp-bridge.cjs"]
    }
  }
}
```

## What you can do

| Tool | What it does |
|------|--------------|
| `listAccounts` | List all email accounts and their identities |
| `listFolders` | List all mail folders with message counts (filterable by account or folder subtree) |
| `searchMessages` | Find emails by subject, sender, recipient, date range, or within a specific folder |
| `getMessage` | Read full email content with optional attachment saving |
| `getRecentMessages` | Get recent messages with date and unread filtering |
| `updateMessage` | Mark read/unread, flag/unflag, move to folder, or trash |
| `deleteMessages` | Delete messages (drafts are moved to Trash) |
| `createFolder` | Create a new subfolder under an existing folder |
| `sendMail` | Open a compose window with pre-filled content |
| `replyToMessage` | Reply with quoted original and proper threading |
| `forwardMessage` | Forward with original attachments preserved |
| `listFilters` | List all mail filter rules with conditions and actions |
| `createFilter` | Create a new mail filter with conditions and actions |
| `updateFilter` | Modify an existing filter's name, state, conditions, or actions |
| `deleteFilter` | Delete a mail filter by index |
| `reorderFilters` | Change filter execution priority order |
| `applyFilters` | Manually run filters on a folder to organize messages |
| `searchContacts` | Look up contacts |
| `listCalendars` | List your calendars |
| `createEvent` | Open a pre-filled calendar event dialog |

Compose tools open a window for you to review before sending. Nothing gets sent automatically.

Filter tools give full control over Thunderbird's message filters — create rules, modify conditions and actions, reorder priority, and run filters on demand. Changes persist immediately to `msgFilterRules.dat`.

## Security

The extension only listens on localhost, but any local process can access it while Thunderbird is running. Keep this in mind on shared machines.

## Troubleshooting

**Extension not loading?**
Check Tools → Add-ons and Themes. For errors: Tools → Developer Tools → Error Console.

**Connection refused?**
Make sure Thunderbird is running and the extension is enabled.

**Can't find recent emails?**
IMAP folders can be stale. Click on the folder in Thunderbird to sync, or right-click → Properties → Repair Folder.

## Development

```bash
# Build the extension
./scripts/build.sh

# Test the HTTP API directly
curl -X POST http://localhost:8765 \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list"}'

# Test the bridge
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | node mcp-bridge.cjs
```

After changing extension code, you'll need to remove it from Thunderbird, restart, reinstall, and restart again. Thunderbird caches aggressively.

## Known issues

- IMAP folder databases can be stale until you click on them
- Email bodies with control characters get sanitized to avoid breaking JSON
- HTML-only emails are converted to plain text (original formatting is lost)

## Project structure

```
thunderbird-mcp/
├── mcp-bridge.cjs              # stdio-to-HTTP bridge
├── extension/
│   ├── manifest.json
│   ├── background.js           # Extension entry point
│   ├── httpd.sys.mjs           # Mozilla's HTTP server lib
│   └── mcp_server/
│       ├── api.js              # The actual MCP implementation
│       └── schema.json
└── scripts/
    ├── build.sh
    └── install.sh
```

## License

MIT. The bundled `httpd.sys.mjs` is from Mozilla and licensed under MPL-2.0.
