# Thunderbird MCP

[![CI](https://github.com/TKasperczyk/thunderbird-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/TKasperczyk/thunderbird-mcp/actions/workflows/ci.yml)
[![Tools](https://img.shields.io/badge/40_Tools-email%2C_compose%2C_filters%2C_calendar%2C_contacts-blue.svg)](#what-you-can-do)
[![Localhost Only](https://img.shields.io/badge/Privacy-localhost_only-green.svg)](#security)
[![Thunderbird](https://img.shields.io/badge/Thunderbird-102%2B-0a84ff.svg)](https://www.thunderbird.net/)
[![License: MIT](https://img.shields.io/badge/License-MIT-grey.svg)](LICENSE)

Give your AI assistant full access to Thunderbird -- search mail, compose messages, manage filters, and organize your inbox. All through the [Model Context Protocol](https://modelcontextprotocol.io/).

<p align="center">
  <img src="docs/demo.gif" alt="Thunderbird MCP Demo" width="600">
</p>

> Inspired by [bb1/thunderbird-mcp](https://github.com/bb1/thunderbird-mcp). Rewritten from scratch with a bundled HTTP server, proper MIME decoding, and UTF-8 handling throughout.

---

## Why?

Thunderbird has no official API for AI tools. Your AI assistant can't read your email, can't help you draft replies, can't organize your inbox. This extension fixes that -- it exposes 40 tools over MCP so any compatible AI (Claude, GPT, local models) can work with your mail the way you'd expect.

Compose sends and event/task creation require review by default because **Block `skipReview`** starts enabled. `skipReview: true` is honored only after you explicitly disable that safety setting. While enabled, meetings with attendees other than the calendar user are read-only through MCP, and adding attendees is blocked: calendar servers can email invitations, updates, or cancellations without review. Automatic Forward/Reply filter actions have a separate opt-in, disabled by default.

---

## How it works

```
                    stdio              HTTP (localhost:8765-8774)
  MCP Client  <----------->  Bridge  <--------------------->  Thunderbird
  (Claude, etc.)           mcp-bridge.cjs                    Extension + HTTP Server
```

The Thunderbird extension embeds a local HTTP server with session-scoped auth tokens. The Node.js bridge translates between MCP's stdio protocol and HTTP, discovering the port and token automatically via a connection file. The bridge handles MCP lifecycle methods (initialize, ping) locally, so clients can connect even before Thunderbird is fully loaded.

---

## What you can do

### Mail

| Tool | Description |
|------|-------------|
| `listAccounts` | List all email accounts and their identities |
| `listFolders` | Browse folder tree with message counts and `isFavorite` in object or table format. Filter by account, subtree, or `favoritesOnly: true`; favorites nested under ordinary folders are included. |
| `searchMessages` | Search by subject, sender, recipient, body preview, date range, or tags. Multi-word queries are AND-of-tokens (every word must appear somewhere). Prefix with `from:`, `subject:`, `to:`, or `cc:` to restrict to one field. Set `searchBody: true` for full-text body search via Thunderbird's Gloda index. Supports `includeSubfolders`, `countOnly`, and offset-based pagination. Results include `threadId` and `preview` snippet. By default, `dedupByMessageId` collapses the same RFC Message-ID found in multiple folders/labels into one row and reports the other folder paths in `dupLocations`; set `dedupByMessageId: false` to return every location. |
| `getMessage` | Read full email content -- `bodyFormat`: `markdown` (default), `text`, or `html`. Set `rawSource: true` for the complete RFC 2822 source (all headers + MIME parts). Optional attachment saving. Set `includeInlineImages: true` to append supported inline CID images as MCP image blocks (PNG, JPEG, GIF, or WebP; max 1 MiB base64 per image and 4 MiB total). Skipped images are reported in attachment metadata. |
| `getMessages` | Read full email content for up to the configured batch limit in one call (default 10, max 20). Uses the same `bodyFormat`, `rawSource`, and attachment options as `getMessage`; each item supplies `messageId` and `folderPath`. |
| `getRecentMessages` | Get recent messages with date, unread, and tag filtering. Supports pagination. Results include `threadId` and `preview`. |
| `displayMessage` | Open a message in Thunderbird's GUI -- `3pane` (default), `tab`, or `window` mode |
| `updateMessage` | Mark read/unread, flag/unflag, add/remove tags, move or copy between folders, or trash -- supports bulk via `messageIds`. `copyTo` preserves the source and can add a Gmail label. |
| `deleteMessages` | Delete messages -- drafts are safely moved to Trash |
| `createFolder` | Create new subfolders to organize your mail |
| `renameFolder` | Rename an existing mail folder |
| `deleteFolder` | Delete a folder (moves to Trash, or permanently deletes if already in Trash) |
| `moveFolder` | Move a folder to a new parent within the same account |
| `emptyTrash` | Permanently delete all messages in Trash (including subfolders) |
| `emptyJunk` | Permanently delete all messages in Junk/Spam (including subfolders) |

`searchMessages` scans cooperatively with a best-effort 20-second budget. Header searches count and sort collected matches before pagination, including those beyond 10,000; only the returned page loads full result fields. Counts and totals are best-effort when folders change during a long search, even if no truncation is reported. Unscoped searches still request an IMAP refresh for each visited folder, but results use the locally available cache. Opening a native database and taking its key snapshot are synchronous and cannot be interrupted by this budget.

Message results remain a plain array when `offset` is omitted or null, including incomplete searches and `searchBody: true`. To receive completeness information, provide `offset: 0` (or another offset) for a paginated object. Only object responses can include `truncated: true` and a `message` explaining why; `countOnly` retains its existing object response and can also include these fields. Partial counts and `totalMatches` cover only observed matches, and date ordering covers only that subset. `hasMore` describes additional pages within the collected results and becomes false at their end, even if `truncated` is true; an empty returned page also sets it false. To improve coverage, narrow the folder, date range, or query (or set `includeSubfolders: false`).

For `searchBody: true`, object responses are conservatively marked `truncated: true`: Gloda returns candidates ranked by relevance and exposes no reliable completeness indicator. It applies its limit before removing stale index rows, so even a short result can omit other matches. Plain-array responses carry no completeness information. Narrowing the text query improves coverage; use header search when header/preview matching is sufficient.

`updateMessage` takes tag **keys**, not display labels: for example, `addTags: ["my=20project"]` uses a key generated for a label with a space. Keys must be non-empty printable ASCII without spaces, parentheses, brackets, braces, `%`, `*`, double quotes, backslashes, `<`, `>`, or `;`. This follows [Thunderbird's keyword handling](https://searchfox.org/comm-central/source/mailnews/imap/public/nsIImapService.idl), including its extra restrictions beyond [RFC 3501 atoms](https://www.rfc-editor.org/rfc/rfc3501#section-9). `=` and legacy modified UTF-7 keys containing `&` are accepted. Invalid entries in either `addTags` or `removeTags` fail the entire call with an error naming those entries before any messages are changed.

Use `copyTo: "<destination folder URI>"` to copy without removing the original. On Gmail, copying from Sent Mail to a project folder adds that label while preserving Sent Mail and existing labels. `copyTo`, `moveTo`, and `trash: true` are mutually exclusive. Both source and destination must be accessible. Copies and moves within or across accounts use the same Thunderbird native copy service call. On IMAP, completion is asynchronous and later failures are not reported back; success means only that the operation was submitted, so verify the destination. Tags applied together with a copy or move may not transfer on IMAP; verify the destination tags as well.

Message body formats (`getMessage` and `getMessages`):

- `markdown` (default) converts HTML to visible structure, allowing only `http:`, `https:`, and `mailto:` link destinations. Other HTML links become their visible text. HTML images become alt text (or nothing); their source URLs are omitted. Text nodes and alt text escape backslashes, backticks, brackets, angle brackets, and `!` before `[` to prevent literal links, images, or raw HTML. A literal `!` immediately before a generated link is also escaped, including across element boundaries. Other punctuation (including underscores, asterisks, hashes, pipes, and tildes) is preserved; HTML code spans and blocks use code formatting.
- For plain-text MIME bodies, `markdown` escapes Markdown image openers (`![`) while preserving the rest of the text, apart from the existing invisible-character removal. This also applies to coerced text and raw-MIME recovery. Other Markdown, links, and literal HTML in plain-text bodies are **not sanitized**; clients must disable raw HTML and remote-content loading when rendering untrusted message text.
- `text` extracts visible text. Both text formats remove hidden HTML content and invisible control characters.
- For `text` and `markdown`, HTML input over 2 MiB (UTF-8) is truncated at a Unicode character boundary before being parsed with the same HTML privacy rules. The output ends with `[Message body truncated at 2 MiB]`. Content after the cut is omitted; retained content keeps the normal formatting and link rules. Oversized compose fragments use the same capped parsing path. Encryption checks inspect the full source: oversized HTML containing `-----BEGIN PGP MESSAGE-----` anywhere is withheld unless encrypted-message access is enabled, even when the marker is only quoted in prose.
- `html` returns the original HTML unchanged. It is **untrusted** and may contain hidden content, scripts, unsafe links, and remote images; clients must sanitize it before rendering. `rawSource: true` also remains unchanged.

`includeInlineImages: true` is a separate opt-in for supported inline CID images as MCP image blocks; it does not add image URLs to the Markdown body.

### Compose

| Tool | Description |
|------|-------------|
| `sendMail` | Compose a new email -- opens a review window; direct sending requires explicitly disabling the `skipReview` safety block |
| `saveDraft` | Save a new or replacement draft without sending or opening a window; supports threading headers and reports an accessible Drafts folder |
| `replyToMessage` | Reply with quoted original and proper threading -- `skipReview` is subject to the same safety block; `saveAsDraft` saves the threaded reply to Drafts without sending |
| `forwardMessage` | Forward with all original attachments preserved -- `skipReview` is subject to the same safety block |

`sendMail`, `replyToMessage`, and `forwardMessage` open a window for you to review and edit before sending by default. The **Block `skipReview`** preference is on by default, so `skipReview: true` is rejected until you explicitly disable the preference; only then can it send directly. Attachments can be file paths or inline base64 objects.

`replyToMessage` accepts `saveAsDraft: true` to build a native reply with quoted text, the identity's signature and threading headers, save it, and close the compose window without sending. This requires the `saveDraft` tool to be enabled and cannot be combined with `skipReview`. Encrypted originals require the **Allow MCP clients to read encrypted messages** opt-in. Before saving, the current compose identity's configured destination must be accessible under account restrictions and carry the Drafts flag. A timeout returns `saveOutcome: "uncertain"`: the outstanding save may still complete, so check Drafts before retrying. Failures and timeouts restore the window's prior close and save-dialog behavior.

`saveDraft` supports these optional parameters:

- `inReplyTo`: one bracketed Message-ID such as `<original@example.com>`. `references`: up to 100 such IDs, oldest first, separated by single ASCII spaces. Each ID is limited to 998 characters and the full References value to 16,384. Missing brackets, extra tokens, whitespace inside IDs, and control characters (including CR/LF) are rejected, never repaired. References defaults to `inReplyTo` when omitted, and may also be supplied independently. The caller supplies the subject and quoted text.
- `replaceMessageId` and `replaceFolderPath`: replace an existing draft with the supplied content. The folder must be accessible under account restrictions, carry Thunderbird's Drafts flag, and match the selected `from` identity's configured drafts folder. Other folders and missing messages are rejected before saving. Supply all fields and attachments you want retained; content is not merged from the old draft.
- `includeSignature`: append the identity's signature, defaulting to `true` for a new draft and `false` for a replacement so re-saving a fetched draft does not duplicate its signature. Set it explicitly when needed. `sendMail` with `skipReview: true` accepts the same option, defaulting to `true`; review windows use Thunderbird's own signature preferences. Signature files are limited to 1 MiB; unreadable or oversized signature files are omitted.

On success, `saveDraft` returns `folderPath` when Thunderbird exposes the destination and that folder is accessible under account restrictions. Otherwise the save still succeeds without disclosing the folder URI. It does not return the saved message's ID; look up the draft in the returned folder when available.

Compose tools validate the `from` identity strictly -- if the specified sender doesn't match any configured Thunderbird identity, the tool returns an error instead of silently substituting another account.

### Filters

| Tool | Description |
|------|-------------|
| `listFilters` | List all filter rules with human-readable conditions and actions |
| `createFilter` | Create filters with structured conditions (from, subject, date...) and actions (move, tag, flag...) |
| `updateFilter` | Modify a filter's name, enabled state, conditions, or actions |
| `deleteFilter` | Remove a filter by index |
| `reorderFilters` | Change filter execution priority |
| `applyFilters` | Start eligible enabled Manual filters on a folder; report submitted rules and skipped rules with reasons |

Your AI can create sorting rules, adjust priorities, and run them on existing mail. Changes persist after validation; all updates, including name or enabled-state edits, validate a complete candidate before replacing the existing rule. Unparseable rules cannot be updated, but can still be deleted. Filter names, conditions, and action text reject control characters (U+0000–U+001F and U+007F) and backslashes. Custom add-on actions cannot be created, preserved by an update, or submitted for execution.

**Allow automatic Forward/Reply filter actions** is off by default in the extension settings. While off, `createFilter` and `updateFilter` reject any resulting rule containing Forward or Reply, including disabled rules or edits that retain an existing sending action. An update that only sets `enabled: false` can still disable a sending rule, and deletion remains available. Enabling this setting permits automatic sends without a compose review window, independently of **Block `skipReview`**. Turning it off does not disable saved filters or stop Thunderbird's own automatic filtering.

The sending-action setting governs rules MCP creates, modifies, or runs manually. Reordering or deleting a rule containing `StopExecution` can change which of your existing rules Thunderbird runs automatically, including Forward/Reply rules.

Move/Copy destinations must remain accessible under the current account restrictions, including actions retained by an update. An update that only sets `enabled: false`, or deletion, remains available for rules with inaccessible destinations.

`applyFilters` submits only enabled, parseable rules with the Manual type flag and skips rules with inaccessible Move/Copy destinations, as well as Forward/Reply rules while the setting is off. It returns `submittedFilters`, submitted rule names, and skipped rule names with reasons. An eligible rule containing a Custom action rejects the call before any rule is submitted. A successful submission means processing has started, not completed; if no rules are eligible, nothing is submitted.

### Contacts

| Tool | Description |
|------|-------------|
| `searchContacts` | Search contacts across all address books by email or name and return full contact details. Supports `maxResults`. |
| `getContact` | Read full contact details by UID |
| `createContact` | Create a contact with optional email/name, phones, postal addresses, organization, title, note, and birthday. Phone-only contacts are supported. |
| `updateContact` | Update contact fields; omitted fields stay unchanged, while empty phone/address arrays clear those collections |
| `deleteContact` | Delete a contact by UID |

### Calendar

| Tool | Description |
|------|-------------|
| `listCalendars` | List all calendars with read-only, event, and task support flags |
| `createEvent` | Create an event through a review dialog, optionally with RRULE recurrence. Direct creation and non-empty `attendees` require disabling **Block `skipReview`**. Accepts `status: tentative \| confirmed \| cancelled`. |
| `listEvents` | Query events by date range with bounded recurrence expansion. Returns a plain array capped at `maxResults`, including status, recurrence, organizer, attendees, and your participation status. Series that cannot be expanded return a master marked `recurrenceNotExpanded: true`. |
| `updateEvent` | Modify an event or series; `recurrenceId` selects one occurrence. Meetings with other attendees are read-only while **Block `skipReview`** is on, even when `attendees` is omitted. |
| `deleteEvent` | Delete an event or series; `recurrenceId` excludes one occurrence. Meetings with other attendees cannot be deleted while **Block `skipReview`** is on. |
| `createTask` | Open a pre-filled task dialog for review; direct creation via `skipReview` requires explicitly disabling the default safety block |
| `listTasks` | List tasks/to-dos from calendars -- filter by completion status, due date, or calendar |
| `updateTask` | Update a task's title, due date, description, priority, completion status, or percent complete |

`recurrence` accepts a single RRULE, such as `FREQ=WEEKLY;BYDAY=MO,TU` or `RRULE:FREQ=DAILY;COUNT=10`. The optional prefix is case-insensitive. Control characters (including CR/LF), malformed rules, `SECONDLY`/`MINUTELY`, and `HOURLY` on all-day events are rejected; Thunderbird's recurrence parser validates the rule before saving. On update, `recurrence: ""` or `null` clears the rule. Replacing a rule discards existing EXDATEs and modified occurrences.

Use the `recurrenceId` returned by `listEvents` to update or delete one occurrence. Omitted or null IDs select the series; a non-null ID cannot be combined with `recurrence` on update. All-day occurrence IDs preserve their calendar date across timezones. Existing Date-compatible inputs remain accepted for event start/end dates; ISO 8601 is recommended.

Attendees use `{ "email": "alice@example.com", "name": "Alice", "role": "optional" }`; new attendees default to `required`, and a `mailto:` email prefix is accepted. Email addresses must be single mailboxes without control characters or URI headers; names cannot contain control characters either. On update, omitted/null attendees preserve the list, while `[]` removes everyone. Retained attendees match by case-insensitive email and keep their participation status and provider metadata. Only supplied name/role fields change; omitted/null fields preserve existing values, and an empty name clears it. New attendees start with `NEEDS-ACTION`.

Attendee management requires the calendar identity to match an existing organizer. Existing organizers are preserved; a missing organizer is initialized from the calendar identity, which may also set a missing calendar `organizerId` property. Creation with non-empty attendees requires disabling **Block `skipReview`**, even when `skipReview` is false: Exchange/Owl or CalDAV may email invitations containing the event title and description without review.

`listEvents` returns a plain array, capped at `maxResults` (default 100, maximum 500). RRULE generation is limited to `maxResults + 1` candidates per series (hard ceiling 501) and 5,000 per request before expansion can allocate an unbounded occurrence array. Generation limits can leave fewer than `maxResults` results even when more occurrences exist, especially when exclusions remove generated candidates. Use narrower date ranges to reduce this effect.

If a series cannot be expanded safely with a bounded count, expansion fails, or the shared generation budget is exhausted before that series, its master is included once with `recurrenceNotExpanded: true`, subject to the same output cap. This includes EXRULE series: incomplete exclusion expansion could expose excluded dates, so their rules are not expanded. A flagged master carries the series' original dates, which may be outside the requested range; it does not represent a matching occurrence. Ordinary EXDATEs and modified occurrences remain supported during bounded expansion.

Each event includes `organizer: { id, commonName }`, up to 100 `attendees`, `attendeeCount` with the full count, and `myParticipationStatus` (empty when unavailable). Do not use a truncated attendee list as an update replacement, since that would remove the omitted attendees.

### Access Control

| Tool | Description |
|------|-------------|
| `getAccountAccess` | View which accounts the MCP server can access |

Account and tool access are configured via the extension settings page (Tools > Add-ons > Thunderbird MCP > Options). Access control is not MCP-exposed -- only the user can change it.

The same settings page has a "Send Safety" section. **Block `skipReview`** is enabled by default and rejects `skipReview: true` for `sendMail`, `replyToMessage`, `forwardMessage`, `createEvent`, and `createTask`. It also blocks adding attendees and makes meetings with attendees other than the calendar user read-only through MCP: `updateEvent` and `deleteEvent` reject every write, including time/description/recurrence edits, attendee removal, and occurrence deletion, even when `attendees` is omitted. Occurrence writes check both the series and selected occurrence; series writes also check stored exceptions. Events without attendees and events whose only attendee is the identified calendar user remain editable. If the calendar identity cannot be established, existing attendees cannot be treated as self-only. Disabling the setting permits these writes, which may email updates or cancellations without review. Event/task creation review dialogs remain available without attendees.

The separate "Filter Send Actions" section controls **Allow automatic Forward/Reply filter actions**, which defaults to off. Only the user can change this preference in the settings page; MCP clients cannot enable it.

---

## Setup

### 1. Install the extension

```bash
git clone https://github.com/TKasperczyk/thunderbird-mcp.git
```

Install `dist/thunderbird-mcp.xpi` in Thunderbird (Tools > Add-ons > Install from File), then restart. A pre-built XPI is included in the repo -- no build step needed.

**Automatic updates:** From v0.7.3 on, the add-on auto-updates through Thunderbird's add-on update check. Thunderbird downloads updates in the background and applies them on the next restart; because this add-on uses an experiment API, updates are not live hot-swapped. v0.7.3 is the last build you need to install by hand because older builds have no `update_url` and cannot auto-discover it. Thunderbird ships with `xpinstall.signatures.required=false`, so unsigned auto-updates work out of the box; a profile hardened to require signatures blocks both manual and automatic installs. If updates do not arrive, check the Add-ons gear menu and make sure **Update Add-ons Automatically** is enabled.

### 2. Configure your MCP client

Add to your MCP client config (e.g. `~/.claude.json` for Claude Code):

```json
{
  "mcpServers": {
    "thunderbird-mail": {
      "command": "node",
      "args": ["/absolute/path/to/thunderbird-mcp/mcp-bridge.cjs"]
    }
  }
}
```

#### Connecting HTTP-capable clients directly

The extension serves MCP over HTTP itself (`POST /`: `initialize`, `tools/list`, `tools/call`, notifications), so clients that support HTTP transport can skip the bridge:

```bash
claude mcp add --transport http thunderbird http://127.0.0.1:8765/ --header "Authorization: Bearer <token>"
```

- The port defaults to `8765` (8766-8774 if taken); the token is in `connection.json`, or set a stable token in the settings page so it survives restarts.
- Responses are plain JSON; there is no SSE stream and no `Mcp-Session-Id`, both optional in the Streamable HTTP spec.
- Path attachments are then read by the extension itself, under the same attachment policy as the bridge, and inside the Snap/Flatpak sandbox where applicable (the bridge is what makes host paths work for sandboxed installs).
- Keep this on localhost. For a client on another machine or container, prefer an SSH tunnel or a TLS proxy bound to loopback over "Listen on all interfaces": that mode is plain HTTP, so the token and mail content cross the network unencrypted.

### Sandbox-aware connection discovery

The bridge re-discovers `connection.json` on every cache miss. It tries these locations in order:

1. `THUNDERBIRD_MCP_CONNECTION_FILE`, if set
2. Native temp dir: `<os.tmpdir()>/thunderbird-mcp/connection.json`
3. macOS fallback: `/var/folders/*/*/T/thunderbird-mcp/connection.json` owned by the current user
4. Linux Snap: Thunderbird's live `TMPDIR` from `/proc/<pid>/environ`, plus the official snap fallback under `~/Downloads/thunderbird.tmp`
5. Linux Flatpak / Betterbird Flatpak: `$XDG_RUNTIME_DIR/app/<id>/thunderbird-mcp/connection.json` and `~/.var/app/<id>/cache/tmp/thunderbird-mcp/connection.json`

This covers native installs, the official Thunderbird snap, Thunderbird Flatpak, Thunderbird Beta Flatpak, and Betterbird Flatpak without changing the extension side. If multiple sandbox candidates exist at once, the bridge tries the newest file first. Flatpak discovery accepts only these application IDs: `org.mozilla.Thunderbird`, `org.mozilla.thunderbird`, `org.mozilla.thunderbird_esr`, `net.thunderbird.Thunderbird`, and `eu.betterbird.Betterbird`. Set `THUNDERBIRD_MCP_CONNECTION_FILE` to force a single explicit path.

Every connection file, including an explicit pin, must be a regular file of at most 4 KiB with a valid port and token. On Linux and macOS it must also belong to the bridge's user, have no group/other permissions (normally `0600`), and must not be a symlink. Rejected files are reported, never repaired or deleted. An invalid explicit pin does not fall back to discovery.

Example override:

```json
{
  "mcpServers": {
    "thunderbird-mail": {
      "command": "node",
      "args": ["/absolute/path/to/thunderbird-mcp/mcp-bridge.cjs"],
      "env": {
        "THUNDERBIRD_MCP_CONNECTION_FILE": "/absolute/path/to/connection.json"
      }
    }
  }
}
```

That's it. Your AI can now access Thunderbird.

---

### Outgoing attachments

For `sendMail`, `saveDraft`, `replyToMessage`, and `forwardMessage`, pass `attachments` as a JSON array, not a JSON-encoded string. The stdio bridge reads file-path attachments on the host and sends inline base64 to Thunderbird. The existing bridge limit is **18 MiB per path attachment**, now also applied to `saveDraft`. The extension's direct-HTTP path limit remains 50 MiB; inline base64 remains limited to 25 MiB of encoded data. Both transports enforce at most 20 attachments and 50 MiB of decoded attachments per operation.

UNC/network and device paths, dotfiles and dot-directories, application-data directories (Windows `AppData` and the macOS home `Library`), and credential/key filenames are refused. Windows alternate data streams and path components ending in dots or spaces are also refused. Files exported by `getMessage(saveAttachments: true)` can still be attached again, and ordinary local document paths continue to work. If any attachment is refused, missing, invalid, or over a limit, the entire operation fails before sending, saving a draft, or opening a review window. The error identifies refused entries; temporary files created during the failed conversion are removed. Callers that previously relied on partial attachment success must handle this error explicitly.

## Security

- **Auth tokens**: The HTTP server requires a session-scoped bearer token. Generated on startup, written to `<TmpD>/thunderbird-mcp/connection.json` with 0600 permissions. The bridge re-discovers that file automatically across native installs, Snap, Flatpak, Betterbird Flatpak, and macOS temp directories.
- **Stable tokens**: Optional stable tokens persist across restarts. The stored token is cleared on uninstall only while the add-on is enabled (best effort). If you disable the add-on before uninstalling, regenerate the token before using it again.
- **Dynamic port**: Tries ports 8765-8774, records the actual port in the connection file. No hardcoded port dependency.
- **Account access control**: Restrict which email accounts are visible to MCP clients via the settings page. Changes take effect immediately. A corrupt stored setting is shown as blocked, and saving it unchanged does not lift the restriction.
- **Calendars and address books under account restrictions**: While any account restriction is active, calendar and address-book tools (listing, lookups by ID, and writes) are blocked unless you enable "Allow all calendars" or "Allow all address books" in the settings page. Both are off by default and grant access to all calendars or all address books respectively. Without an account restriction, calendars and address books behave as before.
- **Tool access control**: Disable specific tools via the settings page. Disabled tools are hidden from `tools/list` and blocked at dispatch.
- **Encrypted-message access**: "Allow MCP clients to read encrypted messages" is off by default. It gates full-message bodies, attachments, inline images, raw-source reads, and direct reply/forward; Thunderbird may still decrypt S/MIME locally. Opaque-signed S/MIME messages are also withheld while this setting is off. This setting does not redact cached headers/previews in `searchMessages`, `getRecentMessages`, or `displayMessage`, or exclude encrypted mail from `searchBody` matches. Cached subjects may already contain decrypted protected headers. Thunderbird has no reliable encryption signal on every message header: OpenPGP's `enigmail` property is cached display status, while S/MIME's URI registry only tracks currently displayed decrypted messages. In particular, OpenPGP can [update a protected subject during background processing](https://searchfox.org/comm-central/source/mail/extensions/openpgp/content/modules/mimeDecrypt.sys.mjs#563) without the display callback that records that property; the [S/MIME registry is scoped to the current display](https://searchfox.org/comm-central/source/mailnews/base/public/nsIEncryptedMsgURIsService.idl).
- **Localhost only**: By default, the server binds to localhost only. The "Listen on all interfaces" option in settings binds to all IPv4 interfaces for WSL, Docker, or remote access. **This exposes the MCP server to every device on your local network.** Only enable on trusted networks. Auth token is always required.
- **Auto-update integrity**: Auto-update is a code-delivery channel whose integrity depends on continued control of the GitHub repository, the GitHub Actions token, and the `tomaszkasperczyk.name` registration.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Extension not loading | Check Tools > Add-ons and Themes. Errors: Tools > Developer Tools > Error Console |
| Connection refused | Make sure Thunderbird is running and the extension is enabled |
| Bridge can't find `connection.json` | Set `THUNDERBIRD_MCP_CONNECTION_FILE` explicitly if your environment uses a non-standard temp/runtime path |
| Missing recent emails | IMAP folders can be stale. Click the folder in Thunderbird to sync, or right-click > Properties > Repair Folder |
| Tool not found after update | Reconnect MCP (`/mcp` in Claude Code) to pick up new tools |
| `searchBody` returns no results | IMAP accounts need offline sync enabled for Gloda to index message bodies |
| `rawSource` fails on IMAP | Requires local/offline message copy. Enable offline sync or click the message first to cache it. |

### Release channel and Experiment API add-ons

This add-on uses Thunderbird's Experiment APIs. The add-on team has [announced plans to disable them on the Release channel](https://thunderbird.topicbox.com/groups/addons/T5426c1d2b0ba520c); Thunderbird has since [reported a postponement](https://blog.thunderbird.net/2026/06/thunderbird-monthly-development-digest-june-2026/), so do not assume every Release version blocks them.

If Thunderbird disables the add-on for this reason, use [Thunderbird ESR](https://www.thunderbird.net/thunderbird/all/) (Extended Support Release), which this add-on supports. It has been verified working on **140.16.0esr** and **153.3.1esr**. Check **Tools > Add-ons and Themes** to confirm the add-on is enabled, then open its Options page for the server status and any startup error.

---

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md) for build, isolated test, and pull request instructions.

```bash
# Build the extension
./scripts/build.sh

# Test via the bridge (handles auth automatically)
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | node mcp-bridge.cjs

# Test the HTTP API directly.
# On Snap / Flatpak / Betterbird Flatpak / macOS, point CONN_FILE at the
# real file or export THUNDERBIRD_MCP_CONNECTION_FILE first.
CONN_FILE="${THUNDERBIRD_MCP_CONNECTION_FILE:-/tmp/thunderbird-mcp/connection.json}"
TOKEN=$(jq -r .token "$CONN_FILE")
PORT=$(jq -r .port "$CONN_FILE")
curl -X POST http://127.0.0.1:$PORT \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer $TOKEN" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list"}'
```

**Dev-only extension reload:** After changing extension source locally, remove the add-on from Thunderbird, restart, reinstall the XPI, and restart again. Thunderbird caches aggressively. Regular users should install the latest release once and let auto-update handle later releases.

---

## Project structure

```
thunderbird-mcp/
├── mcp-bridge.cjs              # stdio <-> HTTP bridge (auth, port discovery)
├── extension/
│   ├── manifest.json
│   ├── background.js           # Extension entry point
│   ├── httpd.sys.mjs           # Embedded HTTP server (Mozilla)
│   ├── options.html            # Settings page UI
│   ├── options.js              # Settings page logic
│   ├── icons/                  # Extension icons
│   └── mcp_server/
│       ├── api.js              # All 40 MCP tools + auth + access control
│       └── schema.json
├── test/                       # Test suite (node:test, zero dependencies)
└── scripts/
    ├── build.sh
    └── install.sh
```

## Known issues

- IMAP folder databases can be stale until you click on them in Thunderbird
- HTML-only emails are converted to plain text (original formatting is lost)
- IMAP folder operations (rename, delete, move) are async -- verify with `listFolders` after
- Combining tags with move/trash on IMAP may not preserve tags on the moved copy -- use separate calls
- Thunderbird itself still runs pre-existing filters with cross-account move/copy targets automatically; the MCP filter tools refuse to update such rules (except disabling or deleting them) and `applyFilters` skips them
- `searchBody` on IMAP without offline sync only searches headers (Gloda limitation)
- `rawSource` requires offline message copy for IMAP -- online-only messages will error

---

## License

This project uses the [MIT license](LICENSE). The bundled `extension/httpd.sys.mjs` is derived from Mozilla's HTTP server and remains under [MPL-2.0](https://mozilla.org/MPL/2.0/); its license notice is retained in the file.
