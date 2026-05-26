---
name: thunderbird-mail
description: Drive the Thunderbird MCP server to triage an inbox, search and read mail, compose/reply/forward with the built-in review-window safety, and manage filters, contacts, calendars, and tasks. Use whenever the user asks to read, search, organize, or send email through Thunderbird, or to manage their Thunderbird filters, address book, or calendar.
---

# Thunderbird Mail

Agent workflows for the Thunderbird MCP server. The server exposes 50 tools over
MCP across six groups: `system`, `messages`, `folders`, `filters`, `contacts`,
and `calendar`. This guide covers how to combine them safely and efficiently.

The MCP server runs inside Thunderbird and talks to a Node bridge over localhost.
Tool names below are the exact MCP tool names — call them as-is.

## Start of session

Call `getServerCapabilities` once at the beginning. It returns, in one round-trip:
which tools are enabled, which accounts are reachable, which safety gates are on
(skipReview block, contact-write block, mailbox-export block, filter
forward/reply block), and how much per-tool rate-limit budget is left. Plan
around those constraints instead of discovering them through failed calls.

Then call `listAccounts` (accounts and their identities) and `listFolders`
(folder tree with message counts; filter by `accountId` or a `folderPath`
subtree). Folder URIs from `listFolders` are the `folderPath` values every
message tool expects.

## Triage the inbox

1. `getRecentMessages` — newest-first messages from one folder or all inboxes.
   Filter with `unreadOnly`, `flaggedOnly`, `daysBack`; paginate with `offset`.
   Results carry `threadId` and a `preview` snippet. Use `countOnly`-style
   checks via `searchMessages` for "how many unread?" questions.
2. For each candidate, prefer the cheap path: `getMessageHeaders` (single) or
   `batchGetMessageHeaders` (up to 200 IDs, one call) to read subject / sender /
   date / tags without paying the MIME-parse + body-decode cost.
3. Only call `getMessage` when you actually need the body. Use
   `bodyFormat: "markdown"` (default), `"text"`, or `"html"`; set
   `rawSource: true` for the full RFC 2822 source. Set `saveAttachments: true`
   to write attachments to a temp dir and get back file paths.
4. Act with `updateMessage`: mark read/unread, flag/unflag, add/remove tags
   (Thunderbird keywords like `$label1`), move between folders, or trash.
   Supply `messageId` for one message or `messageIds` for bulk. On IMAP, do
   tagging and moving in separate calls — combining them may drop tags on the
   moved copy.
5. `displayMessage` opens a message in Thunderbird's GUI (`3pane`, `tab`, or
   `window`) when the user wants to see it themselves.

## Search mail

`searchMessages` searches headers and returns IDs + folder paths you then feed to
`getMessage`. Key behavior:

- Multi-word queries are AND-of-tokens (every word must appear somewhere).
- Field operators: prefix with `from:`, `subject:`, `to:`, or `cc:` to restrict
  to one field.
- `searchBody: true` does full-text body search via Thunderbird's Gloda index
  (slower; IMAP needs offline sync enabled to index bodies).
- `includeSubfolders`, `countOnly`, date range (`startDate`/`endDate`),
  `unreadOnly`, `flaggedOnly`, `tag`, and offset-based pagination are supported.

Specialized searches:

- `searchByThread` — given any messageId + folderPath, return headers for every
  message in the same thread.
- `getSenderHistory` — "have I corresponded with this person before?" Scans the
  inbox of every accessible account for a sender address.
- `searchAttachments` — find messages by attachment filename substring and/or
  MIME-type prefix.
- `exportMailbox` — bulk export a folder to a JSON-lines file (headers-only by
  default; `includeBody` is slower). Disabled by a pref by default; if it errors,
  tell the user to enable mailbox export in the extension options.

## Compose, reply, forward — the review-window safety

This is the most important safety contract. **By default nothing is sent without
the user seeing it first.**

- `sendMail` — compose a new message. Default: opens a compose window for the
  user to review and edit. Set `skipReview: true` to send directly.
- `replyToMessage` — reply with quoted original and correct threading
  (`replyAll` for reply-all). Default: opens a review window. Supports
  `skipReview`.
- `forwardMessage` — forward with original attachments preserved. Default:
  review window. Supports `skipReview`.
- `saveDraft` — save to the identity's Drafts folder without sending or opening a
  window, for a human to finish later.

Rules to follow:

1. **Default to the review window.** Only pass `skipReview: true` when the user
   has explicitly approved the exact content in this session.
2. The user can turn on a "Block skipReview" gate in the extension options. If it
   is on, any `skipReview: true` call is rejected — retry with `skipReview` omitted
   (the review-window path still works). `getServerCapabilities` reports whether
   this gate is on.
3. The `from` identity is validated strictly. If it does not match a configured
   Thunderbird identity, the tool errors instead of silently substituting another
   account. Resolve the right identity from `listAccounts` first.
4. Use `dryRunCompose` to self-check before committing: it resolves the `from`
   identity, counts recipients, runs every attachment through the same
   path/size/deny-list pipeline `sendMail` uses, and returns the sanitized
   subject — all without sending or saving.
5. Pass an `idempotencyKey` on `sendMail` / `replyToMessage` / `forwardMessage`
   when a retry must not double-send: a prior successful send with the same key
   in the last 24h returns the original result instead of sending again.
6. Attachments can be file paths or inline base64 objects. A deny-list blocks
   sensitive paths (keys, credential stores, profile data); oversized
   attachments are rejected.

Reusable content: `listTemplates` lists user-authored templates (Markdown files
with frontmatter under the profile's `thunderbird-mcp/templates/`), and
`renderTemplate` fills one in with variables, returning `{ subject, body, isHtml }`
ready to feed into `sendMail`. Prefer templates for repeated outreach/reply
patterns so wording stays consistent.

## Folders

`createFolder`, `renameFolder`, `deleteFolder` (moves to Trash, or permanently
deletes if already in Trash), `moveFolder` (within the same account),
`emptyTrash`, and `emptyJunk`. IMAP folder operations are async — confirm with a
follow-up `listFolders` rather than assuming immediate success.

## Filters

`listFilters` shows each rule's conditions and actions in human-readable form.
`createFilter` / `updateFilter` take structured `conditions` (from, subject,
date, ...) and `actions` (move, tag, flag, ...). `deleteFilter` and
`reorderFilters` manage the list; `applyFilters` runs a folder's filters on
demand so the agent can organize existing mail.

Safety: `forward` and `reply` filter actions are gated behind a pref and
refused by default — a filter runs on every incoming message with no review UI,
so a single rule could silently exfiltrate or auto-reply to mail. If a
create/update is rejected for that reason, explain the gate to the user rather
than working around it.

## Contacts

`searchContacts` (by email or name across all address books, with `maxResults`),
`createContact`, `updateContact`, `deleteContact` (by UID).

Safety: contact writes are gated behind a pref and refused by default, because
they have no review UI and could repoint an existing contact (e.g. change
"Boss" to an attacker address) so future replies misroute. If a write is
rejected, tell the user to enable contact writes in the extension options if they
trust this client. `updateContact` audits the old and new email before mutating.

## Calendar and tasks

Events: `listCalendars` (shows read-only / event / task support flags),
`createEvent` (opens a review dialog by default; `skipReview` to add directly;
accepts `status`: tentative | confirmed | cancelled), `listEvents` (date-range
query with recurring-event expansion), `updateEvent`, `deleteEvent`.

Tasks: `createTask` (pre-filled task dialog for review), `listTasks` (filter by
completion / due date / calendar), `updateTask` (title, due date, description,
priority, completion / percent-complete), and `listCategories`.

Recurring caveat: event/task CRUD operates on the whole series, not individual
occurrences. `updateEvent` / `deleteEvent` refuse to touch a recurring event
unless you pass `recurringScope: "series"` — surface that to the user before
rewriting or deleting an entire series.

## Auditing what was done

`getAuditLog` returns parsed audit entries newest-first. Every outbound compose
(`sendMail` / `replyToMessage` / `forwardMessage` / `saveDraft`) and every contact
write is recorded as a JSON line with timestamp, tool, and metadata — never the
body, attachment content, or full recipient addresses (only counts). Use it to
confirm what actually happened or to answer "did that send go out?".

## Access control (read-only to the agent)

`getAccountAccess` shows which accounts the MCP server may touch. Account and
per-tool access are configured only by the user in the extension's options page —
the agent cannot change them. If a tool is missing from the enabled set in
`getServerCapabilities`, it has been disabled by the user; do not try to work
around it.

## Practical tips

- Headers before bodies: `getMessageHeaders` / `batchGetMessageHeaders` are far
  cheaper than `getMessage`. Triage on headers, fetch bodies only when needed.
- Stale IMAP folders: if recent messages look missing, `refreshFolder` (or have
  the user click the folder in Thunderbird) to sync before searching.
- Pagination: `searchMessages` and `getRecentMessages` support `offset`; page
  through large result sets instead of asking for everything at once.
- Respect the gates: skipReview, contact writes, filter forward/reply, and
  mailbox export are all off-by-default safety prefs. When one blocks you,
  explain it to the user — do not look for a bypass.
