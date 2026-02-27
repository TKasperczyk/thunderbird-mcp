# Code Review Findings

Tracked from full codebase review conducted 2026-02-27.

## Critical

- [x] **S1** — HTTP server may bind to all interfaces, not just localhost
  - Verified: `httpd.sys.mjs:489` already calls `_start(port, "localhost")`. **False positive.**

## Important

- [x] **S2** — No authentication on HTTP endpoint
  - Fixed: auth token generated at extension startup, written to `~/.thunderbird-mcp-auth`,
    bridge reads it and sends `Authorization: Bearer <token>`. Cleaned up on shutdown.

- [ ] **S3** — Attachment saves to `/tmp/thunderbird-mcp/` grow unboundedly
  - `api.js:1162-1193` — no cleanup mechanism
  - Fix: document cleanup responsibility, consider on-shutdown cleanup

- [ ] **S4** — `addAttachments` accepts arbitrary file paths
  - `api.js:595-617` — could attach sensitive files like `/etc/shadow`
  - Fix: restrict to home directory or whitelist, or log warnings

- [ ] **R1** — IMAP `updateFolder` is fire-and-forget, results may be stale
  - `api.js:669-677` — no staleness indication in results
  - Fix: add `imapSyncPending` field to search/recent results for IMAP folders

- [x] **R5** — Reply/Forward use raw subject for Re:/Fwd: prefix check
  - Fixed: both now use `msgHdr.mime2DecodedSubject` with fallback to `msgHdr.subject`

- [x] **Q2** — "Find trash folder" logic duplicated
  - Fixed: extracted `findTrashFolder(folder)` helper, used in `deleteMessages` and `updateMessage`

## Suggestions

- [ ] **R3** — 40+ empty `catch {}` blocks suppress debugging info
  - Fix: add `console.debug` to critical paths

- [x] **R6** — `searchContacts` doesn't run output through `sanitizeForJson`
  - Fixed: all contact fields now wrapped with `sanitizeForJson()`

- [x] **P3** — Bridge uses wrong JSON-RPC error code (-32700 for connection failures)
  - Fixed: changed to -32603 (Internal error)

- [ ] **Q3** — `callTool` switch and `tools[]` array can drift out of sync
  - `api.js:2396-2441`
  - Fix: consider registry pattern (lower priority)
