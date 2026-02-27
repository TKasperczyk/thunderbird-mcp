# Code Review Findings

Tracked from full codebase review conducted 2026-02-27.

## Critical

- [x] **S1** — HTTP server may bind to all interfaces, not just localhost
  - Verified: `httpd.sys.mjs:489` already calls `_start(port, "localhost")`. **False positive.**

## Important

- [x] **S2** — No authentication on HTTP endpoint
  - Fixed: auth token generated at extension startup, written to `~/.thunderbird-mcp-auth`,
    bridge reads it and sends `Authorization: Bearer <token>`. Cleaned up on shutdown.

- [x] **S3** — Attachment saves to `/tmp/thunderbird-mcp/` grow unboundedly
  - Fixed: `ATTACHMENT_DIR` constant extracted, `/tmp/thunderbird-mcp/` recursively removed in `onShutdown`.

- [x] **S4** — `addAttachments` accepts arbitrary file paths
  - Fixed: blocked sensitive path prefixes (`/etc/`, `/root/`, `/proc/`, `/sys/`, `/dev/`)
    and patterns (`.ssh/`, `.gnupg/`, `.aws/`). Symlinks resolved before checking.
    Returns `blocked` array in result.

- [x] **R1** — IMAP `updateFolder` is fire-and-forget, results may be stale
  - Fixed: `openFolder` returns `isImap` flag. `searchMessages` and `getRecentMessages`
    return `{ messages, imapSyncPending: true, note: "..." }` when IMAP folders are involved.

- [x] **R5** — Reply/Forward use raw subject for Re:/Fwd: prefix check
  - Fixed: both now use `msgHdr.mime2DecodedSubject` with fallback to `msgHdr.subject`.

- [x] **Q2** — "Find trash folder" logic duplicated
  - Fixed: extracted `findTrashFolder(folder)` helper, used in `deleteMessages` and `updateMessage`.

## Suggestions

- [x] **R3** — 40+ empty `catch {}` blocks suppress debugging info
  - Fixed: added `console.debug` to critical paths (IMAP refresh, folder access,
    search iteration, filter operations, HTTP request parsing). Low-risk folder-walking
    catches left as-is to avoid log noise.

- [x] **R6** — `searchContacts` doesn't run output through `sanitizeForJson`
  - Fixed: all contact fields now wrapped with `sanitizeForJson()`.

- [x] **P3** — Bridge uses wrong JSON-RPC error code (-32700 for connection failures)
  - Fixed: changed to -32603 (Internal error).

- [x] **Q3** — `callTool` switch and `tools[]` array can drift out of sync
  - Fixed: replaced switch with `toolHandlers` registry map. `callTool` does a map lookup.
