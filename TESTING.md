# Testing

Thunderbird MCP includes 160 automated tests covering the bridge, validation, auth, and stress/adversarial scenarios. All tests run without a live Thunderbird instance.

## Running tests

```bash
npm test
```

Tests use Node.js built-in test runner (`node:test`) and require Node.js >= 18.0.0. Tests run sequentially (`--test-concurrency=1`) to avoid port conflicts.

When Thunderbird is running on port 8765, tests that need to bind that port will **skip gracefully** rather than fail.

## Test files

### `test/auth.test.cjs`

Tests for the auth token system and connection file discovery.

| Suite | What it covers |
|-------|---------------|
| Auth: connection info file | Bridge reads port/token from connection.json; falls back to default port when file is missing |
| Auth: token verification | Correct token succeeds; wrong token is rejected with auth error |

### `test/bridge.test.cjs`

Tests for the MCP bridge (stdio-to-HTTP translation layer).

| Suite | What it covers |
|-------|---------------|
| Bridge: MCP lifecycle methods | `initialize`, `ping`, `resources/list`, `prompts/list` all handled locally without forwarding to Thunderbird |
| Bridge: notifications | Notification messages (no `id`) return null (no response expected) |
| Bridge: HTTP forwarding | `tools/list` and `tools/call` are forwarded to the Thunderbird HTTP server and responses are relayed back |
| Bridge: sanitizeJson | Control characters (ASCII 0x00-0x1F except `\n`, `\r`, `\t`) are stripped from responses to prevent JSON parse failures |

### `test/validation.test.cjs`

Parameter validation tests for all 32 tools. The validator runs before any tool is dispatched.

| Suite | What it covers |
|-------|---------------|
| Validation: unknown tools | Rejects calls to non-existent tools |
| Validation: required parameters | Enforces required params; rejects null for required fields |
| Validation: type checking | String, number, boolean, and array type enforcement |
| Validation: unknown parameters | Rejects unexpected parameter names to catch typos |
| Validation: multiple errors | Reports all validation errors in a single response |
| Validation: optional parameters | Optional params can be omitted or provided with correct types |
| Validation: updateMessage with tags | `addTags`/`removeTags` accept arrays, reject wrong types; `tag` filter in searchMessages |
| Validation: account access control | `getAccountAccess` accepts no params; `setAccountAccess` requires array |
| Validation: folder management | `renameFolder`, `deleteFolder`, `moveFolder` param validation |
| Validation: attachment sending | Accepts file paths, inline base64 objects, and mixed arrays; rejects non-array |
| Validation: contact write operations | `createContact`, `updateContact`, `deleteContact` param validation |

### `test/stress.test.cjs`

Adversarial and edge-case tests that exercise boundaries a happy-path suite won't catch.

| Suite | What it covers |
|-------|---------------|
| **Pagination** | |
| Pagination: boundary conditions | offset=0, exact boundaries, beyond-end, negative, NaN, Infinity, fractional, null, undefined |
| Pagination: maxResults edge cases | 0, -1, 1, cap at 200, string/NaN coercion, empty results |
| Pagination: sequential page traversal | Full traversal without gaps or duplicates; worst-case page size |
| **Auth stress** | |
| Auth: connection file corruption | Empty file, malformed JSON, missing port, missing token, binary garbage |
| Auth: token edge cases | Special characters, very long tokens, empty string token |
| **Bridge stress** | |
| Bridge: concurrent lifecycle requests | Rapid-fire lifecycle methods; interleaved notifications and requests |
| Bridge: parallel forwarded requests | 10 simultaneous `tools/call` requests |
| Bridge: server restart recovery | Bridge reconnects when the Thunderbird server restarts on a different port |
| Bridge: large payloads | 50KB response without truncation; large request payload |
| Bridge: malformed input handling | Empty lines don't crash the bridge |
| **Sanitization** | |
| sanitizeJson: adversarial inputs | All 32 ASCII control chars, 100KB strings with embedded controls, strings that are entirely control chars, mixed escaped/unescaped |
| **Validation stress** | |
| Validation: adversarial and edge-case inputs | Prototype pollution (`__proto__`, `constructor`), very long param names/values, deeply nested objects, wrong types, massive unknown param count |
| **Feature-specific stress** | |
| Tagging: validation edge cases | Tags as wrong types, empty arrays, large tag counts, special characters, prototype pollution via tags |
| Account access: validation edge cases | Wrong types for allowedAccountIds, unknown params, prototype pollution |
| Folder management: validation edge cases | Null/array/number/boolean for paths, unknown params, very long names, special chars |
| Attachment sending: validation edge cases | Wrong types, empty arrays, large attachment counts, mixed types |
| Contact write: validation edge cases | Null/array/boolean/number for fields, unknown params, prototype pollution, very long emails |

## Writing new tests

When adding a new tool, include tests in both `validation.test.cjs` and `stress.test.cjs`:

1. **Validation tests** -- Add tool definition to the sample tools array, then test required params, type checking, unknown params, and edge cases
2. **Stress tests** -- Cover boundary values (empty, null, max-length), malformed input (wrong types, prototype pollution), and any feature-specific adversarial scenarios

See the [contribution checklist](CLAUDE.md) for the full requirements.
