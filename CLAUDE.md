# CLAUDE.md — Thunderbird MCP Development Guide

## Project Overview

Thunderbird MCP exposes Thunderbird email, calendar, and contacts to AI assistants via the Model Context Protocol (MCP). Architecture: MCP Client <-> mcp-bridge.cjs (stdio<->HTTP) <-> Thunderbird extension (port 8765).

## Build & Test

```bash
npm run build    # Build dist/thunderbird-mcp.xpi (cross-platform)
npm test         # Run all tests (node:test, no external deps)
```

Install the XPI in Thunderbird: Tools -> Add-ons -> gear icon -> Install Add-on From File.

## Contributing: Checklist for New Commands / Capabilities

Every new tool, feature, or capability change **must** include all of the following:

1. **Schema** — Add/update the tool's `inputSchema` in the `tools` array in `extension/mcp_server/api.js`. All parameters must have `type` and `description`. Mark required params in the `required` array.

2. **Schema validation** — The `validateToolArgs()` function validates against `inputSchema` automatically. If you add a new parameter type pattern (e.g., enum, oneOf), extend the validator to cover it.

3. **Automated tests** — Add tests in `test/` using Node.js built-in `node:test`. Cover:
   - Valid input produces expected output
   - Missing required params are rejected
   - Wrong types are rejected
   - Edge cases and error paths

4. **Documentation** — Update:
   - Tool description in the `tools` array (user-facing, shown to AI assistants)
   - README.md if the feature affects setup, usage, or architecture
   - Inline code comments for non-obvious logic

5. **Commit message** — Use conventional commits (`feat:`, `fix:`, `test:`, `docs:`). Include a clear description of *what* and *why*, not just *how*. Each logical change should be its own commit.

6. **PR description** — At the end of a feature branch, collect all changes into a comprehensive pull request with:
   - Summary of all changes
   - Test plan / verification steps
   - Breaking changes (if any)

## Code Conventions

- No external dependencies — the bridge and build use only Node.js built-ins
- Extension code runs in Thunderbird's privileged context (ChromeUtils, Cc, Ci)
- All HTTP responses use `charset=utf-8` (critical for emoji/unicode)
- The HTTP server must only listen on localhost (loopback)
- Use `Object.prototype.hasOwnProperty.call()` instead of `.hasOwnProperty()`

## Key Files

- `mcp-bridge.cjs` — stdio-to-HTTP bridge (Node.js)
- `extension/mcp_server/api.js` — All tool implementations and HTTP handler
- `extension/manifest.json` — Extension permissions and metadata
- `scripts/build.cjs` — Cross-platform XPI build script
- `test/` — Automated tests (bridge, validation)
