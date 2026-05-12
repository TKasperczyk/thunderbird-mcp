# Security Policy

## Supported Versions

We support the latest released version of `thunderbird-mcp`. Older
versions receive no fixes; the recommended path for any security
finding is to upgrade to the current release.

| Version  | Supported          |
| -------- | ------------------ |
| latest   | ✅                 |
| < latest | ❌                 |

## Reporting a Vulnerability

**Please do not file a public GitHub issue for security findings.**

Use GitHub's [private vulnerability reporting](https://github.com/TKasperczyk/thunderbird-mcp/security/advisories/new)
on this repository. That route:

- keeps the report private until a fix is available,
- generates a CVE if appropriate,
- coordinates disclosure with downstream Thunderbird users.

Acknowledgement target: 72 hours. Triage and a first response on
severity + likely fix path: 7 days. Substantive fixes for confirmed
vulnerabilities ship as a patch release.

## What's in scope

This project bridges three trust boundaries; any of them is in scope:

1. **JSON-RPC over stdin** to `mcp-bridge.cjs`. Adversarial input
   includes malformed JSON, oversized payloads, control characters,
   prototype pollution attempts.
2. **HTTP transport** between the bridge and the Thunderbird
   extension (`http://localhost:<port>` with token auth). Token
   handling, timing-safe comparison, port-binding hygiene.
3. **Tool dispatch inside the extension** (`extension/mcp_server/lib/`).
   The permission engine + per-tool argument validation are the
   primary boundary. Any path that lets a caller exceed their
   declared permission scope, or that accesses Thunderbird state
   outside the permission's intent (e.g. reading another account's
   mailbox via a tool that should only see Inbox), is in scope.
4. **Native messaging / WebExtension experiment surface**. The
   extension runs in Thunderbird's chrome context with XPCOM
   privileges -- any path that leaks XPCOM capabilities to an
   unauthenticated caller is in scope.

## What's out of scope

- DoS against `mcp-bridge.cjs` from a process that already has
  local-user trust (the bridge runs in the user's session; that's
  the threat model's TCB by design).
- Behaviour of `npm audit` / `dependabot` flagged transitive deps
  unless there is a confirmed runtime-exploitable path. We ship
  no production npm `dependencies`; only devDeps are exposed.
- Pre-release / unmerged feature branches. Report against `main`
  or the latest release.

## Coordinated disclosure

We follow the Mozilla / Thunderbird coordinated-disclosure cadence
where the issue intersects upstream Thunderbird internals. If your
finding touches `nsIMsgSend`, `nsIMsgCompose`, or other XPCOM
interfaces, we may forward (with credit) to
[Mozilla's security team](https://www.mozilla.org/en-US/security/)
or open a Bugzilla report on your behalf. We'll always coordinate
disclosure timing with you before doing so.

## Recognition

Researchers who report verified vulnerabilities are credited in the
release notes for the patch release that fixes the issue, unless
they request to remain anonymous.
