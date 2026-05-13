# Security PoC harness

Two scripts that demonstrate the two persistent-modification attack chains
discussed in PR #102 (`F1` silent filter forwarding and `F3` contact spoof).
Each script auto-discovers Thunderbird's connection file the same way
`mcp-bridge.cjs` does, dials the localhost HTTP server with its bearer
token, and runs the attack via `tools/call`.

Run them against a **test Thunderbird profile**, not your daily-driver
inbox. No real attack content is involved -- all recipients are at
`@dummy.invalid` (RFC 6761 reserved TLD, cannot resolve).

## Requirements

- Thunderbird 102+ with the `thunderbird-mcp` extension installed
- Node.js (no npm dependencies; everything is built-in `http` + the bridge module)
- A second Thunderbird profile is recommended:

  ```
  thunderbird -P  # opens the Profile Manager; create "poc-test"
  thunderbird -P poc-test -no-remote
  ```

## What the scripts do

### `f1-filter-forward.cjs`

1. Lists accounts.
2. Asks the extension to create a filter named `POC-F1-FILTER-FORWARD-DELETE-ME`
   on the first account, with action `forward` to `attacker@dummy.invalid`.
3. On a **patched build**, the call returns the
   `Filter action 'forward' is blocked by user preference` error and exits 0.
4. On an **unpatched build**, the filter is created. The script then calls
   `deleteFilter` to clean up and exits 1.

If cleanup fails (e.g. you hit Ctrl-C mid-run), open Thunderbird's
Tools -> Message Filters and delete the `POC-F1-FILTER-FORWARD-DELETE-ME`
row by hand.

### `f3-contact-spoof.cjs`

1. Calls `createContact` to add a test contact `POC-F3-Boss-DELETE-ME` with
   email `boss@dummy.invalid`.
2. Calls `updateContact` to repoint the email at `attacker@dummy.invalid`.
3. On a **patched build**, both calls return `User preference blocks
   contact writes via MCP` and the script exits 0 without touching the
   address book.
4. On an **unpatched build**, the address book entry is mutated. The script
   then calls `deleteContact` to clean up and exits 1.

If cleanup fails, open Thunderbird's Address Book and delete the
`POC-F3-Boss-DELETE-ME` card by hand.

## Running

From the repo root:

```
# Make sure Thunderbird is running with the extension loaded.
node test/poc/f1-filter-forward.cjs
node test/poc/f3-contact-spoof.cjs
```

Exit codes:

- `0` -- fix confirmed: server refused the attack.
- `1` -- attack succeeded: this build is vulnerable.
- `2` -- environment setup problem (no accounts visible, etc.).
- `3` -- unexpected server response shape; inspect the printed JSON.
- `99` -- transport failure (connection file missing, server down, etc.).

## Reproducing both states

To see both branches with a single Thunderbird install:

1. Install the patched build of `dist/thunderbird-mcp.xpi`. Run both scripts.
   They should print `[ FIX CONFIRMED ]` and exit 0.
2. In the extension's options page, flip off both `Block filter
   forward/reply` and `Block contact writes`. Restart Thunderbird so the
   prefs take effect.
3. Re-run both scripts. They will print `[ VULNERABLE ]`, complete the
   attack, clean up, and exit 1.

If you have the upstream (unpatched) `.xpi` handy, install it instead and
run the scripts to see the same `[ VULNERABLE ]` path -- that confirms the
attack is real against the published extension.

## Why these scripts ship in the repo

They are review artifacts for the security PR, not a load-bearing part of
the test suite. They aren't picked up by `node --test test/*.cjs` (the
filenames don't match `*.test.cjs`), and they require a live Thunderbird,
so they're safe to keep in-tree without affecting CI.
