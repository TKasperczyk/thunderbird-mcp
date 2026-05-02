# WSL Network Access — Implementation Log

## Goal
Enable the Thunderbird MCP server to be reachable from WSL2 via the Windows gateway IP, so the WSL bridge can connect to Thunderbird running on Windows.

---

## Changes Made

### 1. `mcp-bridge.wsl.cjs` — Dynamic Windows Host IP
**Status:** ✅ Done

Replaced hardcoded `THUNDERBIRD_HOSTS = ['127.0.0.1']` with a runtime lookup of `/etc/resolv.conf` to extract the Windows gateway IP (`nameserver`), with `127.0.0.1` as fallback.

```js
let windowsHost = '127.0.0.1';
try {
  const resolv = fs.readFileSync('/etc/resolv.conf', 'utf8');
  const m = resolv.match(/^nameserver\s+(\S+)/m);
  if (m) windowsHost = m[1];
} catch { /* fallback to 127.0.0.1 */ }
const THUNDERBIRD_HOSTS = [windowsHost, '127.0.0.1'];
```

### 2. `extension/httpd.sys.mjs` — `startAll()` Method
**Status:** ✅ Done

Added `startAll(port)` method to `nsHttpServer.prototype` that binds to `0.0.0.0` instead of `localhost`.

### 3. `extension/mcp_server/api.js` — Listen-All Setting + Auto-Restart
**Status:** ✅ Done

- Added `PREF_LISTEN_ALL` constant (`extensions.thunderbird-mcp.listenAll`)
- `start()` now reads the pref and calls `server.startAll()` or `server.start()` accordingly
- Added `getListenAll()` / `setListenAll(boolean)` API functions
- `setListenAll()` stops existing server, clears sentinels, removes stale connection file, then calls `start()` to restart

### 4. `extension/mcp_server/schema.json` — New API Schema
**Status:** ✅ Done

Added `getListenAll` and `setListenAll` entries to the experiment API schema.

### 5. `extension/options.html` + `extension/options.js` — UI Checkbox
**Status:** ✅ Done

- Added "Listen on all interfaces" checkbox under Server Status section
- Includes warning text about network exposure
- Save button calls `setListenAll()` which auto-restarts the server
- Load handler calls `getListenAll()` on page load

### 6. Rebuild
**Status:** ✅ Done

- Ran `bash scripts/build.sh` — built `dist/thunderbird-mcp.xpi`
- All 282 tests pass across 43 suites

---

## Testing Status

| Check | Result |
|-------|--------|
| Extension loads in Thunderbird | ✅ After restart |
| "Listen on all interfaces" checkbox works | ✅ Server restarts on toggle |
| Server responds on Windows localhost (`127.0.0.1:8765`) | ✅ Confirmed via PowerShell |
| Server responds from WSL via `172.18.64.1:8765` | ❌ Blocked by Windows Firewall |
| WSL bridge connects to Thunderbird | ⏳ Pending firewall fix |

---

## Remaining Issue: Windows Firewall

Thunderbird binds to `0.0.0.0` correctly, but Windows Firewall blocks inbound TCP on port 8765 from WSL.

**Fix (run in PowerShell as Admin):**
```powershell
New-NetFirewallRule -DisplayName "Thunderbird MCP" -Direction Inbound -Protocol TCP -LocalPort 8765 -Action Allow
```

After applying, the WSL bridge should connect via `172.18.64.1:8765`.

---

## Notes

- Auth token is still enforced — `0.0.0.0` binding doesn't remove authentication
- The `setListenAll` restart works within the extension (no Thunderbird restart needed)
- Connection file path in WSL: `/mnt/c/Users/bb1/AppData/Local/Temp/thunderbird-mcp/connection.json` (note: case-sensitive, `c` not `C`)
