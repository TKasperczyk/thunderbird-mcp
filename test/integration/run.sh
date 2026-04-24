#!/usr/bin/env bash
# End-to-end integration runner.
#
# Assumes greenmail is already reachable on $GREENMAIL_HOST (default localhost)
# ports 3143 (IMAP) and 3025 (SMTP) -- in CI this is provided by the
# greenmail service container; locally you can `docker run greenmail/standalone`.
#
# Steps:
#   1. wait for greenmail to answer on 3143 + 3025
#   2. build XPI (if not present at dist/thunderbird-mcp.xpi)
#   3. build a fresh TB profile pointed at greenmail
#   4. launch TB --headless against that profile in the background
#   5. wait for connection.json to appear (= our MCP server booted)
#   6. seed greenmail with fixture messages
#   7. run test/integration/smoke.cjs
#   8. teardown: kill TB, report exit code
#
# Usage:
#   test/integration/run.sh              # run everything
#   SKIP_BUILD=1 test/integration/run.sh # reuse existing XPI
#   KEEP_TB=1 test/integration/run.sh    # leave TB running after (debugging)
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$ROOT"

: "${GREENMAIL_HOST:=localhost}"
: "${GREENMAIL_IMAP_PORT:=3143}"
: "${GREENMAIL_SMTP_PORT:=3025}"
: "${TB_PROFILE_DIR:=/tmp/tb-mcp-ci-profile}"
: "${CONNECTION_INFO_PATH:=/tmp/thunderbird-mcp/connection.json}"
: "${TB_STARTUP_TIMEOUT:=60}"
: "${SEED_COUNT:=20}"

log()  { printf '\033[36m[run]\033[0m %s\n' "$*"; }
fail() { printf '\033[31m[run FAIL]\033[0m %s\n' "$*" >&2; exit 1; }

TB_PID=""
cleanup() {
  local rc=$?
  if [[ -n "$TB_PID" ]] && [[ "${KEEP_TB:-0}" != "1" ]]; then
    log "tearing down: kill TB pid=$TB_PID"
    kill "$TB_PID" 2>/dev/null || true
    wait "$TB_PID" 2>/dev/null || true
  fi
  rm -f "$CONNECTION_INFO_PATH" 2>/dev/null || true
  exit "$rc"
}
trap cleanup EXIT INT TERM

# ── 1. wait for greenmail ─────────────────────────────────────────────
wait_port() {
  local host=$1 port=$2 name=$3 max=${4:-60}
  log "waiting for $name on $host:$port (up to ${max}s)"
  for i in $(seq 1 "$max"); do
    if (echo > "/dev/tcp/$host/$port") 2>/dev/null; then
      log "  -> $name reachable (after ${i}s)"
      return 0
    fi
    sleep 1
  done
  fail "$name did not become reachable on $host:$port within ${max}s"
}
wait_port "$GREENMAIL_HOST" "$GREENMAIL_IMAP_PORT" greenmail-imap
wait_port "$GREENMAIL_HOST" "$GREENMAIL_SMTP_PORT" greenmail-smtp

# ── 2. build XPI ──────────────────────────────────────────────────────
if [[ "${SKIP_BUILD:-0}" != "1" ]] || [[ ! -f dist/thunderbird-mcp.xpi ]]; then
  log "building XPI"
  node scripts/build-xpi.cjs
fi
[[ -f dist/thunderbird-mcp.xpi ]] || fail "dist/thunderbird-mcp.xpi missing after build"

# ── 3. build TB profile ───────────────────────────────────────────────
log "building TB profile at $TB_PROFILE_DIR"
bash scripts/ci/make-tb-profile.sh "$TB_PROFILE_DIR" dist/thunderbird-mcp.xpi

# ── 4. launch TB ──────────────────────────────────────────────────────
# Fresh dbus session so createDrafts / IMAP sync have one.
# nsIMsgSend/IMAP need it; without it TB throws "NS_ERROR_NOT_AVAILABLE".
if command -v dbus-launch >/dev/null 2>&1; then
  eval "$(dbus-launch --sh-syntax)"
  log "dbus session: $DBUS_SESSION_BUS_ADDRESS"
fi

# Wipe stale connection info so we detect fresh startup reliably.
rm -f "$CONNECTION_INFO_PATH" 2>/dev/null || true
mkdir -p "$(dirname "$CONNECTION_INFO_PATH")"

log "launching Thunderbird headless (profile=$TB_PROFILE_DIR)"
# MOZ_LOG writes IMAP wire traffic + account manager decisions to
# /tmp/tb-moz.log. Set to level 4 (Info) which logs commands + responses
# without message bodies. Useful when "0 messages after sync" shows up.
export MOZ_LOG="IMAP:4,MsgDB:3,timestamp"
export MOZ_LOG_FILE=/tmp/tb-moz.log
thunderbird \
  --headless \
  --profile "$TB_PROFILE_DIR" \
  --no-remote \
  > /tmp/tb-stdout.log 2> /tmp/tb-stderr.log &
TB_PID=$!
log "  -> TB pid=$TB_PID"

# ── 5. wait for connection.json ───────────────────────────────────────
log "waiting up to ${TB_STARTUP_TIMEOUT}s for $CONNECTION_INFO_PATH"
for i in $(seq 1 "$TB_STARTUP_TIMEOUT"); do
  if [[ -s "$CONNECTION_INFO_PATH" ]]; then
    log "  -> connection.json ready after ${i}s"
    break
  fi
  if ! kill -0 "$TB_PID" 2>/dev/null; then
    log "TB exited early; stderr tail:"
    tail -40 /tmp/tb-stderr.log >&2 || true
    fail "Thunderbird died before writing connection.json"
  fi
  sleep 1
done

if [[ ! -s "$CONNECTION_INFO_PATH" ]]; then
  log "TB stderr tail:"
  tail -40 /tmp/tb-stderr.log >&2 || true
  fail "timed out waiting for connection.json"
fi

# ── 6. seed greenmail ─────────────────────────────────────────────────
log "seeding greenmail with $SEED_COUNT fixture messages"
GREENMAIL_HOST="$GREENMAIL_HOST" GREENMAIL_SMTP_PORT="$GREENMAIL_SMTP_PORT" \
  node scripts/ci/seed-greenmail.cjs "$SEED_COUNT"

# Give IMAP a moment to notice the new messages so the first searchMessages
# in smoke.cjs can reliably return them. (updateFolder inside tools forces
# a sync, but letting the server settle avoids flakiness on slow CI.)
sleep 2

# ── 7. run smoke suite ────────────────────────────────────────────────
log "running smoke suite"
node test/integration/smoke.cjs
SMOKE_RC=$?

if [[ $SMOKE_RC -eq 0 ]]; then
  log "integration PASS"
else
  log "integration FAIL (exit $SMOKE_RC)"
  log "TB stderr tail:"
  tail -80 /tmp/tb-stderr.log >&2 || true
fi

exit "$SMOKE_RC"
