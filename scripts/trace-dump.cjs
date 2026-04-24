#!/usr/bin/env node
/**
 * Tail of the structured-trace SQLite emitted by extension/mcp_server/
 * lib/trace.sys.mjs. Used by the CI workflow on failure to surface
 * the most recent N events as readable text without forcing a
 * better-sqlite3 dependency on the runner.
 *
 * Strategy: shell out to the sqlite3 CLI (already installed on the CI
 * runner alongside our other apt packages). Falls back to printing
 * the JSONL file if the SQLite read fails -- keeping parity with the
 * trace module's own SQLite-or-JSONL fallback.
 *
 * Usage:
 *   node scripts/trace-dump.cjs [--db=PATH] [--jsonl=PATH] [--limit=N]
 *                               [--scenario=GLOB] [--min-level=N]
 *
 * Defaults match the extension's defaults:
 *   --db=/tmp/thunderbird-mcp-trace.sqlite
 *   --jsonl=/tmp/thunderbird-mcp-trace.jsonl
 *   --limit=200
 *   --min-level=2     (info+; flip to 1 for debug+, 0 for trace+)
 *
 * Exits 0 even if the trace file is missing -- the script is purely
 * diagnostic and shouldn't fail a CI step on its own.
 */

const fs = require("fs");
const { spawnSync } = require("child_process");

const args = Object.fromEntries(
  process.argv.slice(2)
    .filter(a => a.startsWith("--"))
    .map(a => {
      const m = a.match(/^--([^=]+)(?:=(.*))?$/);
      return [m[1], m[2] ?? "true"];
    })
);

const dbPath    = args.db    || "/tmp/thunderbird-mcp-trace.sqlite";
const jsonlPath = args.jsonl || "/tmp/thunderbird-mcp-trace.jsonl";
const limit     = parseInt(args.limit || "200", 10);
const minLevel  = parseInt(args["min-level"] || "2", 10);
const scenario  = args.scenario || "%";  // SQL LIKE glob

function ts(usStr) {
  const ms = Math.floor(Number(usStr) / 1000);
  return new Date(ms).toISOString();
}

function dumpFromSqlite() {
  if (!fs.existsSync(dbPath)) return false;
  const sql = `
    SELECT ts_us, level, scenario, event, ctx_json, COALESCE(trace_id, '')
    FROM events
    WHERE level >= ${minLevel}
      AND scenario LIKE '${scenario.replace(/'/g, "''")}'
    ORDER BY id DESC
    LIMIT ${Math.max(1, limit)}
  `;
  const r = spawnSync("sqlite3", ["-separator", "\x1f", dbPath, sql], { encoding: "utf8" });
  if (r.status !== 0) {
    process.stderr.write(`sqlite3 read failed (status=${r.status}): ${r.stderr || ""}\n`);
    return false;
  }
  const rows = (r.stdout || "").split("\n").filter(Boolean).reverse();  // oldest first for readability
  if (rows.length === 0) {
    console.log(`(no trace events match level>=${minLevel} scenario LIKE '${scenario}')`);
    return true;
  }
  const lvlName = ["TRACE", "DEBUG", "INFO ", "WARN ", "ERROR"];
  console.log(`# trace events from ${dbPath} (newest ${limit}, oldest first)`);
  for (const line of rows) {
    const [tsUs, level, scen, event, ctxJson, traceId] = line.split("\x1f");
    const lvl = lvlName[Number(level)] || `L${level}`;
    const tid = traceId ? ` trace=${traceId}` : "";
    let ctxOut = "";
    if (ctxJson) {
      try { ctxOut = " " + JSON.stringify(JSON.parse(ctxJson)); }
      catch { ctxOut = " " + ctxJson; }
    }
    console.log(`${ts(tsUs)} ${lvl} ${scen}/${event}${tid}${ctxOut}`);
  }
  return true;
}

function dumpFromJsonl() {
  if (!fs.existsSync(jsonlPath)) return false;
  console.log(`# trace events from ${jsonlPath} (jsonl fallback)`);
  const lines = fs.readFileSync(jsonlPath, "utf8").split("\n").filter(Boolean);
  for (const line of lines.slice(-limit)) {
    try {
      const r = JSON.parse(line);
      if (r.level < minLevel) continue;
      const ctx = r.ctx ? " " + JSON.stringify(r.ctx) : "";
      console.log(`${ts(r.ts_us)} ${(r.level_name || r.level).toUpperCase().padEnd(5)} ${r.scenario}/${r.event}${ctx}`);
    } catch { /* skip malformed */ }
  }
  return true;
}

if (!dumpFromSqlite()) {
  if (!dumpFromJsonl()) {
    console.log(`(no trace data: neither ${dbPath} nor ${jsonlPath} has events)`);
  }
}
