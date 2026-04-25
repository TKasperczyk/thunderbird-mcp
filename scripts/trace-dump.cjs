#!/usr/bin/env node
/**
 * Tail of the structured-trace SQLite emitted by extension/mcp_server/
 * lib/trace.sys.mjs (and trace-write.cjs from any Node-side scripts).
 * Used by the CI workflow on failure to surface the most recent N
 * events in human-readable form, but also useful locally during
 * development.
 *
 * Reads via better-sqlite3 (devDep) -- no shell-out to the sqlite3
 * CLI. Falls back to printing the JSONL co-file if the SQLite read
 * fails (matches trace.sys.mjs's own SQLite-or-JSONL behaviour).
 *
 * Usage:
 *   node scripts/trace-dump.cjs [--db=PATH] [--jsonl=PATH] [--limit=N]
 *                               [--scenario=GLOB] [--min-level=N]
 *                               [--trace-id=ID]
 *
 * Defaults:
 *   --db=/tmp/thunderbird-mcp-trace.sqlite
 *   --jsonl=/tmp/thunderbird-mcp-trace.jsonl
 *   --limit=200
 *   --min-level=2     (info+; flip to 1 for debug+, 0 for trace+)
 *
 * --trace-id pins to a single span correlation across extension and
 *  Node-side events; useful for debugging a specific tool call.
 *
 * Exits 0 even if the trace file is missing -- this is purely a
 * diagnostic tool and should never gate CI on its own.
 */

const fs = require("fs");

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
const scenarioGlob = args.scenario || "%";
const traceId   = args["trace-id"] || null;

const LVL_NAME = ["TRACE", "DEBUG", "INFO ", "WARN ", "ERROR"];

function fmtTs(usStr) {
  const ms = Math.floor(Number(usStr) / 1000);
  return new Date(ms).toISOString();
}

function fmtRow({ ts_us, level, scenario, event, ctx_json, trace_id, caller_file, caller_line }) {
  const lvl = LVL_NAME[Number(level)] || `L${level}`;
  const tid = trace_id ? ` trace=${trace_id}` : "";
  const at  = caller_file ? ` at ${caller_file}:${caller_line}` : "";
  let ctxOut = "";
  if (ctx_json) {
    try { ctxOut = " " + JSON.stringify(JSON.parse(ctx_json)); }
    catch { ctxOut = " " + ctx_json; }
  }
  return `${fmtTs(ts_us)} ${lvl} ${scenario}/${event}${at}${tid}${ctxOut}`;
}

function dumpFromSqlite() {
  if (!fs.existsSync(dbPath)) return false;
  let Database;
  try { Database = require("better-sqlite3"); }
  catch (e) {
    process.stderr.write(`trace-dump: better-sqlite3 not installed (${e.message}); skip SQLite path\n`);
    return false;
  }
  let db;
  try { db = new Database(dbPath, { readonly: true, fileMustExist: true }); }
  catch (e) {
    process.stderr.write(`trace-dump: SQLite open failed: ${e.message}\n`);
    return false;
  }
  let sql = `
    SELECT ts_us, level, scenario, event, ctx_json, trace_id,
           caller_file, caller_line
    FROM events
    WHERE level >= ? AND scenario LIKE ?
  `;
  const params = [minLevel, scenarioGlob];
  if (traceId) { sql += ` AND trace_id = ?`; params.push(traceId); }
  sql += ` ORDER BY id DESC LIMIT ?`;
  params.push(Math.max(1, limit));
  let rows;
  try { rows = db.prepare(sql).all(...params); }
  catch (e) {
    process.stderr.write(`trace-dump: query failed: ${e.message}\n`);
    db.close();
    return false;
  }
  db.close();
  rows.reverse();  // oldest first for readability
  const filterDesc = `level>=${minLevel} scenario LIKE '${scenarioGlob}'${traceId ? ` trace_id=${traceId}` : ""}`;
  if (rows.length === 0) {
    console.log(`(no trace events match ${filterDesc})`);
    return true;
  }
  console.log(`# ${rows.length} trace events from ${dbPath} (${filterDesc}, oldest first)`);
  for (const r of rows) console.log(fmtRow(r));
  return true;
}

function dumpFromJsonl() {
  if (!fs.existsSync(jsonlPath)) return false;
  console.log(`# trace events from ${jsonlPath} (jsonl fallback)`);
  const lines = fs.readFileSync(jsonlPath, "utf8").split("\n").filter(Boolean);
  let printed = 0;
  for (const line of lines.slice(-limit)) {
    try {
      const r = JSON.parse(line);
      if (r.level < minLevel) continue;
      if (traceId && r.trace_id !== traceId) continue;
      // Reuse the same formatter -- jsonl rows have ctx as object, sql
      // rows have ctx_json as string. Normalize.
      const ctxStr = r.ctx === undefined ? null : (r.ctx === null ? null : JSON.stringify(r.ctx));
      console.log(fmtRow({
        ts_us: r.ts_us, level: r.level,
        scenario: r.scenario, event: r.event,
        ctx_json: ctxStr, trace_id: r.trace_id,
        caller_file: r.caller_file, caller_line: r.caller_line,
      }));
      printed++;
    } catch { /* skip malformed */ }
  }
  if (printed === 0) console.log(`(no trace events in JSONL fallback match filters)`);
  return true;
}

if (!dumpFromSqlite()) {
  if (!dumpFromJsonl()) {
    console.log(`(no trace data: neither ${dbPath} nor ${jsonlPath} has events)`);
  }
}
