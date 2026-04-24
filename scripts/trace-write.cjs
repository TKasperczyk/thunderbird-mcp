/**
 * Trace-write: append-only library for writing events into the same
 * SQLite database the extension's lib/trace.sys.mjs writes. Lets any
 * Node-side script (smoke.cjs, seed-greenmail, mcp-bridge debug
 * sessions, ...) emit structured events that correlate end-to-end
 * with extension-side events under a shared trace_id.
 *
 * Schema mirrors trace.sys.mjs exactly:
 *   events(id, ts_us, level, scenario, event, ctx_json, trace_id)
 *   indexes on (scenario, ts_us), (level, ts_us), trace_id
 *
 * Why a tiny dedicated module instead of pino + transport: keeps the
 * dependency footprint small (better-sqlite3 only) and the
 * correlation semantics dead-obvious. If/when the harness logging
 * volume justifies it, swap this out for a pino transport that
 * INSERTs into the same schema -- callers won't change.
 *
 * Usage:
 *
 *   const { makeTraceWriter } = require("./trace-write.cjs");
 *   const tracer = makeTraceWriter({ dbPath: "/tmp/thunderbird-mcp-trace.sqlite" });
 *   tracer.info("smoke.tool", "call.dispatched", { tool: "inbox_inventory" });
 *   await tracer.span("smoke.tool", "call", { tool: "createDrafts" }, async () => {
 *     // ... do the work ...
 *   });
 *   tracer.close();
 *
 * Falls back to JSONL append at fallbackPath on SQLite open failure
 * (mirrors trace.sys.mjs behaviour).
 */

const LEVEL = Object.freeze({ TRACE: 0, DEBUG: 1, INFO: 2, WARN: 3, ERROR: 4 });
const LEVEL_NAME = ["trace", "debug", "info", "warn", "error"];

function newTraceId() {
  const a = Math.floor(Math.random() * 0xffffffff).toString(16).padStart(8, "0");
  const b = Math.floor(Math.random() * 0xffffffff).toString(16).padStart(8, "0");
  return a + b;
}

function nowUs() { return Date.now() * 1000; }

function makeTraceWriter({
  dbPath,
  fallbackPath = null,
  minLevel = LEVEL.DEBUG,
} = {}) {
  let db = null;
  let stmt = null;
  let mode = "noop";
  let fallbackFd = null;

  try {
    const Database = require("better-sqlite3");
    db = new Database(dbPath);
    db.pragma("journal_mode = WAL");        // ext writes too; WAL plays nice
    db.pragma("synchronous = NORMAL");      // we don't need fsync per insert
    db.exec(`
      CREATE TABLE IF NOT EXISTS events (
        id        INTEGER PRIMARY KEY AUTOINCREMENT,
        ts_us     INTEGER NOT NULL,
        level     INTEGER NOT NULL,
        scenario  TEXT    NOT NULL,
        event     TEXT    NOT NULL,
        ctx_json  TEXT,
        trace_id  TEXT
      );
      CREATE INDEX IF NOT EXISTS idx_scenario_ts ON events(scenario, ts_us);
      CREATE INDEX IF NOT EXISTS idx_level_ts    ON events(level, ts_us);
      CREATE INDEX IF NOT EXISTS idx_trace_id    ON events(trace_id) WHERE trace_id IS NOT NULL;
    `);
    stmt = db.prepare(`
      INSERT INTO events(ts_us, level, scenario, event, ctx_json, trace_id)
      VALUES(@ts_us, @level, @scenario, @event, @ctx_json, @trace_id)
    `);
    mode = "sqlite";
  } catch (e) {
    db = null;
    stmt = null;
    if (process.env.THUNDERBIRD_MCP_DEBUG) {
      process.stderr.write(`trace-write: SQLite unavailable (${e.message}), using JSONL fallback\n`);
    }
  }

  function ensureFallback() {
    if (fallbackFd !== null || !fallbackPath) return;
    try {
      const fs = require("fs");
      fallbackFd = fs.openSync(fallbackPath, "a");
    } catch {
      fallbackFd = null;
    }
  }

  function writeFallback(row) {
    ensureFallback();
    if (fallbackFd === null) return;
    try {
      require("fs").writeSync(fallbackFd, JSON.stringify(row) + "\n");
    } catch { /* swallow */ }
  }

  function emit(level, scenario, event, ctx) {
    if (level < minLevel) return;
    const row = {
      ts_us: nowUs(),
      level,
      level_name: LEVEL_NAME[level] || String(level),
      scenario,
      event,
      ctx: ctx || null,
      trace_id: (ctx && ctx.trace_id) || null,
    };
    if (stmt) {
      try {
        stmt.run({
          ts_us: row.ts_us,
          level: row.level,
          scenario: row.scenario,
          event: row.event,
          ctx_json: ctx ? JSON.stringify(ctx) : null,
          trace_id: row.trace_id,
        });
      } catch (e) {
        // Most likely cause: ext + smoke writing concurrently and we
        // hit a busy timeout. WAL mode minimises this; if it still
        // happens, fall through to JSONL so we never lose a row.
        writeFallback(row);
      }
    } else {
      writeFallback(row);
    }
  }

  async function span(scenario, event, ctx, fn) {
    if (typeof ctx === "function") { fn = ctx; ctx = null; }
    const trace_id = (ctx && ctx.trace_id) || newTraceId();
    const startCtx = Object.assign({}, ctx || {}, { trace_id });
    emit(LEVEL.DEBUG, scenario, event + ".start", startCtx);
    const t0 = Date.now();
    try {
      const r = await fn();
      emit(LEVEL.DEBUG, scenario, event + ".end", {
        trace_id, duration_ms: Date.now() - t0, ok: true,
      });
      return r;
    } catch (e) {
      emit(LEVEL.ERROR, scenario, event + ".end", {
        trace_id, duration_ms: Date.now() - t0, ok: false,
        err: e && e.message ? e.message : String(e),
      });
      throw e;
    }
  }

  function close() {
    if (db) { try { db.close(); } catch {} db = null; stmt = null; }
    if (fallbackFd !== null) {
      try { require("fs").closeSync(fallbackFd); } catch {}
      fallbackFd = null;
    }
  }

  return {
    mode,
    emit,
    trace: (s, e, c) => emit(LEVEL.TRACE, s, e, c),
    debug: (s, e, c) => emit(LEVEL.DEBUG, s, e, c),
    info:  (s, e, c) => emit(LEVEL.INFO,  s, e, c),
    warn:  (s, e, c) => emit(LEVEL.WARN,  s, e, c),
    error: (s, e, c) => emit(LEVEL.ERROR, s, e, c),
    span,
    close,
    newTraceId,
  };
}

module.exports = { makeTraceWriter, LEVEL, newTraceId };
