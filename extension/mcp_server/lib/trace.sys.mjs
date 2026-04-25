/**
 * Structured tracing for the Thunderbird MCP extension.
 *
 * Built on Mozilla's chrome-context Sqlite.sys.mjs so we don't need any
 * npm dependency. Everything lives in a single WAL-mode SQLite file at
 * a caller-chosen path (typically /tmp/thunderbird-mcp-trace.sqlite in
 * CI, or the profile directory in production). Queryable post-mortem
 * by scenario / level / trace_id -- the whole point is making auth,
 * caching, and tool-call decisions VISIBLE instead of guessing from
 * behavioural output.
 *
 * Why this exists (instead of console.log sprinkled around):
 *   - Short-circuit on level: `if (level < minLevel) return;` before we
 *     do any JSON.stringify work. Matches pino's hot-path shape.
 *   - Durable + queryable: the CI job can ship the .sqlite file as an
 *     artifact and we can grep events by scenario or join across runs.
 *   - Correlated spans: span(scenario, event, fn) emits start/end rows
 *     sharing a trace_id so we can measure tool-call latency and
 *     error-path timing without hand-stitching timestamps.
 *   - No framework lock-in: emit() takes a plain {level, scenario,
 *     event, ctx} tuple. Drop-in replacement for any other logger.
 *
 * Schema (deliberately flat -- denormalised for append-only write):
 *
 *   events(id, ts_us, level, scenario, event, ctx_json, trace_id,
 *          caller_file, caller_line)
 *
 *   level: 0=trace, 1=debug, 2=info, 3=warn, 4=error (matches pino)
 *   scenario: namespace, e.g. "auth.inject", "imap.login",
 *             "tool-call.dispatch", "registry.load"
 *   event: short stable name, e.g. "login_added", "cache_miss",
 *          "server_password_set", "span_start", "span_end"
 *   ctx_json: arbitrary JSON blob for context (folder URI, tool name,
 *             error message, etc.); kept as TEXT for flexibility
 *   trace_id: short random ID for span correlation; null for one-shot
 *   caller_file/line: parsed from Error().stack at emit() time, points
 *             at the source line that called tracer.info(...) etc.
 *             Lets a failure dump pin a hang to a specific await
 *             instead of "somewhere in compose". Costs ~1us per event.
 *
 * Usage:
 *
 *   const { makeTrace } = ChromeUtils.importESModule(
 *     "resource://thunderbird-mcp/mcp_server/lib/trace.sys.mjs"
 *   );
 *   const tracer = await makeTrace({
 *     ChromeUtils,
 *     dbPath: "/tmp/thunderbird-mcp-trace.sqlite",
 *     minLevel: 1,           // debug and up
 *     fallbackPath: "/tmp/thunderbird-mcp-trace.jsonl",
 *   });
 *   tracer.info("auth.inject", "login_added", { origin: "imap://...", user: "..." });
 *   await tracer.span("imap.login", "fetch_capability", async () => { ... });
 *   await tracer.flush();    // on shutdown
 *
 * If Sqlite.sys.mjs fails to open the DB (disk full / permission / the
 * module isn't available on old TB), the tracer degrades to append-
 * only JSON-lines at `fallbackPath`. Every emit() still runs -- caller
 * never sees the failure. That's the "no worse than before" floor.
 */

export const LEVEL = Object.freeze({
  TRACE: 0, DEBUG: 1, INFO: 2, WARN: 3, ERROR: 4,
});

const LEVEL_NAME = ["trace", "debug", "info", "warn", "error"];

function newTraceId() {
  // 8 bytes hex -- enough collision-resistance for within a single
  // process run, short enough to grep. Avoid crypto.randomUUID here
  // because Date.now + Math.random is chrome-context safe without
  // any Web Crypto import dance.
  const a = Math.floor(Math.random() * 0xffffffff).toString(16).padStart(8, "0");
  const b = Math.floor(Math.random() * 0xffffffff).toString(16).padStart(8, "0");
  return a + b;
}

function nowUs() {
  // Microseconds since epoch. We don't have process.hrtime in chrome
  // JS; Date.now()*1000 is accurate enough for our granularity.
  return Date.now() * 1000;
}

/**
 * Pluck the first stack frame OUTSIDE this trace module, returning
 * `{ file, line }`. Mozilla's stack format looks like
 *   "frameFn@resource://thunderbird-mcp/.../trace.sys.mjs:123:9"
 * with one frame per line. We skip frames whose URL contains the
 * trace module's own filename so the pointer always lands on the
 * caller of emit/info/error/etc, not on emit() itself.
 *
 * Returns { file: null, line: null } if the parse fails -- we never
 * want to crash the logger, observability is strictly additive.
 */
function captureCaller() {
  try {
    const stack = new Error().stack || "";
    const lines = stack.split("\n");
    for (const raw of lines) {
      const line = raw.trim();
      if (!line) continue;
      // Mozilla format: optional "fnName@" prefix, then URL, ":line:col"
      const m = line.match(/(?:[^@]*@)?(.+):(\d+):\d+$/);
      if (!m) continue;
      const url = m[1];
      // Skip frames inside the trace module itself.
      if (url.endsWith("/trace.sys.mjs") || url.endsWith("trace.sys.mjs")) continue;
      // Strip resource://thunderbird-mcp/ prefix if present so the path
      // is relative to the extension root (matches how rest of the
      // codebase reports paths).
      let file = url;
      const RES = "resource://thunderbird-mcp/";
      if (file.startsWith(RES)) file = file.slice(RES.length);
      return { file, line: parseInt(m[2], 10) };
    }
  } catch {
    // Fall through to the null sentinel.
  }
  return { file: null, line: null };
}

/**
 * Factory. Returns a Tracer with { emit, trace, debug, info, warn,
 * error, span, flush, close }. Awaits SQLite open + schema migration.
 * Never throws -- a failed open falls through to the JSONL fallback.
 */
export async function makeTrace({
  ChromeUtils,
  dbPath,
  minLevel = LEVEL.DEBUG,
  fallbackPath = null,
}) {
  let conn = null;
  let fallbackFos = null;
  let mode = "noop";

  // Try SQLite first.
  try {
    const { Sqlite } = ChromeUtils.importESModule("resource://gre/modules/Sqlite.sys.mjs");
    conn = await Sqlite.openConnection({ path: dbPath });
    await conn.execute(`
      CREATE TABLE IF NOT EXISTS events (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        ts_us        INTEGER NOT NULL,
        level        INTEGER NOT NULL,
        scenario     TEXT    NOT NULL,
        event        TEXT    NOT NULL,
        ctx_json     TEXT,
        trace_id     TEXT,
        caller_file  TEXT,
        caller_line  INTEGER
      )
    `);
    // Schema migration: older runs of the trace DB don't have the
    // caller_* columns. Add them so cross-run dumps stay consistent.
    // ALTER TABLE ADD COLUMN is idempotent in SQLite when wrapped to
    // tolerate the "duplicate column" error.
    for (const col of [
      "caller_file TEXT",
      "caller_line INTEGER",
    ]) {
      try { await conn.execute(`ALTER TABLE events ADD COLUMN ${col}`); } catch { /* already present */ }
    }
    await conn.execute(`CREATE INDEX IF NOT EXISTS idx_scenario_ts ON events(scenario, ts_us)`);
    await conn.execute(`CREATE INDEX IF NOT EXISTS idx_level_ts    ON events(level, ts_us)`);
    await conn.execute(`CREATE INDEX IF NOT EXISTS idx_trace_id    ON events(trace_id) WHERE trace_id IS NOT NULL`);
    mode = "sqlite";
  } catch (e) {
    // Fall through to JSONL fallback. Swallow the error -- the whole
    // point of this module is that observability is additive and
    // never degrades the caller.
    conn = null;
  }

  // JSONL fallback. Opens the file lazily on first emit to avoid
  // holding the FD when nothing ever logs (e.g. production install
  // where tracing is opt-in).
  function ensureFallbackOpen() {
    if (fallbackFos || !fallbackPath) return;
    try {
      const Cc = Components.classes;
      const Ci = Components.interfaces;
      const file = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
      file.initWithPath(fallbackPath);
      const fos = Cc["@mozilla.org/network/file-output-stream;1"].createInstance(Ci.nsIFileOutputStream);
      // PR_WRONLY | PR_CREATE_FILE | PR_APPEND
      fos.init(file, 0x02 | 0x08 | 0x10, 0o644, 0);
      const cos = Cc["@mozilla.org/intl/converter-output-stream;1"].createInstance(Ci.nsIConverterOutputStream);
      cos.init(fos, "UTF-8");
      fallbackFos = cos;
    } catch {
      fallbackFos = null;
    }
  }

  function writeFallback(row) {
    ensureFallbackOpen();
    if (!fallbackFos) return;
    try {
      fallbackFos.writeString(JSON.stringify(row) + "\n");
    } catch {
      // Disk full or similar -- we've already done what we can.
    }
  }

  // Primary emit. Short-circuits on level BEFORE serializing ctx or
  // walking the stack -- both are expensive when the level is filtered
  // out (matches pino's hot-path discipline).
  function emit(level, scenario, event, ctx) {
    if (level < minLevel) return;
    const caller = captureCaller();
    const row = {
      ts_us: nowUs(),
      level,
      level_name: LEVEL_NAME[level] || String(level),
      scenario,
      event,
      ctx: ctx || null,
      trace_id: (ctx && ctx.trace_id) || null,
      caller_file: caller.file,
      caller_line: caller.line,
    };
    // Fire-and-forget SQLite write. Errors log to Services.console so
    // we never lose a row silently; callers never have to await.
    if (conn) {
      conn.execute(
        `INSERT INTO events(ts_us, level, scenario, event, ctx_json, trace_id, caller_file, caller_line)
         VALUES(:ts_us, :level, :scenario, :event, :ctx_json, :trace_id, :caller_file, :caller_line)`,
        {
          ts_us: row.ts_us,
          level: row.level,
          scenario: row.scenario,
          event: row.event,
          ctx_json: ctx ? JSON.stringify(ctx) : null,
          trace_id: row.trace_id,
          caller_file: row.caller_file,
          caller_line: row.caller_line,
        }
      ).catch((e) => {
        try { Services.console.logStringMessage(`thunderbird-mcp trace write failed: ${e && e.message}`); } catch {}
        writeFallback(row);
      });
    } else {
      writeFallback(row);
    }
  }

  /**
   * Measure an async function as a single span. Emits two events
   * sharing a trace_id: <event>.start with ctx, and <event>.end with
   * {duration_ms, ok, err}. Re-throws so caller sees the error.
   */
  async function span(scenario, event, ctx, fn) {
    // Two calling shapes: span(scenario, event, fn) or
    // span(scenario, event, ctx, fn). Detect and normalize.
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
        err: (e && e.message) || String(e),
      });
      throw e;
    }
  }

  async function flush() {
    if (conn) {
      try { await conn.execute("PRAGMA wal_checkpoint(TRUNCATE)"); } catch {}
    }
    if (fallbackFos) {
      try { fallbackFos.flush && fallbackFos.flush(); } catch {}
    }
  }

  async function close() {
    await flush();
    if (conn) {
      try { await conn.close(); } catch {}
      conn = null;
    }
    if (fallbackFos) {
      try { fallbackFos.close(); } catch {}
      fallbackFos = null;
    }
  }

  return {
    mode,                      // "sqlite" | "noop" (caller can branch on this)
    emit,
    trace: (s, e, c) => emit(LEVEL.TRACE, s, e, c),
    debug: (s, e, c) => emit(LEVEL.DEBUG, s, e, c),
    info:  (s, e, c) => emit(LEVEL.INFO,  s, e, c),
    warn:  (s, e, c) => emit(LEVEL.WARN,  s, e, c),
    error: (s, e, c) => emit(LEVEL.ERROR, s, e, c),
    span,
    flush,
    close,
    newTraceId,
  };
}
