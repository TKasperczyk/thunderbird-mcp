"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

// The existing marker loaders cover individual helpers, not start/status.
// Load the complete production methods; only Thunderbird services are mocked.
function loadServerLifecycle(overrides = {}) {
  const source = fs.readFileSync(
    path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8"
  );
  function markedSnippet(name) {
    const startMarker = `// BEGIN ${name}`;
    const endMarker = `// END ${name}`;
    const start = source.indexOf(startMarker);
    const end = source.indexOf(endMarker, start + startMarker.length);
    assert.ok(start >= 0, `${startMarker} missing`);
    assert.ok(end > start, `${endMarker} missing`);
    return source.slice(start, end);
  }

  const state = {
    attempts: 0,
    servers: [],
    timers: [],
    pendingStops: [],
    connection: null,
    removals: 0,
    logs: [],
    ...overrides,
  };
  class HttpServer {
    constructor() {
      this.listening = false;
      this.binds = [];
      this.stops = 0;
      state.servers.push(this);
    }
    registerPathHandler() {}
    start(port) {
      this.binds.push(port);
      if (state.bindError) throw state.bindError;
      assert.ok(!state.servers.some(server => server.listening), "previous listener must stop before rebinding");
      this.listening = true;
    }
    stop(callback) {
      this.stops++;
      let result;
      if (!callback) {
        result = new Promise(resolve => { callback = resolve; });
      }
      const finish = () => {
        this.listening = false;
        if (typeof callback === "function") callback();
        else callback.onStopped();
      };
      if (state.holdStops) state.pendingStops.push(finish);
      else finish();
      return result;
    }
  }

  const connectionPath = "/mock-tmp/thunderbird-mcp/connection.json";
  const sandbox = {
    MCP_DEFAULT_PORT: 8765,
    MCP_MAX_PORT_ATTEMPTS: 10,
    CONNECTION_FILE_REFRESH_MS: 30000,
    console: {
      log: (...args) => state.logs.push(args),
      warn: (...args) => state.logs.push(args),
      error: (...args) => state.logs.push(args),
    },
    ChromeUtils: {
      importESModule(uri) {
        if (uri.startsWith("resource://thunderbird-mcp/httpd.sys.mjs")) {
          state.attempts++;
          if (state.importError) throw state.importError;
          return { HttpServer };
        }
        return {};
      },
    },
    Services: {
      appinfo: { OS: "WINNT", processID: 4242 },
      prefs: { clearUserPref() {} },
      io: { newURI() { throw { name: "NS_ERROR_FILE_NOT_FOUND" }; } },
      dirsvc: {
        get: () => ({
          append() {},
          exists: () => true,
          isSymlink: () => false,
          clone: () => ({
            path: connectionPath,
            append() {},
            exists: () => state.connection !== null,
            remove() { state.connection = null; },
          }),
        }),
      },
    },
    Ci: { nsIFile: { DIRECTORY_TYPE: 1 }, nsITimer: { TYPE_REPEATING_SLACK: 1 } },
    Cc: {
      "@mozilla.org/network/file-output-stream;1": {
        createInstance: () => ({
          init() { if (state.writeError) throw state.writeError; },
        }),
      },
      "@mozilla.org/intl/converter-output-stream;1": {
        createInstance: () => ({
          init() {},
          writeString(data) { state.connection = JSON.parse(data); },
          close() {},
        }),
      },
      "@mozilla.org/timer;1": {
        createInstance() {
          const timer = {
            cancelled: false,
            initWithCallback(callback) {
              this.callback = callback;
              if (state.timerError) throw state.timerError;
            },
            cancel() { this.cancelled = true; },
          };
          state.timers.push(timer);
          return timer;
        },
      },
    },
    getStableAuthTokenPref: () => "a".repeat(64),
    isListenAllEnabled: () => false,
    PREF_LISTEN_ALL: "extensions.thunderbird-mcp.listenAll",
    readConnectionInfo: () => ({ path: connectionPath, data: state.connection }),
    removeConnectionInfo() {
      state.removals++;
      state.connection = null;
    },
  };
  vm.createContext(sandbox);
  vm.runInContext([
    markedSnippet("CONNECTION INFO REFRESH HELPERS"),
    `this.api = {
      ${markedSnippet("SERVER LIFECYCLE")}
      ${markedSnippet("LISTEN ALL SETTER")}
    };`,
  ].join("\n"), sandbox);
  return { api: sandbox.api, sandbox, state };
}

describe("server startup lifecycle", () => {
  it("evicts a synchronous failure and allows a successful retry", async () => {
    const { api, sandbox, state } = loadServerLifecycle({ importError: new Error("import failed") });
    assert.equal((await api.getServerInfo()).running, false);

    const result = await api.start();
    assert.equal(result.success, false);
    assert.match(result.error, /import failed/);
    assert.equal(sandbox.__tbMcpStartPromise, null);
    assert.equal((await api.getServerInfo()).running, false);
    assert.equal(state.servers.length, 0);

    state.importError = null;
    assert.equal((await api.start()).success, true);
    assert.equal(state.attempts, 2);
    assert.equal((await api.getServerInfo()).running, true);
    assert.equal((await api.getServerInfo()).port, 8765);
  });

  it("coalesces concurrent starts and memoizes a successful start", async () => {
    const { api, sandbox, state } = loadServerLifecycle();
    const first = api.start();
    const cached = sandbox.__tbMcpStartPromise;
    const second = api.start();
    const [firstResult, secondResult] = await Promise.all([first, second]);

    assert.equal(firstResult.success, true);
    assert.strictEqual(secondResult, firstResult);
    assert.strictEqual(await api.start(), firstResult);
    assert.strictEqual(sandbox.__tbMcpStartPromise, cached);
    assert.equal(state.attempts, 1);
    assert.equal(state.servers.length, 1);
    assert.equal(state.timers.length, 1);
  });

  it("reports stopped and shares a failed attempt until listener teardown finishes", async () => {
    const { api, sandbox, state } = loadServerLifecycle({
      writeError: new Error("disk full"), holdStops: true,
    });
    const first = api.start();
    const cached = sandbox.__tbMcpStartPromise;
    const concurrent = api.start();
    let settled = false;
    first.then(() => { settled = true; }).catch(() => {});

    // Drain promise continuations so awaiting stop(callback)'s undefined return
    // cannot accidentally pass before the premature settlement is observable.
    await new Promise(resolve => setImmediate(resolve));
    assert.equal((await api.getServerInfo()).running, false);
    assert.equal(settled, false, "failed startup must await socket closure");
    assert.strictEqual(sandbox.__tbMcpStartPromise, cached);
    assert.equal(state.attempts, 1);
    assert.equal(state.servers[0].listening, true);
    assert.equal(state.pendingStops.length, 1);

    state.pendingStops.shift()();
    const [result, concurrentResult] = await Promise.all([first, concurrent]);
    assert.equal(result.success, false);
    assert.strictEqual(concurrentResult, result);
    assert.equal(state.servers[0].listening, false);
    assert.equal(state.servers[0].stops, 1);
    assert.equal(state.connection, null);
    assert.equal(sandbox.__tbMcpStartPromise, null);
    assert.equal((await api.getServerInfo()).running, false);

    state.writeError = null;
    assert.equal((await api.start()).success, true);
    assert.deepEqual(state.servers[1].binds, [8765]);
    assert.equal((await api.getServerInfo()).running, true);
  });

  it("cleans up a partially initialized refresh timer before retry", async () => {
    const { api, sandbox, state } = loadServerLifecycle({
      timerError: new Error("timer init failed"), holdStops: true,
    });
    const start = api.start();
    assert.equal((await api.getServerInfo()).running, false);
    assert.equal(state.timers[0].cancelled, true);
    assert.equal(sandbox.__tbMcpConnectionInfoRefreshTimer, null);

    state.pendingStops.shift()();
    assert.equal((await start).success, false);
    assert.equal(state.connection, null);
    assert.equal(state.servers[0].listening, false);

    state.timerError = null;
    assert.equal((await api.start()).success, true);
    assert.equal(state.timers.length, 2);
    assert.equal(state.timers[1].cancelled, false);
  });

  it("allows retry after exhausting the port range", async () => {
    const { api, sandbox, state } = loadServerLifecycle({ bindError: new Error("port busy") });
    assert.equal((await api.start()).success, false);
    assert.deepEqual(state.servers[0].binds, Array.from({ length: 10 }, (_, i) => 8765 + i));
    assert.equal(state.servers[0].listening, false);
    assert.equal(state.timers.length, 0);
    assert.equal(sandbox.__tbMcpStartPromise, null);
    assert.equal((await api.getServerInfo()).running, false);

    state.bindError = null;
    assert.equal((await api.start()).success, true);
  });

  it("evicts a rejected attempt after cleaning up its listener", async () => {
    const rejection = new Error("failure serialization failed");
    const { api, sandbox, state } = loadServerLifecycle({
      writeError: { toString() { throw rejection; } },
    });
    await assert.rejects(api.start(), error => error === rejection);
    assert.equal(sandbox.__tbMcpStartPromise, null);
    assert.equal(state.servers[0].listening, false);
    assert.equal(state.connection, null);
    assert.equal((await api.getServerInfo()).running, false);

    state.writeError = null;
    assert.equal((await api.start()).success, true);
  });

  for (const rejects of [false, true]) {
    it(`serializes two waiting setters when the first restart ${rejects ? "rejects" : "fails"} after binding`, async () => {
      const { api, sandbox, state } = loadServerLifecycle({
        writeError: new Error("initial write failed"), holdStops: true,
      });
      const initial = api.start();
      const firstSetter = api.setListenAll(false);
      const rejection = new Error("failure serialization failed");
      const checkedFirst = rejects
        ? assert.rejects(firstSetter, error => error === rejection)
        : firstSetter;
      const secondSetter = api.setListenAll(false);

      await new Promise(resolve => setImmediate(resolve));
      assert.equal(state.attempts, 1);
      state.writeError = rejects
        ? { toString() { throw rejection; } }
        : new Error("restart write failed");
      state.pendingStops.shift()();
      assert.equal((await initial).success, false);
      await new Promise(resolve => setImmediate(resolve));

      assert.equal(state.attempts, 2, "second setter must wait for the first restart's teardown");
      assert.equal(state.pendingStops.length, 1);
      assert.equal(state.servers[1].listening, true);
      const cached = sandbox.__tbMcpStartPromise;
      assert.ok(cached, "the failed restart must remain cached while stopping");
      const joinedStart = api.start();
      const checkedJoined = rejects
        ? assert.rejects(joinedStart, error => error === rejection)
        : joinedStart;
      assert.strictEqual(sandbox.__tbMcpStartPromise, cached);
      assert.equal((await api.getServerInfo()).running, false);

      state.writeError = null;
      state.pendingStops.shift()();
      const firstResult = await checkedFirst;
      const joinedResult = await checkedJoined;
      if (!rejects) {
        assert.equal(firstResult.success, false);
        assert.strictEqual(joinedResult, firstResult);
      }
      assert.equal((await secondSetter).success, true);
      assert.equal(state.attempts, 3);
      assert.equal(state.servers[1].listening, false);
      assert.deepEqual(state.servers[2].binds, [8765]);
      assert.equal(state.pendingStops.length, 0);
      assert.equal(state.connection.port, 8765);
      assert.equal((await api.getServerInfo()).running, true);
    });

    it(`waits for ${rejects ? "rejected" : "failed"} startup cleanup before a settings restart`, async () => {
      const rejection = new Error("failure serialization failed");
      const { api, sandbox, state } = loadServerLifecycle({
        writeError: rejects ? { toString() { throw rejection; } } : new Error("disk full"),
        holdStops: true,
      });
      const start = api.start();
      const checkedStart = rejects ? assert.rejects(start, error => error === rejection) : start;
      const cached = sandbox.__tbMcpStartPromise;
      state.writeError = null;
      const restart = api.setListenAll(false);

      assert.equal((await api.getServerInfo()).running, false);
      assert.equal(state.attempts, 1, "settings must not bypass pending failure cleanup");
      assert.strictEqual(sandbox.__tbMcpStartPromise, cached);
      state.pendingStops.shift()();
      await checkedStart;
      assert.equal((await restart).success, true);
      assert.equal(state.attempts, 2);
      assert.equal(state.servers[0].listening, false);
      assert.deepEqual(state.servers[1].binds, [8765]);
      assert.equal(state.connection.port, 8765);
      assert.equal((await api.getServerInfo()).running, true);
    });

    it(`preserves a newer attempt's promise and connection after ${rejects ? "rejection" : "failure"}`, async () => {
      const rejection = new Error("failure serialization failed");
      const { api, sandbox, state } = loadServerLifecycle({
        writeError: rejects ? { toString() { throw rejection; } } : new Error("disk full"),
        holdStops: true,
      });
      const oldStart = api.start();
      const newer = Promise.resolve({ success: true, port: 8770 });
      const newerConnection = { port: 8770, token: "new", pid: 4242 };
      sandbox.__tbMcpStartPromise = newer;
      state.connection = newerConnection;
      state.pendingStops.shift()();

      if (rejects) await assert.rejects(oldStart, error => error === rejection);
      else assert.equal((await oldStart).success, false);
      assert.strictEqual(sandbox.__tbMcpStartPromise, newer);
      assert.strictEqual(state.connection, newerConnection);
    });
  }

  it("waits for callback-based listener stop during a settings restart", async () => {
    const { api, state } = loadServerLifecycle({ holdStops: true });
    await api.start();
    const restart = api.setListenAll(false);
    let settled = false;
    restart.then(() => { settled = true; }).catch(() => {});

    await new Promise(resolve => setImmediate(resolve));
    assert.equal(settled, false);
    assert.equal(state.attempts, 1);
    assert.equal((await api.getServerInfo()).running, false);
    state.pendingStops.shift()();
    assert.equal((await restart).success, true);
    assert.equal(state.servers[0].listening, false);
    assert.deepEqual(state.servers[1].binds, [8765]);
  });

  it("follows a direct retry that replaces the startup a setter is waiting for", async () => {
    const { api, sandbox, state } = loadServerLifecycle({
      importError: new Error("transient import failure"), holdStops: true,
    });
    const setter = api.setListenAll(false);
    const initial = api.start();
    await Promise.resolve();
    state.importError = null;
    const retry = api.start();
    assert.equal((await initial).success, false);
    assert.equal((await retry).success, true);
    await new Promise(resolve => setImmediate(resolve));

    assert.equal(state.attempts, 2);
    assert.equal(state.pendingStops.length, 1);
    assert.equal((await api.getServerInfo()).running, false);
    state.pendingStops.shift()();
    assert.equal((await setter).success, true);
    assert.equal(state.attempts, 3);
    assert.equal(state.servers[0].listening, false);
    assert.equal(state.servers[1].listening, true);
    assert.equal(state.connection.port, 8765);
    assert.equal((await api.getServerInfo()).running, true);
    assert.equal(sandbox.__tbMcpRestartPromise, null);
  });

  it("does not clear a newer startup promise or connection while a setter stops its listener", async () => {
    const { api, sandbox, state } = loadServerLifecycle({ holdStops: true });
    await api.start();
    const restart = api.setListenAll(false);
    await new Promise(resolve => setImmediate(resolve));

    const newerResult = { success: true, port: 8770 };
    const newer = Promise.resolve(newerResult);
    const newerConnection = { port: 8770, token: "new", pid: 4242 };
    sandbox.__tbMcpStartPromise = newer;
    state.connection = newerConnection;
    state.pendingStops.shift()();

    assert.strictEqual(await restart, newerResult);
    assert.strictEqual(sandbox.__tbMcpStartPromise, newer);
    assert.strictEqual(state.connection, newerConnection);
    assert.equal(state.attempts, 1);
  });

  it("does not report running from a cached promise or stale connection file", async () => {
    const { api, sandbox, state } = loadServerLifecycle();
    sandbox.__tbMcpStartPromise = Promise.resolve({ success: false });
    state.connection = { port: 8765, token: "old", pid: 1 };
    assert.equal((await api.getServerInfo()).running, false);
  });

  it("refreshes a missing connection file through the production Windows writer", async () => {
    const { api, state } = loadServerLifecycle();
    await api.start();
    const connection = state.connection;
    state.connection = null;
    state.timers[0].callback();
    assert.deepEqual(state.connection, connection);
    assert.equal(state.attempts, 1);
    assert.equal((await api.getServerInfo()).running, true);
  });
});
