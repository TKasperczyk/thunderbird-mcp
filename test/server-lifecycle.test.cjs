"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { HTTPD_PORT, createHttpdHarness } = require("./helpers/httpd.cjs");
const { loadPrivacyRuntime, getConstantDeclaration } = require("./helpers/privacy-access.cjs");
const { createResourceLoader } = require("./helpers/resource-loader.cjs");
const REFRESH_TIMER_TYPE = 1;
const HTTP_ACTIVITY = { ACTIVITY_TYPE_HTTP_TRANSACTION: 2, ACTIVITY_SUBTYPE_RESPONSE_HEADER: 0x5004 };

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

  // Reuse the access harness's preference services and production helpers, but
  // leave initialization to the real getAPI-scope code below, before start().
  const access = overrides.accessRuntime || loadPrivacyRuntime({ initialize: false });
  // Seed the stable token like a configured profile without recording a write;
  // production getStableAuthTokenPref semantics apply on top of the map.
  access.values.set("extensions.thunderbird-mcp.stableAuthToken", "a".repeat(64));

  const state = {
    attempts: 0,
    servers: [],
    timers: [],
    channels: [],
    activityObservers: new Set(),
    nonce: 0,
    connectionWrites: [],
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
      this.handlers = new Map();
      this.bindModes = [];
      this.identityAdds = [];
      this.identity = {
        add: (scheme, host, port) => {
          if (state.identityAddError) throw state.identityAddError;
          this.identityAdds.push({ scheme, host, port });
        },
      };
      state.servers.push(this);
    }
    registerPathHandler(path, handler) {
      if (path !== "/" && handler && state.handlerError) throw state.handlerError;
      if (handler) this.handlers.set(path, handler);
      else this.handlers.delete(path);
    }
    start(port) {
      this.binds.push(port);
      this.bindModes.push("loopback");
      const error = typeof state.bindError === "function" ? state.bindError(port, this.binds.length) : state.bindError;
      if (error) throw error;
      assert.ok(!state.servers.some(server => server.listening), "previous listener must stop before rebinding");
      this.listening = true;
      this.port = port;
    }
    startAll(port) {
      this.start(port);
      this.bindModes[this.bindModes.length - 1] = "listenAll";
    }
    stop(callback) {
      this.stops++;
      if (state.stopError) throw state.stopError;
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

  // Reuse the NetUtil read stub pattern from read-message-stream.test.cjs;
  // model channel events separately so cancellation and late delivery are tested.
  const NetUtil = {
    newChannel(options) {
      if (state.channelError) throw state.channelError;
      const url = new URL(options.uri);
      const mode = typeof state.probeResult === "function"
        ? state.probeResult(Number(url.port), state.channels.length)
        : state.probeResult || "ours";
      const channel = {
        ...options, url, mode, loadFlags: 0, headers: {}, cancelled: false,
        bytesRead: 0, readSizes: [], responseStatusText: "OK",
        get responseStatus() {
          if (state.statusError) throw state.statusError;
          return this.statusCode ?? state.responseStatus ?? (mode === "ours" ? 200 : 404);
        },
        get proxyInfo() {
          if (state.proxyError) throw state.proxyError;
          return state.proxyInfo ?? null;
        },
        getResponseVersion(major, minor) {
          if (state.versionError) throw state.versionError;
          [major.value, minor.value] = state.responseVersion || [1, 1];
        },
        QueryInterface() {
          if (state.queryError) throw state.queryError;
          return this;
        },
        setRequestHeader(name, value) { this.headers[name] = value; },
        visitResponseHeaders(visitor) {
          for (const [name, value] of Object.entries(this.responseHeaders || state.responseHeaders || {})) {
            visitor.visitHeader(name, value);
          }
        },
        cancel() {
          this.cancelled = true;
          if (state.cancelError) throw state.cancelError;
        },
        deliver(bytes) {
          this.listener.onDataAvailable(this, { bytes, channel: this }, this.bytesRead, bytes.length);
        },
        reportResponseHead() {
          for (const observer of state.activityObservers) {
            observer.observeActivity(this, HTTP_ACTIVITY.ACTIVITY_TYPE_HTTP_TRANSACTION,
              HTTP_ACTIVITY.ACTIVITY_SUBTYPE_RESPONSE_HEADER, 0, 0, "normalized headers");
          }
        },
        respond() {
          if (mode === "timeout") return;
          let body = "foreign response";
          if (state.httpd) {
            const headers = Object.entries(this.headers).map(([name, value]) => `${name}: ${value}\r\n`).join("");
            this.requestWire = `${this.requestMethod} ${url.pathname} HTTP/1.1\r\nHost: ${url.host}\r\n${headers}\r\n`;
            this.exchange = state.httpd.request(this.requestWire);
            const [head, responseBody] = this.exchange.wire.split("\r\n\r\n");
            this.statusCode = Number(/^HTTP\/1\.1 (\d+)/.exec(head)?.[1]);
            this.responseHeaders = Object.fromEntries(head.split("\r\n").slice(1).map(line => {
              const colon = line.indexOf(":");
              return [line.slice(0, colon), line.slice(colon + 1).trim()];
            }));
            body = responseBody;
          } else if (mode === "ours") {
            const res = { setStatusLine() {}, setHeader() {}, write(value) { body = value; } };
            this.handler({ path: url.pathname, method: this.requestMethod }, res);
          }
          if (mode !== "inconclusive") {
            if (state.responseHeadReceived !== false) this.reportResponseHead();
            this.listener.onStartRequest(this);
            for (const chunk of state.bodyChunks || [body]) {
              if (!this.cancelled) this.deliver(chunk);
            }
          }
          this.listener.onStopRequest(this, 0);
        },
        asyncOpen(listener) {
          this.listener = listener;
          if (state.asyncOpenError) throw state.asyncOpenError;
          this.handler = state.servers.find(server => server.listening && server.port === Number(url.port))?.handlers.get(url.pathname);
          if (!state.holdProbes) queueMicrotask(() => this.respond());
        },
      };
      state.channels.push(channel);
      return channel;
    },
    readInputStreamToString(stream, count) {
      if (state.readError) throw state.readError;
      stream.channel.readSizes.push(count);
      stream.channel.bytesRead += count;
      const result = stream.bytes.slice(0, count);
      stream.bytes = stream.bytes.slice(count);
      return result;
    },
  };

  const connectionPath = "/mock-tmp/thunderbird-mcp/connection.json";
  const sandbox = {
    TextEncoder,
    ...access.prefNames,
    toolAccess: access.policy,
    readAccessListPref: access.readAccessListPref,
    buildTools: access.buildTools,
    getConfiguredGetMessagesLimit: access.getConfiguredGetMessagesLimit,
    MCP_DEFAULT_PORT: 8765,
    MCP_MAX_PORT_ATTEMPTS: 10,
    CONNECTION_FILE_REFRESH_MS: 30000,
    console: {
      log: (...args) => state.logs.push(args),
      warn: (...args) => state.logs.push(args),
      error: (...args) => state.logs.push(args),
    },
    ChromeUtils: {
      generateQI: () => function () { return this; },
      importESModule(uri) {
        if (uri.startsWith("resource://thunderbird-mcp/httpd.sys.mjs")) {
          state.attempts++;
          if (state.importError) throw state.importError;
          return { HttpServer };
        }
        if (uri === "resource://gre/modules/NetUtil.sys.mjs") return { NetUtil };
        return {};
      },
    },
    Services: {
      appinfo: { OS: "WINNT", processID: 4242 },
      uuid: access.Services.uuid,
      prefs: {
        ...access.Services.prefs,
        clearUserPref(name) {
          access.Services.prefs.clearUserPref(name);
          if (name === sandbox.PREF_LISTEN_ALL) state.listenAll = false;
        },
        setBoolPref(name, value) {
          access.Services.prefs.setBoolPref(name, value);
          if (name === sandbox.PREF_LISTEN_ALL) state.listenAll = value;
        },
      },
      io: { newURI() { throw { name: "NS_ERROR_FILE_NOT_FOUND" }; } },
      dirsvc: {
        get: (key) => {
          if (state.profileDirs?.[key]) return state.profileDirs[key];
          if (key !== "TmpD") throw new Error("Directory unavailable");
          return {
            append() {},
            exists: () => true,
            isSymlink: () => false,
            clone: () => ({
              path: connectionPath,
              append() {},
              exists: () => state.connection !== null,
              remove() { state.connection = null; },
            }),
          };
        },
      },
    },
    Ci: {
      nsIPrefBranch: access.Ci.nsIPrefBranch,
      nsIFile: { DIRECTORY_TYPE: 1 }, nsITimer: { TYPE_ONE_SHOT: 0, TYPE_REPEATING_SLACK: REFRESH_TIMER_TYPE },
      nsIRequest: { LOAD_ANONYMOUS: 1, LOAD_BYPASS_CACHE: 2, INHIBIT_CACHING: 4, LOAD_FRESH_CONNECTION: 8 },
      nsIChannel: { LOAD_INITIAL_DOCUMENT_URI: 16 }, nsIHttpChannel: "http",
      nsIHttpChannelInternal: "httpInternal", nsIProxiedChannel: "proxy",
      nsIHttpActivityObserver: HTTP_ACTIVITY, nsIHttpActivityDistributor: "activity",
    },
    Cr: { NS_BINDING_ABORTED: "aborted" },
    Cc: {
      "@mozilla.org/network/http-activity-distributor;1": {
        getService() {
          if (state.activityServiceError) throw state.activityServiceError;
          return {
            addObserver(observer) {
              if (state.activityAddError) throw state.activityAddError;
              state.activityObservers.add(observer);
            },
            removeObserver(observer) {
              state.activityObservers.delete(observer);
              if (state.activityRemoveError) throw state.activityRemoveError;
            },
          };
        },
      },
      "@mozilla.org/network/file-output-stream;1": {
        createInstance: () => ({
          init() { if (state.writeError) throw state.writeError; },
        }),
      },
      "@mozilla.org/intl/converter-output-stream;1": {
        createInstance: () => ({
          init() {},
          writeString(data) {
            state.connection = JSON.parse(data);
            state.connectionWrites.push(state.connection);
          },
          close() {},
        }),
      },
      "@mozilla.org/timer;1": {
        createInstance() {
          if (state.timerCreateFailures > 0) {
            state.timerCreateFailures--;
            throw new Error("timer creation failed");
          }
          const timer = {
            cancelled: false,
            initWithCallback(callback, delay, type) {
              this.callback = callback;
              this.delay = delay;
              this.type = type;
              const error = type === 0 ? state.probeTimerError : state.timerError;
              if (error) throw error;
            },
            cancel() {
              this.cancelled = true;
              if (this.type === 0 && state.probeTimerCancelError) throw state.probeTimerCancelError;
            },
          };
          state.timers.push(timer);
          return timer;
        },
      },
    },
    getStableAuthTokenPref: () => access.Services.prefs.getStringPref(sandbox.PREF_STABLE_AUTH_TOKEN, ""),
    generateAuthToken() {
      if (state.nonceError) throw state.nonceError;
      return (++state.nonce).toString(16).padStart(64, "0");
    },
    isListenAllEnabled: () => !!state.listenAll,
    AUTH_TOKEN_PATTERN: /^[0-9a-f]{64}$/,
    PREF_STABLE_AUTH_TOKEN: "extensions.thunderbird-mcp.stableAuthToken",
    PREF_LISTEN_ALL: "extensions.thunderbird-mcp.listenAll",
    readConnectionInfo: () => ({ path: connectionPath, data: state.connection }),
    removeConnectionInfo() {
      state.removals++;
      state.connection = null;
    },
  };
  if (state.extensionRoot) {
    // Execute the entire Experiment, including getAPI's import order. The
    // ordinary lifecycle fixtures below intentionally load smaller snippets.
    const rootURI = { resolve: name => state.extensionRoot + name };
    state.resources = createResourceLoader(new Map([[
      rootURI.resolve("tool-access.js"),
      fs.readFileSync(path.resolve(__dirname, "../extension/tool-access.js"), "utf8"),
    ]]));
    sandbox.Cu = { importGlobalProperties() {} };
    sandbox.ExtensionCommon = { ExtensionAPI: class {} };
    sandbox.context = { extension: { id: "thunderbird-mcp@tkasperczyk.dev", rootURI } };
    sandbox.Services.scriptloader = state.resources.scriptloader;
    sandbox.Services.obs = { notifyObservers() {} };
    sandbox.Cc["@mozilla.org/network/protocol;1?name=resource"] = {
      getService: () => state.resources.protocol,
    };
    sandbox.Cc["@mozilla.org/security/random-generator;1"] = {
      createInstance: () => ({ generateRandomBytes: count => new Uint8Array(count).fill(++state.nonce) }),
    };
    const importESModule = sandbox.ChromeUtils.importESModule;
    sandbox.ChromeUtils.importESModule = uri => {
      if (uri === "resource://gre/modules/AddonManager.sys.mjs") {
        return { AddonManager: { addAddonListener() {}, removeAddonListener() {} } };
      }
      if (uri.startsWith("resource://thunderbird-mcp/httpd.sys.mjs")) {
        assert.ok(state.resources.substitutions.has("thunderbird-mcp"));
        assert.equal(state.resources.loads.length, 1, "load Tool Access before starting the listener");
      }
      return importESModule(uri);
    };
    vm.createContext(sandbox);
    vm.runInContext(source + `
      this.experiment = new mcpServer();
      this.api = this.experiment.getAPI(context).mcpServer;
      this.shutdown = () => this.experiment.onShutdown(false);
    `, sandbox, { filename: "api.js" });
    return { api: sandbox.api, sandbox, state, NetUtil, access };
  }
  vm.createContext(sandbox);
  vm.runInContext([
    ...["DEFAULT_GET_MESSAGES_LIMIT", "MAX_GET_MESSAGES_LIMIT", "MAX_REQUEST_BODY",
      "UNDISABLEABLE_TOOLS", "GROUP_ORDER", "GROUP_LABELS", "CRUD_ORDER"].map(getConstantDeclaration),
    // Saved-search initialization reuses the production folder-role constants.
    markedSnippet("OUTBOX DESTINATION GUARD"),
    markedSnippet("CONNECTION INFO REFRESH HELPERS"),
    markedSnippet("SENSITIVE ATTACHMENT PATH HELPERS"),
    markedSnippet("THUNDERBIRD PROFILE ROOTS"),
    markedSnippet("LISTENER OWNERSHIP HELPERS"),
    `this.probeSettings = { maxBytes: LISTENER_PROBE_MAX_BYTES, timeoutMs: LISTENER_PROBE_TIMEOUT_MS };
     this.shutdown = function () { ${markedSnippet("SERVER SHUTDOWN")} };`,
    `this.getAPI = function () {
      ${markedSnippet("SERVER INSTANCE STATE")}
      ${getConstantDeclaration("tools")}
      ${getConstantDeclaration("ALL_TOOL_NAMES")}
      ${markedSnippet("TOOL ACCESS STATE")}
      return {
        ${markedSnippet("SERVER LIFECYCLE")}
        ${markedSnippet("OPTIONS ACCESS API")}
        ${markedSnippet("STABLE TOKEN OPTIONS")}
        ${markedSnippet("LISTEN ALL SETTER")}
      };
    };
    this.api = this.getAPI();`,
  ].join("\n"), sandbox);
  return { api: sandbox.api, sandbox, state, NetUtil, access };
}

// Reuse the production lifecycle above for the options page's browser API.
// The other options test loader covers account/privacy controls, not status.
function loadServerStatus(api) {
  const source = fs.readFileSync(path.resolve(__dirname, "../extension/options.js"), "utf8");
  const start = source.indexOf("// BEGIN OPTIONS SERVER STATUS");
  const end = source.indexOf("// END OPTIONS SERVER STATUS", start);
  assert.ok(start >= 0 && end > start, "options server status markers missing");
  const saveStart = source.indexOf("// BEGIN OPTIONS LISTEN ALL SAVE");
  const saveEnd = source.indexOf("// END OPTIONS LISTEN ALL SAVE", saveStart);
  assert.ok(saveStart >= 0 && saveEnd > saveStart, "options listen-all save markers missing");
  const elements = Object.fromEntries([
    "statusDot", "statusText", "serverPort", "serverPortNotice", "connFile", "buildInfo", "retryServerBtn",
    "listenAllCheckbox", "saveListenAllBtn", "saveListenAllStatus",
  ].map(id => [id, {
    textContent: "", className: "", hidden: true, disabled: false,
    addEventListener(event, listener) {
      assert.equal(event, "click");
      this.click = listener;
    },
  }]));
  const ui = {
    ...elements,
    authRefreshes: 0,
    loadAuthenticationConfig: async () => { ui.authRefreshes++; },
    document: { getElementById: id => elements[id] },
    browser: { mcpServer: api },
  };
  vm.runInNewContext(source.slice(start, end) + "\n" + source.slice(saveStart, saveEnd) +
    "\nthis.loadServerInfo = loadServerInfo;", ui);
  return ui;
}

describe("Experiment startup with trusted script URIs", () => {
  for (const extensionRoot of [
    "jar:file:///mock-profile/extensions/thunderbird-mcp.xpi!/",
    "jar:file:///C:/mock-profile/extensions/thunderbird-mcp.xpi!/",
    "file:///mock-unpacked-extension/",
  ]) {
    it(`starts the production getAPI path from ${extensionRoot}`, async () => {
      const { api, sandbox, state, access } = loadServerLifecycle({ extensionRoot });
      assert.throws(() => state.resources.scriptloader.loadSubScript(extensionRoot + "tool-access.js", {}),
        /Trying to load untrusted URI/);
      const result = await api.start();
      assert.equal(result.success, true, result.error);
      assert.equal(state.servers.length, 1);
      assert.equal(state.connectionWrites.length, 1);
      assert.deepEqual(state.resources.loads.map(load => load.uri), ["resource://thunderbird-mcp/tool-access.js"]);
      const config = await api.getToolAccessConfig();
      assert.ok(config.tools.some(tool => tool.name === "dryRunCompose"));
      assert.ok(config.tools.some(tool => tool.name === "sendMail"));
      assert.ok(access.writes.every(write => write.name.startsWith("extensions.thunderbird-mcp.")));
      sandbox.shutdown();
      assert.equal(state.resources.substitutions.size, 0, "unload releases the resource mapping");
    });
  }
});

describe("Tool Access initialization and server startup", () => {
  it("writes migrated restrictions before tracking and before the listener binds", async () => {
    const events = [];
    const access = loadPrivacyRuntime({
      legacy: true, initialize: false, initial: { PREF_DISABLED_TOOLS: '["getMessage"]' },
      beforeWrite: name => events.push(name),
    });
    assert.deepEqual(access.writes, []);
    const { api, state } = loadServerLifecycle({
      accessRuntime: access,
      bindError() {
        events.push("bind");
        const disabled = JSON.parse(access.values.get(access.prefNames.PREF_DISABLED_TOOLS));
        const tracking = JSON.parse(access.values.get(access.prefNames.PREF_TOOL_ACCESS_STATE));
        assert.ok(disabled.includes("getMessage"));
        assert.ok(disabled.includes("refreshFolder"));
        assert.ok(disabled.includes("dryRunCompose"));
        assert.equal(tracking.customized, true);
        assert.ok(tracking.consideredTools.includes("refreshFolder"));
        return null;
      },
    });
    assert.equal(state.servers.length, 0);
    assert.deepEqual(events, [access.prefNames.PREF_TOOL_ACCESS_STATE, access.prefNames.PREF_DISABLED_TOOLS, access.prefNames.PREF_TOOL_ACCESS_STATE]);
    assert.equal((await api.start()).success, true);
    assert.deepEqual(events, [access.prefNames.PREF_TOOL_ACCESS_STATE, access.prefNames.PREF_DISABLED_TOOLS, access.prefNames.PREF_TOOL_ACCESS_STATE, "bind"]);
  });

  it("does not repeat migration writes across repeated starts, Options contexts or runtime reloads", async () => {
    const access = loadPrivacyRuntime({
      legacy: true, initialize: false, initial: { PREF_DISABLED_TOOLS: '["deleteFolder"]' },
    });
    const first = loadServerLifecycle({ accessRuntime: access });
    const writesAfterInitialization = access.writes.length;
    assert.equal(writesAfterInitialization, 3);
    assert.equal((await first.api.start()).success, true);
    assert.equal((await first.api.start()).success, true);
    await first.sandbox.getAPI().getToolAccessConfig();
    assert.equal(access.writes.length, writesAfterInitialization);
    first.sandbox.shutdown();

    const reloaded = loadPrivacyRuntime({ legacy: true, initialize: false, values: access.values });
    const second = loadServerLifecycle({ accessRuntime: reloaded });
    assert.equal((await second.api.start()).success, true);
    const config = await second.api.getToolAccessConfig();
    assert.ok(config.disabledTools.includes("deleteFolder"));
    assert.ok(config.disabledTools.includes("refreshCalendar"));
    assert.deepEqual(reloaded.writes, []);
  });

  it("exposes and saves migrated Options choices even when server startup fails", async () => {
    const access = loadPrivacyRuntime({
      legacy: true, initialize: false, initial: { PREF_DISABLED_TOOLS: '["getMessage"]' },
    });
    const { api, state } = loadServerLifecycle({ accessRuntime: access, importError: new Error("httpd unavailable") });
    const beforeStart = await api.getToolAccessConfig();
    assert.equal(beforeStart.tools.find(tool => tool.name === "refreshFolder").enabled, false);
    assert.equal((await api.start()).success, false);
    assert.equal(state.servers.length, 0);
    const afterFailure = await api.getToolAccessConfig();
    assert.equal(afterFailure.revision, beforeStart.revision);
    const result = await api.setToolAccess(
      afterFailure.disabledTools.filter(name => name !== "refreshFolder"),
      afterFailure.getMessagesLimit,
      afterFailure.revision,
    );
    assert.equal(result.success, true);
    assert.equal((await api.getToolAccessConfig()).tools.find(tool => tool.name === "refreshFolder").enabled, true);
    state.importError = null;
    assert.equal((await api.start()).success, true);
    assert.equal((await api.getToolAccessConfig()).tools.find(tool => tool.name === "refreshFolder").enabled, true);
  });

  for (const failure of ["corrupt deny-list", "corrupt tracking", "deny-list write", "tracking write"]) {
    it(`keeps optional tools blocked on a running listener after ${failure}`, async () => {
      const initial = { PREF_DISABLED_TOOLS: '["getMessage"]' };
      if (failure === "corrupt deny-list") initial.PREF_DISABLED_TOOLS = "not-json";
      if (failure === "corrupt tracking") initial.PREF_TOOL_ACCESS_STATE = "not-json";
      const access = loadPrivacyRuntime({ legacy: true, initialize: false, initial });
      if (failure === "deny-list write") access.writeErrors.add(access.prefNames.PREF_DISABLED_TOOLS);
      if (failure === "tracking write") access.writeErrors.add(access.prefNames.PREF_TOOL_ACCESS_STATE);
      const { api, state } = loadServerLifecycle({ accessRuntime: access });
      assert.equal((await api.start()).success, true);
      assert.equal(state.servers[0].listening, true);
      const config = await api.getToolAccessConfig();
      const trackingFailure = failure === "tracking write";
      assert.equal(config.mode, trackingFailure ? "restricted" : "error");
      assert.ok(config.tools.every(tool => tool.enabled === (trackingFailure
        ? tool.name !== "getMessage" && !Object.hasOwn(access.policy.introducedIn, tool.name)
        : tool.undisableable)));

      // Drive the production handler registered by start(), reusing the access
      // harness's request/response pattern and this fixture's stream boundary.
      const response = await new Promise(resolve => {
        const body = JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/list" });
        let result = "";
        state.servers[0].handlers.get("/")({
          method: "POST",
          getHeader: name => name === "Authorization" ? `Bearer ${state.connection.token}` : String(Buffer.byteLength(body)),
          bodyInputStream: { bytes: body, available: () => body.length, channel: { bytesRead: 0, readSizes: [] } },
        }, {
          processAsync() {}, setStatusLine() {}, setHeader() {},
          write(value) { result += value; }, finish() { resolve(JSON.parse(result)); },
        });
      });
      assert.deepEqual(response.result.tools.map(tool => tool.name).sort(), Array.from(config.tools).filter(tool => tool.enabled).map(tool => tool.name).sort());
      if (failure.startsWith("corrupt")) {
        assert.deepEqual(access.writes, []);
        const pref = failure === "corrupt tracking" ? "PREF_TOOL_ACCESS_STATE" : "PREF_DISABLED_TOOLS";
        assert.equal(access.values.get(access.prefNames[pref]), "not-json");
      }
    });
  }
});

describe("listener ownership classification", () => {
  const { sandbox } = loadServerLifecycle();
  for (const [name, evidence, expected] of [
    ["handler observation survives response failure", { observed: true }, "ours"],
    ["handler observation wins over a proxy", { observed: true, status: 200, unproxied: false }, "ours"],
    ...[201, 302, 403, 404, 500].map(status => [`complete unproxied status ${status}`, { status, unproxied: true, completeHead: true }, "foreign"]),
    ["possibly synthesized 200", { status: 200, unproxied: true, completeHead: true }, "inconclusive"],
    ["missing head completion", { status: 404, unproxied: true }, "inconclusive"],
    ["incomplete head", { status: 404, unproxied: true, completeHead: false }, "inconclusive"],
    ["proxy", { status: 404, unproxied: false, completeHead: true }, "inconclusive"],
    ["unknown proxy state", { status: 404, completeHead: true }, "inconclusive"],
    ...[undefined, null, 0, 99, 600, "200", NaN].map(status => [`missing/invalid status ${status}`, { status, unproxied: true, completeHead: true }, "inconclusive"]),
  ]) {
    it(name, () => assert.equal(sandbox.classifyListenerProbe(evidence), expected));
  }
});

describe("post-bind listener ownership", () => {
  for (const listenAll of [false, true]) {
    it(`stops a foreign listener before rebinding (${listenAll ? "listenAll" : "loopback"})`, async () => {
      const { api, state } = loadServerLifecycle({
        listenAll, holdStops: true, probeResult: port => port === 8765 ? "foreign" : "ours",
      });
      const pending = api.start();
      await new Promise(setImmediate);
      assert.deepEqual(state.servers[0].binds, [8765]);
      assert.equal(state.pendingStops.length, 1);
      assert.equal(state.connection, null);
      assert.equal((await api.getServerInfo()).running, false);
      state.pendingStops.shift()();
      assert.equal((await pending).port, 8766);
      assert.deepEqual(state.servers[0].binds, [8765, 8766]);
      assert.deepEqual(state.servers[0].bindModes, Array(2).fill(listenAll ? "listenAll" : "loopback"));
      assert.equal(state.servers[0].stops, 1);
      assert.deepEqual(state.connectionWrites.map(info => info.port), [8766]);
      assert.deepEqual([...state.servers[0].handlers.keys()], ["/"]);
      assert.notEqual(state.channels[0].url.pathname, state.channels[1].url.pathname);
    });
  }

  for (const oversized of [false, true]) {
    it(`publishes and refreshes profile metadata only for the final listener (${oversized ? "oversized root" : "normal roots"})`, async () => {
      const roots = {
        ProfD: `C:/profiles/${oversized ? "郵".repeat(1600) : "Łukasz"}`,
        DefProfRt: "D:/Thunderbird/Profiles",
      };
      // Immutable profile files follow the directory-service mock pattern in
      // write-connection-info-windows.test.cjs; real profile helpers run below.
      const profileDirs = Object.fromEntries(Object.entries(roots).map(([key, root]) => [key, {
        path: root, target: root, parent: null,
        clone() { return { ...this }; },
        normalize() {},
        isReadable() { return true; },
      }]));
      const { api, state } = loadServerLifecycle({
        profileDirs, holdProbes: true, holdStops: true,
        probeResult: port => port === 8765 ? "foreign" : "ours",
      });
      const pending = api.start();
      assert.deepEqual(state.connectionWrites, []);
      state.channels[0].respond();
      await new Promise(setImmediate);
      assert.deepEqual(state.servers[0].binds, [8765]);
      assert.equal(state.pendingStops.length, 1);
      assert.deepEqual(state.connectionWrites, [], "do not publish the rejected port while stopping");
      state.pendingStops.shift()();
      await new Promise(setImmediate);
      assert.deepEqual(state.servers[0].binds, [8765, 8766]);
      assert.deepEqual(state.connectionWrites, [], "do not publish before the final probe settles");
      assert.equal(state.timers.filter(timer => timer.type === REFRESH_TIMER_TYPE).length, 0);
      state.channels[1].respond();
      assert.equal((await pending).port, 8766);

      const expected = {
        port: 8766, token: "a".repeat(64), pid: 4242,
        profileRoots: oversized ? [roots.DefProfRt] : Object.values(roots),
      };
      assert.deepEqual(state.connectionWrites, [expected]);
      assert.deepEqual(state.connection, expected);
      const refresh = state.timers.find(timer => timer.type === REFRESH_TIMER_TYPE);
      refresh.callback();
      assert.equal(state.connectionWrites.length, 1, "matching metadata must not cause a refresh write");

      for (const stale of [
        null,
        { ...expected, port: 8765 },
        { ...expected, profileRoots: undefined },
        { ...expected, profileRoots: ["C:/stale-profile"] },
      ]) {
        const writes = state.connectionWrites.length;
        state.connection = stale;
        refresh.callback();
        assert.equal(state.connectionWrites.length, writes + 1);
        assert.deepEqual(state.connection, expected);
      }
      assert.deepEqual(state.connectionWrites, Array(5).fill(expected));
      assert.ok(state.connectionWrites.every(info => Buffer.byteLength(JSON.stringify(info)) <= 4096));
      refresh.callback();
      assert.equal(state.connectionWrites.length, 5, "bounded metadata remains stable after refresh");
    });
  }

  // These fixtures model Gecko's normalized metadata and activity delivery,
  // not a second implementation of its HTTP parser. Close() repairs missing
  // terminators without emitting RESPONSE_HEADER; a version-only line becomes
  // 200 even when a real header terminator was received by ProcessData().
  for (const [name, responseStatus, responseHeadReceived] of [
    ["version-only status line: HTTP/1.1\\r\\n\\r\\n", 200, true],
    ["EOF-truncated status line: HTTP/1.1 404 Not Found", 404, false],
    ["EOF-truncated response head: HTTP/1.1 404 Not Found\\r\\n", 404, false],
    ["EOF-truncated version-only line: HTTP/1.1", 200, false],
  ]) {
    it(`keeps the port inconclusive after a ${name}`, async () => {
      const { api, state } = loadServerLifecycle({
        probeResult: port => port === 8765 ? "foreign" : "ours",
        responseStatus, responseHeadReceived, bodyChunks: [],
      });
      const pending = api.start();
      await new Promise(setImmediate);
      if (!responseHeadReceived && responseStatus !== 200) {
        assert.equal(state.connection, null);
        state.timers[0].callback();
      }
      const result = await pending;
      assert.equal(result.success, true);
      assert.equal(result.port, 8765);
      assert.deepEqual(state.servers[0].binds, [8765]);
      assert.equal(state.servers[0].stops, 0);
      assert.deepEqual(state.connectionWrites.map(info => info.port), [8765]);
      assert.ok(state.logs.some(args => String(args[0]).includes("inconclusive on port 8765")));
      assert.equal(state.activityObservers.size, 0);
    });
  }

  for (const responseStatus of [201, 302, 403, 404, 500]) {
    it(`advances after a complete ${responseStatus} response head without a body`, async () => {
      const { api, state } = loadServerLifecycle({
        probeResult: port => port === 8765 ? "foreign" : "ours",
        responseStatus, bodyChunks: [],
      });
      assert.equal((await api.start()).port, 8766);
      assert.deepEqual(state.servers[0].binds, [8765, 8766]);
      assert.equal(state.servers[0].stops, 1);
      assert.equal(state.activityObservers.size, 0);
    });
  }

  it("keeps an unobserved complete 200 inconclusive because its status may be synthesized", async () => {
    const { api, state } = loadServerLifecycle({
      probeResult: port => port === 8765 ? "foreign" : "ours", responseStatus: 200,
    });
    assert.equal((await api.start()).port, 8765);
    assert.deepEqual(state.servers[0].binds, [8765]);
    assert.ok(state.channels[0].bytesRead > 0, "a body does not prove that a status code was present");
    assert.ok(state.logs.some(args => String(args[0]).includes("inconclusive on port 8765")));
  });

  it("requires the head event to belong to this channel and response", async () => {
    const { api, state } = loadServerLifecycle({ holdProbes: true, probeResult: "foreign", responseHeadReceived: false });
    const pending = api.start();
    const observer = [...state.activityObservers][0];
    const channel = state.channels[0];
    const type = HTTP_ACTIVITY.ACTIVITY_TYPE_HTTP_TRANSACTION;
    const subtype = HTTP_ACTIVITY.ACTIVITY_SUBTYPE_RESPONSE_HEADER;
    observer.observeActivity({ QueryInterface() { return this; } }, type, subtype);
    observer.observeActivity({ QueryInterface() { throw new Error("not a channel"); } }, type, subtype);
    observer.observeActivity(channel, type + 1, subtype);
    observer.observeActivity(channel, type, subtype + 1);
    channel.respond();
    assert.equal(state.connection, null);
    state.timers[0].callback();
    assert.equal((await pending).port, 8765);
    assert.ok(state.logs.some(args => String(args[0]).includes("inconclusive on port 8765")));
    assert.equal(state.activityObservers.size, 0);
    // A delayed event cannot revise a settled probe or its published port.
    observer.observeActivity(channel, type, subtype);
    assert.deepEqual(state.connectionWrites.map(info => info.port), [8765]);
  });

  it("accepts a matching head event delivered after onStartRequest", async () => {
    const { api, state } = loadServerLifecycle({ holdProbes: true,
      probeResult: port => port === 8765 ? "foreign" : "ours" });
    const pending = api.start();
    const channel = state.channels[0];
    channel.listener.onStartRequest(channel);
    channel.reportResponseHead();
    channel.listener.onStopRequest(channel, 0);
    await new Promise(setImmediate);
    state.channels[1].respond();
    assert.equal((await pending).port, 8766);
    assert.equal(state.activityObservers.size, 0);
  });

  it("allows matching head activity to arrive after onStopRequest within the deadline", async () => {
    const { api, state } = loadServerLifecycle({ holdProbes: true,
      probeResult: port => port === 8765 ? "foreign" : "ours" });
    const pending = api.start();
    const channel = state.channels[0];
    channel.listener.onStartRequest(channel);
    channel.listener.onStopRequest(channel, 0);
    assert.equal(state.connection, null);
    assert.equal(state.activityObservers.size, 1);
    channel.reportResponseHead();
    assert.equal(state.activityObservers.size, 1, "do not remove observers during activity delivery");
    state.timers[0].callback();
    await new Promise(setImmediate);
    state.channels[1].respond();
    assert.equal((await pending).port, 8766);
    assert.equal(state.activityObservers.size, 0);
  });

  it("accepts our handler observation without any response-head activity", async () => {
    const { api, state } = loadServerLifecycle({ responseHeadReceived: false });
    assert.equal((await api.start()).port, 8765);
    assert.deepEqual(state.servers[0].binds, [8765]);
    assert.ok(!state.logs.some(args => String(args[0]).includes("inconclusive")));
    assert.equal(state.activityObservers.size, 0);
  });

  it("uses anonymous, uncached, unredirected, fresh token-free requests", async () => {
    const { api, sandbox, state } = loadServerLifecycle();
    await api.start();
    const channel = state.channels[0];
    const flags = sandbox.Ci.nsIRequest;
    assert.equal(channel.loadFlags, flags.LOAD_ANONYMOUS | flags.LOAD_BYPASS_CACHE |
      flags.INHIBIT_CACHING | flags.LOAD_FRESH_CONNECTION | sandbox.Ci.nsIChannel.LOAD_INITIAL_DOCUMENT_URI);
    assert.equal(channel.redirectionLimit, 0);
    assert.equal(channel.requestMethod, "GET");
    assert.equal(channel.loadUsingSystemPrincipal, true);
    assert.deepEqual(channel.headers, { Connection: "close" });
    assert.doesNotMatch(channel.uri, new RegExp(state.connection.token));
    assert.match(channel.url.pathname, /^\/thunderbird-mcp-probe\/[0-9a-f]{64}$/);
    assert.equal(state.timers[0].delay, sandbox.probeSettings.timeoutMs);
    assert.equal(state.timers[0].cancelled, true);
    assert.equal(channel.cancelled, true);
  });

  for (const overrides of [
    { probeResult: "inconclusive" },
    { probeResult: "foreign", proxyInfo: { type: "http" } },
    { probeResult: "foreign", proxyInfo: { type: "socks" } },
    { probeResult: "foreign", responseVersion: [0, 9] },
    { probeResult: "foreign", responseVersion: [0, 0] },
  ]) {
    it(`keeps an inconclusive port: ${JSON.stringify(overrides)}`, async () => {
      const { api, state } = loadServerLifecycle(overrides);
      assert.equal((await api.start()).port, 8765);
      assert.deepEqual(state.servers[0].binds, [8765]);
      assert.ok(state.logs.some(args => String(args[0]).includes("inconclusive")));
    });
  }

  for (const property of ["nonceError", "handlerError", "channelError", "queryError", "asyncOpenError",
    "probeTimerError", "probeTimerCancelError", "versionError", "statusError", "proxyError", "readError", "cancelError",
    "activityServiceError", "activityAddError", "activityRemoveError"]) {
    it(`keeps the bound port after ${property}`, async () => {
      const { api, state } = loadServerLifecycle({ probeResult: "foreign", [property]: new Error(property) });
      assert.equal((await api.start()).success, true);
      assert.deepEqual(state.servers[0].binds, [8765]);
      assert.equal(state.servers[0].stops, 0);
      assert.equal(state.connection.port, 8765);
      assert.deepEqual([...state.servers[0].handlers.keys()], ["/"]);
      assert.ok(state.logs.some(args => String(args[0]).includes("inconclusive")));
      assert.equal(state.activityObservers.size, 0);
    });
  }

  it("contains timer creation and unexpected probe-helper exceptions", async () => {
    for (const helperThrows of [false, true]) {
      const { api, sandbox, state } = loadServerLifecycle({ timerCreateFailures: 1 });
      if (helperThrows) {
        state.timerCreateFailures = 0;
        sandbox.probeListenerOwnership = () => { throw new Error("unexpected probe failure"); };
      }
      assert.equal((await api.start()).success, true);
      assert.deepEqual(state.servers[0].binds, [8765]);
      assert.equal(state.servers[0].stops, 0);
    }
  });

  it("keeps the listening port when stopping a foreign listener throws", async () => {
    const { api, state } = loadServerLifecycle({ probeResult: "foreign", stopError: new Error("stop failed") });
    assert.equal((await api.start()).success, true);
    assert.deepEqual(state.servers[0].binds, [8765]);
    assert.equal(state.servers[0].listening, true);
    assert.equal(state.connection.port, 8765);
    assert.ok(state.logs.some(args => String(args[0]).includes("could not stop the listener")));
  });

  it("times out without requiring channel callbacks and ignores late delivery", async () => {
    const { api, state } = loadServerLifecycle({ holdProbes: true });
    const pending = api.start();
    const channel = state.channels[0];
    state.timers[0].callback();
    assert.equal((await pending).port, 8765);
    const connection = state.connection;
    channel.respond();
    channel.deliver("late data");
    assert.equal(channel.bytesRead, 0);
    assert.strictEqual(state.connection, connection);
    assert.equal(state.servers[0].stops, 0);
    assert.equal(state.activityObservers.size, 0);
  });

  it("does not let an earlier attempt's handler or events claim the next port", async () => {
    const { api, state } = loadServerLifecycle({ holdProbes: true, probeResult: "foreign" });
    const pending = api.start();
    const first = state.channels[0];
    const firstObserver = [...state.activityObservers][0];
    first.respond();
    await new Promise(setImmediate);
    assert.equal(state.channels.length, 2);
    const secondObserver = [...state.activityObservers][0];
    firstObserver.observeActivity(state.channels[1], HTTP_ACTIVITY.ACTIVITY_TYPE_HTTP_TRANSACTION,
      HTTP_ACTIVITY.ACTIVITY_SUBTYPE_RESPONSE_HEADER);
    secondObserver.observeActivity(first, HTTP_ACTIVITY.ACTIVITY_TYPE_HTTP_TRANSACTION,
      HTTP_ACTIVITY.ACTIVITY_SUBTYPE_RESPONSE_HEADER);
    first.handler({ path: first.url.pathname, method: "GET" }, { setStatusLine() {}, setHeader() {}, write() {} });
    first.listener.onStopRequest(first, 0);
    state.timers[0].callback();
    assert.equal(state.connection, null);
    // No response from the second port: it must remain inconclusive.
    state.timers[1].callback();
    assert.equal((await pending).port, 8766);
    assert.ok(state.logs.some(args => String(args[0]).includes("inconclusive on port 8766")));
  });

  for (const headers of [false, true]) {
    it(`caps ${headers ? "headers before body" : "incremental body reads"} and cancels the channel`, async () => {
      const { api, sandbox, state } = loadServerLifecycle({
        probeResult: port => port === 8765 ? "foreign" : "ours",
        ...(headers ? { responseHeaders: { "X-Pad": "a".repeat(8192) } } : { bodyChunks: ["a".repeat(3000), "b".repeat(8192)] }),
      });
      assert.equal((await api.start()).port, 8766);
      assert.equal(state.channels[0].cancelled, true);
      assert.ok(state.channels[0].bytesRead <= sandbox.probeSettings.maxBytes);
      if (headers) assert.equal(state.channels[0].bytesRead, 0);
      else assert.equal(state.channels[0].readSizes.length, 2);
    });
  }

  for (const initialBusy of [0, 2]) {
    it(`falls back to the first bound port when all probes are foreign (${initialBusy} initial busy ports)`, async () => {
      const { api, state } = loadServerLifecycle({
        probeResult: "foreign", bindError: port => port < 8765 + initialBusy ? new Error("busy") : null,
      });
      const result = await api.start();
      assert.equal(result.success, true);
      assert.equal(result.port, 8765 + initialBusy);
      assert.equal(state.servers[0].listening, true);
      assert.equal(state.channels.length, 10 - initialBusy);
      assert.deepEqual(state.connectionWrites.map(info => info.port), [result.port]);
      assert.ok(state.logs.some(args => String(args[0]).includes("normal port fallback")));
    });
  }

  it("uses normal bind fallback if the first port cannot be rebound", async () => {
    const { api, state } = loadServerLifecycle({
      probeResult: "foreign",
      bindError: (port, attempt) => attempt > 10 && port === 8765 ? new Error("busy now") : null,
    });
    assert.equal((await api.start()).port, 8766);
    assert.deepEqual(state.servers[0].binds.slice(-2), [8765, 8766]);
    assert.equal(state.channels.length, 10);
  });

  it("recovers from a mix of foreign responses and native bind failures", async () => {
    const { api, state } = loadServerLifecycle({
      probeResult: "foreign", bindError: port => port === 8765 ? null : new Error("busy"),
    });
    assert.equal((await api.start()).port, 8765);
    assert.equal(state.channels.length, 1);
    assert.deepEqual(state.connectionWrites.map(info => info.port), [8765]);
  });

  it("retains the existing failure when no port can be rebound", async () => {
    const { api, state } = loadServerLifecycle({
      probeResult: "foreign", bindError: (_port, attempt) => attempt > 10 ? new Error("busy now") : null,
    });
    const result = await api.start();
    assert.equal(result.success, false);
    assert.match(result.error, /Could not bind to any port in range 8765-8774.*busy now/);
    assert.equal(state.servers[0].listening, false);
    assert.equal(state.connection, null);
  });

  it("cancels startup during a probe without publishing after shutdown", async () => {
    const { api, sandbox, state } = loadServerLifecycle({ holdProbes: true });
    const pending = api.start();
    sandbox.shutdown();
    const result = await pending;
    assert.equal(result.success, false);
    assert.match(result.error, /cancelled/);
    assert.equal(state.channels[0].cancelled, true);
    assert.equal(state.servers[0].listening, false);
    assert.equal(state.connectionWrites.length, 0);
    assert.equal(state.activityObservers.size, 0);
    assert.equal((await api.getServerInfo()).running, false);
    state.channels[0].respond();
    assert.equal((await api.getServerInfo()).running, false);
    assert.deepEqual([...state.servers[0].handlers.keys()], ["/"]);
  });

  it("does not rebind after shutdown while a foreign listener is stopping", async () => {
    const { api, sandbox, state } = loadServerLifecycle({ probeResult: "foreign", holdStops: true });
    const pending = api.start();
    await new Promise(setImmediate);
    sandbox.shutdown();
    state.pendingStops.shift()();
    assert.equal((await pending).success, false);
    assert.deepEqual(state.servers[0].binds, [8765]);
    assert.equal(state.servers[0].stops, 1);
    assert.equal(state.connection, null);
  });

  it("cancels queued settings restarts when shutdown interrupts a probe", async () => {
    const { api, sandbox, state } = loadServerLifecycle({ holdProbes: true });
    const optionsApi = sandbox.getAPI();
    const pending = api.start();
    const firstSetter = api.setListenAll(false);
    await new Promise(setImmediate);
    const secondSetter = optionsApi.setListenAll(true);
    sandbox.shutdown();
    for (const result of await Promise.all([pending, firstSetter, secondSetter, api.start(), optionsApi.start()])) {
      assert.equal(result.success, false);
      assert.match(result.error, /cancelled/);
    }
    assert.equal(state.attempts, 1);
    assert.equal(state.servers[0].listening, false);
    assert.equal(state.connectionWrites.length, 0);
    assert.equal(sandbox.__tbMcpRestartPromise, null);
  });

  it("cancels queued settings restarts during listener teardown", async () => {
    const { api, sandbox, state } = loadServerLifecycle({ holdStops: true });
    assert.equal((await api.start()).success, true);
    const firstSetter = api.setListenAll(false);
    await new Promise(setImmediate);
    const secondSetter = api.setListenAll(true);
    sandbox.shutdown();
    state.pendingStops.shift()();
    for (const result of await Promise.all([firstSetter, secondSetter])) {
      assert.equal(result.success, false);
      assert.match(result.error, /cancelled/);
    }
    assert.equal(state.attempts, 1);
    assert.equal(state.servers[0].listening, false);
    assert.equal(state.connectionWrites.length, 1, "shutdown must not republish connection info");
    assert.equal(state.connection, null);
  });
});

describe("host.docker.internal Host alias registration", () => {
  it("registers the alias on the bound port in loopback mode", async () => {
    const { api, state } = loadServerLifecycle();
    assert.equal((await api.start()).port, 8765);
    assert.deepEqual(state.servers[0].identityAdds, [
      { scheme: "http", host: "host.docker.internal", port: 8765 },
    ]);
  });

  it("registers the alias only for the final port after rebinding", async () => {
    const { api, state } = loadServerLifecycle({
      probeResult: port => port === 8765 ? "foreign" : "ours",
    });
    assert.equal((await api.start()).port, 8766);
    assert.deepEqual(state.servers[0].binds, [8765, 8766]);
    assert.deepEqual(state.servers[0].identityAdds, [
      { scheme: "http", host: "host.docker.internal", port: 8766 },
    ]);
  });

  it("does not register the alias in listenAll mode", async () => {
    const { api, state } = loadServerLifecycle({ listenAll: true });
    assert.equal((await api.start()).port, 8765);
    assert.deepEqual(state.servers[0].identityAdds, []);
  });

  it("keeps startup successful when the identity API rejects the alias", async () => {
    const { api, state } = loadServerLifecycle({ identityAddError: new Error("identity unavailable") });
    assert.equal((await api.start()).success, true);
    assert.deepEqual(state.servers[0].identityAdds, []);
    assert.equal(state.connection.port, 8765);
    assert.ok(state.logs.some(args => String(args[0]).includes("host.docker.internal")));
  });
});

describe("stable auth token option", () => {
  it("rebinds the running listener so a new token takes effect immediately", async () => {
    const { api, state } = loadServerLifecycle();
    assert.equal((await api.start()).success, true);
    const result = await api.setStableAuthToken("b".repeat(64));
    assert.equal(result.success, true);
    assert.equal(result.stableAuthToken, "b".repeat(64));
    assert.deepEqual(state.connectionWrites.map(info => info.token), ["a".repeat(64), "b".repeat(64)]);
    assert.equal(state.servers[0].stops, 1);
    assert.equal(state.servers.length, 2);
    assert.equal(state.servers[1].listening, true);
    assert.deepEqual(state.servers[1].binds, [8765]);
    assert.equal((await api.getStableAuthToken()).stableAuthToken, "b".repeat(64));
  });

  it("keeps the listener when the saved token is already effective", async () => {
    const { api, state } = loadServerLifecycle();
    assert.equal((await api.start()).success, true);
    const result = await api.setStableAuthToken("a".repeat(64));
    assert.equal(result.success, true);
    assert.equal(result.stableAuthToken, "a".repeat(64));
    assert.equal(state.connectionWrites.length, 1);
    assert.equal(state.servers[0].stops, 0);
    assert.equal(state.servers.length, 1);
  });

  it("rejects malformed tokens before touching the pref or the listener", async () => {
    const { api, state } = loadServerLifecycle();
    assert.equal((await api.start()).success, true);
    const result = await api.setStableAuthToken("Not-A-Token");
    assert.match(result.error, /64 lowercase hex/);
    assert.equal(state.connectionWrites.length, 1);
    assert.equal(state.servers[0].stops, 0);
    assert.equal((await api.getStableAuthToken()).stableAuthToken, "a".repeat(64));
  });

  it("falls back to a session token on the listener restarted after clearing the stable token", async () => {
    const { api, state } = loadServerLifecycle();
    await api.start();
    await api.setStableAuthToken("b".repeat(64));
    const cleared = await api.setStableAuthToken("");
    assert.equal(cleared.success, true);
    assert.equal(cleared.stableAuthToken, "");
    const tokens = state.connectionWrites.map(info => info.token);
    assert.equal(tokens.length, 3);
    assert.equal(tokens[0], "a".repeat(64));
    assert.equal(tokens[1], "b".repeat(64));
    assert.match(tokens[2], /^[0-9a-f]{64}$/);
    assert.notEqual(tokens[2], "a".repeat(64));
    assert.notEqual(tokens[2], "b".repeat(64));
  });
});

describe("probe through S1's real httpd parser", () => {
  it("moves to 8766 when another add-on server answers the probe with its default 404", async () => {
    // The second server has only the harness's own "/" handler. Our probe
    // handler belongs to the lifecycle server and is never registered here.
    const other = createHttpdHarness();
    other.server.identity.add("http", "127.0.0.1", 8765);
    const { api, state } = loadServerLifecycle({ httpd: other, holdProbes: true, holdStops: true });
    const pending = api.start();
    const server = state.servers[0];
    const firstProbe = state.channels[0];
    assert.deepEqual(server.binds, [8765]);
    assert.equal(server.listening, true);
    assert.equal(server.handlers.has(firstProbe.url.pathname), true);
    assert.match(firstProbe.url.pathname, /^\/thunderbird-mcp-probe\/[0-9a-f]{64}$/);

    firstProbe.respond();
    await new Promise(setImmediate);
    assert.match(firstProbe.exchange.wire, /^HTTP\/1\.1 404 Not Found\r\n/);
    assert.ok(firstProbe.exchange.wire.includes("\r\n\r\n"), "the default 404 has a complete response head");
    assert.deepEqual(firstProbe.exchange.errors, [], "the probe passes S1's request parser");
    assert.equal(other.requests.length, 0, "the random probe path does not reach the other server's root handler");
    assert.ok(state.logs.some(args => String(args[0]).includes("port 8765 is answered by another listener")));
    assert.equal(server.stops, 1);
    assert.equal(state.pendingStops.length, 1);
    assert.deepEqual(server.binds, [8765], "rebind must wait for stop completion");
    assert.equal(state.channels.length, 1);
    assert.equal(state.connection, null);
    assert.deepEqual(state.connectionWrites, []);
    assert.equal((await api.getServerInfo()).running, false);

    // After the first listener finishes stopping, the next port reaches our
    // production probe handler through the existing lifecycle channel stub.
    state.httpd = null;
    state.pendingStops.shift()();
    await new Promise(setImmediate);
    assert.deepEqual(server.binds, [8765, 8766]);
    assert.equal(state.channels.length, 2);
    assert.equal(state.channels[1].url.port, "8766");
    assert.equal(state.connection, null, "do not publish before the final probe completes");
    state.channels[1].respond();
    const result = await pending;
    assert.equal(result.success, true);
    assert.equal(result.port, 8766);
    assert.equal(server.listening, true);
    assert.equal(server.port, 8766);
    assert.equal(server.stops, 1);
    assert.deepEqual(state.connectionWrites.map(info => info.port), [8766]);
    assert.equal(state.connection.port, 8766);
    assert.ok(!state.logs.some(args => String(args[0]).includes("inconclusive")));
    for (const channel of state.channels) {
      assert.deepEqual(channel.headers, { Connection: "close" });
      assert.equal(channel.uri.includes(state.connection.token), false);
    }
    assert.doesNotMatch(firstProbe.requestWire, /Authorization|Bearer/i);
    assert.equal(firstProbe.requestWire.includes(state.connection.token), false);
    assert.equal(other.socket.closed, false, "only our listener is stopped");
    assert.match(other.request("GET / HTTP/1.1\r\nHost: 127.0.0.1:8765\r\n\r\n").wire, /^HTTP\/1\.1 200 OK\r\n/);
  });

  for (const mode of ["loopback", "listenAll"]) {
    it(`reaches the production probe handler without authentication (${mode})`, async () => {
      const httpd = createHttpdHarness({ mode });
      const { sandbox, state, NetUtil } = loadServerLifecycle({ httpd });
      const result = await sandbox.probeListenerOwnership({
        server: httpd.server, port: HTTPD_PORT, NetUtil,
        generateNonce: sandbox.generateAuthToken, startup: {},
      });
      assert.equal(result, "ours");
      assert.match(state.channels[0].exchange.wire, /^HTTP\/1\.1 200 OK\r\n/);
      assert.match(state.channels[0].exchange.wire, /\r\n\r\n[0-9a-f]{64}$/);
      assert.doesNotMatch(state.channels[0].requestWire, /Authorization|Bearer|Transfer-Encoding/i);
      assert.equal(httpd.socket.closed, false);
      assert.equal(httpd.server._doQuit, false);
      assert.match(httpd.request(`GET ${state.channels[0].url.pathname} HTTP/1.1\r\nHost: 127.0.0.1:${HTTPD_PORT}\r\n\r\n`).wire, /^HTTP\/1\.1 404 /);
    });
  }
});

describe("options server status", () => {
  it("shows the non-default port notice and clears it as status changes", async () => {
    let info = { running: true, port: 8766 };
    const ui = loadServerStatus({ getServerInfo: async () => {
      if (info instanceof Error) throw info;
      return info;
    } });
    for (const next of [
      { running: true, port: 8765 },
      { running: false, port: 8766 },
      { running: false, lastError: "Failed to start" },
      { running: true, port: null },
      new Error("API unavailable"),
    ]) {
      info = { running: true, port: 8766 };
      await ui.loadServerInfo();
      assert.equal(ui.serverPortNotice.hidden, false);
      assert.match(ui.serverPortNotice.textContent, /8766.*8765/);
      assert.match(ui.serverPortNotice.textContent, /connection\.json.*stdio bridge/);
      info = next;
      await ui.loadServerInfo();
      assert.equal(ui.serverPortNotice.hidden, true);
      assert.equal(ui.serverPortNotice.textContent, "");
    }
  });

  it("shows stopped without a startup error before any attempt", async () => {
    const { api } = loadServerLifecycle();
    const ui = loadServerStatus(api);
    await ui.loadServerInfo();
    assert.equal(ui.statusText.textContent, "Not running");
    assert.equal(ui.retryServerBtn.hidden, true);
  });

  it("shows the stored error on reopening options and refreshes status after Retry", async () => {
    const { api, state } = loadServerLifecycle({ importError: new Error("Cannot load <httpd>") });
    const failed = await api.start();
    const ui = loadServerStatus(api);
    await ui.loadServerInfo();
    assert.equal(ui.statusText.textContent, "Failed to start: " + failed.error);
    assert.equal(ui.statusDot.className, "status-dot stopped");
    assert.equal(ui.retryServerBtn.hidden, false);
    assert.equal(ui.serverPort.textContent, "--");
    assert.equal(ui.connFile.textContent, "--");

    state.importError = null;
    const retry = ui.retryServerBtn.click();
    assert.equal(ui.retryServerBtn.disabled, true);
    assert.equal(ui.statusText.textContent, "Starting...");
    await retry;
    assert.equal(state.attempts, 2);
    assert.equal(ui.retryServerBtn.disabled, false);
    assert.equal(ui.retryServerBtn.hidden, true);
    assert.equal(ui.statusText.textContent, "Running");
    assert.equal(ui.statusDot.className, "status-dot running");
    assert.equal(ui.serverPort.textContent, 8765);
    assert.equal(ui.connFile.textContent, "/mock-tmp/thunderbird-mcp/connection.json");
    assert.equal((await api.getServerInfo()).lastError, null);
    assert.equal(ui.authRefreshes, 1);
  });

  it("replaces an earlier failure with the latest retry error and keeps Retry available", async () => {
    const { api, state } = loadServerLifecycle({ importError: new Error("first failure") });
    await api.start();
    const ui = loadServerStatus(api);
    await ui.loadServerInfo();
    state.importError = null;
    state.writeError = new Error("connection file is not writable");
    await ui.retryServerBtn.click();
    assert.equal((await api.getServerInfo()).lastError, "Error: connection file is not writable");
    assert.equal(ui.statusText.textContent, "Failed to start: Error: connection file is not writable");
    assert.equal(ui.retryServerBtn.hidden, false);
    assert.equal(ui.retryServerBtn.disabled, false);
    assert.equal(ui.statusDot.className, "status-dot stopped");
    assert.equal(ui.authRefreshes, 0);
  });

  it("refreshes a previously running status when a listen-all restart fails", async () => {
    let info = { running: true, port: 8765, lastError: null };
    const ui = loadServerStatus({
      getServerInfo: async () => info,
      setListenAll: async () => {
        info = { running: false, lastError: "Cannot bind to port" };
        return { error: info.lastError };
      },
    });
    await ui.loadServerInfo();
    assert.equal(ui.statusText.textContent, "Running");
    assert.equal(ui.retryServerBtn.hidden, true);
    await ui.saveListenAllBtn.click();
    assert.equal(ui.statusText.textContent, "Failed to start: Cannot bind to port");
    assert.equal(ui.retryServerBtn.hidden, false);
    assert.equal(ui.saveListenAllStatus.textContent, "Cannot bind to port");
    assert.equal(ui.saveListenAllBtn.disabled, false);
  });

  it("reenables Retry if the experiment API rejects the call", async () => {
    const ui = loadServerStatus({
      getServerInfo: async () => ({ running: false, lastError: "Previous failure" }),
      start: async () => { throw new Error("API unavailable"); },
    });
    await ui.loadServerInfo();
    await ui.retryServerBtn.click();
    assert.equal(ui.statusText.textContent, "Failed to start: API unavailable");
    assert.equal(ui.retryServerBtn.hidden, false);
    assert.equal(ui.retryServerBtn.disabled, false);
  });
});

describe("server startup lifecycle", () => {
  it("evicts a synchronous failure and allows a successful retry", async () => {
    const { api, sandbox, state } = loadServerLifecycle({ importError: new Error("import failed") });
    assert.equal((await api.getServerInfo()).running, false);
    assert.equal((await api.getServerInfo()).lastError, null);

    const result = await api.start();
    assert.equal(result.success, false);
    assert.match(result.error, /import failed/);
    assert.equal(sandbox.__tbMcpStartPromise, null);
    assert.equal((await api.getServerInfo()).running, false);
    assert.equal((await api.getServerInfo()).lastError, result.error);
    assert.equal(state.servers.length, 0);

    state.importError = null;
    assert.equal((await api.start()).success, true);
    assert.equal(state.attempts, 2);
    assert.equal((await api.getServerInfo()).running, true);
    assert.equal((await api.getServerInfo()).port, 8765);
    assert.equal((await api.getServerInfo()).lastError, null);
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
    assert.equal(state.timers.filter(timer => timer.type === REFRESH_TIMER_TYPE).length, 1);
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
    await new Promise(setImmediate);
    assert.equal((await api.getServerInfo()).running, false);
    assert.equal(state.timers.find(timer => timer.type === REFRESH_TIMER_TYPE).cancelled, true);
    assert.equal(sandbox.__tbMcpConnectionInfoRefreshTimer, null);

    state.pendingStops.shift()();
    assert.equal((await start).success, false);
    assert.equal(state.connection, null);
    assert.equal(state.servers[0].listening, false);

    state.timerError = null;
    assert.equal((await api.start()).success, true);
    const refreshTimers = state.timers.filter(timer => timer.type === REFRESH_TIMER_TYPE);
    assert.equal(refreshTimers.length, 2);
    assert.equal(refreshTimers[1].cancelled, false);
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
    assert.equal((await api.getServerInfo()).lastError, String(rejection));
    assert.equal(state.servers[0].listening, false);
    assert.equal(state.connection, null);
    assert.equal((await api.getServerInfo()).running, false);

    state.writeError = null;
    assert.equal((await api.start()).success, true);
    assert.equal((await api.getServerInfo()).lastError, null);
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
      await new Promise(setImmediate);
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
      await new Promise(setImmediate);
      const newer = Promise.resolve({ success: true, port: 8770 });
      const newerConnection = { port: 8770, token: "new", pid: 4242 };
      sandbox.__tbMcpStartPromise = newer;
      sandbox.__tbMcpLastStartError = null;
      state.connection = newerConnection;
      state.pendingStops.shift()();

      if (rejects) await assert.rejects(oldStart, error => error === rejection);
      else assert.equal((await oldStart).success, false);
      assert.strictEqual(sandbox.__tbMcpStartPromise, newer);
      assert.equal((await api.getServerInfo()).lastError, null);
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
    state.timers.find(timer => timer.type === REFRESH_TIMER_TYPE).callback();
    assert.deepEqual(state.connection, connection);
    assert.equal(state.attempts, 1);
    assert.equal((await api.getServerInfo()).running, true);
  });
});
