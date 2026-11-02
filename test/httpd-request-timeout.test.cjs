"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

// Mock Firefox XPCOM globals required for httpd initialization
if (typeof globalThis.Components === "undefined") {
  globalThis.Components = {
    Constructor: () => function () {},
    Exception: function (msg, result) {
      this.message = msg;
      this.result = result;
    },
  };
  globalThis.Cr = {
    NS_BINDING_ABORTED: 0x80040006,
    NS_ERROR_INVALID_ARG: 0x80070057,
    NS_ERROR_NOT_AVAILABLE: 0x80040111,
  };
  globalThis.Ci = {};
  globalThis.Cc = {};
  globalThis.ChromeUtils = {
    generateQI: () => function () {},
  };
  globalThis.Services = {
    tm: { mainThread: {}, currentThread: {} },
    io: { allowPort: () => true },
    prefs: { getIntPref: () => 6 },
  };
  globalThis.dumpn = () => {};
  globalThis.dump = () => {};
}

describe("httpd request timeout functionality", async () => {
  const httpdPath = path.resolve(__dirname, "../extension/httpd.sys.mjs");
  const source = fs.readFileSync(httpdPath, "utf8");
  const { HttpServer } = await import("../extension/httpd.sys.mjs");

  it("resolves the XXX comment in extension/httpd.sys.mjs", () => {
    assert.equal(
      source.includes("// XXX add request timeout functionality here!"),
      false,
      "The XXX comment should be resolved and removed"
    );
  });

  function createMockTransport(setTimeoutFn) {
    const mockInput = {
      QueryInterface: () => mockInput,
      asyncWait: () => {},
    };
    return {
      port: 8080,
      host: "127.0.0.1",
      openInputStream: () => mockInput,
      openOutputStream: () => ({}),
      setTimeout: setTimeoutFn,
      close: () => {},
    };
  }

  it("configures read timeout using TIMEOUT_READ_WRITE on accepted sockets", () => {
    globalThis.Ci.nsISocketTransport = { TIMEOUT_READ_WRITE: 1 };
    const server = new HttpServer();
    let recordedTimeout = null;

    const mockTrans = createMockTransport((type, val) => {
      recordedTimeout = { type, val };
    });

    server.onSocketAccepted({ port: 8080 }, mockTrans);
    assert.deepEqual(recordedTimeout, { type: 1, val: 30 });
  });

  it("configures read timeout using TIMEOUT_READ when TIMEOUT_READ_WRITE is absent", () => {
    globalThis.Ci.nsISocketTransport = { TIMEOUT_READ: 2 };
    const server = new HttpServer();
    let recordedTimeout = null;

    const mockTrans = createMockTransport((type, val) => {
      recordedTimeout = { type, val };
    });

    server.onSocketAccepted({ port: 8080 }, mockTrans);
    assert.deepEqual(recordedTimeout, { type: 2, val: 30 });
  });

  it("falls back to timeout type 1 if nsISocketTransport has neither constant", () => {
    globalThis.Ci.nsISocketTransport = {};
    const server = new HttpServer();
    let recordedTimeout = null;

    const mockTrans = createMockTransport((type, val) => {
      recordedTimeout = { type, val };
    });

    server.onSocketAccepted({ port: 8080 }, mockTrans);
    assert.deepEqual(recordedTimeout, { type: 1, val: 30 });
  });

  it("falls back to timeout type 1 if nsISocketTransport is undefined", () => {
    const saved = globalThis.Ci.nsISocketTransport;
    delete globalThis.Ci.nsISocketTransport;
    try {
      const server = new HttpServer();
      let recordedTimeout = null;

      const mockTrans = createMockTransport((type, val) => {
        recordedTimeout = { type, val };
      });

      server.onSocketAccepted({ port: 8080 }, mockTrans);
      assert.deepEqual(recordedTimeout, { type: 1, val: 30 });
    } finally {
      globalThis.Ci.nsISocketTransport = saved;
    }
  });

  it("safely handles transports without setTimeout method", () => {
    const server = new HttpServer();
    const mockTrans = createMockTransport(undefined);
    delete mockTrans.setTimeout;

    assert.doesNotThrow(() => {
      server.onSocketAccepted({ port: 8080 }, mockTrans);
    });
  });
});
