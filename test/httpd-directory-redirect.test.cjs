"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const os = require("node:os");

// Mock Firefox components/globals required for basic httpd logic testing
if (typeof globalThis.Components === "undefined") {
  globalThis.Components = {
    Constructor: () => function () {},
    Exception: function (msg, result) {
      this.message = msg;
      this.result = result;
    },
  };
  globalThis.Cr = {
    NS_ERROR_INVALID_ARG: 0x80070057,
    NS_ERROR_NOT_AVAILABLE: 0x80040111,
  };
  globalThis.Ci = {};
  globalThis.Cc = {};
  globalThis.ChromeUtils = {
    generateQI: () => function () {},
  };
}

describe("httpd directory trailing slash redirect", () => {
  it("redirects directory request missing trailing slash with 301 Moved Permanently", async () => {
    const httpdModule = await import("../extension/httpd.sys.mjs");
    const { HttpServer } = httpdModule;

    const server = new HttpServer();
    const handler = server._handler;

    // Create a temporary directory for mapping
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "httpd-test-"));

    try {
      const mockDirFile = {
        path: tmpDir,
        parent: null,
        exists: () => true,
        isDirectory: () => true,
        clone() {
          return this;
        },
      };

      handler.registerDirectory("/testdir/", mockDirFile);

      // 1. Request without trailing slash: /testdir
      let statusLineSet = null;
      let headersSet = {};

      const mockResponse1 = {
        setStatusLine(version, code, description) {
          statusLineSet = { version, code, description };
        },
        setHeader(name, value) {
          headersSet[name] = value;
        },
      };

      const mockMetadata1 = {
        path: "/testdir",
        httpVersion: "1.1",
        queryString: "",
      };

      handler._handleDefault(mockMetadata1, mockResponse1);

      assert.deepEqual(statusLineSet, {
        version: "1.1",
        code: 301,
        description: "Moved Permanently",
      });
      assert.equal(headersSet["Location"], "/testdir/");

      // 2. Request without trailing slash and with query string: /testdir?foo=bar
      statusLineSet = null;
      headersSet = {};

      const mockResponse2 = {
        setStatusLine(version, code, description) {
          statusLineSet = { version, code, description };
        },
        setHeader(name, value) {
          headersSet[name] = value;
        },
      };

      const mockMetadata2 = {
        path: "/testdir",
        httpVersion: "1.1",
        queryString: "foo=bar&baz=1",
      };

      handler._handleDefault(mockMetadata2, mockResponse2);

      assert.deepEqual(statusLineSet, {
        version: "1.1",
        code: 301,
        description: "Moved Permanently",
      });
      assert.equal(headersSet["Location"], "/testdir/?foo=bar&baz=1");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });
});
