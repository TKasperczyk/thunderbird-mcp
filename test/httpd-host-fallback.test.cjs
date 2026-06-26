"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadHostHeaderIdentityHelpers() {
  const httpdPath = path.resolve(__dirname, "../extension/httpd.sys.mjs");
  const source = fs.readFileSync(httpdPath, "utf8");
  const startMarker = "// BEGIN HOST HEADER IDENTITY HELPERS";
  const endMarker = "// END HOST HEADER IDENTITY HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "host header identity helper start marker missing");
  assert.ok(end > start, "host header identity helper end marker missing");

  const HTTP_400 = { code: 400, description: "Bad Request" };
  const snippet = source.slice(start, end);
  const sandbox = { HTTP_400 };
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.resolveIdentityFromHostHeader = resolveIdentityFromHostHeader;`,
    sandbox
  );
  return {
    HTTP_400,
    resolveIdentityFromHostHeader: sandbox.resolveIdentityFromHostHeader,
  };
}

const { HTTP_400, resolveIdentityFromHostHeader } = loadHostHeaderIdentityHelpers();

function makeIdentity({
  boundHost = "localhost",
  primaryScheme = "http",
  primaryHost = "localhost",
  primaryPort = 8765,
  locations = [{ scheme: "http", host: "localhost", port: 8765 }],
} = {}) {
  const schemes = new Map();
  for (const location of locations) {
    schemes.set(`${location.host}:${location.port}`, location.scheme);
  }
  return {
    _host: boundHost,
    primaryScheme,
    primaryHost,
    primaryPort,
    getScheme(host, port) {
      return schemes.get(`${host}:${port}`) || "";
    },
  };
}

function assertBadRequest(result, reason) {
  assert.equal(result.error, HTTP_400);
  assert.equal(result.error.code, 400);
  assert.equal(result.reason, reason);
}

function assertResolved(result, expected) {
  assert.equal(result.error, undefined);
  assert.equal(result.scheme, expected.scheme);
  assert.equal(result.host, expected.host);
  assert.equal(result.port, expected.port);
}

describe("httpd Host header identity resolution", () => {
  it("keeps a recognized Host identity unchanged", () => {
    const identity = makeIdentity({
      locations: [
        { scheme: "http", host: "localhost", port: 8765 },
        { scheme: "http", host: "127.0.0.1", port: 8765 },
      ],
    });

    assertResolved(resolveIdentityFromHostHeader(identity, "127.0.0.1:8765"), {
      scheme: "http",
      host: "127.0.0.1",
      port: 8765,
    });
  });

  it("rejects an unrecognized Host outside listenAll mode", () => {
    const identity = makeIdentity();

    assertBadRequest(
      resolveIdentityFromHostHeader(identity, "proxy.example:443"),
      "unrecognized"
    );
  });

  it("falls back to the primary identity for arbitrary listenAll Host names", () => {
    const identity = makeIdentity({
      boundHost: "0.0.0.0",
      primaryHost: "localhost",
      primaryPort: 8765,
      locations: [
        { scheme: "http", host: "localhost", port: 8765 },
        { scheme: "http", host: "127.0.0.1", port: 8765 },
      ],
    });

    assertResolved(resolveIdentityFromHostHeader(identity, "reverse-proxy.example:443"), {
      scheme: "http",
      host: "localhost",
      port: 8765,
    });
  });

  it("falls back to the primary identity for port-mapped listenAll requests", () => {
    const identity = makeIdentity({
      boundHost: "0.0.0.0",
      primaryHost: "localhost",
      primaryPort: 8765,
    });

    assertResolved(resolveIdentityFromHostHeader(identity, "127.0.0.1:49152"), {
      scheme: "http",
      host: "localhost",
      port: 8765,
    });
  });

  it("still rejects malformed Host headers in listenAll mode", () => {
    const identity = makeIdentity({ boundHost: "0.0.0.0" });

    assertBadRequest(
      resolveIdentityFromHostHeader(identity, "bad host:8765"),
      "malformed"
    );
    assertBadRequest(
      resolveIdentityFromHostHeader(identity, "localhost:not-a-port"),
      "malformed"
    );
  });

  it("preserves IPv6 bracket parsing and default-port behavior", () => {
    const identity = makeIdentity({
      primaryHost: "[::1]",
      locations: [
        { scheme: "http", host: "[::1]", port: 8765 },
        { scheme: "http", host: "example.com", port: 80 },
      ],
    });

    assertResolved(resolveIdentityFromHostHeader(identity, "[::1]:8765"), {
      scheme: "http",
      host: "[::1]",
      port: 8765,
    });
    assertResolved(resolveIdentityFromHostHeader(identity, "example.com"), {
      scheme: "http",
      host: "example.com",
      port: 80,
    });
    assertResolved(resolveIdentityFromHostHeader(identity, "example.com:"), {
      scheme: "http",
      host: "example.com",
      port: 80,
    });
  });
});
