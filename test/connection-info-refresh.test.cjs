"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadConnectionInfoRefreshHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN CONNECTION INFO REFRESH HELPERS";
  const endMarker = "// END CONNECTION INFO REFRESH HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "connection info refresh helper start marker missing");
  assert.ok(end > start, "connection info refresh helper end marker missing");

  const snippet = source.slice(start, end);
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.ensureFreshConnectionInfo = ensureFreshConnectionInfo;`,
    sandbox
  );
  return {
    ensureFreshConnectionInfo: sandbox.ensureFreshConnectionInfo,
  };
}

const { ensureFreshConnectionInfo } = loadConnectionInfoRefreshHelpers();

function makeOptions(overrides) {
  return {
    port: 8765,
    token: "token",
    expectedPid: 1234,
    onCheckError() {},
    ...overrides,
  };
}

describe("connection info refresh decision", () => {
  it("does not rewrite when connection.json already matches the running server", () => {
    const result = ensureFreshConnectionInfo(makeOptions({
      readConnectionInfo: () => ({
        path: "/tmp/thunderbird-mcp/connection.json",
        data: { port: 8765, token: "token", pid: 1234 },
      }),
      writeConnectionInfo() {
        throw new Error("writeConnectionInfo should not be called");
      },
    }));

    assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
  });

  it("rewrites when connection.json is missing", () => {
    const writes = [];
    const result = ensureFreshConnectionInfo(makeOptions({
      readConnectionInfo: () => ({
        path: "/tmp/thunderbird-mcp/connection.json",
        data: null,
      }),
      writeConnectionInfo(port, token) {
        writes.push({ port, token });
        return "/tmp/thunderbird-mcp/connection.json";
      },
    }));

    assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
    assert.deepEqual(writes, [{ port: 8765, token: "token" }]);
  });

  for (const staleCase of [
    {
      name: "port changed",
      data: { port: 8766, token: "token", pid: 1234 },
    },
    {
      name: "token changed",
      data: { port: 8765, token: "old-token", pid: 1234 },
    },
    {
      name: "pid changed",
      data: { port: 8765, token: "token", pid: 4321 },
    },
  ]) {
    it(`rewrites when connection.json is stale: ${staleCase.name}`, () => {
      let writeCount = 0;
      const result = ensureFreshConnectionInfo(makeOptions({
        readConnectionInfo: () => ({
          path: "/tmp/thunderbird-mcp/connection.json",
          data: staleCase.data,
        }),
        writeConnectionInfo(port, token) {
          writeCount++;
          assert.equal(port, 8765);
          assert.equal(token, "token");
          return "/tmp/thunderbird-mcp/connection.json";
        },
      }));

      assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
      assert.equal(writeCount, 1);
    });
  }

  it("rewrites when reading or parsing connection.json fails", () => {
    const errors = [];
    let writeCount = 0;
    const parseError = new SyntaxError("Unexpected token");

    const result = ensureFreshConnectionInfo(makeOptions({
      readConnectionInfo() {
        throw parseError;
      },
      writeConnectionInfo() {
        writeCount++;
        return "/tmp/thunderbird-mcp/connection.json";
      },
      onCheckError(error) {
        errors.push(error);
      },
    }));

    assert.equal(result, "/tmp/thunderbird-mcp/connection.json");
    assert.deepEqual(errors, [parseError]);
    assert.equal(writeCount, 1);
  });

  it("propagates write failures", () => {
    const writeError = new Error("disk full");

    assert.throws(
      () => ensureFreshConnectionInfo(makeOptions({
        readConnectionInfo: () => ({
          path: "/tmp/thunderbird-mcp/connection.json",
          data: null,
        }),
        writeConnectionInfo() {
          throw writeError;
        },
      })),
      (error) => error === writeError
    );
  });
});
