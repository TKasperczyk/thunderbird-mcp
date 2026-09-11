"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadWriteConnectionInfo(sandbox) {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN CONNECTION INFO WRITER";
  const endMarker = "// END CONNECTION INFO WRITER";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "connection info writer start marker missing");
  assert.ok(end > start, "connection info writer end marker missing");

  const snippet = source.slice(start, end);
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.writeConnectionInfo = writeConnectionInfo;`,
    sandbox
  );
  return sandbox.writeConnectionInfo;
}

/**
 * Minimal nsIFile stand-in for <TmpD>/thunderbird-mcp. `permissions` mimics
 * platform behaviour: on Windows nsIFile reports a synthesised mode (0o777
 * for directories) and assignments are ignored (NTFS ACLs); on POSIX the
 * assignment takes effect unless `chmodFails` is set.
 */
function makeTmpDir({ exists = true, permissions = 0o777, chmodWorks = true }) {
  let mode = permissions;
  const dir = {
    segments: [],
    chmodCalls: [],
    append(name) { this.segments.push(name); },
    exists() { return exists; },
    isSymlink() { return false; },
    create() { exists = true; },
    get permissions() { return mode; },
    set permissions(value) {
      this.chmodCalls.push(value);
      if (chmodWorks) mode = value;
    },
    clone() {
      const file = {
        segments: [...this.segments],
        append(name) { this.segments.push(name); },
        exists() { return false; },
        remove() {},
        get path() { return this.segments.join("/"); },
      };
      return file;
    },
  };
  return dir;
}

function makeSandbox({ os, tmpDir }) {
  const written = [];
  return {
    written,
    sandbox: {
      Services: {
        dirsvc: { get: () => tmpDir },
        appinfo: { OS: os, processID: 4242 },
      },
      Ci: { nsIFile: { DIRECTORY_TYPE: 1 } },
      Cc: {
        "@mozilla.org/network/file-output-stream;1": {
          createInstance: () => ({ init() {} }),
        },
        "@mozilla.org/intl/converter-output-stream;1": {
          createInstance: () => ({
            init() {},
            writeString(data) { written.push(data); },
            close() {},
          }),
        },
      },
    },
  };
}

describe("writeConnectionInfo directory hardening", () => {
  it("skips the POSIX permission check on Windows where nsIFile reports 0o777", () => {
    const tmpDir = makeTmpDir({ permissions: 0o777, chmodWorks: false });
    const { sandbox, written } = makeSandbox({ os: "WINNT", tmpDir });
    const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

    const result = writeConnectionInfo(8765, "token");

    assert.equal(result, "thunderbird-mcp/connection.json");
    assert.deepEqual(tmpDir.chmodCalls, []);
    assert.deepEqual(JSON.parse(written[0]), { port: 8765, token: "token", pid: 4242 });
  });

  it("still refuses a group/world-accessible directory on POSIX when chmod fails", () => {
    const tmpDir = makeTmpDir({ permissions: 0o777, chmodWorks: false });
    const { sandbox } = makeSandbox({ os: "Linux", tmpDir });
    const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

    assert.throws(
      () => writeConnectionInfo(8765, "token"),
      /group\/world permissions/
    );
    assert.deepEqual(tmpDir.chmodCalls, [0o700]);
  });

  it("repairs a permissive directory on POSIX when chmod succeeds", () => {
    const tmpDir = makeTmpDir({ permissions: 0o755, chmodWorks: true });
    const { sandbox, written } = makeSandbox({ os: "Linux", tmpDir });
    const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

    const result = writeConnectionInfo(8770, "token");

    assert.equal(result, "thunderbird-mcp/connection.json");
    assert.equal(tmpDir.permissions, 0o700);
    assert.equal(written.length, 1);
  });
});
