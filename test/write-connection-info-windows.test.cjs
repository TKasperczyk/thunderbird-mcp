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
 * platform behaviour: Windows can report synthesised group/other bits that
 * chmod cannot clear; on POSIX the assignment takes effect when `chmodWorks`.
 */
function makeTmpDir({
  exists = true,
  permissions = 0o777,
  chmodWorks = true,
  symlink = false,
  connectionFileExists = false,
}) {
  let mode = permissions;
  const dir = {
    segments: [],
    chmodCalls: [],
    createCalls: [],
    removeCalls: [],
    events: [],
    permissionReads: 0,
    append(name) { this.segments.push(name); },
    exists() { return exists; },
    isSymlink() { return symlink; },
    create(type, permissions) {
      this.createCalls.push({ type, permissions });
      exists = true;
    },
    get permissions() {
      this.permissionReads++;
      return mode;
    },
    set permissions(value) {
      this.chmodCalls.push(value);
      if (chmodWorks) mode = value;
    },
    clone() {
      const file = {
        segments: [...this.segments],
        append(name) { this.segments.push(name); },
        exists() { return connectionFileExists; },
        remove(recursive) {
          dir.events.push("remove");
          dir.removeCalls.push(recursive);
          connectionFileExists = false;
        },
        get path() { return this.segments.join("/"); },
      };
      return file;
    },
  };
  return dir;
}

function makeSandbox({ os, tmpDir }) {
  const written = [];
  const opened = [];
  return {
    written,
    opened,
    sandbox: {
      Services: {
        dirsvc: { get: () => tmpDir },
        appinfo: { OS: os, processID: 4242 },
      },
      Ci: { nsIFile: { DIRECTORY_TYPE: 1 } },
      Cc: {
        "@mozilla.org/network/file-output-stream;1": {
          createInstance: () => ({
            init(file, flags, permissions, behavior) {
              tmpDir.events.push("open");
              opened.push({
                path: file.path,
                flags,
                permissions,
                behavior,
                fileExisted: file.exists(),
              });
            },
          }),
        },
        "@mozilla.org/intl/converter-output-stream;1": {
          createInstance: () => ({
            init() {},
            writeString(data) {
              tmpDir.events.push("write");
              written.push(data);
            },
            close() { tmpDir.events.push("close"); },
          }),
        },
      },
    },
  };
}

describe("writeConnectionInfo directory hardening", () => {
  for (const permissions of [0o666, 0o777]) {
    it(`skips the POSIX permission check on Windows with synthesised 0o${permissions.toString(8)}`, () => {
      const tmpDir = makeTmpDir({ permissions, chmodWorks: false });
      const { sandbox, written } = makeSandbox({ os: "WINNT", tmpDir });
      const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

      const result = writeConnectionInfo(8765, "token");

      assert.equal(result, "thunderbird-mcp/connection.json");
      assert.equal(tmpDir.permissionReads, 0);
      assert.deepEqual(tmpDir.chmodCalls, []);
      assert.deepEqual(JSON.parse(written[0]), { port: 8765, token: "token", pid: 4242 });
    });
  }

  for (const os of ["Linux", "Darwin"]) {
    for (const permissions of [0o770, 0o707, 0o777]) {
      it(`refuses 0o${permissions.toString(8)} on ${os} when chmod does not clear group/other bits`, () => {
        const tmpDir = makeTmpDir({ permissions, chmodWorks: false });
        const { sandbox, opened, written } = makeSandbox({ os, tmpDir });
        const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

        assert.throws(
          () => writeConnectionInfo(8765, "token"),
          /group\/world permissions/
        );
        assert.deepEqual(tmpDir.chmodCalls, [0o700]);
        assert.deepEqual(opened, []);
        assert.deepEqual(written, []);
      });
    }

    it(`repairs a permissive directory on ${os} when chmod succeeds`, () => {
      const tmpDir = makeTmpDir({ permissions: 0o755, chmodWorks: true });
      const { sandbox, written } = makeSandbox({ os, tmpDir });
      const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

      const result = writeConnectionInfo(8770, "token");

      assert.equal(result, "thunderbird-mcp/connection.json");
      assert.deepEqual(tmpDir.chmodCalls, [0o700]);
      assert.equal(tmpDir.permissions, 0o700);
      assert.equal(written.length, 1);
    });

    for (const permissions of [0o700, 0]) {
      it(`preserves existing 0o${permissions.toString(8)} handling on ${os} without chmod`, () => {
        const tmpDir = makeTmpDir({ permissions });
        const { sandbox, written } = makeSandbox({ os, tmpDir });
        const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

        assert.equal(writeConnectionInfo(8765, "token"), "thunderbird-mcp/connection.json");
        assert.equal(tmpDir.permissionReads, 1);
        assert.deepEqual(tmpDir.chmodCalls, []);
        assert.equal(written.length, 1);
      });
    }
  }

  for (const os of ["WINNT", "Linux", "Darwin"]) {
    it(`creates a missing directory with 0700 on ${os}`, () => {
      const tmpDir = makeTmpDir({ exists: false });
      const { sandbox, written } = makeSandbox({ os, tmpDir });
      const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

      assert.equal(writeConnectionInfo(8765, "token"), "thunderbird-mcp/connection.json");
      assert.deepEqual(tmpDir.createCalls, [{ type: 1, permissions: 0o700 }]);
      assert.equal(tmpDir.exists(), true);
      assert.equal(tmpDir.permissionReads, 0);
      assert.deepEqual(tmpDir.chmodCalls, []);
      assert.equal(written.length, 1);
    });

    it(`rejects a symlink directory before writing on ${os}`, () => {
      const tmpDir = makeTmpDir({ symlink: true });
      const { sandbox, opened, written } = makeSandbox({ os, tmpDir });
      const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

      assert.throws(() => writeConnectionInfo(8765, "token"), /directory is a symlink/);
      assert.deepEqual(tmpDir.createCalls, []);
      assert.equal(tmpDir.permissionReads, 0);
      assert.deepEqual(tmpDir.chmodCalls, []);
      assert.deepEqual(opened, []);
      assert.deepEqual(written, []);
    });

    for (const connectionFileExists of [false, true]) {
      it(`writes ${connectionFileExists ? "a replacement" : "a new"} file with exclusive creation and 0600 on ${os}`, () => {
        const tmpDir = makeTmpDir({ permissions: 0o700, connectionFileExists });
        const { sandbox, opened, written } = makeSandbox({ os, tmpDir });
        const writeConnectionInfo = loadWriteConnectionInfo(sandbox);

        const result = writeConnectionInfo(8770, "new-token");

        assert.equal(result, "thunderbird-mcp/connection.json");
        assert.deepEqual(tmpDir.removeCalls, connectionFileExists ? [false] : []);
        assert.deepEqual(tmpDir.events, [
          ...(connectionFileExists ? ["remove"] : []),
          "open", "write", "close",
        ]);
        assert.deepEqual(opened, [{
          path: "thunderbird-mcp/connection.json",
          flags: 0x02 | 0x08 | 0x80,
          permissions: 0o600,
          behavior: 0,
          fileExisted: false,
        }]);
        assert.deepEqual(JSON.parse(written[0]), { port: 8770, token: "new-token", pid: 4242 });
      });
    }
  }
});
