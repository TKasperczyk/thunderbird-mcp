"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadTmpDirHardeningHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN TMP DIR HARDENING HELPERS";
  const endMarker = "// END TMP DIR HARDENING HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "tmp dir hardening helper start marker missing");
  assert.ok(end > start, "tmp dir hardening helper end marker missing");

  const snippet = source.slice(start, end);
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.tmpDirModeIsPosix = tmpDirModeIsPosix;
this.hardenTmpDirPermissions = hardenTmpDirPermissions;`,
    sandbox
  );
  return {
    tmpDirModeIsPosix: sandbox.tmpDirModeIsPosix,
    hardenTmpDirPermissions: sandbox.hardenTmpDirPermissions,
  };
}

const { tmpDirModeIsPosix, hardenTmpDirPermissions } = loadTmpDirHardeningHelpers();

// Minimal nsIFile stand-in. `sticky: false` models Windows, where assigning to
// .permissions only toggles the read-only attribute and cannot clear the
// group/world bits of the synthesized mode.
function makeDirWithSetter(initialMode, { sticky = true } = {}) {
  let mode = initialMode;
  const assignments = [];
  const dir = {
    get permissions() { return mode; },
    set permissions(v) {
      assignments.push(v);
      if (sticky) { mode = v; }
    },
  };
  Object.defineProperty(dir, "assignments", { value: assignments });
  return dir;
}

describe("tmpDirModeIsPosix", () => {
  it("treats Windows as non-POSIX", () => {
    assert.equal(tmpDirModeIsPosix("WINNT"), false);
  });

  for (const os of ["Linux", "Darwin", "FreeBSD", "OpenBSD"]) {
    it(`treats ${os} as POSIX`, () => {
      assert.equal(tmpDirModeIsPosix(os), true);
    });
  }
});

describe("hardenTmpDirPermissions", () => {
  // Regression: nsIFile.permissions on Windows reports a synthesized mode with
  // group/world bits set and the chmod never sticks, so the pre-existing
  // directory branch used to throw on every call. Once the connection file was
  // deleted but the directory survived, the MCP server could never write it
  // again -- fatal at startup, and silently unrecoverable in the refresh timer.
  it("skips hardening on Windows even when the synthesized mode looks permissive", () => {
    const dir = makeDirWithSetter(0o777, { sticky: false });
    const result = hardenTmpDirPermissions({ dir, osName: "WINNT" });
    assert.equal(result.hardened, false);
    assert.equal(result.reason, "non-posix-permissions");
    assert.deepEqual([...dir.assignments], [], "must not attempt a chmod on Windows");
  });

  it("does not touch an already-private POSIX directory", () => {
    const dir = makeDirWithSetter(0o700);
    const result = hardenTmpDirPermissions({ dir, osName: "Linux" });
    assert.equal(result.hardened, true);
    assert.equal(result.reason, "already-private");
    assert.deepEqual([...dir.assignments], []);
  });

  it("treats a zero mode as an unsupported accessor rather than a permissive dir", () => {
    const dir = makeDirWithSetter(0);
    const result = hardenTmpDirPermissions({ dir, osName: "Linux" });
    assert.equal(result.hardened, true);
    assert.equal(result.reason, "already-private");
    assert.deepEqual([...dir.assignments], []);
  });

  it("chmods a group/world-readable POSIX directory back to 0700", () => {
    const dir = makeDirWithSetter(0o755);
    const result = hardenTmpDirPermissions({ dir, osName: "Linux" });
    assert.equal(result.hardened, true);
    assert.equal(result.reason, "chmod");
    assert.deepEqual([...dir.assignments], [0o700]);
    assert.equal(dir.permissions, 0o700);
  });

  it("refuses to write when the POSIX chmod does not stick", () => {
    const dir = makeDirWithSetter(0o777, { sticky: false });
    assert.throws(
      () => hardenTmpDirPermissions({ dir, osName: "Linux" }),
      /group\/world permissions/
    );
  });
});
