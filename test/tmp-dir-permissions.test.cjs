"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadTmpDirPermissionHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN TMP DIR PERMISSION HELPERS";
  const endMarker = "// END TMP DIR PERMISSION HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "tmp dir permission helper start marker missing");
  assert.ok(end > start, "tmp dir permission helper end marker missing");

  const snippet = source.slice(start, end);
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.tmpDirModeUnsafe = tmpDirModeUnsafe;`,
    sandbox
  );
  return { tmpDirModeUnsafe: sandbox.tmpDirModeUnsafe };
}

const { tmpDirModeUnsafe } = loadTmpDirPermissionHelpers();

describe("tmp dir permission check", () => {
  for (const osName of ["Linux", "Darwin"]) {
    it(`rejects group/world access on ${osName}`, () => {
      assert.equal(tmpDirModeUnsafe(0o777, osName), true);
      assert.equal(tmpDirModeUnsafe(0o750, osName), true);
      assert.equal(tmpDirModeUnsafe(0o705, osName), true);
    });

    it(`accepts an owner-only directory on ${osName}`, () => {
      assert.equal(tmpDirModeUnsafe(0o700, osName), false);
      assert.equal(tmpDirModeUnsafe(0o600, osName), false);
    });

    it(`accepts mode 0 on ${osName} (accessor unsupported)`, () => {
      assert.equal(tmpDirModeUnsafe(0, osName), false);
    });
  }

  // Regression: nsIFile.permissions reports 0o777 for every writable directory
  // on Windows and assigning 0o700 back is a no-op, so enforcing the POSIX mode
  // there made the server start exactly once, on the run that created the
  // directory, and throw on every start afterwards.
  it("does not reject the 0o777 Windows reports for a normal directory", () => {
    assert.equal(tmpDirModeUnsafe(0o777, "WINNT"), false);
    assert.equal(tmpDirModeUnsafe(0o555, "WINNT"), false);
    assert.equal(tmpDirModeUnsafe(0o700, "WINNT"), false);
  });
});
