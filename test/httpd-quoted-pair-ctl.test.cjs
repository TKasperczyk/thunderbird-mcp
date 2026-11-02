"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadNormalizeFieldValue() {
  const httpdPath = path.resolve(__dirname, "../extension/httpd.sys.mjs");
  const source = fs.readFileSync(httpdPath, "utf8");

  const isCTLMatch = source.match(/function isCTL[\s\S]*?\n}/);
  const headerUtilsMatch = source.match(/const headerUtils = {[\s\S]*?\n};/);
  assert.ok(isCTLMatch, "isCTL helper function missing in httpd.sys.mjs");
  assert.ok(headerUtilsMatch, "headerUtils object missing in httpd.sys.mjs");

  const Components = { Exception: (msg, code) => ({ msg, code }) };
  const Cr = { NS_ERROR_INVALID_ARG: "NS_ERROR_INVALID_ARG" };
  const dumpn = () => {};
  const dump = () => {};

  const sandbox = { Components, Cr, dumpn, dump };
  vm.createContext(sandbox);
  vm.runInContext(
    isCTLMatch[0] +
      "\n" +
      headerUtilsMatch[0] +
      "\nthis.normalizeFieldValue = headerUtils.normalizeFieldValue;",
    sandbox
  );
  return sandbox.normalizeFieldValue;
}

const normalizeFieldValue = loadNormalizeFieldValue();

describe("httpd headerUtils.normalizeFieldValue quoted-pair CTL handling", () => {
  it("allows quoted-pairs where the escaped character is a CTL", () => {
    // Escaped CTL characters (0-31, 127) should be allowed as part of quoted-pair
    for (const ctlCode of [0, 1, 7, 12, 27, 31, 127]) {
      const escaped = "text \\" + String.fromCharCode(ctlCode) + " end";
      const result = normalizeFieldValue(escaped);
      assert.equal(result, escaped);
    }
  });

  it("disallows unquoted CTL characters", () => {
    for (const ctlCode of [0, 1, 7, 12, 27, 31, 127]) {
      const unescaped = "text " + String.fromCharCode(ctlCode) + " end";
      assert.throws(
        () => normalizeFieldValue(unescaped),
        (err) => err && err.code === "NS_ERROR_INVALID_ARG"
      );
    }
  });

  it("handles trailing backslash at end of string without erroring", () => {
    // A single trailing backslash has i + 1 >= len, so it shouldn't skip or panic
    const trailingSlash = "text \\";
    const result = normalizeFieldValue(trailingSlash);
    assert.equal(result, trailingSlash);
  });
});
