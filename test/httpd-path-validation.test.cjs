"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadPathValidationHelper() {
  const httpdPath = path.resolve(__dirname, "../extension/httpd.sys.mjs");
  const source = fs.readFileSync(httpdPath, "utf8");
  const startMarker = "// BEGIN PATH VALIDATION HELPERS";
  const endMarker = "// END PATH VALIDATION HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "path validation helper start marker missing");
  assert.ok(end > start, "path validation helper end marker missing");

  const snippet = source.slice(start, end);
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}\nthis._isValidPath = _isValidPath;`,
    sandbox
  );
  return sandbox._isValidPath;
}

const _isValidPath = loadPathValidationHelper();

describe("httpd path validation (_isValidPath)", () => {
  it("accepts valid paths", () => {
    assert.equal(_isValidPath("/"), true);
    assert.equal(_isValidPath("/foo"), true);
    assert.equal(_isValidPath("/foo/bar"), true);
    assert.equal(_isValidPath("/foo/bar/"), true);
    assert.equal(_isValidPath("/a/b/c.html"), true);
    assert.equal(_isValidPath("/path%20with%20spaces"), true);
  });

  it("handles CONNECT special case when allowConnect is enabled", () => {
    assert.equal(_isValidPath("CONNECT", true), true);
    assert.equal(_isValidPath("CONNECT", false), false);
  });

  it("rejects non-string or empty paths", () => {
    assert.equal(_isValidPath(""), false);
    assert.equal(_isValidPath(null), false);
    assert.equal(_isValidPath(undefined), false);
    assert.equal(_isValidPath(123), false);
  });

  it("rejects relative paths not starting with /", () => {
    assert.equal(_isValidPath("foo"), false);
    assert.equal(_isValidPath("foo/bar"), false);
    assert.equal(_isValidPath("http://example.com/foo"), false);
  });

  it("rejects path traversal attempts", () => {
    assert.equal(_isValidPath("/.."), false);
    assert.equal(_isValidPath("/../"), false);
    assert.equal(_isValidPath("/foo/.."), false);
    assert.equal(_isValidPath("/foo/../bar"), false);
    assert.equal(_isValidPath("/a/b/../../c"), false);
  });

  it("rejects query strings, fragments, and matrix parameters", () => {
    assert.equal(_isValidPath("/foo?bar=1"), false);
    assert.equal(_isValidPath("/foo#section"), false);
    assert.equal(_isValidPath("/foo;param=val"), false);
  });

  it("rejects backslashes and null bytes", () => {
    assert.equal(_isValidPath("/foo\\bar"), false);
    assert.equal(_isValidPath("\\foo"), false);
    assert.equal(_isValidPath("/foo\0bar"), false);
  });

  it("rejects malformed URI percent-encodings", () => {
    assert.equal(_isValidPath("/foo%2"), false);
    assert.equal(_isValidPath("/foo%GH"), false);
  });
});
