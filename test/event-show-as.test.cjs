"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

// Mirrors createEvent's showAs/status decision logic. XPCOM is unavailable in
// node:test, so this covers only the input validation and derivation rules.
function validateAndDeriveCreate(status, showAs) {
  if (showAs !== undefined && showAs !== null && showAs !== "busy" && showAs !== "free") {
    return { error: `Invalid showAs: "${showAs}". Expected "busy" or "free".` };
  }
  const effectiveStatus = (status !== undefined && status !== null && status !== "")
    ? status
    : (showAs === "free" ? "tentative" : "confirmed");
  const transp = showAs === "free" ? "TRANSPARENT" : "OPAQUE";
  return { effectiveStatus, transp };
}

// Mirrors updateEvent's showAs decision tree.
function applyUpdateShowAs(showAs) {
  if (showAs === undefined) return { action: "noop" };
  if (showAs === null || showAs === "") return { action: "clear" };
  if (showAs === "free") return { action: "set", transp: "TRANSPARENT" };
  if (showAs === "busy") return { action: "set", transp: "OPAQUE" };
  return { error: `Invalid showAs: "${showAs}". Expected "busy" or "free".` };
}

describe("createEvent showAs validation and derivation", () => {
  it("rejects non-enum showAs values", () => {
    const result = validateAndDeriveCreate(undefined, "bogus");
    assert.match(result.error, /^Invalid showAs/);
  });

  it("accepts omitted showAs (defaults to busy)", () => {
    assert.deepStrictEqual(validateAndDeriveCreate(undefined, undefined),
      { effectiveStatus: "confirmed", transp: "OPAQUE" });
  });

  it("derives tentative+TRANSPARENT for showAs=free", () => {
    assert.deepStrictEqual(validateAndDeriveCreate(undefined, "free"),
      { effectiveStatus: "tentative", transp: "TRANSPARENT" });
  });

  it("derives confirmed+OPAQUE for showAs=busy", () => {
    assert.deepStrictEqual(validateAndDeriveCreate(undefined, "busy"),
      { effectiveStatus: "confirmed", transp: "OPAQUE" });
  });

  it("explicit status overrides showAs-derived status (TRANSP still follows showAs)", () => {
    assert.deepStrictEqual(validateAndDeriveCreate("cancelled", "free"),
      { effectiveStatus: "cancelled", transp: "TRANSPARENT" });
    assert.deepStrictEqual(validateAndDeriveCreate("confirmed", "free"),
      { effectiveStatus: "confirmed", transp: "TRANSPARENT" });
  });

  it("empty-string status falls back to showAs-derived", () => {
    assert.deepStrictEqual(validateAndDeriveCreate("", "free"),
      { effectiveStatus: "tentative", transp: "TRANSPARENT" });
  });
});

describe("updateEvent showAs action tree", () => {
  it("undefined is no-op (preserves existing TRANSP)", () => {
    assert.deepStrictEqual(applyUpdateShowAs(undefined), { action: "noop" });
  });

  it("null clears TRANSP", () => {
    assert.deepStrictEqual(applyUpdateShowAs(null), { action: "clear" });
  });

  it("empty string clears TRANSP", () => {
    assert.deepStrictEqual(applyUpdateShowAs(""), { action: "clear" });
  });

  it("busy sets OPAQUE", () => {
    assert.deepStrictEqual(applyUpdateShowAs("busy"), { action: "set", transp: "OPAQUE" });
  });

  it("free sets TRANSPARENT", () => {
    assert.deepStrictEqual(applyUpdateShowAs("free"), { action: "set", transp: "TRANSPARENT" });
  });

  it("non-enum value rejected with error", () => {
    const result = applyUpdateShowAs("bogus");
    assert.match(result.error, /^Invalid showAs/);
  });
});
