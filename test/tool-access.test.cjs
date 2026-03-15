"use strict";
const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

// ── Helpers that mirror production logic ──────────────────────────────

const UNDISABLEABLE_TOOLS = new Set(["listAccounts", "listFolders", "getAccountAccessConfig"]);

/**
 * Mirrors the production getDisabledTools() fail-closed behavior.
 * @param {string} prefValue - raw preference string (simulates Services.prefs)
 */
function getDisabledTools(prefValue) {
  try {
    if (!prefValue) return [];
    const parsed = JSON.parse(prefValue);
    if (!Array.isArray(parsed)) {
      return ["__all__"];
    }
    return parsed;
  } catch {
    return ["__all__"];
  }
}

/**
 * Mirrors the production isToolEnabled() logic.
 */
function isToolEnabled(toolName, prefValue) {
  if (UNDISABLEABLE_TOOLS.has(toolName)) return true;
  const disabled = getDisabledTools(prefValue);
  if (disabled.includes("__all__")) return false;
  return !disabled.includes(toolName);
}

/**
 * Mirrors tools/list filtering.
 */
function filterTools(allTools, prefValue) {
  return allTools.filter(t => isToolEnabled(t.name, prefValue));
}

/**
 * Mirrors callTool guard - throws if tool is disabled.
 */
function callToolGuard(toolName, prefValue) {
  if (!isToolEnabled(toolName, prefValue)) {
    throw new Error(`Tool is disabled: ${toolName}`);
  }
  return { success: true };
}

// ── Test data ─────────────────────────────────────────────────────────

const ALL_TOOLS = [
  { name: "listAccounts" },
  { name: "listFolders" },
  { name: "getAccountAccessConfig" },
  { name: "searchMessages" },
  { name: "getMessage" },
  { name: "sendMail" },
  { name: "replyToMessage" },
  { name: "forwardMessage" },
  { name: "deleteMessages" },
  { name: "updateMessage" },
  { name: "getRecentMessages" },
  { name: "createFolder" },
  { name: "searchContacts" },
  { name: "createContact" },
  { name: "deleteContact" },
];

// ── Tests ─────────────────────────────────────────────────────────────

describe("Tool access: getDisabledTools", () => {
  it("empty pref returns empty array (all enabled)", () => {
    assert.deepStrictEqual(getDisabledTools(""), []);
    assert.deepStrictEqual(getDisabledTools(undefined), []);
    assert.deepStrictEqual(getDisabledTools(null), []);
  });

  it("valid JSON array returns tool names", () => {
    const pref = JSON.stringify(["sendMail", "deleteMessages"]);
    assert.deepStrictEqual(getDisabledTools(pref), ["sendMail", "deleteMessages"]);
  });

  it("FAIL-CLOSED: non-array JSON returns __all__ sentinel", () => {
    assert.deepStrictEqual(getDisabledTools('{"sendMail": true}'), ["__all__"]);
    assert.deepStrictEqual(getDisabledTools('"sendMail"'), ["__all__"]);
    assert.deepStrictEqual(getDisabledTools("42"), ["__all__"]);
    assert.deepStrictEqual(getDisabledTools("true"), ["__all__"]);
  });

  it("FAIL-CLOSED: invalid JSON returns __all__ sentinel", () => {
    assert.deepStrictEqual(getDisabledTools("not json at all"), ["__all__"]);
    assert.deepStrictEqual(getDisabledTools("{broken"), ["__all__"]);
    assert.deepStrictEqual(getDisabledTools("[unclosed"), ["__all__"]);
  });

  it("empty array pref means all enabled", () => {
    assert.deepStrictEqual(getDisabledTools("[]"), []);
  });
});

describe("Tool access: isToolEnabled", () => {
  it("all tools enabled when pref is empty", () => {
    assert.ok(isToolEnabled("sendMail", ""));
    assert.ok(isToolEnabled("deleteMessages", ""));
    assert.ok(isToolEnabled("searchContacts", ""));
  });

  it("disabled tool returns false", () => {
    const pref = JSON.stringify(["sendMail", "deleteMessages"]);
    assert.ok(!isToolEnabled("sendMail", pref));
    assert.ok(!isToolEnabled("deleteMessages", pref));
  });

  it("non-disabled tool returns true", () => {
    const pref = JSON.stringify(["sendMail"]);
    assert.ok(isToolEnabled("searchMessages", pref));
    assert.ok(isToolEnabled("getMessage", pref));
  });

  it("undisableable tools ALWAYS return true even when in disabled list", () => {
    const pref = JSON.stringify(["listAccounts", "listFolders", "getAccountAccessConfig", "sendMail"]);
    assert.ok(isToolEnabled("listAccounts", pref));
    assert.ok(isToolEnabled("listFolders", pref));
    assert.ok(isToolEnabled("getAccountAccessConfig", pref));
    // But sendMail IS disabled
    assert.ok(!isToolEnabled("sendMail", pref));
  });

  it("undisableable tools survive __all__ sentinel (corrupt pref)", () => {
    const pref = "not valid json";
    assert.ok(isToolEnabled("listAccounts", pref));
    assert.ok(isToolEnabled("listFolders", pref));
    assert.ok(isToolEnabled("getAccountAccessConfig", pref));
    // Regular tools are blocked
    assert.ok(!isToolEnabled("sendMail", pref));
    assert.ok(!isToolEnabled("deleteMessages", pref));
  });
});

describe("Tool access: tools/list filtering", () => {
  it("returns all tools when pref is empty", () => {
    const result = filterTools(ALL_TOOLS, "");
    assert.equal(result.length, ALL_TOOLS.length);
  });

  it("hides disabled tools from the list", () => {
    const pref = JSON.stringify(["sendMail", "deleteMessages"]);
    const result = filterTools(ALL_TOOLS, pref);
    const names = result.map(t => t.name);
    assert.ok(!names.includes("sendMail"));
    assert.ok(!names.includes("deleteMessages"));
    assert.ok(names.includes("searchMessages"));
    assert.ok(names.includes("getMessage"));
    assert.equal(result.length, ALL_TOOLS.length - 2);
  });

  it("undisableable tools always appear in list", () => {
    const pref = JSON.stringify(["listAccounts", "listFolders", "getAccountAccessConfig"]);
    const result = filterTools(ALL_TOOLS, pref);
    const names = result.map(t => t.name);
    assert.ok(names.includes("listAccounts"));
    assert.ok(names.includes("listFolders"));
    assert.ok(names.includes("getAccountAccessConfig"));
  });

  it("corrupt pref shows only undisableable tools", () => {
    const result = filterTools(ALL_TOOLS, "corrupt!");
    const names = result.map(t => t.name);
    assert.equal(names.length, 3);
    assert.ok(names.includes("listAccounts"));
    assert.ok(names.includes("listFolders"));
    assert.ok(names.includes("getAccountAccessConfig"));
  });
});

describe("Tool access: callTool guard (defense in depth)", () => {
  it("allows enabled tools", () => {
    const result = callToolGuard("searchMessages", "");
    assert.deepStrictEqual(result, { success: true });
  });

  it("rejects disabled tools with clear error", () => {
    const pref = JSON.stringify(["sendMail"]);
    assert.throws(
      () => callToolGuard("sendMail", pref),
      { message: "Tool is disabled: sendMail" }
    );
  });

  it("allows undisableable tools even when explicitly disabled", () => {
    const pref = JSON.stringify(["listAccounts"]);
    const result = callToolGuard("listAccounts", pref);
    assert.deepStrictEqual(result, { success: true });
  });

  it("rejects all regular tools on corrupt pref (fail-closed)", () => {
    assert.throws(
      () => callToolGuard("sendMail", "{not an array}"),
      { message: "Tool is disabled: sendMail" }
    );
    assert.throws(
      () => callToolGuard("deleteMessages", "{not an array}"),
      { message: "Tool is disabled: deleteMessages" }
    );
  });

  it("allows undisableable tools on corrupt pref", () => {
    const result = callToolGuard("listAccounts", "{not an array}");
    assert.deepStrictEqual(result, { success: true });
  });
});

describe("Tool access: adversarial inputs", () => {
  it("tool name with prototype pollution attempt", () => {
    const pref = JSON.stringify(["constructor"]);
    // "constructor" is not undisableable, so it should be blocked
    assert.ok(!isToolEnabled("constructor", pref));
    // But __proto__ is not in the disabled list
    assert.ok(isToolEnabled("__proto__", pref));
  });

  it("disabled list with __proto__ and constructor entries", () => {
    const pref = JSON.stringify(["__proto__", "constructor", "toString", "sendMail"]);
    assert.ok(!isToolEnabled("sendMail", pref));
    assert.ok(!isToolEnabled("__proto__", pref));
    assert.ok(!isToolEnabled("constructor", pref));
    // Undisableable tools still work
    assert.ok(isToolEnabled("listAccounts", pref));
  });

  it("empty string tool name in disabled list", () => {
    const pref = JSON.stringify([""]);
    // Empty string tool is disabled
    assert.ok(!isToolEnabled("", pref));
    // Real tools still work
    assert.ok(isToolEnabled("sendMail", pref));
  });

  it("disabled list with duplicate entries", () => {
    const pref = JSON.stringify(["sendMail", "sendMail", "sendMail"]);
    assert.ok(!isToolEnabled("sendMail", pref));
    assert.ok(isToolEnabled("deleteMessages", pref));
  });

  it("disabled list with non-existent tool names", () => {
    const pref = JSON.stringify(["fakeToolThatDoesNotExist", "anotherFake"]);
    // These fake tools are disabled (defense in depth)
    assert.ok(!isToolEnabled("fakeToolThatDoesNotExist", pref));
    // Real tools unaffected
    assert.ok(isToolEnabled("sendMail", pref));
    assert.ok(isToolEnabled("deleteMessages", pref));
  });

  it("case sensitivity: tool names are case-sensitive", () => {
    const pref = JSON.stringify(["SendMail", "SENDMAIL"]);
    // Only exact match is disabled
    assert.ok(isToolEnabled("sendMail", pref));
    assert.ok(!isToolEnabled("SendMail", pref));
    assert.ok(!isToolEnabled("SENDMAIL", pref));
  });

  it("pref with nested arrays (not flat strings)", () => {
    // JSON.parse will succeed, Array.isArray will pass, but includes() works on elements
    const pref = JSON.stringify([["sendMail"], "deleteMessages"]);
    // The nested array element won't match string comparison
    assert.ok(isToolEnabled("sendMail", pref));
    // But the string entry does match
    assert.ok(!isToolEnabled("deleteMessages", pref));
  });

  it("pref with null and number entries", () => {
    const pref = JSON.stringify([null, 42, "sendMail"]);
    assert.ok(!isToolEnabled("sendMail", pref));
    assert.ok(isToolEnabled("deleteMessages", pref));
  });

  it("very large disabled list", () => {
    const bigList = Array.from({ length: 10000 }, (_, i) => `tool_${i}`);
    bigList.push("sendMail");
    const pref = JSON.stringify(bigList);
    assert.ok(!isToolEnabled("sendMail", pref));
    assert.ok(isToolEnabled("deleteMessages", pref));
  });

  it("tools/list with all tools disabled shows only undisableable", () => {
    const allNames = ALL_TOOLS.map(t => t.name);
    const pref = JSON.stringify(allNames);
    const result = filterTools(ALL_TOOLS, pref);
    assert.equal(result.length, 3);
    for (const t of result) {
      assert.ok(UNDISABLEABLE_TOOLS.has(t.name), `${t.name} should be undisableable`);
    }
  });

  it("callTool guard blocks unknown tool names that happen to be disabled", () => {
    const pref = JSON.stringify(["nonExistentTool"]);
    assert.throws(
      () => callToolGuard("nonExistentTool", pref),
      { message: "Tool is disabled: nonExistentTool" }
    );
  });
});
