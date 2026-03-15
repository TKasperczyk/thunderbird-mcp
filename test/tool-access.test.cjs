"use strict";
const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

// ── Helpers that mirror production logic ──────────────────────────────

const UNDISABLEABLE_TOOLS = new Set(["listAccounts", "listFolders", "getAccountAccess"]);

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

// Valid metadata values (mirrors production constants)
const VALID_GROUPS = ["messages", "folders", "contacts", "calendar", "filters", "system"];
const VALID_CRUD = ["create", "read", "update", "delete"];
const CRUD_ORDER = { read: 0, create: 1, update: 2, delete: 3 };
const GROUP_ORDER = { system: 0, messages: 1, folders: 2, contacts: 3, calendar: 4, filters: 5 };

const ALL_TOOLS = [
  { name: "listAccounts", group: "system", crud: "read" },
  { name: "listFolders", group: "system", crud: "read" },
  { name: "getAccountAccess", group: "system", crud: "read" },
  { name: "searchMessages", group: "messages", crud: "read" },
  { name: "getMessage", group: "messages", crud: "read" },
  { name: "getRecentMessages", group: "messages", crud: "read" },
  { name: "displayMessage", group: "messages", crud: "read" },
  { name: "sendMail", group: "messages", crud: "create" },
  { name: "replyToMessage", group: "messages", crud: "create" },
  { name: "forwardMessage", group: "messages", crud: "create" },
  { name: "updateMessage", group: "messages", crud: "update" },
  { name: "deleteMessages", group: "messages", crud: "delete" },
  { name: "createFolder", group: "folders", crud: "create" },
  { name: "renameFolder", group: "folders", crud: "update" },
  { name: "moveFolder", group: "folders", crud: "update" },
  { name: "deleteFolder", group: "folders", crud: "delete" },
  { name: "emptyTrash", group: "folders", crud: "delete" },
  { name: "emptyJunk", group: "folders", crud: "delete" },
  { name: "searchContacts", group: "contacts", crud: "read" },
  { name: "createContact", group: "contacts", crud: "create" },
  { name: "updateContact", group: "contacts", crud: "update" },
  { name: "deleteContact", group: "contacts", crud: "delete" },
  { name: "listCalendars", group: "calendar", crud: "read" },
  { name: "listEvents", group: "calendar", crud: "read" },
  { name: "createEvent", group: "calendar", crud: "create" },
  { name: "createTask", group: "calendar", crud: "create" },
  { name: "updateEvent", group: "calendar", crud: "update" },
  { name: "deleteEvent", group: "calendar", crud: "delete" },
  { name: "listFilters", group: "filters", crud: "read" },
  { name: "createFilter", group: "filters", crud: "create" },
  { name: "updateFilter", group: "filters", crud: "update" },
  { name: "reorderFilters", group: "filters", crud: "update" },
  { name: "applyFilters", group: "filters", crud: "update" },
  { name: "deleteFilter", group: "filters", crud: "delete" },
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
    const pref = JSON.stringify(["listAccounts", "listFolders", "getAccountAccess", "sendMail"]);
    assert.ok(isToolEnabled("listAccounts", pref));
    assert.ok(isToolEnabled("listFolders", pref));
    assert.ok(isToolEnabled("getAccountAccess", pref));
    // But sendMail IS disabled
    assert.ok(!isToolEnabled("sendMail", pref));
  });

  it("undisableable tools survive __all__ sentinel (corrupt pref)", () => {
    const pref = "not valid json";
    assert.ok(isToolEnabled("listAccounts", pref));
    assert.ok(isToolEnabled("listFolders", pref));
    assert.ok(isToolEnabled("getAccountAccess", pref));
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
    const pref = JSON.stringify(["listAccounts", "listFolders", "getAccountAccess"]);
    const result = filterTools(ALL_TOOLS, pref);
    const names = result.map(t => t.name);
    assert.ok(names.includes("listAccounts"));
    assert.ok(names.includes("listFolders"));
    assert.ok(names.includes("getAccountAccess"));
  });

  it("corrupt pref shows only undisableable tools", () => {
    const result = filterTools(ALL_TOOLS, "corrupt!");
    const names = result.map(t => t.name);
    assert.equal(names.length, 3);
    assert.ok(names.includes("listAccounts"));
    assert.ok(names.includes("listFolders"));
    assert.ok(names.includes("getAccountAccess"));
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

  it("__all__ sentinel in disabled list blocks all regular tools", () => {
    const pref = JSON.stringify(["__all__"]);
    assert.ok(!isToolEnabled("sendMail", pref));
    assert.ok(!isToolEnabled("deleteMessages", pref));
    assert.ok(!isToolEnabled("searchContacts", pref));
    // Undisableable tools survive
    assert.ok(isToolEnabled("listAccounts", pref));
    assert.ok(isToolEnabled("listFolders", pref));
    assert.ok(isToolEnabled("getAccountAccess", pref));
  });

  it("__all__ sentinel mixed with other entries still blocks all", () => {
    const pref = JSON.stringify(["sendMail", "__all__", "deleteMessages"]);
    assert.ok(!isToolEnabled("searchContacts", pref));
    assert.ok(!isToolEnabled("getMessage", pref));
  });
});

describe("Tool access: setToolAccess validation", () => {
  /**
   * Mirrors the production setToolAccess validation logic.
   */
  function validateSetToolAccess(disabledTools) {
    if (!Array.isArray(disabledTools)) {
      return { error: "disabledTools must be an array" };
    }
    // Validate types first, then semantic checks
    if (!disabledTools.every(t => typeof t === "string")) {
      return { error: "All tool names must be strings" };
    }
    if (disabledTools.includes("__all__")) {
      return { error: "Invalid tool name: __all__" };
    }
    const blocked = disabledTools.filter(t => UNDISABLEABLE_TOOLS.has(t));
    if (blocked.length > 0) {
      return { error: `Cannot disable infrastructure tools: ${blocked.join(", ")}` };
    }
    return { success: true };
  }

  it("accepts empty array (enable all)", () => {
    assert.deepStrictEqual(validateSetToolAccess([]), { success: true });
  });

  it("accepts valid tool names", () => {
    assert.deepStrictEqual(validateSetToolAccess(["sendMail", "deleteMessages"]), { success: true });
  });

  it("rejects non-array input", () => {
    assert.ok(validateSetToolAccess("sendMail").error);
    assert.ok(validateSetToolAccess(42).error);
    assert.ok(validateSetToolAccess(null).error);
  });

  it("rejects undisableable tools", () => {
    const result = validateSetToolAccess(["listAccounts", "sendMail"]);
    assert.ok(result.error);
    assert.match(result.error, /cannot disable/i);
    assert.match(result.error, /listAccounts/);
  });

  it("rejects non-string entries", () => {
    const result = validateSetToolAccess([42, "sendMail"]);
    assert.ok(result.error);
    assert.match(result.error, /strings/i);
  });

  it("SECURITY: rejects __all__ sentinel injection", () => {
    const result = validateSetToolAccess(["__all__"]);
    assert.ok(result.error);
    assert.match(result.error, /__all__/);
  });

  it("SECURITY: rejects __all__ even mixed with valid tools", () => {
    const result = validateSetToolAccess(["sendMail", "__all__", "deleteMessages"]);
    assert.ok(result.error);
    assert.match(result.error, /__all__/);
  });
});

// ── Tool metadata validation tests ───────────────────────────────────

describe("Tool metadata: group and crud validation", () => {
  it("every tool has a valid group", () => {
    for (const tool of ALL_TOOLS) {
      assert.ok(
        VALID_GROUPS.includes(tool.group),
        `Tool "${tool.name}" has invalid group: "${tool.group}"`
      );
    }
  });

  it("every tool has a valid crud type", () => {
    for (const tool of ALL_TOOLS) {
      assert.ok(
        VALID_CRUD.includes(tool.crud),
        `Tool "${tool.name}" has invalid crud: "${tool.crud}"`
      );
    }
  });

  it("no duplicate tool names", () => {
    const names = ALL_TOOLS.map(t => t.name);
    const unique = new Set(names);
    assert.equal(names.length, unique.size, `Duplicate tool names found: ${names.filter((n, i) => names.indexOf(n) !== i)}`);
  });

  it("all undisableable tools exist in the tools array", () => {
    const names = new Set(ALL_TOOLS.map(t => t.name));
    for (const name of UNDISABLEABLE_TOOLS) {
      assert.ok(names.has(name), `Undisableable tool "${name}" not found in tools array`);
    }
  });

  it("GROUP_ORDER covers all VALID_GROUPS", () => {
    for (const group of VALID_GROUPS) {
      assert.ok(
        GROUP_ORDER[group] !== undefined,
        `GROUP_ORDER missing entry for group: "${group}"`
      );
    }
  });

  it("CRUD_ORDER covers all VALID_CRUD", () => {
    for (const crud of VALID_CRUD) {
      assert.ok(
        CRUD_ORDER[crud] !== undefined,
        `CRUD_ORDER missing entry for crud: "${crud}"`
      );
    }
  });

  it("detects tool with missing group", () => {
    const badTool = { name: "testTool", crud: "read" };
    assert.ok(!VALID_GROUPS.includes(badTool.group));
  });

  it("detects tool with invalid crud value", () => {
    const badTool = { name: "testTool", group: "messages", crud: "execute" };
    assert.ok(!VALID_CRUD.includes(badTool.crud));
  });
});

describe("Tool metadata: getToolAccessConfig sorting", () => {
  function sortTools(toolList) {
    return [...toolList].sort((a, b) => {
      const gA = GROUP_ORDER[a.group] ?? 99;
      const gB = GROUP_ORDER[b.group] ?? 99;
      if (gA !== gB) return gA - gB;
      return (CRUD_ORDER[a.crud] ?? 99) - (CRUD_ORDER[b.crud] ?? 99);
    });
  }

  it("system tools sort before all other groups", () => {
    const sorted = sortTools(ALL_TOOLS);
    const firstGroup = sorted[0].group;
    assert.equal(firstGroup, "system");
  });

  it("within each group, read sorts before create before update before delete", () => {
    const sorted = sortTools(ALL_TOOLS);
    // Check each group's internal CRUD ordering
    const groups = {};
    for (const t of sorted) {
      if (!groups[t.group]) groups[t.group] = [];
      groups[t.group].push(t.crud);
    }
    for (const [group, cruds] of Object.entries(groups)) {
      for (let i = 1; i < cruds.length; i++) {
        assert.ok(
          CRUD_ORDER[cruds[i]] >= CRUD_ORDER[cruds[i - 1]],
          `Group "${group}": "${cruds[i]}" should not come before "${cruds[i - 1]}"`
        );
      }
    }
  });

  it("group order is: system, messages, folders, contacts, calendar, filters", () => {
    const sorted = sortTools(ALL_TOOLS);
    const seenGroups = [];
    for (const t of sorted) {
      if (!seenGroups.includes(t.group)) seenGroups.push(t.group);
    }
    assert.deepStrictEqual(seenGroups, ["system", "messages", "folders", "contacts", "calendar", "filters"]);
  });

  it("unknown group sorts to end", () => {
    const withUnknown = [...ALL_TOOLS, { name: "mystery", group: "alien", crud: "read" }];
    const sorted = sortTools(withUnknown);
    assert.equal(sorted[sorted.length - 1].name, "mystery");
  });

  it("unknown crud sorts to end within group", () => {
    const withUnknown = [...ALL_TOOLS, { name: "mystery", group: "messages", crud: "execute" }];
    const sorted = sortTools(withUnknown);
    const msgTools = sorted.filter(t => t.group === "messages");
    assert.equal(msgTools[msgTools.length - 1].name, "mystery");
  });
});

describe("Tool metadata: tools/list stripping", () => {
  /**
   * Mirrors the production tools/list stripping logic.
   * Input tools have internal fields; output should only have MCP-visible fields.
   */
  function stripForToolsList(toolsArray) {
    return toolsArray.map(({ name, description, inputSchema }) => ({ name, description, inputSchema }));
  }

  const ALLOWED_MCP_FIELDS = new Set(["name", "description", "inputSchema"]);

  it("strips group, crud, and title from tools with all internal fields", () => {
    // Create tools that HAVE internal metadata (unlike the old test which stripped first)
    const toolsWithMetadata = ALL_TOOLS.map(t => ({
      name: t.name,
      description: "test description",
      inputSchema: { type: "object" },
      group: t.group,
      crud: t.crud,
      title: "Some Title",
    }));

    const stripped = stripForToolsList(toolsWithMetadata);

    for (const tool of stripped) {
      assert.ok(!("group" in tool), `tools/list leaked "group" for ${tool.name}`);
      assert.ok(!("crud" in tool), `tools/list leaked "crud" for ${tool.name}`);
      assert.ok(!("title" in tool), `tools/list leaked "title" for ${tool.name}`);
    }
  });

  it("output contains only MCP-allowed fields (allowlist check)", () => {
    const toolsWithMetadata = [{
      name: "testTool", description: "desc", inputSchema: {},
      group: "messages", crud: "read", title: "Test", _internal: true, sortOrder: 5,
    }];

    const stripped = stripForToolsList(toolsWithMetadata);
    for (const tool of stripped) {
      for (const key of Object.keys(tool)) {
        assert.ok(ALLOWED_MCP_FIELDS.has(key), `Unexpected field "${key}" leaked in tools/list output`);
      }
    }
  });

  it("preserves name, description, and inputSchema values", () => {
    const schema = { type: "object", properties: { x: { type: "string" } } };
    const input = [{ name: "myTool", description: "my desc", inputSchema: schema, group: "system", crud: "read", title: "My" }];
    const stripped = stripForToolsList(input);
    assert.equal(stripped[0].name, "myTool");
    assert.equal(stripped[0].description, "my desc");
    assert.deepStrictEqual(stripped[0].inputSchema, schema);
  });
});

describe("Tool access: tag keyword validation", () => {
  // Mirrors production: allowlist of safe IMAP atom characters
  const VALID_TAG = /^[a-zA-Z0-9_$.\-]+$/;
  function sanitizeTags(tags) {
    return (tags || []).filter(t => typeof t === "string" && VALID_TAG.test(t));
  }

  it("rejects tags containing spaces", () => {
    const result = sanitizeTags(["$label1", "$label1 \\Deleted", "clean"]);
    assert.deepStrictEqual(result, ["$label1", "clean"]);
  });

  it("rejects tags containing tabs", () => {
    const result = sanitizeTags(["valid", "has\ttab"]);
    assert.deepStrictEqual(result, ["valid"]);
  });

  it("rejects tags containing newlines", () => {
    const result = sanitizeTags(["valid", "has\nnewline", "has\rnewline"]);
    assert.deepStrictEqual(result, ["valid"]);
  });

  it("accepts valid single-token tags", () => {
    const result = sanitizeTags(["$label1", "$label2", "project-x", "custom_tag"]);
    assert.deepStrictEqual(result, ["$label1", "$label2", "project-x", "custom_tag"]);
  });

  it("rejects tags with IMAP special characters", () => {
    const specials = [
      "tag(paren", "tag)paren", "tag{brace", "tag}brace",
      "tag*wild", "tag%wild", "tag\\backslash", 'tag"quote',
      "tag[bracket", "tag]bracket",
    ];
    const result = sanitizeTags(specials);
    assert.deepStrictEqual(result, [], "All IMAP special chars should be rejected");
  });

  it("rejects null bytes", () => {
    const result = sanitizeTags(["valid", "has\0null"]);
    assert.deepStrictEqual(result, ["valid"]);
  });

  it("rejects zero-width spaces and other unicode whitespace", () => {
    const result = sanitizeTags([
      "valid",
      "has\u200Bzwsp",     // zero-width space (not caught by \s)
      "has\u00A0nbsp",     // non-breaking space
      "has\uFEFFbom",      // BOM
    ]);
    assert.deepStrictEqual(result, ["valid"]);
  });

  it("rejects empty strings", () => {
    const result = sanitizeTags(["", "valid", ""]);
    assert.deepStrictEqual(result, ["valid"]);
  });

  it("rejects non-string entries", () => {
    const result = sanitizeTags([42, null, undefined, true, "valid", {}, []]);
    assert.deepStrictEqual(result, ["valid"]);
  });

  it("accepts tags with dots and hyphens", () => {
    const result = sanitizeTags(["project.v2", "work-item", "$label1", "tag_name"]);
    assert.deepStrictEqual(result, ["project.v2", "work-item", "$label1", "tag_name"]);
  });
});

// ── displayMessage validation tests ──────────────────────────────────

describe("displayMessage: input validation", () => {
  const VALID_DISPLAY_MODES = ["3pane", "tab", "window"];

  /**
   * Mirrors the production displayMode validation logic.
   * Returns { error } for invalid modes, { mode } for valid ones.
   */
  function validateDisplayMode(displayMode) {
    const mode = displayMode || "3pane";
    if (!VALID_DISPLAY_MODES.includes(mode)) {
      return { error: `Invalid displayMode: "${mode}". Must be one of: ${VALID_DISPLAY_MODES.join(", ")}` };
    }
    return { mode };
  }

  it("defaults to 3pane when displayMode is omitted", () => {
    assert.deepStrictEqual(validateDisplayMode(undefined), { mode: "3pane" });
    assert.deepStrictEqual(validateDisplayMode(null), { mode: "3pane" });
    assert.deepStrictEqual(validateDisplayMode(""), { mode: "3pane" });
  });

  it("accepts all three valid display modes", () => {
    for (const mode of VALID_DISPLAY_MODES) {
      const result = validateDisplayMode(mode);
      assert.equal(result.mode, mode);
      assert.ok(!result.error);
    }
  });

  it("rejects invalid displayMode values", () => {
    const invalid = ["popup", "fullscreen", "WINDOW", "Tab", "3PANE", "anything", "3 pane"];
    for (const mode of invalid) {
      const result = validateDisplayMode(mode);
      assert.ok(result.error, `Should reject displayMode: "${mode}"`);
      assert.match(result.error, /Invalid displayMode/);
    }
  });

  it("error message lists valid modes", () => {
    const result = validateDisplayMode("invalid");
    assert.match(result.error, /3pane/);
    assert.match(result.error, /tab/);
    assert.match(result.error, /window/);
  });

  it("displayMode is case-sensitive", () => {
    assert.ok(validateDisplayMode("3pane").mode);
    assert.ok(validateDisplayMode("Tab").error);
    assert.ok(validateDisplayMode("Window").error);
  });
});

// ── Build version parsing tests ──────────────────────────────────────

describe("Build version: options.js display parsing", () => {
  /**
   * Mirrors the production version parsing regex from options.js.
   * Input: git-describe string like "v0.2.0-7-g1461f1a+dirty"
   * Output: formatted display string
   */
  function parseBuildVersion(buildVersion, buildDate) {
    const m = buildVersion.match(/^(v[\d.]+)(?:-(\d+)-g([0-9a-f]+))?(\+dirty)?$/);
    let display;
    if (m) {
      const [, tag, commits, hash, dirty] = m;
      display = tag;
      if (commits && commits !== "0") display += ` +${commits}`;
      display += ` (${hash || tag})`;
      if (dirty) {
        display += " +dirty";
        if (buildDate) {
          display += " " + buildDate.replace("T", " ").replace(/\.\d+Z$/, " UTC");
        }
      }
    } else {
      display = buildVersion;
    }
    return display;
  }

  it("parses tag-only version (on exact tag)", () => {
    const result = parseBuildVersion("v0.2.0");
    assert.equal(result, "v0.2.0 (v0.2.0)");
  });

  it("parses tag with commits past (dev build)", () => {
    const result = parseBuildVersion("v0.2.0-7-g1461f1a");
    assert.equal(result, "v0.2.0 +7 (1461f1a)");
  });

  it("parses tag with zero commits past", () => {
    const result = parseBuildVersion("v0.2.0-0-gabcdef0");
    assert.equal(result, "v0.2.0 (abcdef0)");
  });

  it("parses dirty build without date", () => {
    const result = parseBuildVersion("v0.2.0-3-g1461f1a+dirty");
    assert.equal(result, "v0.2.0 +3 (1461f1a) +dirty");
  });

  it("parses dirty build with date", () => {
    const result = parseBuildVersion("v0.2.0-3-g1461f1a+dirty", "2026-03-15T10:30:00.000Z");
    assert.equal(result, "v0.2.0 +3 (1461f1a) +dirty 2026-03-15 10:30:00 UTC");
  });

  it("parses dirty tag-only (exact tag, dirty worktree)", () => {
    const result = parseBuildVersion("v0.2.0+dirty");
    assert.equal(result, "v0.2.0 (v0.2.0) +dirty");
  });

  it("falls back to raw string for non-matching format", () => {
    assert.equal(parseBuildVersion("1461f1a"), "1461f1a");
    assert.equal(parseBuildVersion("unknown"), "unknown");
    assert.equal(parseBuildVersion(""), "");
  });

  it("handles multi-digit version numbers", () => {
    const result = parseBuildVersion("v1.12.345-99-gabcdef0");
    assert.equal(result, "v1.12.345 +99 (abcdef0)");
  });
});

// ── Structural: ALL_TOOLS sync with production ───────────────────────

describe("Structural: test data matches production source", () => {
  it("ALL_TOOLS count matches tool definitions in api.js", () => {
    const apiPath = path.join(__dirname, "..", "extension", "mcp_server", "api.js");
    const src = fs.readFileSync(apiPath, "utf8");
    // Count tool name declarations in the tools array: name: "toolName"
    // Each tool object in the `const tools = [...]` array has a name field
    const nameMatches = src.match(/{\s*name:\s*"[a-zA-Z]+"/g);
    assert.ok(nameMatches, "Could not find tool name declarations in api.js");
    assert.equal(
      ALL_TOOLS.length,
      nameMatches.length,
      `Test ALL_TOOLS has ${ALL_TOOLS.length} tools but api.js has ${nameMatches.length} tool definitions. ` +
      `Update ALL_TOOLS when adding/removing tools.`
    );
  });

  it("ALL_TOOLS names match production tool names", () => {
    const apiPath = path.join(__dirname, "..", "extension", "mcp_server", "api.js");
    const src = fs.readFileSync(apiPath, "utf8");
    const nameMatches = src.match(/{\s*name:\s*"([a-zA-Z]+)"/g);
    const prodNames = nameMatches.map(m => m.match(/"([a-zA-Z]+)"/)[1]).sort();
    const testNames = ALL_TOOLS.map(t => t.name).sort();
    assert.deepStrictEqual(testNames, prodNames, "Test ALL_TOOLS names don't match production");
  });

  it("VALID_GROUPS matches production VALID_GROUPS", () => {
    const apiPath = path.join(__dirname, "..", "extension", "mcp_server", "api.js");
    const src = fs.readFileSync(apiPath, "utf8");
    const m = src.match(/const VALID_GROUPS\s*=\s*\[([^\]]+)\]/);
    assert.ok(m, "Could not find VALID_GROUPS in api.js");
    const prodGroups = m[1].match(/"([^"]+)"/g).map(s => s.replace(/"/g, "")).sort();
    const testGroups = [...VALID_GROUPS].sort();
    assert.deepStrictEqual(testGroups, prodGroups, "Test VALID_GROUPS doesn't match production");
  });

  it("VALID_CRUD matches production VALID_CRUD", () => {
    const apiPath = path.join(__dirname, "..", "extension", "mcp_server", "api.js");
    const src = fs.readFileSync(apiPath, "utf8");
    const m = src.match(/const VALID_CRUD\s*=\s*\[([^\]]+)\]/);
    assert.ok(m, "Could not find VALID_CRUD in api.js");
    const prodCrud = m[1].match(/"([^"]+)"/g).map(s => s.replace(/"/g, "")).sort();
    const testCrud = [...VALID_CRUD].sort();
    assert.deepStrictEqual(testCrud, prodCrud, "Test VALID_CRUD doesn't match production");
  });
});
