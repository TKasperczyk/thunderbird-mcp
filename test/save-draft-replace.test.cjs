"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

// saveDraft itself needs XPCOM and cannot run here. What CAN be checked without
// Thunderbird is the wiring, and that is where the silent defect lives: a
// parameter declared in inputSchema but never forwarded at the dispatch switch
// is accepted by the client, validated, and then dropped -- the call succeeds
// and quietly does the old thing. Reading the real api.js is the point; a
// mirrored copy of the dispatch line would move with the bug.
const API_PATH = path.join(__dirname, "..", "extension", "mcp_server", "api.js");
const source = fs.readFileSync(API_PATH, "utf8");

function saveDraftToolDefinition() {
  const start = source.indexOf('name: "saveDraft"');
  assert.notEqual(start, -1, "saveDraft tool definition not found in api.js");
  // Definitions are separated by the next tool's name key.
  const next = source.indexOf("name: \"", start + 20);
  return source.slice(start, next === -1 ? source.length : next);
}

function saveDraftDispatchLine() {
  const marker = 'case "saveDraft":';
  const start = source.indexOf(marker);
  assert.notEqual(start, -1, "saveDraft dispatch case not found in api.js");
  const end = source.indexOf(";", start + marker.length);
  return source.slice(start, end);
}

describe("saveDraft replace-draft wiring", () => {
  const params = ["replaceMessageId", "replaceFolderPath"];

  it("declares both replace parameters in the tool schema", () => {
    const def = saveDraftToolDefinition();
    for (const p of params) {
      assert.ok(def.includes(`${p}:`), `inputSchema is missing ${p}`);
    }
  });

  it("forwards both replace parameters at the dispatch switch", () => {
    const line = saveDraftDispatchLine();
    for (const p of params) {
      assert.ok(
        line.includes(`args.${p}`),
        `dispatch drops ${p}: a caller can set it and nothing happens`
      );
    }
  });

  // Known-negative control: proves the assertions above discriminate rather
  // than matching anything that happens to be in the file. A parameter that
  // was never added must NOT be reported as wired.
  it("does not report an absent parameter as wired", () => {
    const def = saveDraftToolDefinition();
    const line = saveDraftDispatchLine();
    assert.ok(!def.includes("replaceNothingAtAll:"));
    assert.ok(!line.includes("args.replaceNothingAtAll"));
  });

  it("scopes the reads to saveDraft, not the whole file", () => {
    // A slice that ran to EOF would make every assertion pass for the wrong
    // reason -- some other tool's schema would satisfy it.
    const def = saveDraftToolDefinition();
    assert.ok(def.length < source.length / 4, "tool definition slice is too wide to be saveDraft's");
    assert.ok(!def.includes('name: "listCalendars"'), "slice leaked into the next tool definition");
  });
});
