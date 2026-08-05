"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const apiSource = fs.readFileSync(
  path.resolve(__dirname, "../extension/mcp_server/api.js"),
  "utf8"
);
const organizerSource = fs.readFileSync(
  path.resolve(__dirname, "../scripts/categorize-mail.cjs"),
  "utf8"
);

describe("folder-local message key targeting", () => {
  it("returns messageKey in every search/recent result path", () => {
    const occurrences = apiSource.match(/messageKey: msgHdr\.messageKey/g) || [];
    assert.ok(occurrences.length >= 3);
  });

  it("declares single and bulk integer targets on updateMessage", () => {
    assert.match(apiSource, /messageKey: \{ type: "integer"/);
    assert.match(
      apiSource,
      /messageKeys: \{ type: "array", items: \{ type: "integer" \}/
    );
  });

  it("rejects mixed ID/key modes and resolves keys directly in the folder database", () => {
    assert.match(apiSource, /Use messageId\/messageIds or messageKey\/messageKeys, not both/);
    assert.match(apiSource, /db\.getMsgHdrForKey\(key\)/);
    assert.match(apiSource, /notFoundKeys/);
  });

  it("makes the organizer prefer unique message keys with an ID fallback", () => {
    assert.match(organizerSource, /Number\.isInteger\(message\.messageKey\)/);
    assert.match(organizerSource, /\? "messageKeys"/);
    assert.match(organizerSource, /: "messageIds"/);
    assert.match(organizerSource, /\[targetField\]: batch/);
  });
});
