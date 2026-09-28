"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { loadFolderTools } = require("./helpers/folder-tools.cjs");

function makeHarness() {
  const server = { type: "imap" };
  const folder = {
    server, prettyName: "INBOX", URI: "imap://account/INBOX", flags: 0x00001000,
    hasSubFolders: false, getTotalMessages: () => 42, getNumUnread: () => 3,
  };
  server.rootFolder = { hasSubFolders: true, subFolders: [folder] };
  return loadFolderTools({ accounts: [{ key: "account1", incomingServer: server }], folders: [folder] });
}

describe("production listFolders table format", () => {
  const columns = ["accountId", "depth", "isFavorite", "name", "path", "totalMessages", "type", "unreadMessages"];

  it("returns the object array for omitted, null, and objects format", () => {
    const runtime = makeHarness();
    for (const format of [undefined, null, "objects"]) {
      const result = runtime.listFolders(undefined, undefined, format);
      assert.equal(result.length, 1);
      assert.equal(result[0].isFavorite, false);
      assert.equal(result[0].type, "inbox");
      assert.equal(result[0].accountId, "account1");
    }
  });

  it("sorts stable columns and orders row values to match", () => {
    const runtime = makeHarness();
    const result = runtime.listFolders(undefined, undefined, "table");
    const objects = runtime.listFolders();
    assert.deepEqual(Array.from(result.columns), columns);
    assert.deepEqual(Array.from(result.rows[0]), columns.map(column => objects[0][column]));
    assert.deepEqual(Object.keys(result), ["columns", "rows"]);
  });

  it("preserves the columns in an empty table", () => {
    const result = makeHarness().listFolders(undefined, undefined, "table", true);
    assert.deepEqual(Array.from(result.columns), columns);
    assert.equal(result.rows.length, 0);
  });

  it("rejects an unknown format", () => {
    assert.match(makeHarness().listFolders(undefined, undefined, "compact").error, /Invalid format/);
  });
});
