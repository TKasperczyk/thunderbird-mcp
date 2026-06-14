const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

// Mirrors the pure listFolders table formatting logic from api.js. The
// production module runs in Thunderbird/XPCOM and cannot be required directly.
const FOLDER_KEYS = ["name", "path", "type", "accountId", "totalMessages", "unreadMessages", "depth"];

function toColumnarTable(items, keys) {
  const columns = Array.from(keys).sort();
  return {
    columns,
    rows: items.map(item => columns.map(column => item[column])),
  };
}

function formatFolderResults(results, format) {
  const outputFormat = format == null ? "objects" : format;
  if (outputFormat !== "objects" && outputFormat !== "table") {
    return { error: `Invalid format: "${outputFormat}". Must be one of: objects, table` };
  }
  return outputFormat === "table" ? toColumnarTable(results, FOLDER_KEYS) : results;
}

describe("listFolders table format", () => {
  const folders = [
    {
      name: "INBOX",
      path: "imap://user@example.com/INBOX",
      type: "inbox",
      accountId: "account1",
      totalMessages: 42,
      unreadMessages: 3,
      depth: 0,
    },
    {
      name: "Projects",
      path: "imap://user@example.com/INBOX/Projects",
      type: "folder",
      accountId: "account1",
      totalMessages: 7,
      unreadMessages: 1,
      depth: 1,
    },
  ];

  it("returns the existing array unchanged when format is omitted, null, or objects", () => {
    assert.strictEqual(formatFolderResults(folders, undefined), folders);
    assert.strictEqual(formatFolderResults(folders, null), folders);
    assert.strictEqual(formatFolderResults(folders, "objects"), folders);
  });

  it("returns the table shape when requested", () => {
    const result = formatFolderResults(folders, "table");
    assert.deepStrictEqual(Object.keys(result), ["columns", "rows"]);
    assert.equal(result.rows.length, 2);
  });

  it("sorts columns alphabetically from the stable folder key set", () => {
    const result = formatFolderResults(folders, "table");
    assert.deepStrictEqual(result.columns, [
      "accountId",
      "depth",
      "name",
      "path",
      "totalMessages",
      "type",
      "unreadMessages",
    ]);
  });

  it("orders row values to match the columns", () => {
    const result = formatFolderResults(folders, "table");
    assert.deepStrictEqual(result.rows[0], result.columns.map(column => folders[0][column]));
    assert.deepStrictEqual(result.rows[1], result.columns.map(column => folders[1][column]));
  });

  it("keeps stable columns for an empty table", () => {
    const result = formatFolderResults([], "table");
    assert.deepStrictEqual(result.columns, [
      "accountId",
      "depth",
      "name",
      "path",
      "totalMessages",
      "type",
      "unreadMessages",
    ]);
    assert.deepStrictEqual(result.rows, []);
  });

  it("rejects unknown format values", () => {
    assert.deepStrictEqual(formatFolderResults(folders, "compact"), {
      error: 'Invalid format: "compact". Must be one of: objects, table',
    });
  });
});
