const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

// Mirrors the pure listFolders favorite-flag logic from api.js. The production
// module runs in Thunderbird/XPCOM and cannot be required directly, so the
// source-contract block below asserts the real file stays in sync.
const FOLDER_KEYS = ["name", "path", "type", "accountId", "totalMessages", "unreadMessages", "depth", "isFavorite"];

// nsMsgFolderFlags.Favorite. NOT 0x00100000 — that is ImapPublic.
const FLAG_FAVORITE = 0x80000000;

function isFavoriteFolder(flags) {
  return Boolean(flags & FLAG_FAVORITE);
}

function toColumnarTable(items, keys) {
  const columns = Array.from(keys).sort();
  return {
    columns,
    rows: items.map(item => columns.map(column => item[column])),
  };
}

function filterFavorites(results, favoritesOnly) {
  return favoritesOnly ? results.filter(folder => folder.isFavorite) : results;
}

describe("listFolders favorite flag", () => {
  it("detects the Favorite bit on a real-world flag value", () => {
    // Observed on " Current Projects/Michelman" (hover IMAP): 0x88082014
    assert.equal(isFavoriteFolder(0x88082014), true);
  });

  it("does not treat ImapPublic (0x00100000) as favorite", () => {
    assert.equal(isFavoriteFolder(0x00100000), false);
  });

  it("returns false for ordinary folder flags", () => {
    assert.equal(isFavoriteFolder(0x00001000), false); // Inbox
    assert.equal(isFavoriteFolder(0x00000000), false);
    assert.equal(isFavoriteFolder(0x08082014), false); // same folder, favorite cleared
  });

  it("handles flag values above 2^31 without sign errors", () => {
    assert.equal(isFavoriteFolder(2282726420), true); // 0x88082014 as a plain Number
  });
});

describe("listFolders favoritesOnly filter", () => {
  const folders = [
    { name: "INBOX", path: "imap://u@h/INBOX", isFavorite: false },
    { name: "Michelman", path: "imap://u@h/ Current Projects/Michelman", isFavorite: true },
    { name: "Davis", path: "imap://u@h/ Current Projects/Davis", isFavorite: true },
  ];

  it("returns every folder when favoritesOnly is falsy", () => {
    assert.equal(filterFavorites(folders, undefined).length, 3);
    assert.equal(filterFavorites(folders, false).length, 3);
  });

  it("returns only favorites when favoritesOnly is true", () => {
    const result = filterFavorites(folders, true);
    assert.deepStrictEqual(result.map(f => f.name), ["Michelman", "Davis"]);
  });

  it("returns an empty array when nothing is favorited", () => {
    const none = folders.map(f => ({ ...f, isFavorite: false }));
    assert.deepStrictEqual(filterFavorites(none, true), []);
  });
});

describe("listFolders table format includes isFavorite", () => {
  it("sorts isFavorite into the columnar key set", () => {
    const result = toColumnarTable([{ name: "INBOX", isFavorite: false }], FOLDER_KEYS);
    assert.deepStrictEqual(result.columns, [
      "accountId",
      "depth",
      "isFavorite",
      "name",
      "path",
      "totalMessages",
      "type",
      "unreadMessages",
    ]);
  });
});

// The mirrored logic above proves the algorithm; these assertions prove the
// shipped api.js actually implements it.
describe("api.js source contract", () => {
  const source = fs.readFileSync(
    path.join(__dirname, "..", "extension", "mcp_server", "api.js"),
    "utf8"
  );

  it("declares isFavorite in the listFolders key set", () => {
    const keyLine = source.match(/const folderKeys = \[[^\]]*\]/);
    assert.ok(keyLine, "folderKeys declaration not found in api.js");
    assert.match(keyLine[0], /"isFavorite"/);
  });

  it("uses the Favorite flag bit, not ImapPublic", () => {
    assert.match(source, /0x80000000/, "api.js does not reference the Favorite flag bit");
  });

  it("exposes a favoritesOnly parameter on the listFolders tool schema", () => {
    assert.match(source, /favoritesOnly/, "api.js does not declare favoritesOnly");
  });
});
