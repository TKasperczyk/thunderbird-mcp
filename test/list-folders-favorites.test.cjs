"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { loadFolderTools } = require("./helpers/folder-tools.cjs");

function makeHarness() {
  const server = { type: "imap" };
  const hiddenServer = { type: "imap" };
  function folder(name, flags = 0, subFolders = [], folderServer = server) {
    return {
      name, prettyName: name, flags, URI: `imap://${folderServer === server ? "allowed" : "hidden"}/${name}`,
      server: folderServer, subFolders, hasSubFolders: subFolders.length > 0,
      getTotalMessages: () => 42, getNumUnread: () => 3,
    };
  }
  const favorite = folder("ne\u202Ested", 0x88082014);
  const parent = folder("Projects", 0, [favorite]);
  const publicFolder = folder("Public", 0x00100000);
  const inbox = folder("INBOX", 0x00001000);
  const virtual = folder("Saved search", 0x80000020);
  server.rootFolder = folder("root", 0, [inbox, parent, publicFolder, virtual]);
  const hidden = folder("Hidden favorite", 0x80000000, [], hiddenServer);
  hiddenServer.rootFolder = folder("Hidden root", 0, [hidden], hiddenServer);
  const accounts = [
    { key: "allowed", incomingServer: server },
    { key: "hidden", incomingServer: hiddenServer },
  ];
  const folders = [server.rootFolder, inbox, parent, favorite, publicFolder, virtual, hiddenServer.rootFolder, hidden];
  const runtime = loadFolderTools({ accounts, folders, isAccountAllowed: key => key === "allowed" });
  return { runtime, favorite, parent, publicFolder, inbox, hidden };
}

describe("production listFolders favorites", () => {
  it("walks non-favorite parents and returns favorites at their original depth", async () => {
    const { runtime, favorite } = makeHarness();
    const result = await runtime.callTool("listFolders", { favoritesOnly: true });
    assert.equal(result.length, 1);
    assert.equal(result[0].path, favorite.URI);
    assert.equal(result[0].depth, 1);
    assert.equal(result[0].isFavorite, true);
    assert.equal(result[0].name, "nested");
  });

  it("includes every accessible non-virtual folder and reports both flag states by default", async () => {
    const { runtime, favorite, publicFolder } = makeHarness();
    for (const favoritesOnly of [undefined, false]) {
      const result = await runtime.callTool("listFolders", { favoritesOnly });
      assert.equal(result.length, 4);
      assert.equal(result.find(row => row.path === favorite.URI).isFavorite, true);
      assert.equal(result.find(row => row.path === publicFolder.URI).isFavorite, false);
      assert.equal(result.filter(row => row.isFavorite).length, 1);
    }
  });

  it("supports folder and account scopes without exposing a restricted account", async () => {
    const { runtime, parent, hidden } = makeHarness();
    assert.equal((await runtime.callTool("listFolders", { folderPath: parent.URI, favoritesOnly: true })).length, 1);
    assert.equal((await runtime.callTool("listFolders", { accountId: "allowed", favoritesOnly: true })).length, 1);
    assert.match((await runtime.callTool("listFolders", { accountId: "hidden", favoritesOnly: true })).error, /not accessible/);
    assert.match((await runtime.callTool("listFolders", { folderPath: hidden.URI, favoritesOnly: true })).error, /not accessible/);
    assert.match((await runtime.callTool("listFolders", { folderPath: "missing", favoritesOnly: true })).error, /not found/);
  });

  it("returns an empty table with stable columns when no favorites match", async () => {
    const { runtime, inbox } = makeHarness();
    const result = await runtime.callTool("listFolders", { folderPath: inbox.URI, favoritesOnly: true, format: "table" });
    assert.equal(result.rows.length, 0);
    assert.ok(result.columns.includes("isFavorite"));
  });

  it("returns the favorite column in table format and sanitizes only display text", async () => {
    const { runtime, favorite } = makeHarness();
    const result = await runtime.callTool("listFolders", { favoritesOnly: true, format: "table" });
    assert.equal(result.rows.length, 1);
    assert.equal(result.rows[0][result.columns.indexOf("isFavorite")], true);
    assert.equal(result.rows[0][result.columns.indexOf("path")], favorite.URI);
    assert.equal(result.rows[0][result.columns.indexOf("name")], "nested");
  });
});
