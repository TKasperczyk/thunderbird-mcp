"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const Ci = {
  nsIImapIncomingServer: Symbol("nsIImapIncomingServer"),
  nsMsgFolderFlags: {
    Trash: 0x00000100,
    Drafts: 0x00000400,
  },
  nsMsgImapDeleteModels: {
    IMAPDelete: 0,
    MoveToTrash: 1,
    DeleteNoTrash: 2,
  },
};

function loadDeleteMessages(dependencies) {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "function isTrashOrDescendant(folder) {";
  const endMarker = "function updateMessage(";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start);
  assert.ok(start >= 0, "deleteMessages start marker missing");
  assert.ok(end > start, "deleteMessages end marker missing");

  const snippet = source.slice(start, end);
  const sandbox = {
    openFolder: dependencies.openFolder,
    findTrashFolder: dependencies.findTrashFolder,
    MailServices: dependencies.MailServices,
    Ci,
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.deleteMessages = deleteMessages;`,
    sandbox
  );
  return sandbox.deleteMessages;
}

function makeHarness({
  isDrafts = false,
  isTrash = false,
  isTrashDescendant = false,
  trashFolder = null,
  serverType = "imap",
  deleteModel = Ci.nsMsgImapDeleteModels.MoveToTrash,
} = {}) {
  const header = { messageId: "message-1@example.com" };
  const deleteCalls = [];
  const copyCalls = [];
  const trashLookups = [];
  const specialFolderChecks = [];
  const queryInterfaceCalls = [];
  const trashAncestor = isTrashDescendant ? {
    getFlag(flag) {
      return flag === Ci.nsMsgFolderFlags.Trash;
    },
    isSpecialFolder(flag, checkAncestors) {
      return this.getFlag(flag) || Boolean(checkAncestors && this.parent?.isSpecialFolder(flag, true));
    },
    parent: null,
  } : null;
  const folder = {
    getFlag(flag) {
      return (flag === Ci.nsMsgFolderFlags.Drafts && isDrafts) ||
        (flag === Ci.nsMsgFolderFlags.Trash && isTrash);
    },
    isSpecialFolder(flag, checkAncestors) {
      specialFolderChecks.push([flag, checkAncestors]);
      return this.getFlag(flag) || Boolean(checkAncestors && this.parent?.isSpecialFolder(flag, true));
    },
    parent: trashAncestor,
    server: {
      type: serverType,
      deleteModel,
      QueryInterface(interfaceType) {
        queryInterfaceCalls.push(interfaceType);
        if (interfaceType !== Ci.nsIImapIncomingServer) {
          throw new Error("Unexpected interface");
        }
        return this;
      },
    },
    deleteMessages(...args) {
      deleteCalls.push(args);
    },
  };
  const db = {
    getMsgHdrForMessageID(messageId) {
      return messageId === header.messageId ? header : null;
    },
    enumerateMessages() {
      return [header];
    },
  };
  const deleteMessages = loadDeleteMessages({
    openFolder: () => ({ folder, db }),
    findTrashFolder(sourceFolder) {
      trashLookups.push(sourceFolder);
      return trashFolder;
    },
    MailServices: {
      copy: {
        copyMessages(...args) {
          copyCalls.push(args);
        },
      },
    },
  });

  return {
    copyCalls,
    deleteCalls,
    deleteMessages,
    folder,
    header,
    queryInterfaceCalls,
    specialFolderChecks,
    trashLookups,
  };
}

function toPlainObject(value) {
  return JSON.parse(JSON.stringify(value));
}

describe("deleteMessages trash lookup", () => {
  it("fails without deleting when a non-Drafts folder has no reachable Trash", () => {
    const harness = makeHarness();

    const result = harness.deleteMessages(
      [harness.header.messageId],
      "imap://account/INBOX"
    );

    assert.deepStrictEqual(toPlainObject(result), { error: "Trash folder not found" });
    assert.deepStrictEqual(harness.queryInterfaceCalls, [Ci.nsIImapIncomingServer]);
    assert.equal(harness.trashLookups.length, 1);
    assert.equal(harness.trashLookups[0], harness.folder);
    assert.equal(harness.deleteCalls.length, 0);
    assert.equal(harness.copyCalls.length, 0);
  });

  it("keeps the regular move-style delete when Trash is reachable", () => {
    const trashFolder = { URI: "imap://account/Trash" };
    const harness = makeHarness({ trashFolder });

    const result = harness.deleteMessages(
      [harness.header.messageId],
      "imap://account/INBOX"
    );

    assert.deepStrictEqual(toPlainObject(result), { success: true, deleted: 1 });
    assert.deepStrictEqual(harness.queryInterfaceCalls, [Ci.nsIImapIncomingServer]);
    assert.equal(harness.trashLookups.length, 1);
    assert.equal(harness.deleteCalls.length, 1);
    assert.equal(harness.deleteCalls[0][0][0], harness.header);
    assert.deepStrictEqual(harness.deleteCalls[0].slice(1), [null, false, true, null, false]);
    assert.equal(harness.copyCalls.length, 0);
  });

  for (const [location, options, folderPath] of [
    ["Trash", { isTrash: true }, "imap://account/Trash"],
    ["a Trash descendant", { isTrashDescendant: true }, "imap://account/Trash/Nested"],
  ]) {
    it(`allows permanent deletion from ${location} without another Trash folder`, () => {
      const harness = makeHarness(options);

      const result = harness.deleteMessages(
        [harness.header.messageId],
        folderPath
      );

      assert.deepStrictEqual(toPlainObject(result), { success: true, deleted: 1 });
      assert.deepStrictEqual(harness.specialFolderChecks, [[Ci.nsMsgFolderFlags.Trash, true]]);
      assert.equal(harness.queryInterfaceCalls.length, 0);
      assert.equal(harness.trashLookups.length, 0);
      assert.equal(harness.deleteCalls.length, 1);
      assert.deepStrictEqual(harness.deleteCalls[0].slice(1), [null, false, true, null, false]);
      assert.equal(harness.copyCalls.length, 0);
    });
  }

  for (const [modelName, deleteModel] of [
    ["IMAPDelete", Ci.nsMsgImapDeleteModels.IMAPDelete],
    ["DeleteNoTrash", Ci.nsMsgImapDeleteModels.DeleteNoTrash],
  ]) {
    it(`allows the IMAP ${modelName} model to delete without Trash`, () => {
      const harness = makeHarness({ deleteModel });

      const result = harness.deleteMessages(
        [harness.header.messageId],
        "imap://account/INBOX"
      );

      assert.deepStrictEqual(toPlainObject(result), { success: true, deleted: 1 });
      assert.deepStrictEqual(harness.queryInterfaceCalls, [Ci.nsIImapIncomingServer]);
      assert.equal(harness.trashLookups.length, 0);
      assert.equal(harness.deleteCalls.length, 1);
      assert.deepStrictEqual(harness.deleteCalls[0].slice(1), [null, false, true, null, false]);
      assert.equal(harness.copyCalls.length, 0);
    });
  }

  it("keeps the Drafts fallback delete when Trash is unavailable", () => {
    const harness = makeHarness({ isDrafts: true });

    const result = harness.deleteMessages(
      [harness.header.messageId],
      "imap://account/Drafts"
    );

    assert.deepStrictEqual(toPlainObject(result), { success: true, deleted: 1 });
    assert.equal(harness.trashLookups.length, 1);
    assert.equal(harness.deleteCalls.length, 1);
    assert.equal(harness.copyCalls.length, 0);
  });

  it("keeps the explicit Drafts move when Trash is reachable", () => {
    const trashFolder = { URI: "imap://account/Trash" };
    const harness = makeHarness({ isDrafts: true, trashFolder });

    const result = harness.deleteMessages(
      [harness.header.messageId],
      "imap://account/Drafts"
    );

    assert.deepStrictEqual(toPlainObject(result), {
      success: true,
      deleted: 1,
      movedToTrash: true,
    });
    assert.equal(harness.deleteCalls.length, 0);
    assert.equal(harness.copyCalls.length, 1);
    assert.equal(harness.copyCalls[0][0], harness.folder);
    assert.equal(harness.copyCalls[0][1][0], harness.header);
    assert.equal(harness.copyCalls[0][2], trashFolder);
    assert.deepStrictEqual(harness.copyCalls[0].slice(3), [true, null, null, false]);
  });
});
