"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { loadFolderTools } = require("./helpers/folder-tools.cjs");

const OUTBOX_FLAG = 0x800; // nsMsgFolderFlags.Queue

function makeHarness({ crossAccount = false, denied = [], copyError = false } = {}) {
  const mutations = [];
  const copies = [];
  const folderMoves = [];
  const server = { type: "imap" };
  const otherServer = { type: "imap" };
  const headers = [{ messageId: "one" }, { messageId: "two" }];
  const folder = {
    server, URI: "imap://one/Sent", hasSubFolders: false,
    updateFolder() {},
    msgDatabase: {
      getMsgHdrForMessageID: id => headers.find(header => header.messageId === id),
      enumerateMessages: () => headers,
    },
    markMessagesRead: (hdrs, value) => mutations.push(["read", hdrs, value]),
    markMessagesFlagged: (hdrs, value) => mutations.push(["flagged", hdrs, value]),
    addKeywordsToMessages: (hdrs, value) => mutations.push(["addTags", hdrs, value]),
    removeKeywordsFromMessages: (hdrs, value) => mutations.push(["removeTags", hdrs, value]),
  };
  const target = { server: crossAccount ? otherServer : server, URI: "imap://target/Project", hasSubFolders: false };
  const trash = { server, URI: "imap://one/Trash", hasSubFolders: false, getFlag: flag => flag === 0x100 };
  const outbox = { server, URI: "mailbox://one/Unsent%20Messages", hasSubFolders: false, isSpecialFolder: flag => flag === OUTBOX_FLAG };
  const nestedOutbox = {
    server, URI: "mailbox://one/Unsent%20Messages/Nested", hasSubFolders: false,
    isSpecialFolder: (flag, ancestors) => ancestors === true && flag === OUTBOX_FLAG,
  };
  const root = { server, URI: "imap://one", isServer: true, hasSubFolders: true, subFolders: [folder, trash, outbox] };
  server.rootFolder = root;
  const accounts = [{ key: "one", incomingServer: server }, { key: "two", incomingServer: otherServer }];
  const runtime = loadFolderTools({
    accounts, folders: [root, folder, target, trash, outbox, nestedOutbox], isAccountAllowed: key => !denied.includes(key),
    copyMessages(...args) {
      if (copyError) throw new Error("Copy refused");
      copies.push(args);
    },
    copyFolder(...args) { folderMoves.push(args); },
  });
  async function update(args) {
    return runtime.callTool("updateMessage", { messageId: "one", folderPath: folder.URI, ...args });
  }
  return { runtime, update, mutations, copies, folderMoves, root, folder, target, trash, outbox, nestedOutbox, headers };
}

describe("production updateMessage tag keys", () => {
  it("passes Thunderbird escape keys and the other permitted ASCII symbols unchanged", async () => {
    const h = makeHarness();
    const keys = ["my=20project", "=c3=a9", "$label1", "&AOQ-", "a:b/c@d+e!f'g,h?i#j$k^l_m`n|o~p.q-r"];
    const result = await h.update({ addTags: keys, removeTags: ["old=20tag"] });
    assert.equal(result.success, true);
    assert.equal(result.updated, 1);
    assert.equal(h.mutations[0][2], keys.join(" "));
    assert.equal(h.mutations[1][2], "old=20tag");
    assert.deepEqual(Array.from(result.actions[0].value), keys);
  });

  it("checks every ASCII character against Thunderbird's keyword contract", async () => {
    const blocked = '()[]{}%*"\\<>;';
    for (let code = 0; code < 128; code++) {
      const h = makeHarness();
      const char = String.fromCharCode(code);
      const result = await h.update({ addTags: [`a${char}b`] });
      const allowed = code > 32 && code < 127 && !blocked.includes(char);
      assert.equal(result.success === true, allowed, `ASCII ${code}`);
      assert.equal(h.mutations.length, allowed ? 1 : 0, `ASCII ${code}`);
    }
  });

  it("rejects mixed valid/invalid arrays before read, flags, tags, or copy mutate anything", async () => {
    for (const field of ["addTags", "removeTags"]) {
      for (const tag of ["", "has space", "tag%bad", "tag]bad", "tag\n", "tag\0", "é", "\u200B", "😀", 42, null, {}, []]) {
        const h = makeHarness();
        const result = await h.update({ read: true, flagged: true, addTags: ["valid"], [field]: ["valid", tag], copyTo: h.target.URI });
        assert.match(result.error, new RegExp(`${field} contains invalid Thunderbird tag keys`));
        assert.ok(result.error.includes(h.runtime.stripInvisibleCharacters(JSON.stringify(tag))), `Rejected value must be named: ${JSON.stringify(tag)}`);
        assert.equal(result.success, undefined);
        assert.equal(h.mutations.length, 0);
        assert.equal(h.copies.length, 0);
      }
    }
  });

  it("supports JSON-encoded arrays and reports invalid array inputs", async () => {
    const h = makeHarness();
    assert.equal((await h.update({ addTags: '["my=20project"]', removeTags: '[]' })).success, true);
    for (const field of ["addTags", "removeTags"]) {
      for (const value of ["not-json", "{}", null, 42]) {
        assert.match((await h.update({ [field]: value })).error, /must be an array/);
      }
    }
  });

  it("handles empty tag arrays without inventing a tag action", async () => {
    const h = makeHarness();
    const result = await h.update({ addTags: [], removeTags: [], read: true });
    assert.equal(result.actions.length, 1);
    assert.equal(result.actions[0].type, "read");
  });
});

describe("production updateMessage copyTo", () => {
  for (const crossAccount of [false, true]) {
    it(`submits a non-moving ${crossAccount ? "cross-account" : "same-account"} bulk copy`, async () => {
      const h = makeHarness({ crossAccount });
      const result = await h.update({ messageId: undefined, messageIds: ["one", "two", "missing"], copyTo: h.target.URI });
      assert.equal(result.success, true);
      assert.equal(result.updated, 2);
      assert.deepEqual(Array.from(result.notFound), ["missing"]);
      assert.equal(result.actions[0].type, "copy");
      assert.equal(result.actions[0].to, h.target.URI);
      assert.equal(h.copies.length, 1);
      assert.equal(h.copies[0][0], h.folder);
      assert.deepEqual(Array.from(h.copies[0][1]), h.headers);
      assert.equal(h.copies[0][2], h.target);
      assert.deepEqual(h.copies[0].slice(3), [false, null, null, false]);
    });
  }

  it("keeps existing move and trash calls moving", async () => {
    for (const action of ["moveTo", "trash"]) {
      const h = makeHarness();
      const result = await h.update(action === "moveTo" ? { moveTo: h.target.URI } : { trash: true });
      assert.equal(result.success, true);
      assert.equal(result.actions[0].type, "move");
      assert.equal(h.copies[0][3], true);
      assert.equal(h.copies[0][2], action === "moveTo" ? h.target : h.trash);
    }
  });

  it("refuses source or destination account restrictions before any mutation", async () => {
    for (const denied of [["one"], ["two"]]) {
      const h = makeHarness({ crossAccount: true, denied });
      const result = await h.update({ copyTo: h.target.URI, read: true, addTags: ["valid"] });
      assert.match(result.error, /not accessible/);
      assert.equal(h.copies.length, 0);
      assert.equal(h.mutations.length, 0);
    }
  });

  it("refuses a missing destination for both copy and move before changing tags", async () => {
    for (const field of ["copyTo", "moveTo"]) {
      const h = makeHarness();
      assert.match((await h.update({ [field]: "missing", addTags: ["valid"] })).error, /not found/);
      assert.equal(h.mutations.length, 0);
      assert.equal(h.copies.length, 0);
    }
  });

  it("rejects conflicting operations and malformed copy destinations", async () => {
    for (const args of [
      { copyTo: "target", moveTo: "target" }, { copyTo: "target", trash: true },
      { copyTo: "target", trash: "true" }, { moveTo: "target", trash: true },
      { copyTo: "" }, { copyTo: null }, { copyTo: 3 },
    ]) {
      const h = makeHarness();
      assert.ok((await h.update({ read: true, ...args })).error);
      assert.equal(h.mutations.length, 0);
      assert.equal(h.copies.length, 0);
    }
  });

  it("does not copy when no messages match and returns copy service errors", async () => {
    const h = makeHarness();
    assert.match((await h.update({ messageId: "missing", copyTo: h.target.URI })).error, /No matching messages/);
    assert.equal(h.copies.length, 0);
    const failing = makeHarness({ copyError: true });
    assert.match((await failing.update({ copyTo: failing.target.URI })).error, /Copy refused/);
  });

  it("allows read/tag changes with copying and reports the IMAP tag caveat", async () => {
    const h = makeHarness();
    const result = await h.update({ read: true, addTags: ["my=20project"], copyTo: h.target.URI, trash: false });
    assert.deepEqual(Array.from(result.actions, action => action.type), ["read", "addTags", "copy"]);
    assert.match(result.warning, /destination copy/);
    assert.equal(h.copies[0][3], false);
  });
});

describe("the Outbox is never a destination", () => {
  for (const action of ["moveTo", "copyTo"]) {
    it(`refuses ${action} into the Outbox or a folder below it before any change`, async () => {
      for (const key of ["outbox", "nestedOutbox"]) {
        const h = makeHarness();
        const result = await h.update({ [action]: h[key].URI, read: true, flagged: true, addTags: ["valid"] });
        assert.match(result.error, /Outbox \(Unsent Messages\) cannot be used as a move or copy destination/);
        assert.equal(result.success, undefined);
        assert.equal(h.mutations.length, 0);
        assert.equal(h.copies.length, 0);
      }
    });
  }

  it("refuses creating a folder inside the Outbox", () => {
    for (const key of ["outbox", "nestedOutbox"]) {
      const h = makeHarness();
      let created = 0;
      h[key].createSubfolder = () => { created++; };
      const result = h.runtime.createFolder(h[key].URI, "Queued");
      assert.match(result.error, /Cannot create a folder inside the Outbox/);
      assert.equal(created, 0);
    }
    const h = makeHarness();
    let created = 0;
    h.target.createSubfolder = () => { created++; };
    assert.equal(h.runtime.createFolder(h.target.URI, "Project B").success, true);
    assert.equal(created, 1);
  });

  it("refuses moving a folder into the Outbox", () => {
    for (const key of ["outbox", "nestedOutbox"]) {
      const h = makeHarness();
      const result = h.runtime.moveFolder(h.target.URI, h[key].URI);
      assert.match(result.error, /Cannot move a folder into the Outbox/);
      assert.equal(h.folderMoves.length, 0);
    }
    const h = makeHarness();
    assert.equal(h.runtime.moveFolder(h.target.URI, h.folder.URI).success, true);
    assert.equal(h.folderMoves.length, 1);
  });
});

describe("folders never take the Outbox role by name", () => {
  const outboxNames = ["Unsent Messages", "unsent messages", "Outbox", "OUTBOX", " Outbox "];

  it("refuses creating a top-level folder with an Outbox name", () => {
    for (const name of outboxNames) {
      const h = makeHarness();
      let created = 0;
      h.root.createSubfolder = () => { created++; };
      assert.match(h.runtime.createFolder(h.root.URI, name).error, /would become the Outbox/, name);
      assert.equal(created, 0);
    }
  });

  it("refuses renaming a top-level folder to an Outbox name", () => {
    for (const name of outboxNames) {
      const h = makeHarness();
      let renamed = 0;
      h.target.parent = h.root;
      h.target.rename = () => { renamed++; };
      assert.match(h.runtime.renameFolder(h.target.URI, name).error, /would become the Outbox/, name);
      assert.equal(renamed, 0);
    }
  });

  it("refuses moving a folder with an Outbox name to the top level", () => {
    for (const name of ["Unsent Messages", "outbox"]) {
      const h = makeHarness();
      h.target.name = name;
      h.target.prettyName = name;
      assert.match(h.runtime.moveFolder(h.target.URI, h.root.URI).error, /would become the Outbox/, name);
      assert.equal(h.folderMoves.length, 0);
    }
  });

  it("allows the same names below the top level and other names at the top level", () => {
    const h = makeHarness();
    let created = 0;
    let renamed = 0;
    h.folder.createSubfolder = () => { created++; };
    h.root.createSubfolder = () => { created++; };
    h.target.parent = h.folder;
    h.target.rename = () => { renamed++; };
    h.target.name = "Outbox";
    assert.equal(h.runtime.createFolder(h.folder.URI, "Outbox").success, true);
    assert.equal(h.runtime.createFolder(h.root.URI, "Outbox archive").success, true);
    assert.equal(h.runtime.renameFolder(h.target.URI, "Unsent Messages").success, true);
    assert.equal(h.runtime.moveFolder(h.target.URI, h.trash.URI).success, true);
    assert.equal(created, 2);
    assert.equal(renamed, 1);
    assert.equal(h.folderMoves.length, 1);
  });
});
