"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
const DRAFTS_URI = "imap://user@example.test/Drafts";
const OTHER_DRAFTS_URI = "imap://other@example.test/Drafts";
const Ci = {
  nsIMsgCompType: { New: 0, Draft: 9 },
  nsIMsgCompDeliverMode: { Now: 0, SaveAsDraft: 4 },
  nsMsgFolderFlags: { Drafts: 0x400 },
  nsMsgMessageFlags: { IMAPDeleted: 0x200000 },
  nsITimer: { TYPE_ONE_SHOT: 0 },
};
const NEW_DRAFT_ID = "new-draft-id@example.test";

/**
 * An IMAP Drafts folder whose local database catches up with the server only
 * on updateFolderWithListener, as Thunderbird's does. serverEffects says what
 * saving the new draft did on the server: store it (newStored) and what
 * happened to the replaced draft ("expunged", "marked-deleted", "kept").
 */
function makeImapDrafts(calls, { newStored = true, oldOnServer = "expunged", moveRemovesOld = true, syncStatus = 0, syncStatuses = null, dbThrowsAfterSyncs = null, droppedSyncs = 0 } = {}) {
  const local = new Map();
  const server = new Map();
  let syncCount = 0;
  const database = {
    containsKey: (key) => local.has(key),
    getMsgHdrForKey: (key) => local.get(key) || null,
    getMsgHdrForMessageID: (id) => [...local.values()].find(h => h.messageId === id) || null,
  };
  const folder = {
    URI: DRAFTS_URI,
    server: { type: "imap" },
    getFlag(flag) { calls.flags.push(flag); return true; },
    get msgDatabase() {
      if (dbThrowsAfterSyncs !== null && syncCount >= dbThrowsAfterSyncs) throw new Error("database is being rebuilt");
      return database;
    },
    updateFolderWithListener(_window, listener) {
      calls.syncs++;
      // Measured live: requests right after a save are never answered.
      if (calls.syncs <= droppedSyncs) return;
      const status = syncStatuses ? syncStatuses[syncCount] ?? 0 : syncStatus;
      syncCount++;
      if (status === 0) {
        local.clear();
        for (const [key, h] of server) local.set(key, { ...h });
      }
      listener.OnStopRunningUrl(null, status);
    },
    // No-Trash removal: on an IMAP "mark as deleted" account it only flags.
    deleteMessages(hdrs, _window, _deleteStorage, _isMove, listener) {
      calls.deletes = (calls.deletes || 0) + 1;
      for (const h of hdrs) server.get(h.messageKey).flags |= Ci.nsMsgMessageFlags.IMAPDeleted;
      listener.onStopCopy(0);
    },
    store(hdr) { local.set(hdr.messageKey, hdr); server.set(hdr.messageKey, { ...hdr }); },
    applySave() {
      if (newStored) server.set(2, { messageKey: 2, messageId: NEW_DRAFT_ID, flags: 0 });
      if (oldOnServer === "expunged") server.delete(1);
      if (oldOnServer === "marked-deleted") server.get(1).flags |= Ci.nsMsgMessageFlags.IMAPDeleted;
      if (oldOnServer === "rekeyed") {
        // The old draft survives under a new key; its old key now holds another message.
        server.set(7, { ...server.get(1), messageKey: 7 });
        server.set(1, { messageKey: 1, messageId: "unrelated@example.test", flags: 0 });
      }
    },
    applyMove(key) { if (moveRemovesOld) server.delete(key); },
  };
  return folder;
}

// Same source-marker sandbox pattern as privacy-messages.test.cjs. The complete
// saveDraft -> sendMessageDirectly path runs, including modern/legacy send slots.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function loadDraftTools(options = {}) {
  const calls = { finds: [], accessible: [], flags: [], sends: [], attachments: [], bodies: [], timerCancels: 0, syncs: 0, moves: [], closedViewsFor: [] };
  const folder = options.folder || makeImapDrafts(calls, options.store);
  const hdr = {
    messageId: "old-draft-id", messageKey: 1, flags: 0, subject: "Old subject", recipients: "old@example.test",
    folder: Object.hasOwn(options, "headerFolder") ? options.headerFolder : folder,
  };
  if (folder.store) folder.store(hdr);
  const identity = Object.hasOwn(options, "identity") ? options.identity : {
    key: "identity-1", email: "sender@example.test", fullName: "Sender", draftFolder: DRAFTS_URI,
  };
  const destination = options.destination || folder;
  const sandbox = {
    Ci,
    Components: { isSuccessCode: status => status === 0 },
    ChromeUtils: { generateQI: () => () => {} },
    MailServices: {
      accounts: { accounts: [{ key: "account-1", identities: [identity] }] },
      copy: {
        copyMessages(source, hdrs, destination, isMove, listener) {
          calls.moves.push({ source, keys: hdrs.map(h => h.messageKey), destination, isMove });
          source.applyMove(hdrs[0].messageKey);
          listener.onStopCopy(0);
        },
      },
    },
    console,
    Cc: {
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({
        headers: new Map(), attachments: [],
        setHeader(name, value) { this.headers.set(name, value); },
        addAttachment(value) { this.attachments.push(value); },
      }) },
      "@mozilla.org/timer;1": { createInstance: () => {
        const timer = {
          cancelled: false,
          // With timersFire, a one-shot timer fires on the next turn unless
          // cancelled first; otherwise timers never fire.
          initWithCallback(callback) {
            if (options.timersFire) setImmediate(() => { if (!timer.cancelled) callback.notify(); });
          },
          cancel() { timer.cancelled = true; calls.timerCancels++; },
        };
        return timer;
      } },
      "@mozilla.org/messengercompose/send;1": { createInstance: () => ({
        createAndSendMessage(...args) {
          calls.sends.push(args);
          if (options.legacy && args.length === 16) {
            throw Object.assign(new Error("Not enough arguments"), { result: 0x80570001 });
          }
          const listener = args[args.length === 16 ? 12 : 14];
          // Like MimeMessage._gatherMimeHeaders, the send writes the
          // generated Message-ID back into the fields it was given.
          args[3].messageId = `<${NEW_DRAFT_ID}>`;
          if (folder.applySave) folder.applySave();
          listener.onStopCopy(0);
        },
      }) },
    },
    closeViewsShowingMessages(msgHdrs) {
      calls.closedViewsFor.push(msgHdrs);
      return options.openViewIds || [];
    },
    findTrashFolder: () => (options.noTrash ? null : { URI: "imap://user@example.test/Trash" }),
    findMessage(messageId, folderPath) {
      calls.finds.push({ messageId, folderPath });
      if (options.findThrows) throw options.findThrows;
      if (options.findError) return { error: options.findError };
      return { msgHdr: hdr, folder };
    },
    getAccessibleFolder(uri) {
      calls.accessible.push(uri);
      if (options.accessThrows) throw options.accessThrows;
      if (options.accessError) return { error: options.accessError };
      return { folder: destination };
    },
    setComposeIdentity(params, from) {
      calls.from = from;
      params.identity = identity;
      return options.identityError ? { error: options.identityError } : null;
    },
    resolveComposeFormat: (_identity, isHtml) => ({ useHtml: isHtml === true, format: 0 }),
    buildBodyWithSignature(body, selectedIdentity, useHtml, isHtml, includeSignature) {
      calls.bodies.push({ body, selectedIdentity, useHtml, isHtml, includeSignature });
      return body || "";
    },
    filePathsToAttachDescs(attachments) {
      calls.attachments.push(attachments);
      if (options.attachmentError) throw new Error(options.attachmentError);
      return { descs: attachments || [], failed: [] };
    },
    descsToMsgAttachments: descs => descs,
  };
  vm.createContext(sandbox);
  vm.runInContext([
    snippet("FOLDER SYNC HELPER"), snippet("DRAFT HELPERS"), snippet("DRAFT REPLACE VERIFICATION"), snippet("SAVE DRAFT TOOL"),
    snippet("DIRECT SEND HELPER"), snippet("TOOL CALL DISPATCH"),
  ].join("\n"), sandbox);
  return {
    api: sandbox, calls, hdr, identity,
    save(args = {}) {
      return sandbox.dispatchTool("saveDraft", {
        to: "new@example.test", subject: "New subject", body: "New body", isHtml: false,
        replaceMessageId: "old-draft-id", replaceFolderPath: DRAFTS_URI, ...args,
      });
    },
  };
}

function assertNoComposeSideEffects(calls) {
  assert.equal(calls.sends.length, 0, "must not submit a draft or replacement header");
  assert.equal(calls.attachments.length, 0, "must reject before attachment processing");
  assert.equal(calls.bodies.length, 0, "must reject before reading or appending a signature");
}

describe("saveDraft replacement is proven in the Drafts folder", () => {
  it("reports replaced only once the new draft is stored and the old one is gone", async () => {
    const { save, calls, hdr } = loadDraftTools({ openViewIds: ["tab:1:3"] });
    const result = await save();
    assert.equal(result.success, true);
    assert.equal(result.message, "Draft replaced");
    assert.equal(result.messageId, NEW_DRAFT_ID);
    assert.equal(result.replaceVerified, true);
    assert.equal(result.previousDraft, "gone");
    assert.equal(result.previousDraftMovedToTrash, undefined);
    assert.equal(calls.moves.length, 0);
    assert.equal(calls.syncs, 1);
    assert.equal(calls.closedViewsFor[0][0], hdr);
    assert.deepEqual([...result.closedViews], ["tab:1:3"]);
  });

  it("moves a predecessor Thunderbird left in place to Trash and says so", async () => {
    const { save, calls } = loadDraftTools({ store: { oldOnServer: "kept" } });
    const result = await save();
    assert.equal(result.success, true);
    assert.equal(result.replaceVerified, true);
    assert.equal(result.previousDraft, "gone");
    assert.equal(result.previousDraftMovedToTrash, true);
    assert.match(result.message, /left the previous draft in place, so it was moved to Trash/);
    assert.deepEqual(JSON.parse(JSON.stringify(calls.moves.map(m => ({ keys: m.keys, isMove: m.isMove, to: m.destination.URI })))),
      [{ keys: [1], isMove: true, to: "imap://user@example.test/Trash" }]);
  });

  it("never touches the predecessor while the new draft is not confirmed", async () => {
    const { save, calls } = loadDraftTools({ store: { newStored: false, oldOnServer: "kept" } });
    const result = await save();
    assert.equal(result.success, true, "the save itself was reported done by Thunderbird");
    assert.equal(result.replaceVerified, false);
    assert.equal(result.previousDraft, "present");
    assert.match(result.message, /new draft was not found .* previous draft was left untouched/);
    assert.doesNotMatch(result.message, /replaced/);
    assert.equal(calls.moves.length, 0);
  });

  it("reports a predecessor that is only marked deleted without moving it", async () => {
    const { save, calls } = loadDraftTools({ store: { oldOnServer: "marked-deleted" } });
    const result = await save();
    assert.equal(result.replaceVerified, true);
    assert.equal(result.previousDraft, "marked-deleted");
    assert.match(result.message, /marked deleted/);
    assert.equal(calls.moves.length, 0);
  });

  it("does not claim a replacement when the predecessor survives the move", async () => {
    const { save } = loadDraftTools({ store: { oldOnServer: "kept", moveRemovesOld: false } });
    const result = await save();
    assert.equal(result.success, true);
    assert.equal(result.replaceVerified, false);
    assert.equal(result.previousDraft, "present");
    assert.match(result.message, /still in the Drafts folder/);
  });

  it("treats an unconfirmed folder sync as no proof and leaves the predecessor untouched", async () => {
    const { save, calls } = loadDraftTools({ store: { syncStatus: 0x80004005, oldOnServer: "kept" } });
    const result = await save();
    assert.equal(result.success, true);
    assert.equal(result.folderSyncConfirmed, false);
    assert.equal(result.replaceVerified, false);
    assert.match(result.message, /could not be synced .* unconfirmed/);
    assert.equal(calls.moves.length, 0);
  });

  it("re-issues the sync request until one is answered, then proves and completes the replacement", async () => {
    const { save, calls } = loadDraftTools({ timersFire: true, store: { oldOnServer: "kept", droppedSyncs: 3 } });
    const result = await save();
    assert.equal(calls.syncs, 5, "three dropped requests, one answered before the check, one after the move");
    assert.equal(result.folderSyncConfirmed, undefined);
    assert.equal(result.previousDraftMovedToTrash, true);
    assert.equal(result.replaceVerified, true);
    assert.equal(calls.moves.length, 1);
  });

  it("gives up after its budget when no request is ever answered", async () => {
    const { save, calls } = loadDraftTools({ timersFire: true, store: { oldOnServer: "kept", droppedSyncs: 1000 } });
    const result = await save();
    assert.equal(calls.syncs, 10, "30 s budget in 3 s attempts");
    assert.equal(result.folderSyncConfirmed, false);
    assert.equal(result.replaceVerified, false);
    assert.equal(calls.moves.length, 0);
  });

  it("does not report a move to Trash when, without a Trash folder, the draft was only marked deleted", async () => {
    const { save, calls } = loadDraftTools({ noTrash: true, store: { oldOnServer: "kept" } });
    const result = await save();
    assert.equal(calls.deletes, 1);
    assert.equal(result.previousDraft, "marked-deleted");
    assert.equal(result.previousDraftMovedToTrash, undefined);
    assert.doesNotMatch(result.message, /Trash/);
    assert.match(result.message, /marked deleted in the Drafts folder/);
  });

  it("never reads an unreadable database as the predecessor being gone", async () => {
    const { save } = loadDraftTools({ store: { oldOnServer: "kept", dbThrowsAfterSyncs: 2 } });
    const result = await save();
    assert.equal(result.previousDraft, "unknown");
    assert.equal(result.replaceVerified, false);
    assert.equal(result.previousDraftMovedToTrash, undefined);
    assert.match(result.message, /could not be confirmed/);
  });

  it("finds a predecessor whose key was reassigned and moves that one, not the key's new occupant", async () => {
    const { save, calls } = loadDraftTools({ store: { oldOnServer: "rekeyed" } });
    const result = await save();
    assert.deepEqual(calls.moves.map(m => m.keys[0]), [7]);
    assert.equal(result.previousDraftMovedToTrash, true);
    assert.equal(result.replaceVerified, true);
  });

  it("does not confirm a move whose follow-up sync fails on every attempt", async () => {
    const { save } = loadDraftTools({ store: { oldOnServer: "kept", syncStatuses: [0, ...Array(10).fill(0x80004005)] } });
    const result = await save();
    assert.equal(result.folderSyncConfirmed, false);
    assert.equal(result.replaceVerified, false);
    assert.equal(result.previousDraftMovedToTrash, undefined);
  });

  it("reports the new draft's Message-ID for plain saves without syncing", async () => {
    const { save, calls } = loadDraftTools();
    const result = await save({ replaceMessageId: undefined, replaceFolderPath: undefined });
    assert.equal(result.message, "Draft saved");
    assert.equal(result.messageId, NEW_DRAFT_ID);
    assert.equal(result.replaceVerified, undefined);
    assert.equal(calls.syncs, 0);
    assert.equal(calls.closedViewsFor.length, 0);
  });
});

describe("saveDraft replacement", () => {
  for (const legacy of [false, true]) {
    it(`passes only an accessible identity Drafts header to send slot 8, legacy=${legacy}`, async () => {
      const { save, calls, hdr, identity } = loadDraftTools({ legacy });
      const result = await save({ from: "sender@example.test" });
      assert.equal(result.success, true);
      assert.equal(result.message, "Draft replaced");
      assert.equal(result.folderPath, DRAFTS_URI);
      assert.deepEqual(calls.finds, [{ messageId: "old-draft-id", folderPath: DRAFTS_URI }]);
      assert.deepEqual(calls.flags, [Ci.nsMsgFolderFlags.Drafts]);
      assert.ok(calls.accessible.includes(DRAFTS_URI));
      assert.equal(calls.from, "sender@example.test");
      assert.equal(calls.sends.length, legacy ? 2 : 1);
      const args = calls.sends.at(-1);
      assert.equal(args.length, legacy ? 18 : 16);
      assert.equal(args[1], identity);
      assert.equal(args[2], "account-1");
      assert.equal(args[6], Ci.nsIMsgCompDeliverMode.SaveAsDraft);
      assert.equal(args[7], hdr);
      assert.equal(args[8], "text/plain");
      assert.equal(args[9], "New body");
      assert.equal(args.at(-1), Ci.nsIMsgCompType.Draft);
      assert.equal(calls.bodies[0].includeSignature, false);
      // The send timeout, then the Drafts sync of the replacement check.
      assert.equal(calls.timerCancels, 2);
    });
  }

  it("replaces all supplied content and attachments while preserving supplied threading", async () => {
    const { save, calls, hdr } = loadDraftTools();
    const attachment = { name: "new.txt", content: "bmV3" };
    const result = await save({
      to: "new@example.test", subject: "Corrected", body: "Corrected body",
      cc: "cc@example.test", bcc: "bcc@example.test", attachments: [attachment],
      inReplyTo: "<parent@example.test>", references: "<root@example.test> <parent@example.test>",
      includeSignature: true,
    });
    assert.equal(result.success, true);
    const fields = calls.sends[0][3];
    assert.equal(fields.to, "new@example.test");
    assert.equal(fields.cc, "cc@example.test");
    assert.equal(fields.bcc, "bcc@example.test");
    assert.equal(fields.subject, "Corrected");
    assert.equal(fields.body, "Corrected body");
    assert.equal(fields.headers.get("In-Reply-To"), "<parent@example.test>");
    assert.equal(fields.references, "<root@example.test> <parent@example.test>");
    assert.deepEqual(fields.attachments, [attachment]);
    assert.equal(calls.bodies[0].includeSignature, true);
    assert.equal(hdr.subject, "Old subject", "the old header must not be edited directly");
  });

  it("does not silently retain omitted recipients or attachments from the old draft", async () => {
    const { save, calls } = loadDraftTools();
    const result = await save({ to: undefined, subject: undefined, body: undefined });
    assert.equal(result.success, true);
    const fields = calls.sends[0][3];
    for (const key of ["to", "cc", "bcc", "subject", "body"]) assert.equal(fields[key], "");
    assert.deepEqual(fields.attachments, []);
  });

  it("still creates a new draft when no replacement is requested", async () => {
    const { save, calls } = loadDraftTools();
    const result = await save({ replaceMessageId: undefined, replaceFolderPath: undefined });
    assert.equal(result.success, true);
    assert.equal(result.message, "Draft saved");
    assert.equal(calls.finds.length, 0);
    assert.equal(calls.sends[0][7], null);
    assert.equal(calls.sends[0].at(-1), Ci.nsIMsgCompType.New);
    assert.equal(calls.bodies[0].includeSignature, true);
  });

  it("requires the replacement folder instead of silently appending a new draft", async () => {
    const { save, calls } = loadDraftTools();
    assert.match((await save({ replaceFolderPath: undefined })).error, /requires replaceFolderPath/);
    assert.equal(calls.finds.length, 0);
    assertNoComposeSideEffects(calls);
  });

  for (const replaceMessageId of [null, "", "  ", 42, {}, []]) {
    it(`rejects malformed replacement ID ${JSON.stringify(replaceMessageId)}`, async () => {
      const { save, calls } = loadDraftTools();
      assert.match((await save({ replaceMessageId })).error, /replaceMessageId.*non-empty/i);
      assertNoComposeSideEffects(calls);
    });
  }

  for (const name of ["Inbox", "Trash", "Drafts"]) {
    it(`refuses an unflagged ${name} folder even when its URI matches the identity`, async () => {
      const folder = { URI: `imap://user@example.test/${name}`, getFlag: () => false };
      const { save, calls } = loadDraftTools({ folder, identity: { draftFolder: folder.URI } });
      assert.match((await save({ replaceFolderPath: folder.URI })).error, /Drafts flag/);
      assertNoComposeSideEffects(calls);
    });
  }

  for (const getFlag of [undefined, () => { throw new Error("flag unavailable"); }]) {
    it(`fails closed when the folder Drafts flag is ${getFlag ? "unreadable" : "missing"}`, async () => {
      const { save, calls } = loadDraftTools({ folder: { URI: DRAFTS_URI, getFlag } });
      assert.match((await save()).error, /Drafts flag/);
      assertNoComposeSideEffects(calls);
    });
  }

  for (const headerFolder of [undefined, { URI: "imap://user@example.test/Inbox", getFlag: () => true }]) {
    it(`rejects an actual message folder that is ${headerFolder ? "different from the lookup folder" : "missing"}`, async () => {
      const { save, calls } = loadDraftTools({ headerFolder });
      assert.match((await save()).error, /message is not stored in the requested Drafts folder/);
      assertNoComposeSideEffects(calls);
    });
  }

  it("checks the actual message folder's Drafts flag even when the lookup folder is flagged", async () => {
    const { save, calls } = loadDraftTools({ headerFolder: { URI: DRAFTS_URI, getFlag: () => false } });
    assert.match((await save()).error, /Drafts flag/);
    assertNoComposeSideEffects(calls);
  });

  for (const findError of ["Message not found", "Account not accessible for folder"]) {
    it(`honors findMessage rejection: ${findError}`, async () => {
      const { save, calls } = loadDraftTools({ findError });
      assert.equal((await save()).error, findError);
      assertNoComposeSideEffects(calls);
    });
  }

  it("does not submit a replacement after a lookup exception", async () => {
    const { save, calls } = loadDraftTools({ findThrows: new Error("lookup failed") });
    assert.match((await save()).error, /lookup failed/);
    assertNoComposeSideEffects(calls);
  });

  it("requires a configured Drafts destination for the selected identity", async () => {
    const { save, calls } = loadDraftTools({ identity: { fccFolder: DRAFTS_URI, draftsFolder: DRAFTS_URI } });
    assert.match((await save()).error, /no configured Drafts folder/);
    assertNoComposeSideEffects(calls);
  });

  for (const options of [
    { accessError: `Account not accessible for folder: ${OTHER_DRAFTS_URI}` },
    { accessThrows: new Error(`Cannot open ${OTHER_DRAFTS_URI}`) },
  ]) {
    it(`rejects an inaccessible identity destination without exposing it (${options.accessError ? "result" : "exception"})`, async () => {
      const { save, calls } = loadDraftTools({ ...options, identity: { draftFolder: OTHER_DRAFTS_URI } });
      const result = await save();
      assert.match(result.error, /Drafts folder is not accessible/);
      assert.doesNotMatch(JSON.stringify(result), /other@example/);
      assertNoComposeSideEffects(calls);
    });
  }

  it("refuses a flagged Drafts folder belonging to a different identity", async () => {
    const { save, calls } = loadDraftTools({
      identity: { draftFolder: OTHER_DRAFTS_URI }, destination: { URI: OTHER_DRAFTS_URI },
    });
    assert.match((await save()).error, /outside the selected identity.*Drafts folder/);
    assertNoComposeSideEffects(calls);
  });

  it("stops replacement if identity selection fails", async () => {
    const { save, calls } = loadDraftTools({ identityError: "No permitted identity" });
    assert.equal((await save()).error, "No permitted identity");
    assertNoComposeSideEffects(calls);
  });

  it("does not pass a replacement header to Thunderbird after attachment preparation fails", async () => {
    const { save, calls } = loadDraftTools({ attachmentError: "Attachment preparation failed" });
    assert.match((await save()).error, /Attachment preparation failed/);
    assert.equal(calls.sends.length, 0);
  });
});
