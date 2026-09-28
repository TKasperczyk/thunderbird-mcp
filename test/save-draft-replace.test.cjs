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
  nsITimer: { TYPE_ONE_SHOT: 0 },
};

// Same source-marker sandbox pattern as privacy-messages.test.cjs. The complete
// saveDraft -> sendMessageDirectly path runs, including modern/legacy send slots.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function loadDraftTools(options = {}) {
  const calls = { finds: [], accessible: [], flags: [], sends: [], attachments: [], bodies: [], timerCancels: 0 };
  const folder = options.folder || {
    URI: DRAFTS_URI,
    getFlag(flag) { calls.flags.push(flag); return options.draftsFlag !== false; },
  };
  const hdr = {
    messageId: "old-draft-id", subject: "Old subject", recipients: "old@example.test",
    folder: Object.hasOwn(options, "headerFolder") ? options.headerFolder : folder,
  };
  const identity = Object.hasOwn(options, "identity") ? options.identity : {
    key: "identity-1", email: "sender@example.test", fullName: "Sender", draftFolder: DRAFTS_URI,
  };
  const destination = options.destination || folder;
  const sandbox = {
    Ci,
    Components: { isSuccessCode: status => status === 0 },
    ChromeUtils: { generateQI: () => () => {} },
    MailServices: { accounts: { accounts: [{ key: "account-1", identities: [identity] }] } },
    console,
    Cc: {
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({
        headers: new Map(), attachments: [],
        setHeader(name, value) { this.headers.set(name, value); },
        addAttachment(value) { this.attachments.push(value); },
      }) },
      "@mozilla.org/timer;1": { createInstance: () => ({
        initWithCallback() {}, cancel() { calls.timerCancels++; },
      }) },
      "@mozilla.org/messengercompose/send;1": { createInstance: () => ({
        createAndSendMessage(...args) {
          calls.sends.push(args);
          if (options.legacy && args.length === 16) {
            throw Object.assign(new Error("Not enough arguments"), { result: 0x80570001 });
          }
          const listener = args[args.length === 16 ? 12 : 14];
          listener.onStopCopy(0);
        },
      }) },
    },
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
    snippet("DRAFT HELPERS"), snippet("SAVE DRAFT TOOL"),
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
      assert.equal(calls.timerCancels, 1);
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
