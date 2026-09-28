"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
// Use the production marker+VM pattern from privacy-messages.test.cjs.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

const DRAFTS_URI = "imap://allowed@example.test/Drafts";
const RESTRICTED_URI = "imap://restricted@example.test/Drafts";

function loadReply(options = {}) {
  const { enabled = true, blocked = true, attachmentError, mode = "success" } = options;
  const calls = { finds: 0, sends: 0, opens: [], saves: 0, timers: [], windows: [], mime: [], lookups: [], flags: [], dialogs: 0 };
  const identity = {
    draftsFolderURI: DRAFTS_URI,
    get showSaveMsgDlg() { return true; },
    set showSaveMsgDlg(_value) { assert.fail("shared identity preference must not change"); },
  };
  const composeIdentity = Object.hasOwn(options, "composeIdentity") ? options.composeIdentity : identity;
  const folder = { server: {}, getUriForMsg: () => "message-uri" };
  const destination = Object.hasOwn(options, "destination") ? options.destination : {
    URI: DRAFTS_URI,
    server: { accountKey: "allowed" },
    getFlag(flag) { calls.flags.push(flag); return true; },
  };
  const sandbox = {
    TextDecoder,
    Uint8Array,
    atob,
    Services: { prefs: { getBoolPref: () => options.allowEncrypted === true } },
    Components: { isSuccessCode: status => status === 0 },
    ChromeUtils: {
      generateQI: () => () => {},
      importESModule: () => ({
        MsgHdrToMimeMessage(hdr, _listener, callback, _download, mimeOptions) {
          calls.mime.push(mimeOptions);
          callback(hdr, options.mime || { contentType: "text/plain", body: "original text" });
        },
      }),
    },
    Ci: { nsIMsgCompType: { Reply: 1, ReplyAll: 2 }, nsITimer: { TYPE_ONE_SHOT: 0 }, nsMsgFolderFlags: { Drafts: 0x400 } },
    MailServices: {
      folderLookup: { getFolderForURL(uri) {
        calls.lookups.push(uri);
        if (options.lookupError) throw new Error(`Lookup failed: ${uri}`);
        if (uri === RESTRICTED_URI) return { URI: uri, server: { accountKey: "restricted" }, getFlag: () => true };
        return destination && uri === destination.URI ? destination : null;
      } },
      accounts: { findAccountForServer: server => ({ key: server.accountKey }) },
    },
    isAccountAllowed: key => key === "allowed",
    Cc: {
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({}) },
      "@mozilla.org/timer;1": { createInstance() {
        const timer = {
          cancelled: false,
          initWithCallback(callback) { this.callback = callback; },
          cancel() { this.cancelled = true; },
        };
        calls.timers.push(timer);
        return timer;
      } },
    },
    isToolEnabled(name) { assert.equal(name, "saveDraft"); return enabled; },
    isSkipReviewBlocked: () => blocked,
    findMessage() { calls.finds++; return { msgHdr: {}, folder }; },
    filePathsToAttachDescs(attachments) {
      if (attachmentError) throw new Error(attachmentError);
      return { descs: attachments || [] };
    },
    setComposeIdentity(params) { params.identity = identity; },
    resolveComposeFormat: () => ({ useHtml: false, format: 0 }),
    sendMessageDirectly() { calls.sends++; return Promise.resolve({ success: true }); },
    async openComposeWindowWithCustomizations(...args) {
      calls.opens.push(args);
      const afterInsert = args[10];
      if (!afterInsert) return { success: true };
      let listener;
      const win = {
        gCurrentIdentity: Object.hasOwn(options, "windowIdentity") ? options.windowIdentity : composeIdentity,
        gCloseWindowAfterSave: options.originalCloseFlag ?? false,
        gMsgCompose: {
          identity: composeIdentity,
          RegisterStateListener(value) { listener = value; },
          UnregisterStateListener(value) { assert.equal(value, listener); win.unregistered = true; },
        },
        DisplaySaveFolderDlg() { calls.dialogs++; },
        completeSave(status = 0) {
          listener.ComposeProcessDone(status);
          // Thunderbird's native listener closes the window after a late success
          // if our close flag is still set, even after our listener is removed.
          if (status === 0 && win.gCloseWindowAfterSave) win.closed = true;
        },
        SaveAsDraft() {
          calls.saves++;
          assert.equal(win.gCloseWindowAfterSave, true);
          win.DisplaySaveFolderDlg();
          assert.equal(calls.dialogs, 0, "native draft dialog should be suppressed per window");
          if (mode === "reject") return Promise.reject(new Error("native save rejected"));
          if (mode === "throw") throw new Error("native save threw");
          if (mode === "pending") return undefined;
          if (mode === "timeout") {
            calls.timers.at(-1).callback.notify();
          } else {
            win.completeSave(mode === "failure" ? 0x80004005 : 0);
          }
        },
      };
      win.originalDialog = win.DisplaySaveFolderDlg;
      if (options.absentWindowState) {
        delete win.gCloseWindowAfterSave;
        delete win.DisplaySaveFolderDlg;
      }
      if (mode === "unsupported") win.SaveAsDraft = undefined;
      calls.windows.push(win);
      return afterInsert(win);
    },
  };
  // Exercise the real access gate too, so a restricted account with a Drafts
  // flag cannot pass merely because a stub forgot to enforce account access.
  const accessStart = source.indexOf("function isFolderAccessible(folder) {");
  const accessEnd = source.indexOf("function getAccessibleAccounts() {", accessStart);
  assert.ok(accessStart >= 0 && accessEnd > accessStart, "Missing folder access helper boundaries");
  vm.createContext(sandbox);
  vm.runInContext([
    source.match(/^const PREF_ALLOW_ENCRYPTED_MESSAGES = .+;$/m)[0],
    snippet("PRIVACY PREFERENCE HELPERS"), snippet("MCP TEXT SANITIZATION"),
    snippet("MESSAGE TEXT CONVERSION"), snippet("RAW MIME PARSING HELPERS"),
    snippet("ENCRYPTED MESSAGE GUARD"), snippet("DRAFT HELPERS"),
    source.slice(accessStart, accessEnd),
    snippet("COMPOSE WINDOW DRAFT HELPER"),
    snippet("REPLY TOOL"),
    snippet("TOOL CALL DISPATCH"),
  ].join("\n"), sandbox);
  const reply = args => sandbox.dispatchTool("replyToMessage", {
    messageId: "original@example.com", folderPath: "imap://example/INBOX", body: "reply text", saveAsDraft: true, ...args,
  });
  return { reply, calls, identity };
}

describe("replyToMessage saveAsDraft", () => {
  it("uses native reply composition and saves without sending even while skipReview is blocked", async () => {
    const { reply, calls, identity } = loadReply();
    const attachment = { name: "note.txt", base64: "bm90ZQ==" };
    const result = await reply({ replyAll: true, to: "recipient@example.com", cc: "cc@example.com", attachments: [attachment] });
    assert.equal(result.success, true);
    assert.equal(result.message, "Reply saved as draft");
    assert.equal(calls.sends, 0);
    assert.equal(calls.saves, 1);
    const args = calls.opens[0];
    assert.equal(args[0].originalMsgURI, "message-uri");
    assert.equal(args[2], 2);
    assert.equal(args[3], identity);
    assert.equal(args[4], "reply text");
    assert.equal(args[6], "recipient@example.com");
    assert.equal(args[7], "cc@example.com");
    assert.deepEqual(args[9], [attachment]);
    assert.equal(calls.mime[0].examineEncryptedParts, false);
    assert.deepEqual(calls.lookups, [DRAFTS_URI]);
    assert.deepEqual(calls.flags, [0x400]);
    assert.equal(calls.timers[0].cancelled, true);
    assert.equal(calls.windows[0].unregistered, true);
  });

  it("rejects skipReview plus saveAsDraft before lookup or composition", async () => {
    const { reply, calls } = loadReply({ blocked: false });
    assert.match((await reply({ skipReview: true })).error, /cannot be combined/);
    assert.equal(calls.finds, 0);
    assert.equal(calls.opens.length, 0);
    assert.equal(calls.sends, 0);
  });

  it("requires the saveDraft tool to be enabled", async () => {
    const { reply, calls } = loadReply({ enabled: false });
    assert.match((await reply()).error, /requires the saveDraft tool/);
    assert.equal(calls.finds, 0);
    assert.equal(calls.saves, 0);
  });

  it("stops before opening or saving if any attachment is refused", async () => {
    const { reply, calls } = loadReply({ attachmentError: "Attachments refused: bad.txt" });
    assert.match((await reply()).error, /Attachments refused/);
    assert.equal(calls.opens.length, 0);
    assert.equal(calls.saves, 0);
    assert.equal(calls.sends, 0);
  });

  for (const [label, options] of [
    ["restricted current window identity", { windowIdentity: { draftsFolderURI: RESTRICTED_URI } }],
    ["restricted compose identity", { windowIdentity: null, composeIdentity: { draftsFolderURI: RESTRICTED_URI } }],
    ["restricted destination despite encrypted opt-in", { allowEncrypted: true, mime: { contentType: "multipart/encrypted" }, windowIdentity: { draftFolder: RESTRICTED_URI } }],
    ["missing current identity", { windowIdentity: null, composeIdentity: null }],
    ["missing configured destination", { windowIdentity: {} }],
    ["unreadable destination", { windowIdentity: { get draftsFolderURI() { throw new Error(RESTRICTED_URI); } } }],
    ["unknown destination", { windowIdentity: { draftFolder: "imap://unknown@example.test/Drafts" } }],
    ["lookup exception", { lookupError: true }],
    ["missing folder", { destination: null }],
    ["unflagged folder named Drafts", { destination: { URI: DRAFTS_URI, server: { accountKey: "allowed" }, getFlag: () => false } }],
    ["missing folder flags", { destination: { URI: DRAFTS_URI, server: { accountKey: "allowed" } } }],
    ["unreadable folder flags", { destination: { URI: DRAFTS_URI, server: { accountKey: "allowed" }, getFlag() { throw new Error(RESTRICTED_URI); } } }],
  ]) {
    it(`refuses saving to a ${label} without exposing destination details or changing window state`, async () => {
      const { reply, calls, identity } = loadReply(options);
      const result = await reply();
      assert.equal(identity.draftsFolderURI, DRAFTS_URI, "requested identity stays allowed");
      assert.match(result.error, /accessible Drafts-flagged destination/);
      assert.doesNotMatch(JSON.stringify(result), /imap:\/\//);
      assert.equal(calls.saves, 0);
      assert.equal(calls.sends, 0);
      const win = calls.windows[0];
      assert.equal(win.gCloseWindowAfterSave, false);
      assert.equal(win.DisplaySaveFolderDlg, win.originalDialog);
      assert.equal(calls.timers[0].cancelled, true);
    });
  }

  it("accepts the current window identity's accessible legacy Drafts folder", async () => {
    const legacyURI = "mailbox://local@example.test/CustomDrafts";
    const { reply, calls } = loadReply({
      windowIdentity: { draftFolder: { URI: legacyURI } },
      destination: { URI: legacyURI, server: { accountKey: "allowed" }, getFlag: flag => flag === 0x400 },
    });
    assert.equal((await reply()).success, true);
    assert.equal(calls.saves, 1);
    assert.deepEqual(calls.lookups, [legacyURI]);
  });

  it("retains the ordinary review path when saveAsDraft is omitted", async () => {
    const { reply, calls } = loadReply({ windowIdentity: { draftsFolderURI: RESTRICTED_URI } });
    const result = await reply({ saveAsDraft: undefined });
    assert.equal(result.message, "Reply window opened");
    assert.equal(calls.opens.length, 1);
    assert.equal(calls.opens[0][10], undefined);
    assert.equal(calls.saves, 0);
    assert.equal(calls.mime.length, 0);
    assert.equal(calls.lookups.length, 0);
  });

  it("keeps shared signature/dialog preferences intact across concurrent saves", async () => {
    const { reply, calls, identity } = loadReply();
    const results = await Promise.all([reply(), reply()]);
    assert.ok(results.every(result => result.success));
    assert.equal(identity.showSaveMsgDlg, true);
    assert.equal(calls.saves, 2);
    assert.notEqual(calls.windows[0], calls.windows[1]);
  });

  for (const [mode, error] of [
    ["failure", /Saving reply draft failed/],
    ["reject", /native save rejected/],
    ["throw", /native save threw/],
    ["unsupported", /does not support SaveAsDraft/],
    ["timeout", /Timed out saving reply draft/],
  ]) {
    it(`reports native save ${mode} and cleans up its listener and timer`, async () => {
      const { reply, calls } = loadReply({ mode });
      const result = await reply();
      assert.match(result.error, error);
      assert.equal(result.success, undefined);
      assert.equal(calls.sends, 0);
      assert.equal(calls.timers[0].cancelled, true);
      assert.equal(calls.windows[0].unregistered, mode === "unsupported" ? undefined : true);
      assert.equal(calls.windows[0].gCloseWindowAfterSave, false);
      assert.equal(calls.windows[0].DisplaySaveFolderDlg, calls.windows[0].originalDialog);
      if (mode === "timeout") {
        assert.equal(result.saveOutcome, "uncertain");
        assert.match(result.error, /uncertain.*may still complete/);
        assert.doesNotMatch(result.error, /was left open/);
      }
    });
  }

  for (const completion of ["timeout", "failure"]) {
    it(`restores window state before returning ${completion} and ignores late successful completion`, async () => {
      const { reply, calls } = loadReply({ mode: "pending" });
      const pending = reply();
      const win = calls.windows[0];
      assert.equal(win.gCloseWindowAfterSave, true);
      assert.notEqual(win.DisplaySaveFolderDlg, win.originalDialog);

      if (completion === "timeout") calls.timers[0].callback.notify();
      else win.completeSave(0x80004005);
      assert.equal(win.gCloseWindowAfterSave, false);
      assert.equal(win.DisplaySaveFolderDlg, win.originalDialog);
      const result = await pending;
      assert.equal(typeof result.error, "string");

      win.completeSave(0);
      calls.timers[0].callback.notify();
      assert.notEqual(win.closed, true, "late success must not close the window using our override");
      assert.equal(win.gCloseWindowAfterSave, false);
      assert.equal(win.DisplaySaveFolderDlg, win.originalDialog);
      assert.equal(result.success, undefined);
      assert.equal(calls.saves, 1);
    });
  }

  it("restores an originally true close flag after failure", async () => {
    const { reply, calls } = loadReply({ mode: "failure", originalCloseFlag: true });
    assert.match((await reply()).error, /Saving reply draft failed/);
    assert.equal(calls.windows[0].gCloseWindowAfterSave, true);
    assert.equal(calls.windows[0].DisplaySaveFolderDlg, calls.windows[0].originalDialog);
  });

  it("removes temporary window properties that did not originally exist", async () => {
    const { reply, calls } = loadReply({ mode: "timeout", absentWindowState: true });
    assert.equal((await reply()).saveOutcome, "uncertain");
    assert.equal(Object.hasOwn(calls.windows[0], "gCloseWindowAfterSave"), false);
    assert.equal(Object.hasOwn(calls.windows[0], "DisplaySaveFolderDlg"), false);
    calls.windows[0].completeSave(0);
    assert.notEqual(calls.windows[0].closed, true);
  });
});
