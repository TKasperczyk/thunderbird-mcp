"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
function snippet(startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${startMarker}`);
  return source.slice(start, end);
}

// Fires the poll callback on the next turn until cancelled, so read-back
// deadlines expire after their poll count instead of in wall-clock time.
function makeImmediateTimer() {
  let cancelled = false;
  return {
    initWithCallback(callback) {
      const tick = () => {
        if (cancelled) return;
        callback.notify();
        setImmediate(tick);
      };
      setImmediate(tick);
    },
    cancel() { cancelled = true; },
  };
}

const OPEN = { URI: "imap://open/INBOX", allowed: true };
const DRAFTS = { URI: "imap://open/Drafts", allowed: true };
const RESTRICTED = { URI: "imap://restricted/INBOX", allowed: false };

function hdr(messageId, folder, messageKey, subject = `Subject ${messageId}`) {
  return { messageId, folder, messageKey, mime2DecodedSubject: subject, encrypted: false };
}

/**
 * One 3-pane window (folder tab, message tabs, a calendar tab), standalone
 * message windows and compose windows, shaped like tabmail/mailTabs.js.
 */
function makeGui({ messageTabs = [], messageWindows = [], composeWindows = [], folderSelection = null, stuckIds = [] } = {}) {
  let nextTabId = 0;
  const tabmail = { tabInfo: [], currentTabInfo: null, closed: [] };
  const addTab = (mode, props) => {
    const tab = { tabId: nextTabId++, mode: { name: mode }, canClose: mode !== "mail3PaneTab", ...props };
    tabmail.tabInfo.push(tab);
    return tab;
  };
  addTab("mail3PaneTab", { folder: OPEN, message: folderSelection, title: "INBOX" });
  for (const message of messageTabs) addTab("mailMessageTab", { message, title: message.mime2DecodedSubject });
  addTab("calendar", { title: "Calendar" });
  tabmail.currentTabInfo = tabmail.tabInfo[1] || tabmail.tabInfo[0];
  tabmail.closeTab = (tab) => {
    if (!tab.canClose || stuckIds.includes(tab.tabId)) return;
    tabmail.tabInfo.splice(tabmail.tabInfo.indexOf(tab), 1);
    tabmail.closed.push(tab.tabId);
  };
  const mainWindow = {
    docShell: { outerWindowID: 7 },
    document: { getElementById: (id) => (id === "tabmail" ? tabmail : null) },
  };

  let nextWindowId = 100;
  const windows = { "mail:3pane": [mainWindow], "mail:messageWindow": [], msgcompose: [] };
  for (const message of messageWindows) {
    const win = {
      closed: false,
      docShell: { outerWindowID: nextWindowId++ },
      document: { getElementById: (id) => (id === "messageBrowser" ? { contentWindow: { gMessage: message } } : null) },
      close() {
        this.closed = true;
        windows["mail:messageWindow"].splice(windows["mail:messageWindow"].indexOf(this), 1);
      },
    };
    windows["mail:messageWindow"].push(win);
  }
  const addComposeWindow = ({ draft, subject }) => {
    const win = {
      docShell: { outerWindowID: nextWindowId++ },
      gMsgCompose: { compFields: { draftId: draft ? `imap-message://${draft.folder.URI}#${draft.messageKey}?fetchCompleteMessage=true` : "" } },
      document: { getElementById: (id) => (id === "msgSubject" ? { value: subject } : null) },
      focused: 0,
      focus() { this.focused++; },
      close() { throw new Error("compose windows must not be closed"); },
    };
    windows.msgcompose.push(win);
    return win;
  };
  composeWindows.forEach(addComposeWindow);
  const drafts = composeWindows.map(c => c.draft).filter(Boolean);
  return { tabmail, windows, drafts, addComposeWindow };
}

function loadViews(gui, { allowEncrypted = false, messages = [], composeOpens = true } = {}) {
  const composeCalls = [];
  const sandbox = {
    composeCalls,
    findMessage: (messageId, folderPath) => {
      const msgHdr = messages.find(m => m.messageId === messageId && m.folder.URI === folderPath);
      return msgHdr ? { msgHdr, folder: msgHdr.folder } : { error: `Message not found: ${messageId}` };
    },
    ENCRYPTED_DISPLAY_ERROR: "Encrypted messages cannot be displayed while encrypted message access is off.",
    Services: { wm: { getEnumerator: (type) => [...(gui.windows[type] || [])] } },
    Cc: { "@mozilla.org/timer;1": { createInstance: () => makeImmediateTimer() } },
    Ci: {
      nsITimer: { TYPE_REPEATING_SLACK: 1 },
      nsMsgFolderFlags: { Drafts: 0x400 },
      nsIMsgCompType: { Draft: 9 },
      nsIMsgCompFormat: { Default: 0 },
    },
    MailServices: {
      accounts: { getFirstIdentityForServer: () => ({ key: "id1" }) },
      compose: {
        OpenComposeWindow(...args) {
          composeCalls.push(args);
          if (!composeOpens) return;
          gui.drafts.push(args[1]);
          gui.addComposeWindow({ draft: args[1], subject: args[1].mime2DecodedSubject });
        },
      },
      messageServiceFromURI: () => ({
        messageURIToMsgHdr: (uri) => {
          assert.doesNotMatch(uri, /\?/, "the query is stripped before the lookup");
          return gui.drafts.find(d => uri === `imap-message://${d.folder.URI}#${d.messageKey}`) || null;
        },
      }),
    },
    isFolderAccessible: (folder) => !!folder?.allowed,
    getAccessibleFolder: (uri) => ([OPEN, DRAFTS].some(f => f.URI === uri) ? { folder: {} } : { error: `Folder not found: ${uri}` }),
    isPrivacyOptInEnabled: () => allowEncrypted,
    PREF_ALLOW_ENCRYPTED_MESSAGES: "allowEncryptedMessages",
    hasEncryptedOuterStructure: (msgHdr) => msgHdr.encrypted,
    withholdProtectedSubject: (row) => { row.subject = "[Encrypted message]"; row.encryptedContentWithheld = true; return row; },
  };
  vm.createContext(sandbox);
  vm.runInContext(snippet("// BEGIN MESSAGE VIEW TOOLS", "// END MESSAGE VIEW TOOLS"), sandbox);
  return sandbox;
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

describe("listTabs", () => {
  it("lists tabs, message windows and compose windows live, with what each shows", () => {
    const a = hdr("a@x", OPEN, 1);
    const b = hdr("b@x", OPEN, 2);
    const c = hdr("c@x", OPEN, 3);
    const draft = hdr("d@x", DRAFTS, 9);
    const gui = makeGui({ folderSelection: a, messageTabs: [b], messageWindows: [c], composeWindows: [{ draft, subject: "Draft subject" }, { subject: "New" }] });
    const { views } = plain(loadViews(gui).listTabs());

    assert.deepStrictEqual(views, [
      { id: "tab:7:0", kind: "tab", mode: "mail3PaneTab", selected: false, folderPath: OPEN.URI, messageId: "a@x", messageFolderPath: OPEN.URI, subject: "Subject a@x" },
      { id: "tab:7:1", kind: "tab", mode: "mailMessageTab", selected: true, messageId: "b@x", messageFolderPath: OPEN.URI, subject: "Subject b@x" },
      { id: "tab:7:2", kind: "tab", mode: "calendar", selected: false, title: "Calendar" },
      { id: "window:100", kind: "window", mode: "messageWindow", messageId: "c@x", messageFolderPath: OPEN.URI, subject: "Subject c@x" },
      { id: "compose:101", kind: "compose", mode: "compose", subject: "Draft subject", draftMessageId: "d@x", draftFolderPath: DRAFTS.URI },
      { id: "compose:102", kind: "compose", mode: "compose", subject: "New" },
    ]);
  });

  it("reveals nothing about views in restricted accounts", () => {
    const hidden = hdr("secret@x", RESTRICTED, 1, "Secret subject");
    const gui = makeGui({ messageTabs: [hidden], messageWindows: [hidden] });
    const { views } = plain(loadViews(gui).listTabs());
    const restricted = views.filter(v => v.restricted);
    assert.equal(restricted.length, 2);
    assert.doesNotMatch(JSON.stringify(views), /secret|Secret/);
  });

  it("withholds the subject of encrypted messages unless access is enabled", () => {
    const encrypted = { ...hdr("e@x", OPEN, 1, "Decrypted subject"), encrypted: true };
    const off = plain(loadViews(makeGui({ messageTabs: [encrypted] })).listTabs()).views[1];
    assert.equal(off.subject, "[Encrypted message]");
    assert.equal(off.encryptedContentWithheld, true);
    const on = plain(loadViews(makeGui({ messageTabs: [encrypted] }), { allowEncrypted: true }).listTabs()).views[1];
    assert.equal(on.subject, "Decrypted subject");
  });
});

describe("closeTabs", () => {
  it("closes every message tab and window showing a messageId, but not the folder tab", async () => {
    const a = hdr("a@x", OPEN, 1);
    const gui = makeGui({ folderSelection: a, messageTabs: [a, hdr("b@x", OPEN, 2)], messageWindows: [a] });
    const result = plain(await loadViews(gui).closeTabs(undefined, "a@x"));

    assert.deepStrictEqual(result.closed, ["tab:7:1", "window:100"]);
    assert.deepStrictEqual(result.skipped, [{ id: "tab:7:0", reason: "folder tab only selects the message" }]);
    assert.deepStrictEqual(gui.tabmail.tabInfo.map(t => t.message?.messageId), ["a@x", "b@x", undefined]);
    assert.equal(gui.windows["mail:messageWindow"].length, 0);
  });

  it("limits a messageId close to one folder when folderPath is given", async () => {
    const gui = makeGui({ messageTabs: [hdr("a@x", OPEN, 1), hdr("a@x", DRAFTS, 5)] });
    const result = plain(await loadViews(gui).closeTabs(undefined, "a@x", DRAFTS.URI));
    assert.deepStrictEqual(result.closed, ["tab:7:2"]);
  });

  it("closes by id and explains every id it did not close", async () => {
    const gui = makeGui({ messageTabs: [hdr("a@x", OPEN, 1), hdr("s@x", RESTRICTED, 2)], composeWindows: [{ subject: "x" }] });
    const result = plain(await loadViews(gui).closeTabs(["tab:7:1", "tab:7:0", "tab:7:2", "tab:7:3", "compose:100", "tab:9:9"]));

    assert.deepStrictEqual(result.closed, ["tab:7:1"]);
    assert.deepStrictEqual(result.skipped, [
      { id: "tab:7:0", reason: "tab cannot be closed" },
      { id: "tab:7:2", reason: "account not accessible" },
      { id: "tab:7:3", reason: "only mail tabs and message windows are closed by this tool" },
      { id: "compose:100", reason: "compose windows are not closed by this tool" },
      { id: "tab:9:9", reason: "not found" },
    ]);
    assert.ok(gui.tabmail.tabInfo.some(t => t.mode.name === "calendar"), "the calendar tab stays open");
  });

  it("reports a view that is still open after closing as stillOpen, never as closed", async () => {
    const gui = makeGui({ messageTabs: [hdr("a@x", OPEN, 1)], stuckIds: [1] });
    const result = plain(await loadViews(gui).closeTabs(["tab:7:1"]));
    assert.deepStrictEqual(result.closed, []);
    assert.deepStrictEqual(result.stillOpen, ["tab:7:1"]);
  });

  it("requires exactly one of viewIds and messageId", async () => {
    const api = loadViews(makeGui());
    assert.match((await api.closeTabs(undefined, undefined)).error, /either viewIds or messageId/);
    assert.match((await api.closeTabs(["tab:7:0"], "a@x")).error, /either viewIds or messageId/);
    assert.match((await api.closeTabs([])).error, /non-empty array/);
    assert.match((await api.closeTabs(undefined, "a@x", "imap://nowhere")).error, /Folder not found/);
  });
});

describe("closeViewsShowingMessages", () => {
  it("closes only message views of the same stored message", () => {
    const target = hdr("a@x", OPEN, 1);
    const sameIdOtherFolder = hdr("a@x", DRAFTS, 1);
    const sameMessageOtherWrapper = hdr("a@x", OPEN, 1);
    const gui = makeGui({ folderSelection: target, messageTabs: [sameIdOtherFolder, sameMessageOtherWrapper], messageWindows: [target] });
    const closed = plain(loadViews(gui).closeViewsShowingMessages([target]));

    assert.deepStrictEqual(closed, ["tab:7:2", "window:100"]);
    assert.deepStrictEqual(gui.tabmail.tabInfo.map(t => t.tabId), [0, 1, 3]);
  });
});

describe("editDraft", () => {
  function draftFolder() {
    return { ...DRAFTS, getFlag: (flag) => flag === 0x400, getUriForMsg: (m) => `imap-message://${DRAFTS.URI}#${m.messageKey}` };
  }

  it("opens the draft in a compose window and reports the window read back", async () => {
    const draft = hdr("d@x", draftFolder(), 9);
    const gui = makeGui();
    const api = loadViews(gui, { messages: [draft] });
    const result = plain(await api.editDraft("d@x", DRAFTS.URI));

    assert.deepStrictEqual(result, { success: true, viewId: "compose:100", draftMessageId: "d@x" });
    assert.equal(api.composeCalls.length, 1);
    const [, origHdr, uri, type, format, identity] = api.composeCalls[0];
    assert.equal(origHdr, draft);
    assert.equal(uri, `imap-message://${DRAFTS.URI}#9`);
    assert.equal(type, 9);
    assert.equal(format, 0);
    assert.equal(identity.key, "id1");
  });

  it("focuses a compose window already editing the draft instead of opening another", async () => {
    const draft = hdr("d@x", draftFolder(), 9);
    const gui = makeGui({ composeWindows: [{ draft, subject: "Draft" }] });
    const api = loadViews(gui, { messages: [draft] });
    const result = plain(await api.editDraft("d@x", DRAFTS.URI));

    assert.equal(result.alreadyOpen, true);
    assert.equal(result.viewId, "compose:100");
    assert.equal(api.composeCalls.length, 0);
    assert.equal(gui.windows.msgcompose[0].focused, 1);
  });

  it("reports an unconfirmed state when no compose window appears", async () => {
    const draft = hdr("d@x", draftFolder(), 9);
    const api = loadViews(makeGui(), { messages: [draft], composeOpens: false });
    const result = await api.editDraft("d@x", DRAFTS.URI);
    assert.match(result.error, /No compose window .* unconfirmed/);
  });

  it("refuses messages outside a Drafts folder and encrypted drafts", async () => {
    const inbox = { ...OPEN, getFlag: () => false };
    const notDraft = hdr("m@x", inbox, 1);
    const encrypted = { ...hdr("e@x", draftFolder(), 2), encrypted: true };
    const api = loadViews(makeGui(), { messages: [notDraft, encrypted] });

    assert.match((await api.editDraft("m@x", OPEN.URI)).error, /Drafts flag only/);
    assert.match((await api.editDraft("e@x", DRAFTS.URI)).error, /Encrypted messages/);
    assert.match((await api.editDraft("missing@x", DRAFTS.URI)).error, /Message not found/);
    assert.equal(api.composeCalls.length, 0);
  });
});
