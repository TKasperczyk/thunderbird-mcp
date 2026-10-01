"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
const modes = { Now: 0, Later: 1, SaveAsDraft: 4 };
const failure = 0x80004005;

// Searched test/ for DIRECT SEND HELPER, listener mocks and disposition tests.
// save-draft-replace supplies the marker+VM pattern; no shared send harness exists.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function loadSend(options = {}) {
  const state = { calls: [], dispositions: [], result: undefined };
  const identity = { key: "sender", email: "sender@example.test" };
  const timer = {
    cancelled: false,
    initWithCallback(callback, delay) { this.callback = callback; this.delay = delay; },
    cancel() { this.cancelled = true; },
  };
  const creation = new Promise((resolve, reject) => {
    state.deliveryStarted = resolve;
    state.creationFailed = reject;
  });
  const folder = {
    server: {}, getUriForMsg: () => "message-uri",
    addMessageDispositionState(hdr, disposition) { state.dispositions.push({ hdr, disposition }); },
  };
  const hdr = { folder, messageId: "original", author: "author@example.test", subject: "Original", date: 0 };
  const sandbox = {
    Ci: {
      nsIMsgCompDeliverMode: modes, nsITimer: { TYPE_ONE_SHOT: 0 },
      nsIMsgCompType: { New: 0, Reply: 1, ReplyAll: 2, ForwardInline: 3 },
      nsIMsgFolder: { nsMsgDispositionState_Replied: 10, nsMsgDispositionState_Forwarded: 11 },
    },
    Cc: {
      "@mozilla.org/timer;1": { createInstance: () => timer },
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({ setHeader() {}, addAttachment() {} }) },
      "@mozilla.org/messengercompose/send;1": { createInstance: () => ({
        createAndSendMessage(...args) {
          state.calls.push(args);
          if (options.legacy && args.length === 16) {
            throw Object.assign(new Error("Not enough arguments"), { result: 0x80570001 });
          }
          if (options.createError) throw options.createError;
          state.listener = args[args.length === 16 ? 12 : 14];
          return options.legacy ? undefined : creation;
        },
      }) },
    },
    ChromeUtils: {
      generateQI: () => () => {},
      importESModule: () => ({ MsgHdrToMimeMessage: (msg, _listener, callback) => callback(msg, {}) }),
    },
    Components: { isSuccessCode: status => status === 0 },
    MailServices: { accounts: { accounts: [{ key: "account", identities: [identity] }] } },
    descsToMsgAttachments(descs) {
      if (options.attachmentError) throw options.attachmentError;
      return descs;
    },
    isSkipReviewBlocked: () => !!options.blocked,
    findMessage: () => ({ msgHdr: hdr, folder }),
    filePathsToAttachDescs: () => ({ descs: [] }),
    setComposeIdentity(params) { params.identity = identity; },
    resolveComposeFormat: () => ({ useHtml: false, format: 0 }),
    PREF_ALLOW_ENCRYPTED_MESSAGES: "encrypted",
    isPrivacyOptInEnabled: () => false,
    extractPlainTextBody: () => "Original body",
    isEncryptedMimeMessage: () => false,
    withSignedContent: (_fetch, _hdr, _allowEncrypted, callback) => callback,
    hasInlinePgpBodyArmor: () => false,
  };
  const dispositionStart = source.indexOf("function markMessageDispositionState(");
  const dispositionEnd = source.indexOf("// BEGIN DIRECT SEND HELPER", dispositionStart);
  assert.ok(dispositionStart >= 0 && dispositionEnd > dispositionStart, "Missing disposition helper boundaries");
  vm.createContext(sandbox);
  vm.runInContext([
    source.slice(dispositionStart, dispositionEnd), snippet("DIRECT SEND HELPER"),
    snippet("REPLY TOOL"), snippet("FORWARD TOOL"),
  ].join("\n"), sandbox);
  return {
    state, timer, hdr,
    start(tool = "direct", mode) {
      let result;
      if (tool === "reply") {
        result = sandbox.replyToMessage("original", "folder", "Reply", false, false,
          undefined, undefined, undefined, undefined, undefined, true);
      } else if (tool === "forward") {
        result = sandbox.forwardMessage("original", "folder", "recipient@example.test", "Forward", false,
          undefined, undefined, undefined, undefined, true);
      } else {
        result = sandbox.sendMessageDirectly({ body: "Body", addAttachment() {} },
          options.noIdentity ? null : identity, [], "message-uri", 0, mode, "text/plain");
      }
      state.promise = result.then(value => { state.result = value; return value; });
      return state.promise;
    },
  };
}

async function assertPending(harness) {
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(harness.state.result, undefined, "SMTP outcome must still be pending");
  assert.equal(harness.timer.cancelled, false, "delivery timer must stay active");
  assert.equal(harness.state.dispositions.length, 0);
}

const errors = [
  ["SMTP failure", h => h.state.listener.onStopSending("id", failure), /Send failed.*80004005/],
  ["send not performed", h => h.state.listener.onSendNotPerformed("id", failure), /Send was not performed/],
  ["transport security failure", h => h.state.listener.onTransportSecurityError("id", failure, {}, "smtp.example.test"), /Transport security error: smtp\.example\.test/],
  ["promise rejection", h => h.state.creationFailed(new Error("create rejected")), /create rejected/],
  ["timeout", h => h.timer.callback.notify(), /outcome is unknown; check Sent and the Outbox before retrying/],
];

describe("direct send SMTP completion", () => {
  for (const legacy of [false, true]) {
    for (const copyStatus of [0, failure]) {
      it(`waits through creation and Sent copy until SMTP succeeds, legacy=${legacy}, copy=${copyStatus}`, async () => {
        const h = loadSend({ legacy });
        const result = h.start(); // omitted mode also means Now
        h.state.deliveryStarted();
        h.state.listener.onStopCopy(copyStatus);
        h.state.listener.onStartSending();
        h.state.listener.onProgress();
        h.state.listener.onSendProgress();
        h.state.listener.onStatus();
        h.state.listener.onStartCopy();
        h.state.listener.setMessageKey();
        h.state.listener.onGetDraftFolderURI();
        await assertPending(h);
        assert.equal(h.timer.delay, 120000);
        assert.equal(h.state.calls.at(-1).length, legacy ? 18 : 16);
        h.state.listener.onStopSending("id", 0);
        assert.equal((await result).success, true);
        assert.equal(h.state.result.message, "Message sent");
        assert.equal(h.timer.cancelled, true);
      });
    }
  }

  for (const [label, trigger, pattern] of errors) {
    it(`reports ${label} after an early Sent copy, ignoring later success`, async () => {
      const h = loadSend();
      const result = h.start("direct", modes.Now);
      h.state.listener.onStopCopy(0);
      if (label !== "promise rejection") h.state.deliveryStarted();
      await assertPending(h);
      trigger(h);
      assert.match((await result).error, pattern);
      assert.equal(h.state.result.success, undefined);
      const original = h.state.result;
      h.state.listener.onStopSending("id", 0);
      h.state.listener.onStopCopy(0);
      h.state.deliveryStarted();
      await new Promise(resolve => setImmediate(resolve));
      assert.equal(h.state.result, original);
    });
  }

  it("retains confirmed SMTP success when copy or promise fails afterward", async () => {
    const h = loadSend();
    const result = h.start();
    h.state.listener.onStopSending("id", 0);
    h.state.listener.onStopCopy(failure);
    h.state.creationFailed(new Error("late rejection"));
    h.state.listener.onSendNotPerformed("id", failure);
    assert.equal((await result).success, true);
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(h.state.result.error, undefined);
  });

  for (const [label, options, pattern] of [
    ["missing identity", { noIdentity: true }, /No identity/],
    ["synchronous creation failure", { createError: new Error("create threw") }, /create threw/],
    ["attachment failure", { attachmentError: new Error("attachments invalid") }, /attachments invalid/],
  ]) {
    it(`reports ${label} without success`, async () => {
      const h = loadSend(options);
      const result = await h.start();
      assert.match(result.error, pattern);
      assert.equal(result.success, undefined);
      if (!options.createError) assert.equal(h.state.calls.length, 0);
    });
  }
});

describe("draft and queued completion", () => {
  for (const mode of [modes.SaveAsDraft, modes.Later]) {
    for (const completion of ["promise", "copy"]) {
      it(`still completes ${mode} via ${completion}`, async () => {
        const h = loadSend();
        const result = h.start("direct", mode);
        if (completion === "promise") h.state.deliveryStarted();
        else h.state.listener.onStopCopy(0);
        assert.equal((await result).success, true);
        assert.equal(h.timer.cancelled, true);
      });
    }
    it(`reports copy failure for ${mode} even if creation later fulfills`, async () => {
      const h = loadSend();
      const result = h.start("direct", mode);
      h.state.listener.onStopCopy(failure);
      h.state.deliveryStarted();
      assert.match((await result).error, /Save failed.*80004005/);
    });
  }
});

describe("reply and forward dispositions follow the SMTP result", () => {
  for (const tool of ["reply", "forward"]) {
    it(`marks ${tool} only after SMTP success`, async () => {
      const h = loadSend();
      const result = h.start(tool);
      h.state.deliveryStarted();
      h.state.listener.onStopCopy(0);
      await assertPending(h);
      h.state.listener.onStopSending("id", 0);
      assert.equal((await result).success, true);
      assert.deepEqual(h.state.dispositions, [{ hdr: h.hdr, disposition: tool === "reply" ? 10 : 11 }]);
      h.state.listener.onStopSending("id", 0);
      await new Promise(resolve => setImmediate(resolve));
      assert.equal(h.state.dispositions.length, 1);
    });
    for (const [label, trigger, pattern] of errors) {
      it(`does not mark ${tool} on ${label}, including late SMTP completion`, async () => {
        const h = loadSend();
        const result = h.start(tool);
        h.state.listener.onStopCopy(0);
        trigger(h);
        assert.match((await result).error, pattern);
        h.state.listener.onStopSending("id", 0);
        h.state.deliveryStarted();
        await new Promise(resolve => setImmediate(resolve));
        assert.equal(h.state.dispositions.length, 0);
        assert.equal(h.state.result.success, undefined);
      });
    }
    it(`keeps the skipReview block ahead of ${tool} delivery`, async () => {
      const h = loadSend({ blocked: true });
      assert.match((await h.start(tool)).error, /blocks skipReview/);
      assert.equal(h.state.calls.length, 0);
      assert.equal(h.state.dispositions.length, 0);
    });
  }
});
