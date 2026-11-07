"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { makeDirectSendPrefs, prefTypes, clearOriginalRawFixture } = require("./helpers/direct-send-prefs.cjs");

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
  const state = { calls: [], composes: [], attached: [], result: undefined };
  const identity = { key: "sender", email: "sender@example.test" };
  const timer = {
    cancelled: false,
    initWithCallback(callback, delay) {
      this.cancelled = false;
      this.callback = callback;
      this.delay = delay;
      if (options.timerInitError) throw options.timerInitError;
    },
    cancel() { this.cancelled = true; },
  };
  const creation = new Promise((resolve, reject) => {
    state.deliveryStarted = resolve;
    state.creationFailed = reject;
  });
  const folder = { server: {}, getUriForMsg: () => "message-uri" };
  const hdr = { folder, messageId: "original", author: "author@example.test", subject: "Original", date: 0 };
  // Windowless nsIMsgCompose: it shares the listener slot and creation promise
  // with the nsIMsgSend mock so one error table drives both send paths.
  const makeCompose = () => {
    const compose = {
      stateListener: null, unregistered: 0,
      sendListeners: new Set(), stateListeners: new Set(),
      initialize(params, win, editor) {
        compose.init = { params, win, editor };
        if (options.initError) throw options.initError;
      },
      addMsgSendListener(listener) {
        // Keep callback handles for late-notification tests, independently of
        // the native object's strong listener registrations.
        state.listener = listener;
        compose.sendListeners.add(listener);
        if (options.sendListenerError) throw options.sendListenerError;
      },
      removeMsgSendListener(listener) { compose.sendListeners.delete(listener); },
      RegisterStateListener(listener) {
        compose.stateListener = listener;
        compose.stateListeners.add(listener);
        if (options.stateListenerError) throw options.stateListenerError;
      },
      UnregisterStateListener(listener) {
        if (compose.stateListeners.delete(listener)) compose.unregistered++;
      },
      sendMsg(...args) {
        compose.sendArgs = args;
        if (options.createError) throw options.createError;
        return options.legacy ? undefined : creation;
      },
    };
    state.composes.push(compose);
    return compose;
  };
  const sandbox = {
    ...clearOriginalRawFixture,
    Ci: {
      nsIPrefBranch: prefTypes,
      nsIMsgCompDeliverMode: modes, nsITimer: { TYPE_ONE_SHOT: 0 },
      nsIMsgCompType: { New: 0, Reply: 1, ReplyAll: 2, ForwardInline: 3 },
      nsIMsgCompFormat: { Default: 0, HTML: 1, PlainText: 2, OppositeOfDefault: 3 },
      nsIMsgFolder: { nsMsgDispositionState_Replied: 10, nsMsgDispositionState_Forwarded: 11 },
    },
    Cc: {
      "@mozilla.org/timer;1": { createInstance() {
        if (options.timerCreateError) throw options.timerCreateError;
        state.timerCreated = true;
        return timer;
      } },
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({
        headers: {}, setHeader(name, value) { this.headers[name] = value; }, addAttachment(att) { state.attached.push(att); },
      }) },
      "@mozilla.org/messengercompose/compose;1": { createInstance: makeCompose },
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
      importESModule: () => ({ MsgHdrToMimeMessage: (msg, _listener, callback) => callback(msg, { contentType: "text/plain", body: "Original body" }) }),
    },
    Components: { isSuccessCode: status => status === 0 },
    Services: { prefs: makeDirectSendPrefs() },
    MailServices: { accounts: { get accounts() {
      if (options.accountError) throw options.accountError;
      return [
        { key: "other", identities: [{ key: "someone-else" }] },
        ...(options.noAccount ? [] : [{
          key: Object.hasOwn(options, "accountKey") ? options.accountKey : "account", identities: [identity],
        }]),
      ];
    } } },
    descsToMsgAttachments(descs) {
      if (options.attachmentError) throw options.attachmentError;
      return descs;
    },
    isSkipReviewBlocked: () => !!options.blocked,
    findMessage: () => ({ msgHdr: hdr, folder }),
    filePathsToAttachDescs: () => ({ descs: [] }),
    materializeAttachmentDescs: descs => descs,
    buildBodyWithSignature: body => body || "",
    buildSignatureFragment: () => "",
    setComposeIdentity(params) { params.identity = identity; },
    resolveComposeFormat: () => ({ useHtml: false, format: 0 }),
    PREF_ALLOW_ENCRYPTED_MESSAGES: "encrypted",
    isPrivacyOptInEnabled: () => false,
    getDirectQuoteContent: () => ({ body: "Original body" }),
  };
  vm.createContext(sandbox);
  vm.runInContext([
    snippet("DIRECT SEND ENCRYPTION GUARD"), snippet("DIRECT SEND ORIGINAL GUARD"),
    snippet("DIRECT SEND HELPER"), snippet("COMPOSE SEND HELPER"),
    snippet("REPLY TOOL"), snippet("FORWARD TOOL"),
  ].join("\n"), sandbox);
  return {
    state, timer, hdr, identity,
    get compose() { return state.composes.at(-1); },
    start(tool = "direct", mode) {
      let result;
      if (tool === "compose") {
        const composeFields = sandbox.Cc["@mozilla.org/messengercompose/composefields;1"].createInstance();
        result = sandbox.sendThroughCompose(
          { identity: options.noIdentity ? null : identity, composeFields, type: 0 },
          options.attachments || [], options.useHtml ?? false);
      } else if (tool === "reply") {
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
  for (const compose of harness.state.composes) {
    assert.equal(compose.sendListeners.size, 1, "send listener must stay registered until settlement");
    assert.equal(compose.stateListeners.size, 1, "state listener must stay registered until settlement");
  }
}

// Searched for listener-cleanup assertions: compose-window-claim has a local
// window-specific assertClean, but no shared send-lifecycle assertion exists.
function assertComposeReleased(harness) {
  for (const compose of harness.state.composes) {
    assert.equal(compose.sendListeners.size, 0, "send listener is released");
    assert.equal(compose.stateListeners.size, 0, "state listener is released");
  }
  assert.equal(harness.timer.cancelled, !!harness.state.timerCreated, "created timer is cancelled");
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

describe("windowless nsIMsgCompose send", () => {
  it("initializes a windowless compose and sends Now with the identity's account", async () => {
    const h = loadSend({ attachments: [{ url: "file:///a.pdf" }] });
    const result = h.start("compose");
    const compose = h.compose;
    assert.equal(compose.init.win, null);
    assert.equal(compose.init.editor, null);
    assert.equal(compose.init.params.identity, h.identity);
    assert.deepEqual(compose.sendArgs, [modes.Now, h.identity, "account", null, null]);
    assert.deepEqual(h.state.attached, [{ url: "file:///a.pdf" }]);
    assert.equal(h.state.calls.length, 0, "must not create a bare nsIMsgSend");
    h.state.deliveryStarted();
    await assertPending(h);
    h.state.listener.onStopSending("id", 0);
    assert.equal((await result).success, true);
    assert.equal(h.timer.cancelled, true);
    assert.equal(compose.unregistered, 1, "state listener is released");
    assertComposeReleased(h);
  });

  it("succeeds on ComposeProcessDone when the send listener stays silent", async () => {
    const h = loadSend();
    const result = h.start("compose");
    h.compose.stateListener.ComposeProcessDone(0);
    assert.equal((await result).success, true);
    assertComposeReleased(h);
  });

  it("reports a failed ComposeProcessDone", async () => {
    const h = loadSend();
    const result = h.start("compose");
    h.compose.stateListener.ComposeProcessDone(failure);
    assert.match((await result).error, /Send failed.*80004005/);
    assertComposeReleased(h);
  });

  it("retains confirmed delivery when ComposeProcessDone or the promise fails afterward", async () => {
    const h = loadSend();
    const result = h.start("compose");
    h.state.listener.onStopSending("id", 0);
    h.compose.stateListener.ComposeProcessDone(failure);
    h.state.creationFailed(new Error("late rejection"));
    h.state.listener.onSendNotPerformed("id", failure);
    assert.equal((await result).success, true);
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(h.state.result.error, undefined);
    assert.equal(h.compose.unregistered, 1);
    h.timer.callback.notify();
    assertComposeReleased(h);
  });

  for (const [label, options, pattern] of [
    ["missing owning account", { noAccount: true }, /account/i],
    ["empty account key", { accountKey: "" }, /account/i],
    ["missing account key", { accountKey: undefined }, /account/i],
    ["account lookup exception", { accountError: new Error("account lookup failed") }, /account lookup failed/],
  ]) {
    it(`refuses a ${label} before constructing or sending a compose`, async () => {
      const h = loadSend(options);
      const result = await h.start("compose");
      assert.match(result.error, pattern);
      assert.equal(result.success, undefined);
      assert.equal(h.state.composes.length, 0);
      assert.equal(h.state.calls.length, 0);
      assertComposeReleased(h);
    });
  }

  for (const [label, trigger, pattern] of errors) {
    it(`reports ${label} and ignores later completion`, async () => {
      const h = loadSend();
      const result = h.start("compose");
      if (label !== "promise rejection") h.state.deliveryStarted();
      await assertPending(h);
      trigger(h);
      assert.match((await result).error, pattern);
      assertComposeReleased(h);
      const original = h.state.result;
      h.state.listener.onStopSending("id", 0);
      h.compose.stateListener.ComposeProcessDone(0);
      h.timer.callback.notify();
      await new Promise(resolve => setImmediate(resolve));
      assert.equal(h.state.result, original);
      assert.equal(h.compose.unregistered, 1);
      assertComposeReleased(h);
    });
  }

  for (const [label, options, pattern] of [
    ["missing identity", { noIdentity: true }, /No identity/],
    ["timer creation failure", { timerCreateError: new Error("timer creation threw") }, /timer creation threw/],
    ["timer initialization failure", { timerInitError: new Error("timer init threw") }, /timer init threw/],
    ["initialize failure", { initError: new Error("init threw") }, /init threw/],
    ["send listener registration failure", { sendListenerError: new Error("send listener threw") }, /send listener threw/],
    ["state listener registration failure", { stateListenerError: new Error("state listener threw") }, /state listener threw/],
    ["sendMsg failure", { createError: new Error("sendMsg threw") }, /sendMsg threw/],
    ["attachment failure", { attachmentError: new Error("attachments invalid") }, /attachments invalid/],
  ]) {
    it(`reports ${label} without success`, async () => {
      const h = loadSend(options);
      const result = await h.start("compose");
      assert.match(result.error, pattern);
      assert.equal(result.success, undefined);
      assert.equal(h.state.calls.length, 0);
      assertComposeReleased(h);
      // Even a callback retained before a partially successful registration
      // must not change the failure or re-register either listener.
      h.state.listener?.onStopSending("late", 0);
      h.compose?.stateListener?.ComposeProcessDone(0);
      h.timer.callback?.notify();
      assert.equal(h.state.result, result);
      assertComposeReleased(h);
    });
  }

  it("does not depend on sendMsg returning a promise", async () => {
    const h = loadSend({ legacy: true });
    const result = h.start("compose");
    h.state.listener.onStopSending("id", 0);
    assert.equal((await result).success, true);
    assertComposeReleased(h);
  });
});

describe("reply and forward send through nsIMsgCompose", () => {
  for (const tool of ["reply", "forward"]) {
    it(`sends ${tool} with the original's URI and type so Thunderbird threads and flags it`, async () => {
      const h = loadSend();
      const result = h.start(tool);
      await assertPending(h);
      const { params } = h.compose.init;
      assert.equal(params.originalMsgURI, "message-uri");
      assert.equal(params.type, tool === "reply" ? 1 : 3);
      assert.equal(params.composeFields.subject, undefined, "initialize sets the Re:/Fwd: subject");
      assert.equal(h.state.calls.length, 0);
      h.state.listener.onStopSending("id", 0);
      const value = await result;
      assert.equal(value.success, true);
      assert.match(value.message, tool === "reply" ? /^Reply sent$/ : /^Forward sent/);
    });

    if (tool === "reply") {
      it("keeps the reply threading headers", async () => {
        const h = loadSend();
        const result = h.start(tool);
        await assertPending(h);
        const fields = h.compose.init.params.composeFields;
        assert.equal(fields.references, "<original>");
        assert.equal(fields.headers["In-Reply-To"], "<original>");
        h.state.listener.onStopSending("id", 0);
        assert.equal((await result).success, true);
      });
    }

    for (const [label, trigger, pattern] of errors) {
      it(`reports ${tool} ${label}`, async () => {
        const h = loadSend();
        const result = h.start(tool);
        await assertPending(h);
        trigger(h);
        const value = await result;
        assert.match(value.error, pattern);
        assert.equal(value.success, undefined);
      });
    }

    it(`keeps the skipReview block ahead of ${tool} delivery`, async () => {
      const h = loadSend({ blocked: true });
      assert.match((await h.start(tool)).error, /blocks skipReview/);
      assert.equal(h.state.calls.length, 0);
      assert.equal(h.state.composes.length, 0);
    });
  }
});
