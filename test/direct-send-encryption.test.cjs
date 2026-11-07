"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { makeDirectSendPrefs, prefTypes, clearOriginalRawFixture, composeHeaderParser } = require("./helpers/direct-send-prefs.cjs");
const { makeDirectSendAttachments, assertNoAttachmentEffects } = require("./helpers/direct-send-attachments.cjs");
const { SEND_LATER_NOW, SEND_LATER_READY, SEND_LATER_HEADER_PARSER, makeSendLaterClock } = require("./helpers/send-later.cjs");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
const tools = ["sendMail", "replyToMessage", "forwardMessage"];
const identityPrefix = "mail.identity.sender.";
const defaultPrefix = "mail.identity.default.";
const autoEnable = "mail.e2ee.auto_enable";

// Searched for shared marker extraction -> found per-suite snippet loaders.
// Reuse the direct-send-outcome marker assertions and node:vm pattern.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

// Searched for native compose construction counters and identity selection
// harnesses -> found direct-send-outcome's native mocks, but no runtime that
// combines real identity selection with typed preferences and all three tools.
function loadEncryptionTools(options = {}) {
  const attachments = makeDirectSendAttachments();
  const calls = {
    created: { params: 0, fields: 0, compose: 0, sender: 0 },
    sends: [], drafts: 0, draftSaves: 0, reviews: 0, mime: 0, preferences: [],
    attachments: attachments.calls,
  };
  const identity = {
    key: "sender", email: "sender@example.test", composeHtml: false,
    signMail: true, attachPgpKey: true, autoEncryptDrafts: true, protectSubject: true,
    get encryptionPolicy() { throw new Error("The silently defaulting identity getter must not be trusted"); },
    ...options.identity,
  };
  if (options.keyError) {
    Object.defineProperty(identity, "key", { get() { throw new Error("Identity key unavailable"); } });
  }
  const prefs = makeDirectSendPrefs(options.prefs);
  for (const method of ["getPrefType", "getIntPref", "getStringPref", "getBoolPref"]) {
    const read = prefs[method];
    prefs[method] = name => {
      calls.preferences.push({ method, name });
      if (options.unreadable?.includes(name) || options.readFailure === `${method}:${name}`) {
        throw new Error("Preference unavailable");
      }
      if (Object.hasOwn(options.returnValues || {}, `${method}:${name}`)) {
        return options.returnValues[`${method}:${name}`];
      }
      return read(name);
    };
  }
  const folder = { server: {}, getUriForMsg: () => "message-uri" };
  const hdr = { folder, messageId: "original@example.test", subject: "Subject", author: "author@example.test", date: 0 };
  const account = { key: "account", defaultIdentity: options.defaultIdentity || identity, identities: [identity] };
  const sandbox = {
    ...clearOriginalRawFixture,
    Date: makeSendLaterClock().Date,
    __tbMcpSendLaterInfo: Object.hasOwn(options, "info") ? options.info : SEND_LATER_READY,
    Ci: {
      nsIPrefBranch: prefTypes,
      nsIMsgCompType: { New: 0, Reply: 1, ReplyAll: 2, ForwardInline: 3 },
      nsIMsgCompFormat: { Default: 0, HTML: 1, PlainText: 2 },
      nsIMsgCompDeliverMode: { Now: 0, SaveAsDraft: 4 },
      nsITimer: { TYPE_ONE_SHOT: 0 },
    },
    Services: { prefs },
    Components: { isSuccessCode: status => status === 0 },
    ChromeUtils: {
      generateQI: () => () => {},
      importESModule: () => ({
        MsgHdrToMimeMessage(msg, _listener, callback) {
          calls.mime++;
          const finish = () => callback(msg, { contentType: "text/plain", body: "Original text" });
          if (options.deferMime) calls.finishMime = finish;
          else finish();
        },
      }),
    },
    Cc: {
      "@mozilla.org/timer;1": { createInstance: () => ({ initWithCallback() {}, cancel() {} }) },
      "@mozilla.org/messengercompose/composeparams;1": { createInstance() {
        calls.created.params++;
        return {};
      } },
      "@mozilla.org/messengercompose/composefields;1": { createInstance() {
        calls.created.fields++;
        return { addAttachment() {}, setHeader() {} };
      } },
      "@mozilla.org/messengercompose/compose;1": { createInstance() {
        calls.created.compose++;
        return {
          initialize(params) { this.params = params; },
          addMsgSendListener(listener) { this.listener = listener; },
          removeMsgSendListener() {},
          RegisterStateListener() {},
          UnregisterStateListener() {},
          sendMsg() {
            calls.sends.push(this.params);
            this.listener.onStopSending("sent@example.test", 0);
            return Promise.resolve();
          },
        };
      } },
      "@mozilla.org/messengercompose/send;1": { createInstance() {
        calls.created.sender++;
        return {
          createAndSendMessage(...args) {
            assert.equal(args[6], 4, "Only saveDraft uses the direct sender");
            calls.drafts++;
            args[12].onStopCopy(0);
            return Promise.resolve();
          },
        };
      } },
      "@mozilla.org/messengercompose;1": { getService: () => ({
        OpenComposeWindowWithParams() { calls.reviews++; },
      }) },
    },
    MailServices: { headerParser: { ...composeHeaderParser, ...SEND_LATER_HEADER_PARSER }, accounts: {
      accounts: [account], defaultAccount: account,
      findAccountForServer: () => account,
    } },
    findIdentity: from => from === identity.email ? identity : null,
    findIdentityIn: (_accounts, from) => from === identity.email ? identity : null,
    getAccessibleAccounts: () => [account],
    isAccountAllowed: () => true,
    isSkipReviewBlocked: () => options.blocked === true,
    isToolEnabled: name => !(options.disabledTools || []).includes(name),
    findMessage: () => ({ msgHdr: hdr, folder }),
    filePathsToAttachDescs: attachments.convert,
    materializeAttachmentDescs: attachments.materialize,
    descsToMsgAttachments: attachments.toNative,
    buildBodyWithSignature: body => body,
    buildSignatureFragment: () => "",
    getDirectQuoteContent: () => ({ body: "Original text" }),
    getReplyAllCcRecipients: () => "",
    rawMessageHasForwardableParts: () => false,
    PREF_ALLOW_ENCRYPTED_MESSAGES: "encrypted",
    isPrivacyOptInEnabled: () => false,
    async openComposeWindowWithCustomizations(...args) {
      calls.reviews++;
      return args[10] ? args[10]({}) : { success: true };
    },
    async saveComposeWindowAsDraft() {
      calls.draftSaves++;
      return { success: true };
    },
  };
  vm.createContext(sandbox);
  vm.runInContext([
    snippet("TWO DIGIT PADDING"),
    snippet("COMPOSE IDENTITY HELPER"), snippet("COMPOSE FORMAT HELPER"),
    snippet("DIRECT SEND ENCRYPTION GUARD"), snippet("DIRECT SEND ORIGINAL GUARD"),
    snippet("DIRECT SEND HELPER"),
    snippet("COMPOSE SEND HELPER"), snippet("OUTBOUND MAIL TOOLS"),
    snippet("TOOL CALL DISPATCH"),
  ].join("\n"), sandbox);
  return {
    calls,
    identity,
    prefs,
    run: (tool, args = {}) => sandbox.dispatchTool(tool, {
      to: "recipient@example.test", subject: "Subject", body: "Body", isHtml: false,
      messageId: hdr.messageId, folderPath: "folder", skipReview: true,
      attachments: attachments.entries, ...args,
    }),
  };
}

// Searched for zero-allocation refusal assertions -> found send-only checks in
// direct-send-outcome, but none checks all native constructors and review text.
function assertRefused(harness, result, mimeReads = 0) {
  assert.equal(result.success, undefined);
  assert.match(result.error, /skipReview: false/);
  assert.match(result.error, /review window/);
  assert.match(result.error, /Thunderbird.*encryption/);
  assert.doesNotMatch(result.error, /enable.*read encrypted/i);
  assert.deepEqual(harness.calls.created, { params: 0, fields: 0, compose: 0, sender: 0 });
  assert.equal(harness.calls.sends.length, 0);
  assert.equal(harness.calls.drafts, 0);
  assert.equal(harness.calls.reviews, 0);
  assert.equal(harness.calls.mime, mimeReads, "An initial identity refusal precedes reading the original");
  assertNoAttachmentEffects(harness.calls.attachments);
}

const invalidPolicies = [
  ["positive override 1", { prefs: { [`${identityPrefix}encryptionpolicy`]: 1 } }],
  ["positive override 2", { prefs: { [`${identityPrefix}encryptionpolicy`]: 2 } }],
  ["future positive override", { prefs: { [`${identityPrefix}encryptionpolicy`]: 3 } }],
  ["inherited positive policy 1", { prefs: { [`${defaultPrefix}encryptionpolicy`]: 1 } }],
  ["inherited positive policy 2", { prefs: { [`${defaultPrefix}encryptionpolicy`]: 2 } }],
  ["OpenPGP encryption requirement", { prefs: { [`${identityPrefix}encryptionpolicy`]: 2, [`${identityPrefix}openpgp_key_id`]: "ABCD" } }],
  ["S/MIME encryption requirement", { prefs: { [`${identityPrefix}encryptionpolicy`]: 2, [`${identityPrefix}encryption_cert_name`]: "certificate" } }],
  ["wrong-type override", { prefs: { [`${identityPrefix}encryptionpolicy`]: "0" } }],
  ["wrong-type inherited policy", { prefs: { [`${defaultPrefix}encryptionpolicy`]: false } }],
  ["missing inherited policy", { prefs: { [`${defaultPrefix}encryptionpolicy`]: undefined } }],
  ["negative policy", { prefs: { [`${identityPrefix}encryptionpolicy`]: -1 } }],
  ["fractional policy", { prefs: { [`${identityPrefix}encryptionpolicy`]: 0.5 } }],
  ["NaN policy", { prefs: { [`${identityPrefix}encryptionpolicy`]: NaN } }],
  ["infinite policy", { prefs: { [`${identityPrefix}encryptionpolicy`]: Infinity } }],
  ["unreadable override type", { readFailure: `getPrefType:${identityPrefix}encryptionpolicy` }],
  ["unreadable inherited type", { readFailure: `getPrefType:${defaultPrefix}encryptionpolicy` }],
  ["unreadable override value", { prefs: { [`${identityPrefix}encryptionpolicy`]: 0 }, readFailure: `getIntPref:${identityPrefix}encryptionpolicy` }],
  ["unreadable inherited value", { readFailure: `getIntPref:${defaultPrefix}encryptionpolicy` }],
  ["silently zero identity getter cannot hide an unreadable policy", {
    identity: { encryptionPolicy: 0 }, readFailure: `getIntPref:${defaultPrefix}encryptionpolicy`,
  }],
  ["integer accessor returning string", { returnValues: { [`getIntPref:${defaultPrefix}encryptionpolicy`]: "0" } }],
  ["missing identity key", { identity: { key: undefined } }],
  ["empty identity key", { identity: { key: "" } }],
  ["blank identity key", { identity: { key: " " } }],
  ["wrong-type identity key", { identity: { key: 42 } }],
  ["throwing identity key", { keyError: true }],
];

describe("direct send identity encryption refuses before native composition", () => {
  for (const tool of tools) {
    for (const [label, options] of invalidPolicies) {
      it(`${tool}: ${label}`, async () => {
        const h = loadEncryptionTools(options);
        assertRefused(h, await h.run(tool));
      });
    }
    for (const capability of ["openpgp_key_id", "encryption_cert_name"]) {
      for (const [label, options] of [
        ["wrong-type override", { prefs: { [`${identityPrefix}${capability}`]: false } }],
        ["wrong-type default", { prefs: { [`${defaultPrefix}${capability}`]: 0 } }],
        ["missing default", { prefs: { [`${defaultPrefix}${capability}`]: undefined } }],
        ["unreadable override type", { readFailure: `getPrefType:${identityPrefix}${capability}` }],
        ["unreadable default value", { readFailure: `getStringPref:${defaultPrefix}${capability}` }],
        ["string accessor returning non-string", { returnValues: { [`getStringPref:${defaultPrefix}${capability}`]: null } }],
      ]) {
        it(`${tool}: ${capability} ${label}`, async () => {
          const h = loadEncryptionTools(options);
          assertRefused(h, await h.run(tool));
        });
      }
      for (const [label, options] of [
        ["automatic encryption enabled", { prefs: { [autoEnable]: true } }],
        ["automatic encryption pref missing", { prefs: { [autoEnable]: undefined } }],
        ["automatic encryption wrong-type", { prefs: { [autoEnable]: "false" } }],
        ["automatic encryption type unreadable", { readFailure: `getPrefType:${autoEnable}` }],
        ["automatic encryption value unreadable", { readFailure: `getBoolPref:${autoEnable}` }],
        ["boolean accessor returning non-boolean", { returnValues: { [`getBoolPref:${autoEnable}`]: 0 } }],
      ]) {
        it(`${tool}: configured ${capability}, ${label}`, async () => {
          const h = loadEncryptionTools({
            ...options,
            prefs: { [`${identityPrefix}${capability}`]: "configured", ...options.prefs },
          });
          assertRefused(h, await h.run(tool));
        });
      }
      it(`${tool}: inherited ${capability} activates automatic encryption protection`, async () => {
        const h = loadEncryptionTools({ prefs: { [`${defaultPrefix}${capability}`]: "configured", [autoEnable]: true } });
        assertRefused(h, await h.run(tool));
      });
    }
  }
});

describe("direct send identity encryption pass-through cases", () => {
  for (const tool of tools) {
    for (const [label, options] of [
      ["absent override inherits zero", {}],
      ["explicit zero overrides encryption-required default", { prefs: { [`${identityPrefix}encryptionpolicy`]: 0, [`${defaultPrefix}encryptionpolicy`]: 2 } }],
      ["signing and unrelated encryption flags do not require encryption", { prefs: {
        [`${identityPrefix}sign_mail`]: true, [`${identityPrefix}attachPgpKey`]: true,
        [`${identityPrefix}autoEncryptDrafts`]: true, [`${identityPrefix}protectSubject`]: true,
      } }],
      ["OpenPGP configured, automatic encryption off", { prefs: { [`${identityPrefix}openpgp_key_id`]: "ABCD" } }],
      ["S/MIME configured, automatic encryption off", { prefs: { [`${identityPrefix}encryption_cert_name`]: "certificate" } }],
      ["both configured, automatic encryption off", { prefs: { [`${identityPrefix}openpgp_key_id`]: "ABCD", [`${identityPrefix}encryption_cert_name`]: "certificate" } }],
      ["no encryption capability, automatic encryption on", { prefs: { [autoEnable]: true } }],
      ["no encryption capability, automatic encryption missing", { prefs: { [autoEnable]: undefined } }],
      ["no encryption capability, automatic encryption wrong-type", { prefs: { [autoEnable]: "true" } }],
      ["no encryption capability, automatic encryption unreadable", { unreadable: [autoEnable] }],
      ["empty capability overrides configured default", { prefs: {
        [`${identityPrefix}openpgp_key_id`]: "", [`${defaultPrefix}openpgp_key_id`]: "ABCD",
        [`${identityPrefix}encryption_cert_name`]: "", [`${defaultPrefix}encryption_cert_name`]: "certificate",
        [autoEnable]: true,
      } }],
    ]) {
      it(`${tool}: ${label}`, async () => {
        const h = loadEncryptionTools(options);
        const result = await h.run(tool);
        assert.equal(result.success, true, result.error);
        assert.deepEqual(h.calls.created, { params: 1, fields: 1, compose: 1, sender: 0 });
        assert.equal(h.calls.sends.length, 1);
        assert.equal(h.calls.reviews, 0);
        assert.equal(h.calls.sends[0].identity, h.identity);
        assert.equal(h.calls.attachments.tempCreates.length, 1);
        assert.equal(h.calls.attachments.writes.length, 1);
        assert.ok(h.calls.attachments.fileAccesses.length > 0);
        assert.equal(h.calls.sends[0].composeFields.composeSecure, undefined, "No cryptographic signing or encryption is requested");
        assert.ok(h.calls.preferences.every(({ name }) => !/sign_mail|attachPgpKey|autoEncryptDrafts|protectSubject/.test(name)));
        if (label.startsWith("no encryption capability")) {
          assert.ok(h.calls.preferences.every(({ name }) => name !== autoEnable));
        }
      });
    }
    it(`${tool}: explicit from selects the identity before checking its policy`, async () => {
      const h = loadEncryptionTools({
        defaultIdentity: { key: "other", email: "other@example.test" },
        prefs: { [`${identityPrefix}encryptionpolicy`]: 1 },
      });
      assertRefused(h, await h.run(tool, { from: h.identity.email }));
    });
    for (const options of [
      { prefs: { [`${identityPrefix}encryptionpolicy`]: 2 } },
      { prefs: { [`${identityPrefix}openpgp_key_id`]: "ABCD", [autoEnable]: true } },
      { unreadable: [`${identityPrefix}encryptionpolicy`] },
    ]) {
      it(`${tool}: review bypasses the direct-send identity gate, ${JSON.stringify(options)}`, async () => {
        const h = loadEncryptionTools(options);
        const result = await h.run(tool, { skipReview: false });
        assert.equal(result.success, true, result.error);
        assert.equal(h.calls.reviews, 1);
        assert.equal(h.calls.sends.length, 0);
        assert.equal(h.calls.preferences.filter(({ name }) =>
          /encryptionpolicy|openpgp_key_id|encryption_cert_name|auto_enable/.test(name)).length, 0);
      });
    }
    it(`${tool}: Block skipReview still wins before identity preferences or native creation`, async () => {
      const h = loadEncryptionTools({ blocked: true });
      const result = await h.run(tool);
      assert.match(result.error, /User preference blocks skipReview/);
      assert.deepEqual(h.calls.created, { params: 0, fields: 0, compose: 0, sender: 0 });
      assert.equal(h.calls.preferences.length, 0);
      assert.equal(h.calls.mime, 0);
      assertNoAttachmentEffects(h.calls.attachments);
    });
  }
  for (const tool of ["saveDraft", "replyToMessage"]) {
    it(`${tool}: draft creation bypasses unreadable identity encryption settings`, async () => {
      const h = loadEncryptionTools({ unreadable: [`${identityPrefix}encryptionpolicy`] });
      const result = await h.run(tool, { skipReview: false, saveAsDraft: true });
      assert.equal(result.success, true, result.error);
      assert.equal(h.calls.preferences.length, 0);
      assert.equal(h.calls.sends.length, 0);
      assert.equal(h.calls.drafts + h.calls.draftSaves, 1);
    });
  }
});

describe("direct send identity policy changes while the original is being classified", () => {
  for (const tool of ["replyToMessage", "forwardMessage"]) {
    for (const change of ["requires encryption", "automatic encryption enabled", "policy becomes unreadable"]) {
      it(`${tool}: refuses before native creation when ${change}`, async () => {
        const h = loadEncryptionTools({
          deferMime: true,
          prefs: { [`${identityPrefix}openpgp_key_id`]: "ABCD" },
        });
        const pending = h.run(tool);
        await Promise.resolve();
        assert.equal(h.calls.mime, 1);
        assert.deepEqual(h.calls.created, { params: 0, fields: 0, compose: 0, sender: 0 });
        assert.equal(h.calls.sends.length, 0);
        assertNoAttachmentEffects(h.calls.attachments);
        if (change === "requires encryption") h.prefs.getIntPref = () => 1;
        else if (change === "automatic encryption enabled") h.prefs.getBoolPref = () => true;
        else h.prefs.getIntPref = () => { throw new Error("Preference unavailable"); };
        h.calls.finishMime();
        assertRefused(h, await pending, 1);
      });
    }
  }
});


describe("scheduled draft attachment boundaries", () => {
  const sendAt = new Date(SEND_LATER_NOW + 3600000).toISOString();
  for (const [label, options] of [
    ["encryption policy", { prefs: { [`${identityPrefix}encryptionpolicy`]: 1 } }],
    ["safety block", { blocked: true }],
    ["disabled sender", { disabledTools: ["sendMail"] }],
    ["unavailable Send Later", { info: null }],
    ["disabled unsent sending", { info: { ...SEND_LATER_READY, sendUnsentMsgs: false } }],
  ]) {
    it(`refuses ${label} before real attachment conversion`, async () => {
      const h = loadEncryptionTools(options);
      assert.equal(typeof (await h.run("saveDraft", { sendAt })).error, "string");
      assert.equal(h.calls.drafts, 0);
      assert.equal(h.calls.created.sender, 0);
      assertNoAttachmentEffects(h.calls.attachments);
    });
  }

  it("retains S2's sensitive-path refusal for scheduled drafts", async () => {
    const h = loadEncryptionTools();
    const result = await h.run("saveDraft", { sendAt, attachments: ["/fixture/.ssh/id_rsa"] });
    assert.equal(typeof result.error, "string");
    assert.equal(h.calls.drafts, 0);
    assert.equal(h.calls.created.sender, 0);
  });

  it("saves allowed attachments through the existing draft sender", async () => {
    const h = loadEncryptionTools();
    const result = await h.run("saveDraft", { sendAt });
    assert.equal(result.success, true, result.error);
    assert.equal(result.sendAt, sendAt);
    assert.equal(h.calls.drafts, 1);
    assert.equal(h.calls.sends.length, 0);
    assert.equal(h.calls.attachments.nativeCreates, 2);
  });
});
