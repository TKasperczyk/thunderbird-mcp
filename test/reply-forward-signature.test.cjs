"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { makeDirectSendPrefs, prefTypes, clearOriginalRawFixture, composeHeaderParser } = require("./helpers/direct-send-prefs.cjs");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
const formats = { Default: 0, HTML: 1, PlainText: 2, OppositeOfDefault: 3 };
const types = { New: 0, Reply: 1, ReplyAll: 2, ForwardAsAttachment: 3, ForwardInline: 4, Draft: 9 };
const hdrId = "original@example.test";
const originalBody = "Original line\nSecond line";
const signature = "Sender signature";

function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function loadCompose(options = {}) {
  const calls = { sends: [], reviews: [], sizePlans: [] };
  const identity = {
    key: "sender", email: "sender@example.test", composeHtml: options.composeHtml ?? false,
    htmlSigText: options.hasSignature === false ? "" : signature,
    sig_bottom: options.sigBottom === true,
  };
  const folder = { server: {}, getUriForMsg: () => "message-uri" };
  const hdr = {
    messageId: hdrId, author: "Author <author@example.test>", mime2DecodedAuthor: "Author <author@example.test>",
    subject: "Original subject", mime2DecodedSubject: "Original subject", recipients: "sender@example.test",
    mime2DecodedRecipients: "sender@example.test", date: 0,
  };
  const sandbox = {
    ...clearOriginalRawFixture,
    Ci: {
      nsIPrefBranch: prefTypes,
      nsIMsgCompType: types, nsIMsgCompFormat: formats,
      nsIMsgCompDeliverMode: { Now: 0, SaveAsDraft: 4 }, nsITimer: { TYPE_ONE_SHOT: 0 },
    },
    ChromeUtils: {
      generateQI: () => () => {},
      importESModule: () => ({ MsgHdrToMimeMessage: (_msg, _listener, callback) => callback(_msg, { contentType: "text/plain", body: originalBody }) }),
    },
    Components: { isSuccessCode: status => status === 0 },
    Services: { prefs: makeDirectSendPrefs() },
    MailServices: { headerParser: composeHeaderParser, accounts: { accounts: [{ key: "account", identities: [identity] }] } },
    Cc: {
      "@mozilla.org/timer;1": { createInstance: () => ({ initWithCallback() {}, cancel() {} }) },
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({ setHeader() {}, addAttachment() {} }) },
      "@mozilla.org/messengercompose/compose;1": { createInstance() {
        return {
          initialize(params, win, editor) {
            assert.equal(win, null);
            assert.equal(editor, null);
            const useHtml = params.format === formats.HTML ||
              (params.format === formats.Default && params.identity.composeHtml) ||
              (params.format === formats.OppositeOfDefault && !params.identity.composeHtml);
            this.sent = { body: params.composeFields.body, mimeType: useHtml ? "text/html" : "text/plain" };
          },
          addMsgSendListener(listener) { this.listener = listener; },
          removeMsgSendListener(listener) { if (this.listener === listener) this.listener = null; },
          RegisterStateListener() {}, UnregisterStateListener() {},
          sendMsg(mode) {
            assert.equal(mode, 0);
            calls.sends.push(this.sent);
            this.listener.onStopSending("sent@example.test", 0);
            return Promise.resolve();
          },
        };
      } },
      "@mozilla.org/messengercompose/send;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose;1": { getService: () => ({ OpenComposeWindowWithParams(_parent, params) { calls.reviews.push(params.composeFields.body); } }) },
    },
    setComposeIdentity(params) { params.identity = identity; },
    isSkipReviewBlocked: () => false,
    isToolEnabled: () => true,
    materializeAttachmentDescs: descs => descs,
    filePathsToAttachDescs: () => ({ descs: [] }),
    descsToMsgAttachments: descs => descs,
    findMessage: () => ({ msgHdr: hdr, folder }),
    PREF_ALLOW_ENCRYPTED_MESSAGES: "encrypted",
    isPrivacyOptInEnabled: () => false,
    DIRECT_SEND_REVIEW_INSTRUCTION: "Use skipReview: false to review.",
    DIRECT_SEND_ORIGINAL_ERROR: "Direct original unavailable",
    DIRECT_SEND_RAW_ERROR: "Direct original unavailable",
    getDirectSendIdentityError: () => null,
    getDirectSendOriginal: async () => ({
      mime: { allUserAttachments: [] }, quoted: { body: originalBody }, rawBytes: 0, deadline: Date.now() + 60000,
    }),
    getReplyAllCcRecipients: () => "",
    rawMessageHasForwardableParts: () => false,
    assessComposeSize(plan) {
      calls.sizePlans.push(plan);
      if (options.rejectIfSigned && plan.body.includes(signature)) return { warning: "encoded size exceeds threshold" };
      return { warning: null };
    },
    async openComposeWindowWithCustomizations(params, _uri, _type, _identity, intro, isHtml, _to, _cc, _bcc, _descs, afterInsert) {
      calls.reviews.push(intro);
      return afterInsert ? afterInsert({}) : { success: true };
    },
    saveComposeWindowAsDraft: async () => ({ success: true }),
  };
  const escapeStart = source.indexOf("function escapeHtml(s) {");
  const escapeEnd = source.indexOf("// BEGIN MESSAGE TEXT CONVERSION", escapeStart);
  assert.ok(escapeStart >= 0 && escapeEnd > escapeStart, "escapeHtml boundaries missing");
  vm.createContext(sandbox);
  vm.runInContext([
    source.slice(escapeStart, escapeEnd),
    snippet("MESSAGE TEXT CONVERSION"), snippet("COMPOSE HTML FRAGMENT"),
    snippet("COMPOSE SIGNATURE HELPERS"), snippet("COMPOSE FORMAT HELPER"),
    snippet("COMPOSE SEND HELPER"),
    snippet("DIRECT SEND ENCRYPTION GUARD"),
    snippet("DIRECT SEND HELPER"),
    snippet("OUTBOUND MAIL TOOLS"), snippet("TOOL DISPATCH"),
  ].join("\n"), sandbox);
  sandbox.assessComposeSize = plan => {
    calls.sizePlans.push(plan);
    return options.rejectIfSigned && plan.body.includes(signature)
      ? { warning: "encoded size exceeds threshold" }
      : { warning: null };
  };
  return { api: sandbox, calls, identity };
}

function sentBody(tool, calls) {
  assert.equal(calls.sends.length, 1);
  assert.equal(calls.reviews.length, 0);
  return calls.sends[0].body;
}

for (const [tool, subject] of [["replyToMessage", "Reply"], ["forwardMessage", "Forward"]]) {
  describe(`${subject} identity signature on direct sends`, () => {
    for (const isHtml of [false, true]) {
      it(`includes the signature by default and honors includeSignature=false (${isHtml ? "HTML" : "plain"})`, async () => {
        for (const includeSignature of [undefined, false]) {
          const { api, calls } = loadCompose({ composeHtml: isHtml });
          const args = { messageId: hdrId, folderPath: "folder", body: "Reply text", isHtml, skipReview: true, includeSignature };
          if (tool === "forwardMessage") args.to = "recipient@example.test";
          const result = await api.dispatchTool(tool, args);
          assert.equal(Boolean(result.success), true, result.error);
          if (includeSignature !== false) assert.ok(sentBody(tool, calls).includes(signature));
          else assert.ok(!sentBody(tool, calls).includes(signature));
        }
      });

      it(`keeps the body unchanged when the identity has no signature (${isHtml ? "HTML" : "plain"})`, async () => {
        const { api, calls } = loadCompose({ composeHtml: isHtml, hasSignature: false });
        const args = { messageId: hdrId, folderPath: "folder", body: "Reply text", isHtml, skipReview: true };
        if (tool === "forwardMessage") args.to = "recipient@example.test";
        const result = await api.dispatchTool(tool, args);
        assert.equal(result.success, true, result.error);
        assert.ok(!sentBody(tool, calls).includes(signature));
      });

      for (const sigBottom of [false, true]) {
        it(`places signature ${sigBottom ? "after" : "before"} the quoted block when sig_bottom=${sigBottom} (${isHtml ? "HTML" : "plain"})`, async () => {
          const { api, calls } = loadCompose({ composeHtml: isHtml, sigBottom: sigBottom });
          const args = { messageId: hdrId, folderPath: "folder", body: "Reply text", isHtml, skipReview: true };
          if (tool === "forwardMessage") args.to = "recipient@example.test";
          const result = await api.dispatchTool(tool, args);
          assert.equal(result.success, true, result.error);
          const body = sentBody(tool, calls);
          const signatureIndex = body.indexOf(signature);
          assert.ok(signatureIndex >= 0, "identity signature should be present");
          const blockIndex = body.indexOf(isHtml ? "blockquote" : "-------- Forwarded Message --------");
          if (tool === "replyToMessage") {
            const quoteIndex = body.indexOf(isHtml ? "blockquote" : "On ");
            assert.ok(quoteIndex >= 0);
            assert.equal(signatureIndex < quoteIndex, !sigBottom);
          } else {
            assert.ok(blockIndex >= 0);
            assert.equal(signatureIndex < blockIndex, !sigBottom);
          }
        });
      }
    }

    it("uses the signed body in the size plan and refuses a send that exceeds the limit only with the signature", async () => {
      const { api, calls } = loadCompose({ composeHtml: false, rejectIfSigned: true });
      const args = { messageId: hdrId, folderPath: "folder", body: "Reply text", isHtml: false, skipReview: true };
      if (tool === "forwardMessage") args.to = "recipient@example.test";
      const result = await api.dispatchTool(tool, args);
      assert.ok(calls.sizePlans.some(plan => plan.body.includes(signature)), "size plan must include the body that will be sent");
      assert.match(result.error, /encoded size exceeds threshold/);
      assert.equal(calls.sends.length, 0);
    });

    it("adds the signature exactly once", async () => {
      const { api, calls } = loadCompose({ composeHtml: false });
      const args = { messageId: hdrId, folderPath: "folder", body: "Reply text", isHtml: false, skipReview: true };
      if (tool === "forwardMessage") args.to = "recipient@example.test";
      const result = await api.dispatchTool(tool, args);
      assert.equal(result.success, true, result.error);
      assert.equal(sentBody(tool, calls).split(signature).length - 1, 1);
    });
  });
}
