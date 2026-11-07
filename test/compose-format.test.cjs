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
const tools = [["sendMail", types.New], ["replyToMessage", types.Reply], ["forwardMessage", types.ForwardInline]];
const body = "Hello <b>there</b>\n&bye";
const hdrId = "original@example.test";

// Searched test/ for snippet/extraction helpers and compose-format harnesses:
// reuse compose-signature/direct-send-outcome's local marker+VM pattern; no
// shared harness integrates the real resolver, body builders and both senders.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function loadComposeFormat(composeHtml) {
  const calls = { sends: [], draftSends: [], reviews: [], draftSaves: 0 };
  const identity = { key: "sender", email: "sender@example.test", composeHtml, htmlSigText: "Sender signature" };
  const folder = { server: {}, getUriForMsg: () => "message-uri" };
  const hdr = {
    messageId: hdrId, author: "Author <author@example.test>",
    subject: "Original subject", recipients: "sender@example.test", date: 0,
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
      importESModule: () => ({ MsgHdrToMimeMessage: (msg, _listener, callback) => callback(msg, { contentType: "text/plain", body: "Original <text>\nSecond line" }) }),
    },
    Components: { isSuccessCode: status => status === 0 },
    Services: { prefs: makeDirectSendPrefs() },
    MailServices: { headerParser: composeHeaderParser, accounts: { accounts: [{ key: "account", identities: [identity] }] } },
    Cc: {
      "@mozilla.org/timer;1": { createInstance: () => ({ initWithCallback() {}, cancel() {} }) },
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({ setHeader() {}, addAttachment() {} }) },
      "@mozilla.org/messengercompose/compose;1": { createInstance() {
        // Reuse #253's fake native compose. As in the verification probe,
        // initialize chooses MIME solely from the received format + identity,
        // never by inspecting body markup or the caller's isHtml argument.
        return {
          initialize(params, win, editor) {
            assert.equal(win, null);
            assert.equal(editor, null);
            const useHtml = params.format === formats.HTML ||
              (params.format === formats.Default && params.identity.composeHtml) ||
              (params.format === formats.OppositeOfDefault && !params.identity.composeHtml);
            this.sent = { body: params.composeFields.body, mimeType: useHtml ? "text/html" : "text/plain", format: params.format };
          },
          addMsgSendListener(listener) { this.listener = listener; },
          removeMsgSendListener(listener) { if (this.listener === listener) this.listener = null; },
          RegisterStateListener() {},
          UnregisterStateListener() {},
          sendMsg(mode, sender, accountKey) {
            assert.equal(mode, 0);
            assert.equal(sender, identity);
            assert.equal(accountKey, "account");
            calls.sends.push(this.sent);
            this.listener.onStopSending("sent@example.test", 0);
            return Promise.resolve();
          },
        };
      } },
      "@mozilla.org/messengercompose/send;1": { createInstance: () => ({
        createAndSendMessage(...args) {
          calls.draftSends.push({ body: args[9], mimeType: args[8], mode: args[6], msgToReplace: args[7] });
          args[12].onStopCopy(0);
          return Promise.resolve();
        },
      }) },
      "@mozilla.org/messengercompose;1": { getService: () => ({
        OpenComposeWindowWithParams(_parent, params) {
          calls.reviews.push({ format: params.format, body: params.composeFields.body });
        },
      }) },
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
    getDirectQuoteContent: () => ({ body: "Original <text>\nSecond line" }),
    getReplyAllCcRecipients: () => "",
    rawMessageHasForwardableParts: () => false,
    async openComposeWindowWithCustomizations(params, _uri, _type, _identity, intro, isHtml, _to, _cc, _bcc, _descs, afterInsert) {
      calls.reviews.push({ format: params.format, body: sandbox.formatBodyFragmentHtml(intro, isHtml) });
      return afterInsert ? afterInsert({}) : { success: true };
    },
    saveComposeWindowAsDraft() {
      calls.draftSaves++;
      return Promise.resolve({ success: true });
    },
  };
  const escapeStart = source.indexOf("function escapeHtml(s) {");
  const escapeEnd = source.indexOf("// BEGIN MESSAGE TEXT CONVERSION", escapeStart);
  assert.ok(escapeStart >= 0 && escapeEnd > escapeStart, "escapeHtml boundaries missing");
  vm.createContext(sandbox);
  vm.runInContext([
    source.slice(escapeStart, escapeEnd),
    snippet("MESSAGE TEXT CONVERSION"), snippet("COMPOSE HTML FRAGMENT"),
    snippet("COMPOSE SIGNATURE HELPERS"), snippet("COMPOSE FORMAT HELPER"),
    snippet("DIRECT SEND ENCRYPTION GUARD"), snippet("DIRECT SEND ORIGINAL GUARD"),
    snippet("DIRECT SEND HELPER"), snippet("COMPOSE SEND HELPER"),
    snippet("OUTBOUND MAIL TOOLS"), snippet("TOOL DISPATCH"),
  ].join("\n"), sandbox);
  return { api: sandbox, calls, identity };
}

describe("direct compose format matrix (24 cases)", () => {
  for (const composeHtml of [true, false]) {
    for (const isHtml of [undefined, null, true, false]) {
      for (const [tool, compType] of tools) {
        it(`${tool}: identity.composeHtml=${composeHtml}, isHtml=${isHtml}`, async () => {
          const { api, calls, identity } = loadComposeFormat(composeHtml);
          const expectedHtml = typeof isHtml === "boolean" ? isHtml : composeHtml;
          const resolved = api.resolveComposeFormat(identity, isHtml, compType);
          assert.equal(resolved.useHtml, expectedHtml);
          if (isHtml === null || isHtml === undefined) {
            // This is part of the matrix contract, even though the direct
            // sender fixes MIME to the already-resolved body mode. Reverting
            // null handling must still fail instead of masking review bugs.
            assert.equal(resolved.format, formats.Default, "null and omitted both use the identity default");
          }

          const result = await api.dispatchTool(tool, {
            to: "recipient@example.test", subject: "Subject", body, isHtml,
            messageId: hdrId, folderPath: "folder", skipReview: true,
          });
          assert.equal(result.success, true, result.error);
          assert.equal(calls.sends.length, 1);
          assert.equal(calls.draftSends.length, 0);
          assert.equal(calls.reviews.length, 0);
          const sent = calls.sends[0];
          assert.equal(sent.mimeType, expectedHtml ? "text/html" : "text/plain");
          assert.equal(/^<html>/.test(sent.body), expectedHtml, "body envelope must match the outgoing MIME type");
          assert.equal(sent.format, expectedHtml ? formats.HTML : formats.PlainText, "windowless initialization uses the resolved body mode");
          if (expectedHtml) {
            assert.ok(sent.body.includes(isHtml === true ? "<b>there</b>" : "&lt;b&gt;there&lt;/b&gt;"));
          } else {
            assert.ok(sent.body.startsWith(body));
          }
          assert.equal(sent.body.includes("Sender signature"), true, "direct sends include the identity signature by default");
        });
      }
    }
  }
});

describe("review compose formats", () => {
  for (const composeHtml of [true, false]) {
    for (const isHtml of [undefined, null, true, false]) {
      for (const [tool, compType] of tools) {
        it(`${tool}: identity.composeHtml=${composeHtml}, isHtml=${isHtml}`, async () => {
          const { api, calls } = loadComposeFormat(composeHtml);
          const result = await api.dispatchTool(tool, {
            to: "recipient@example.test", subject: "Subject", body, isHtml,
            messageId: hdrId, folderPath: "folder",
          });
          assert.equal(result.success, true, result.error);
          assert.equal(calls.sends.length, 0);
          assert.equal(calls.draftSends.length, 0);
          assert.equal(calls.reviews.length, 1);
          let expectedFormat = formats.Default;
          if (typeof isHtml === "boolean") {
            expectedFormat = compType === types.ForwardInline
              ? (isHtml === composeHtml ? formats.Default : formats.OppositeOfDefault)
              : (isHtml ? formats.HTML : formats.PlainText);
          }
          const review = calls.reviews[0];
          assert.equal(review.format, expectedFormat);
          if (tool === "sendMail") {
            const expectedHtml = typeof isHtml === "boolean" ? isHtml : composeHtml;
            assert.equal(/^<html>/.test(review.body), expectedHtml);
            assert.ok(!review.body.includes("Sender signature"), "Thunderbird remains responsible for review signatures");
          }
          if (isHtml === null) {
            await api.dispatchTool(tool, {
              to: "recipient@example.test", subject: "Subject", body,
              messageId: hdrId, folderPath: "folder",
            });
            assert.deepEqual(calls.reviews[0], calls.reviews[1], "null matches omitted format and caller-body insertion");
          }
        });
      }
    }
  }
});

describe("draft compose null equivalence", () => {
  for (const composeHtml of [true, false]) {
    it(`saveDraft uses the same body and MIME for null/omitted, identity.composeHtml=${composeHtml}`, async () => {
      const { api, calls } = loadComposeFormat(composeHtml);
      for (const isHtml of [undefined, null]) {
        const result = await api.dispatchTool("saveDraft", {
          to: "recipient@example.test", subject: "Subject", body, isHtml,
        });
        assert.equal(result.success, true, result.error);
      }
      assert.equal(calls.draftSends.length, 2);
      assert.deepEqual(calls.draftSends[0], calls.draftSends[1]);
      assert.equal(calls.draftSends[0].mimeType, composeHtml ? "text/html" : "text/plain");
      assert.equal(/^<html>/.test(calls.draftSends[0].body), composeHtml);
      assert.equal(calls.draftSends[0].mode, 4);
      assert.equal(calls.sends.length, 0);
      assert.equal(calls.reviews.length, 0);
    });

    it(`reply saveAsDraft uses the same native format for null/omitted, identity.composeHtml=${composeHtml}`, async () => {
      const { api, calls } = loadComposeFormat(composeHtml);
      for (const isHtml of [undefined, null]) {
        const result = await api.dispatchTool("replyToMessage", {
          messageId: hdrId, folderPath: "folder", body, isHtml, saveAsDraft: true,
        });
        assert.equal(result.success, true, result.error);
        assert.equal(result.message, "Reply saved as draft");
      }
      assert.equal(calls.reviews.length, 2);
      assert.equal(calls.reviews[0].format, formats.Default);
      assert.deepEqual(calls.reviews[0], calls.reviews[1]);
      assert.equal(calls.draftSaves, 2);
      assert.equal(calls.sends.length, 0);
      assert.equal(calls.draftSends.length, 0);
    });
  }
});
