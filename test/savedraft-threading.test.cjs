"use strict";

/**
 * saveDraft writes into the Drafts folder without a compose window, so it
 * composes as nsIMsgCompType.New and has no originalMsgURI for Thunderbird to
 * derive threading from. These tests pin the In-Reply-To / References headers
 * it sets from the caller's inReplyTo/references arguments, and that a draft
 * without them stays an ordinary new message.
 */

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const Ci = {
  nsIMsgComposeParams: Symbol("nsIMsgComposeParams"),
  nsIMsgCompFields: Symbol("nsIMsgCompFields"),
  nsIMsgCompType: { New: 0 },
  nsIMsgCompDeliverMode: { SaveAsDraft: 4 },
};

function loadSaveDraft() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "function saveDraft(to, subject,";
  const endMarker = "function replyToMessage(";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start);
  assert.ok(start >= 0, "saveDraft start marker missing");
  assert.ok(end > start, "saveDraft end marker missing");

  const sent = [];
  const composeFields = {
    to: "",
    cc: "",
    bcc: "",
    subject: "",
    body: "",
    references: undefined,
    headers: new Map(),
    setHeader(name, value) {
      this.headers.set(name, value);
    },
  };

  const sandbox = {
    Ci,
    Cc: {
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => composeFields },
    },
    setComposeIdentity: () => null,
    resolveComposeFormat: () => ({ useHtml: false, format: 1 }),
    formatBodyHtml: (body) => body || "",
    filePathsToAttachDescs: () => ({ descs: [], failed: [] }),
    sendMessageDirectly: (fields, identity, descs, _unused, compType, deliverMode) => {
      sent.push({ fields, compType, deliverMode });
      return Promise.resolve({ success: true });
    },
    console,
  };

  vm.createContext(sandbox);
  vm.runInContext(`${source.slice(start, end)}\nthis.saveDraft = saveDraft;`, sandbox);

  return { saveDraft: sandbox.saveDraft, composeFields, sent };
}

describe("saveDraft threading", () => {
  it("saves as a draft, never sends", async () => {
    const { saveDraft, sent } = loadSaveDraft();

    await saveDraft("a@example.com", "Hi", "body");

    assert.equal(sent.length, 1);
    assert.equal(sent[0].deliverMode, Ci.nsIMsgCompDeliverMode.SaveAsDraft);
    assert.equal(sent[0].compType, Ci.nsIMsgCompType.New);
  });

  it("sets no threading headers when inReplyTo is omitted", async () => {
    const { saveDraft, composeFields } = loadSaveDraft();

    await saveDraft("a@example.com", "Hi", "body");

    assert.equal(composeFields.headers.has("In-Reply-To"), false);
    assert.equal(composeFields.references, undefined);
  });

  it("brackets a bare Message-ID and mirrors it into References", async () => {
    const { saveDraft, composeFields } = loadSaveDraft();

    await saveDraft("a@example.com", "Re: Hi", "body", null, null, false, null, null,
      "orig123@mail.example.com");

    assert.equal(composeFields.headers.get("In-Reply-To"), "<orig123@mail.example.com>");
    assert.equal(composeFields.references, "<orig123@mail.example.com>");
  });

  it("keeps an already bracketed Message-ID as it is", async () => {
    const { saveDraft, composeFields } = loadSaveDraft();

    await saveDraft("a@example.com", "Re: Hi", "body", null, null, false, null, null,
      "  <orig123@mail.example.com>  ");

    assert.equal(composeFields.headers.get("In-Reply-To"), "<orig123@mail.example.com>");
  });

  it("prefers an explicit References chain over the mirrored default", async () => {
    const { saveDraft, composeFields } = loadSaveDraft();

    await saveDraft("a@example.com", "Re: Hi", "body", null, null, false, null, null,
      "<c@example.com>", "<a@example.com> <b@example.com> <c@example.com>");

    assert.equal(composeFields.headers.get("In-Reply-To"), "<c@example.com>");
    assert.equal(composeFields.references, "<a@example.com> <b@example.com> <c@example.com>");
  });
});
