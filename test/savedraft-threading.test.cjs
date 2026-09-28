"use strict";

/**
 * saveDraft writes into the Drafts folder without a compose window, so it
 * composes as nsIMsgCompType.New and has no originalMsgURI for Thunderbird to
 * derive threading from. These tests pin the In-Reply-To / References headers
 * it sets from the caller's inReplyTo/references arguments, rejects malformed
 * headers before saving, and leaves ordinary new messages unthreaded.
 */

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");

function snippet(name) {
  const startMarker = `// BEGIN ${name}`;
  const endMarker = `// END ${name}`;
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

const Ci = {
  nsIMsgComposeParams: Symbol("nsIMsgComposeParams"),
  nsIMsgCompFields: Symbol("nsIMsgCompFields"),
  nsIMsgCompType: { New: 0 },
  nsIMsgCompDeliverMode: { SaveAsDraft: 4 },
};

function loadSaveDraft() {
  const sent = [];
  const calls = { attachments: 0 };
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
    setComposeIdentity(params) { params.identity = {}; },
    resolveComposeFormat: () => ({ useHtml: false, format: 1 }),
    buildBodyWithSignature: body => body || "",
    getAccessibleFolder: uri => ({ folder: { URI: uri } }),
    filePathsToAttachDescs() {
      calls.attachments++;
      return { descs: [], failed: [] };
    },
    sendMessageDirectly: (fields, identity, descs, _unused, compType, deliverMode) => {
      sent.push({ fields, compType, deliverMode });
      return Promise.resolve({ success: true });
    },
    console,
  };

  vm.createContext(sandbox);
  vm.runInContext([
    snippet("DRAFT HELPERS"),
    snippet("SAVE DRAFT TOOL"),
    "this.saveDraft = saveDraft;",
  ].join("\n"), sandbox);

  return { saveDraft: sandbox.saveDraft, composeFields, sent, calls };
}

function saveThreaded(harness, inReplyTo, references) {
  return harness.saveDraft("a@example.com", "Re: Hi", "body", undefined, undefined,
    false, undefined, undefined, inReplyTo, references);
}

const PARENT = "<orig123@mail.example.com>";
const messageIdOfLength = length => `<${"a".repeat(length - 4)}@x>`;

describe("saveDraft threading", () => {
  it("saves as a draft, never sends", async () => {
    const { saveDraft, sent } = loadSaveDraft();

    const result = await saveDraft("a@example.com", "Hi", "body");

    assert.equal(result.success, true);
    assert.equal(sent.length, 1);
    assert.equal(sent[0].deliverMode, Ci.nsIMsgCompDeliverMode.SaveAsDraft);
    assert.equal(sent[0].compType, Ci.nsIMsgCompType.New);
  });

  it("sets no threading headers when inReplyTo is omitted", async () => {
    const { saveDraft, composeFields } = loadSaveDraft();

    const result = await saveDraft("a@example.com", "Hi", "body");

    assert.equal(result.success, true);
    assert.equal(composeFields.headers.has("In-Reply-To"), false);
    assert.equal(composeFields.references, undefined);
  });

  it("preserves a bracketed Message-ID and mirrors it into References", async () => {
    const harness = loadSaveDraft();
    const result = await saveThreaded(harness, PARENT);

    assert.equal(result.success, true);
    assert.equal(harness.composeFields.headers.get("In-Reply-To"), PARENT);
    assert.equal(harness.composeFields.references, PARENT);
  });

  it("prefers an explicit References chain over the mirrored default", async () => {
    const harness = loadSaveDraft();
    const references = "<a@example.com> <b@example.com> <c@example.com>";
    const result = await saveThreaded(harness, "<c@example.com>", references);

    assert.equal(result.success, true);
    assert.equal(harness.composeFields.headers.get("In-Reply-To"), "<c@example.com>");
    assert.equal(harness.composeFields.references, references);
  });

  it("accepts References independently of In-Reply-To", async () => {
    const harness = loadSaveDraft();
    const result = await saveThreaded(harness, undefined, PARENT);

    assert.equal(result.success, true);
    assert.equal(harness.composeFields.headers.has("In-Reply-To"), false);
    assert.equal(harness.composeFields.references, PARENT);
  });

  const invalidIds = [
    ["missing brackets", "orig123@mail.example.com"],
    ["missing closing bracket", "<orig123@mail.example.com"],
    ["missing opening bracket", "orig123@mail.example.com>"],
    ["nested brackets", "<<orig123@mail.example.com>>"],
    ["empty token", "<>"],
    ["missing domain separator", "<orig123>"],
    ["empty local part", "<@example.com>"],
    ["empty domain", "<orig123@>"],
    ["multiple domain separators", "<orig@other@example.com>"],
    ["leading space", ` ${PARENT}`],
    ["trailing space", `${PARENT} `],
    ["inner space", "<orig 123@mail.example.com>"],
    ["inner tab", "<orig\t123@mail.example.com>"],
    ["inner CR", "<orig\r123@mail.example.com>"],
    ["inner LF", "<orig\n123@mail.example.com>"],
    ["NUL", "<orig\u0000123@mail.example.com>"],
    ["C0 control", "<orig\u0001123@mail.example.com>"],
    ["DEL control", "<orig\u007f123@mail.example.com>"],
    ["C1 control", "<orig\u0085123@mail.example.com>"],
    ["Unicode whitespace", "<orig\u00a0123@mail.example.com>"],
    ["header injection", `${PARENT}\r\nBcc: other@example.com`],
    ["extra text", `${PARENT} unexpected`],
    ["empty string", ""],
    ["null", null],
    ["number", 123],
    ["boolean", false],
    ["array", [PARENT]],
    ["object", { id: PARENT }],
    ["overlong token", messageIdOfLength(999)],
  ];

  for (const [label, value] of invalidIds) {
    for (const field of ["inReplyTo", "references"]) {
      it(`rejects ${label} in ${field} without saving or preparing attachments`, async () => {
        const harness = loadSaveDraft();
        const result = field === "inReplyTo"
          ? await saveThreaded(harness, value)
          : await saveThreaded(harness, undefined, value);

        assert.equal(typeof result.error, "string");
        assert.match(result.error, new RegExp(field, "i"));
        assert.equal(harness.sent.length, 0);
        assert.equal(harness.calls.attachments, 0);
      });
    }
  }

  it("rejects multiple In-Reply-To IDs", async () => {
    const harness = loadSaveDraft();
    const result = await saveThreaded(harness, `${PARENT} <other@example.com>`);

    assert.match(result.error, /inReplyTo/i);
    assert.equal(harness.sent.length, 0);
  });

  for (const separator of ["", "  ", "\t", "\r", "\n", "\u00a0"]) {
    it(`rejects References separated by ${JSON.stringify(separator)}`, async () => {
      const harness = loadSaveDraft();
      const result = await saveThreaded(harness, PARENT, `${PARENT}${separator}<other@example.com>`);

      assert.match(result.error, /references/i);
      assert.equal(harness.sent.length, 0);
    });
  }

  it("accepts a token at the 998-character limit", async () => {
    const harness = loadSaveDraft();
    const id = messageIdOfLength(998);
    const result = await saveThreaded(harness, id);

    assert.equal(result.success, true);
    assert.equal(harness.composeFields.headers.get("In-Reply-To"), id);
  });

  it("accepts 100 References IDs and rejects a 101st", async () => {
    const references = Array.from({ length: 100 }, (_, i) => `<id${i}@example.com>`).join(" ");
    const valid = loadSaveDraft();
    assert.equal((await saveThreaded(valid, PARENT, references)).success, true);
    assert.equal(valid.composeFields.references, references);

    const invalid = loadSaveDraft();
    assert.match((await saveThreaded(invalid, PARENT, `${references} <extra@example.com>`)).error, /references/i);
    assert.equal(invalid.sent.length, 0);
  });

  it("accepts 16,384 header characters and rejects a longer References header", async () => {
    const prefix = Array(16).fill(messageIdOfLength(998));
    const references = [...prefix, messageIdOfLength(400)].join(" ");
    assert.equal(references.length, 16384);
    const valid = loadSaveDraft();
    assert.equal((await saveThreaded(valid, PARENT, references)).success, true);
    assert.equal(valid.composeFields.references, references);

    const invalid = loadSaveDraft();
    const tooLong = [...prefix, messageIdOfLength(401)].join(" ");
    assert.match((await saveThreaded(invalid, PARENT, tooLong)).error, /references/i);
    assert.equal(invalid.sent.length, 0);
  });
});
