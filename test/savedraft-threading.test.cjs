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
const { makeDirectSendPrefs, prefTypes } = require("./helpers/direct-send-prefs.cjs");
const { SEND_LATER_NOW, SEND_LATER_READY, SEND_LATER_HEADER_PARSER, makeSendLaterClock } = require("./helpers/send-later.cjs");

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
  nsIPrefBranch: prefTypes,
  nsIMsgComposeParams: Symbol("nsIMsgComposeParams"),
  nsIMsgCompFields: Symbol("nsIMsgCompFields"),
  nsIMsgCompType: { New: 0 },
  nsIMsgCompDeliverMode: { SaveAsDraft: 4 },
};

function loadSaveDraft(options = {}) {
  const clock = makeSendLaterClock();
  const sent = [];
  const calls = { attachments: 0, signatures: 0 };
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
    Services: { prefs: makeDirectSendPrefs(options.prefs) },
    MailServices: { headerParser: options.headerParser || SEND_LATER_HEADER_PARSER },
    Date: clock.Date,
    isSkipReviewBlocked: () => options.blocked === true,
    isToolEnabled: name => !(options.disabledTools || []).includes(name),
    __tbMcpSendLaterInfo: Object.hasOwn(options, "info") ? options.info : SEND_LATER_READY,
    Cc: {
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => composeFields },
    },
    setComposeIdentity(params, from) {
      if (options.identityError) return { error: options.identityError };
      params.identity = { key: from ? "selected" : "default", signMail: true };
    },
    resolveComposeFormat: () => ({ useHtml: false, format: 1 }),
    buildBodyWithSignature(body) {
      calls.signatures++;
      return body || "";
    },
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
    snippet("TWO DIGIT PADDING"),
    snippet("DIRECT SEND ENCRYPTION GUARD"),
    snippet("DRAFT HELPERS"),
    snippet("OUTBOUND ATTACHMENT LIMITS"), snippet("CUSTOM HEADER HELPERS"), snippet("SAVE DRAFT TOOL"),
    snippet("TOOL CALL DISPATCH"),
    "this.saveDraft = saveDraft;",
  ].join("\n"), sandbox);

  return {
    saveDraft: sandbox.saveDraft, composeFields, sent, calls, clock,
    save: args => sandbox.dispatchTool("saveDraft", { to: "a@example.com", subject: "Hi", body: "body", ...args }),
  };
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

describe("saveDraft scheduling permissions and identity policy", () => {
  const future = () => new Date(SEND_LATER_NOW + 3600000).toISOString();
  const identityPrefix = "mail.identity.default.";
  const autoEnable = "mail.e2ee.auto_enable";
  const unavailableParser = { parseDecodedHeader() { throw new Error("Address parser unavailable"); } };
  const refusals = [
    ["Block skipReview", { blocked: true }, {}, /blocks skipReview/],
    ["disabled sendMail", { disabledTools: ["sendMail"] }, {}, /requires the sendMail tool/],
    ["unavailable add-on", { info: null }, {}, /not installed/],
    ["disabled draft sending", { info: { ...SEND_LATER_READY, sendDrafts: false } }, {}, /does not send scheduled drafts/],
    ["disabled scan", { info: { ...SEND_LATER_READY, checkTimePref: 0 } }, {}, /check interval/],
    ["unsent sending off", { info: { ...SEND_LATER_READY, sendUnsentMsgs: false } }, {}, /send unsent messages/],
    ["stale information", { info: { ...SEND_LATER_READY, checkedAt: SEND_LATER_NOW - 180001 } }, {}, /older than 3 minutes/],
    ["calendar rollover", {}, { sendAt: "2027-02-30T10:00:00Z" }, /valid calendar/],
    ["past date", {}, { sendAt: "2026-12-31T23:00:00Z" }, /in the past/],
    ["distant date", {}, { sendAt: "2028-02-01T00:00:00Z" }, /366 days/],
    ["missing recipients", {}, { to: "", cc: " \t", bcc: "" }, /at least one recipient/],
    ["unreadable recipients", { headerParser: unavailableParser }, {}, /Address parser unavailable/],
    ["restricted identity", { identityError: "identity belongs to a restricted account" }, {}, /restricted account/],
    ["invalid threading", {}, { inReplyTo: "bad" }, /inReplyTo/],
    ["malformed date", {}, { sendAt: "tomorrow" }, /ISO 8601/],
  ];
  for (const [label, options, args, error] of refusals) {
    it(`refuses ${label} before saving, signatures or attachments`, async () => {
      const h = loadSaveDraft(options);
      assert.match((await h.save({ sendAt: future(), ...args })).error, error);
      assert.equal(h.sent.length, 0);
      assert.deepEqual(h.calls, { attachments: 0, signatures: 0 });
    });
  }

  for (const [recipient, value, parsed] of [
    ["to", "undisclosed-recipients:;", []],
    ["cc", ",", []],
    ["bcc", "Nobody <>", [{ name: "Nobody", email: "" }]],
  ]) {
    it(`refuses a nonempty ${recipient} header with no parsed mailbox before preparing a draft`, async () => {
      let parses = 0;
      const h = loadSaveDraft({ headerParser: {
        parseDecodedHeader(header, preserveGroups) {
          parses++;
          assert.equal(header, value);
          assert.equal(preserveGroups, false);
          return parsed;
        },
      } });
      const result = await h.save({ to: "", [recipient]: value, sendAt: future() });
      assert.match(result.error, /at least one recipient/);
      assert.equal(parses, 1);
      assert.equal(h.sent.length, 0);
      assert.deepEqual(h.calls, { attachments: 0, signatures: 0 });
    });
  }

  const encrypted = [
    ["required encryption", { [`${identityPrefix}encryptionpolicy`]: 1 }],
    ["unreadable encryption", { [`${identityPrefix}encryptionpolicy`]: undefined }],
    ["wrong-type policy", { [`${identityPrefix}encryptionpolicy`]: "0" }],
    ["OpenPGP automatic encryption", { [`${identityPrefix}openpgp_key_id`]: "key", [autoEnable]: true }],
    ["S/MIME automatic encryption", { [`${identityPrefix}encryption_cert_name`]: "cert", [autoEnable]: true }],
    ["unreadable automatic encryption", { [`${identityPrefix}openpgp_key_id`]: "key", [autoEnable]: undefined }],
  ];
  for (const [label, prefs] of encrypted) {
    it(`refuses ${label} with saveDraft-specific review guidance`, async () => {
      const h = loadSaveDraft({ prefs });
      const result = await h.save({ sendAt: future() });
      assert.match(result.error, /encryption/);
      assert.match(result.error, /Omit sendAt to save a plain draft/);
      assert.match(result.error, /sendMail with skipReview: false.*review window/);
      assert.equal(h.sent.length, 0);
      assert.deepEqual(h.calls, { attachments: 0, signatures: 0 });
    });
  }

  it("checks the explicitly selected identity, including inherited defaults", async () => {
    const h = loadSaveDraft({ prefs: { "mail.identity.selected.encryptionpolicy": 2 } });
    assert.match((await h.save({ from: "selected@example.test", sendAt: future() })).error, /requires encryption/);
    assert.equal(h.sent.length, 0);
    assert.equal(h.calls.attachments, 0);
    const inherited = loadSaveDraft({ prefs: { [`${identityPrefix}encryptionpolicy`]: 1 } });
    assert.match((await inherited.save({ from: "selected@example.test", sendAt: future() })).error, /requires encryption/);
    assert.equal(inherited.sent.length, 0);
  });

  for (const recipient of ["to", "cc", "bcc"]) {
    it(`schedules with only ${recipient}, preserving headers and returning sendAt through dispatch`, async () => {
      const h = loadSaveDraft();
      const sendAt = future();
      const result = await h.save({ to: "", [recipient]: "recipient@example.test", sendAt, inReplyTo: PARENT });
      assert.equal(result.success, true, result.error);
      assert.equal(result.sendAt, sendAt);
      assert.equal(h.sent.length, 1);
      assert.equal(h.sent[0].deliverMode, Ci.nsIMsgCompDeliverMode.SaveAsDraft);
      assert.deepEqual([...h.composeFields.headers.keys()], ["In-Reply-To", "X-Send-Later-Uuid", "X-Send-Later-At"]);
      assert.equal(h.composeFields.headers.get("X-Send-Later-Uuid"), "draft-test-instance");
      assert.ok(Math.abs(Date.parse(h.composeFields.headers.get("X-Send-Later-At")) - Date.parse(sendAt)) < 1000);
      assert.equal(h.composeFields.references, PARENT);
      assert.equal(result.uuid, undefined);
    });
  }

  for (const sendAt of [undefined, null]) {
    it(`keeps a plain draft unchanged with sendAt=${sendAt}`, async () => {
      const h = loadSaveDraft({ blocked: true, disabledTools: ["sendMail"], info: null,
        prefs: { [`${identityPrefix}encryptionpolicy`]: undefined }, headerParser: unavailableParser });
      const result = await h.save({ to: "", sendAt });
      assert.equal(result.success, true, result.error);
      assert.equal(result.sendAt, undefined);
      assert.equal(h.composeFields.headers.size, 0);
      assert.equal(h.sent.length, 1);
    });
  }
});

describe("saveDraft custom headers", () => {
  it("adds caller X- headers to the draft", async () => {
    const harness = loadSaveDraft();
    const result = await harness.save({ headers: [{ name: "X-AI-Assisted", value: "Drafted with AI" }] });

    assert.equal(result.success, true);
    assert.equal(harness.composeFields.headers.get("X-AI-Assisted"), "Drafted with AI");
  });

  it("refuses a Send Later header so scheduling stays behind sendAt", async () => {
    const harness = loadSaveDraft();
    const result = await harness.save({ headers: [{ name: "X-Send-Later-At", value: "Mon, 1 Feb 2027 10:00:00 +0000" }] });

    assert.match(result.error, /reserved/);
    assert.equal(harness.sent.length, 0);
    assert.equal(harness.composeFields.headers.size, 0);
  });
});
