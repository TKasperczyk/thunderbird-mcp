"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
const MAX_SIGNATURE_BYTES = 1024 * 1024;
const DRAFTS_URI = "imap://user@example.test/Drafts";

// Reuse the production-marker VM pattern from privacy-messages.test.cjs and
// read-message-stream.test.cjs; Thunderbird services are the only substitutes.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function loadSignatureHelpers(options = {}) {
  const bytes = Buffer.from(options.bytes ?? "signature", "utf8").toString("latin1");
  const calls = { opens: 0, closes: 0, reads: [], warnings: [], conversions: [] };
  let cursor = 0;
  const file = {
    fileSize: Object.hasOwn(options, "fileSize") ? options.fileSize : bytes.length,
    leafName: options.leafName || "signature.txt",
    exists: () => true,
    isFile: () => true,
  };
  const stream = {
    init(openedFile) {
      assert.equal(openedFile, file);
      calls.opens++;
      if (options.initError) throw options.initError;
    },
    available() {
      if (options.streamError) throw options.streamError;
      return bytes.length - cursor;
    },
    read(count) {
      calls.reads.push(count);
      const chunk = bytes.slice(cursor, cursor + count);
      cursor += chunk.length;
      return chunk;
    },
    close() {
      calls.closes++;
      if (options.closeError) throw options.closeError;
    },
  };
  const sandbox = {
    TextDecoder,
    Uint8Array,
    Cr: { NS_BASE_STREAM_CLOSED: "NS_BASE_STREAM_CLOSED" },
    console: { warn: (...args) => calls.warnings.push(args), error() {} },
    NetUtil: { readInputStreamToString: (input, count) => input.read(count) },
    Ci: {
      nsIDocumentEncoder: { OutputFormatted: 1, OutputLFLineBreak: 2 },
      nsIMsgCompType: { New: 0, Draft: 9 },
      nsIMsgCompDeliverMode: { Now: 0, SaveAsDraft: 4 },
      nsMsgFolderFlags: { Drafts: 0x400 },
    },
    Cc: {
      "@mozilla.org/network/file-input-stream;1": { createInstance: () => stream },
      "@mozilla.org/parserutils;1": { getService: () => ({
        convertToPlainText(...args) {
          calls.conversions.push(args);
          return options.plainText ?? "Jane\n";
        },
      }) },
    },
  };
  if (options.DOMParser) sandbox.DOMParser = options.DOMParser;
  const escapeStart = source.indexOf("function escapeHtml(s) {");
  const escapeEnd = source.indexOf("// BEGIN MESSAGE TEXT CONVERSION", escapeStart);
  assert.ok(escapeStart >= 0 && escapeEnd > escapeStart, "escapeHtml boundaries missing");
  vm.createContext(sandbox);
  vm.runInContext([
    source.slice(escapeStart, escapeEnd),
    snippet("RAW MIME ATTACHMENT HELPERS"),
    snippet("COMPOSE SIGNATURE HELPERS"),
  ].join("\n"), sandbox);
  return { api: sandbox, calls, file, bytesRead: () => cursor };
}

function loadComposeTools(identity = { htmlSigText: "Team signature" }, options = {}) {
  const loaded = loadSignatureHelpers(options);
  const { api, calls } = loaded;
  calls.sends = [];
  calls.reviews = [];
  const folder = { URI: DRAFTS_URI, getFlag: () => true };
  Object.assign(api, {
    setComposeIdentity(params) { params.identity = { draftsFolderURI: DRAFTS_URI, ...identity }; },
    resolveComposeFormat: (_identity, isHtml) => ({ useHtml: isHtml === true, format: 0 }),
    filePathsToAttachDescs: () => ({ descs: [], failed: [] }),
    descsToMsgAttachments: () => [],
    getAccessibleFolder: () => ({ folder }),
    findMessage: () => ({ folder, msgHdr: { folder } }),
    isSkipReviewBlocked: () => options.blocked === true,
    sendMessageDirectly: async (fields, _identity, _descs, _uri, compType, deliverMode) => {
      calls.sends.push({ fields, compType, deliverMode });
      return { success: true };
    },
  });
  Object.assign(api.Cc, {
    "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
    "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({ setHeader() {}, addAttachment() {} }) },
    "@mozilla.org/messengercompose;1": { getService: () => ({
      OpenComposeWindowWithParams(_parent, params) { calls.reviews.push(params); },
    }) },
  });
  vm.runInContext([
    snippet("OUTBOUND MAIL TOOLS"), snippet("TOOL CALL DISPATCH"),
  ].join("\n"), api);
  return loaded;
}

describe("Direct compose identity signatures", () => {
  it("appends the configured plain signature with one separator by default", () => {
    const { api } = loadSignatureHelpers();
    assert.equal(api.buildBodyWithSignature("Hello", { htmlSigText: "Jane\nSupport" }, false, false), "Hello\n\n-- \nJane\nSupport");
    assert.equal(api.buildBodyWithSignature("Hello", { htmlSigText: "-- \nJane" }, false, false), "Hello\n\n-- \nJane");
  });

  it("escapes plain signatures and bodies when composing HTML", () => {
    const { api } = loadSignatureHelpers();
    const body = api.buildBodyWithSignature("<Hi>\n&bye", { htmlSigText: '<Jane> & "support"' }, true, false);
    assert.match(body, /&lt;Hi&gt;<br>&amp;bye/);
    assert.match(body, /<div class="moz-signature">-- <br>&lt;Jane&gt; &amp; &quot;support&quot;<\/div>/);
    assert.doesNotMatch(body, /<Jane>/);
  });

  it("converts HTML signatures with Thunderbird parserUtils for a plain body", () => {
    const html = "<p>Jane &amp; Support</p><p>Second line</p>";
    const { api, calls } = loadSignatureHelpers({ plainText: "Jane & Support\nSecond line\n\n" });
    assert.equal(api.buildBodyWithSignature("Hello", { htmlSigText: html, htmlSigFormat: true }, false, false), "Hello\n\n-- \nJane & Support\nSecond line");
    assert.deepEqual(calls.conversions, [[html, 3, 0]]);
  });

  it("unwraps signature documents before putting them in the message body", () => {
    const html = '<!DOCTYPE html><html><head><title>Signature</title></head><body><b>Jane</b></body></html>';
    for (const DOMParser of [undefined, class {
      parseFromString(input, type) {
        assert.equal(input, html);
        assert.equal(type, "text/html");
        return { body: { innerHTML: "<b>Jane</b>" } };
      }
    }]) {
      const { api } = loadSignatureHelpers({ DOMParser });
      const body = api.buildBodyWithSignature("<p>Hello</p>", { htmlSigText: html, htmlSigFormat: true }, true, true);
      assert.equal((body.match(/<html>/g) || []).length, 1);
      assert.match(body, /<div class="moz-signature">-- <br><b>Jane<\/b><\/div>/);
      assert.doesNotMatch(body, /DOCTYPE|<title>/);
    }
  });

  it("inserts into an existing HTML body without treating signature text as replacement syntax", () => {
    const { api } = loadSignatureHelpers({ plainText: "$& $` $'" });
    const signature = "<b>$& $` $'</b>";
    const body = api.buildBodyWithSignature("<html><body>Hello</body></html>", { htmlSigText: signature, htmlSigFormat: true }, true, true);
    assert.equal(body, `<html><body>Hello<br><div class="moz-signature">-- <br>${signature}</div></body></html>`);
  });

  it("includeSignature false skips signature lookup entirely", () => {
    const { api } = loadSignatureHelpers();
    let lookups = 0;
    const identity = { get attachSignature() { lookups++; return false; } };
    assert.equal(api.buildBodyWithSignature("Already signed", identity, false, false, false), "Already signed");
    assert.equal(lookups, 0);
  });

  it("new drafts include a signature by default and honor an explicit opt-out through dispatch", async () => {
    for (const includeSignature of [undefined, false, true]) {
      const { api, calls } = loadComposeTools();
      const result = await api.dispatchTool("saveDraft", { to: "to@example.test", subject: "Subject", body: "Hello", isHtml: false, includeSignature });
      assert.equal(result.success, true);
      assert.equal(calls.sends[0].fields.body, includeSignature === false ? "Hello" : "Hello\n\n-- \nTeam signature");
    }
  });

  it("replacing a fetched draft preserves its signature unless explicitly requested", async () => {
    for (const includeSignature of [undefined, false, true]) {
      const { api, calls } = loadComposeTools();
      const body = "Hello\n\n-- \nTeam signature";
      const result = await api.dispatchTool("saveDraft", {
        to: "to@example.test", subject: "Subject", body, isHtml: false,
        replaceMessageId: "draft-id", replaceFolderPath: DRAFTS_URI, includeSignature,
      });
      assert.equal(result.success, true);
      assert.equal(calls.sends[0].fields.body, includeSignature === true ? `${body}\n\n-- \nTeam signature` : body);
    }
  });

  it("direct send honors includeSignature through the production dispatch", async () => {
    for (const includeSignature of [undefined, false, true]) {
      const { api, calls } = loadComposeTools();
      const result = await api.dispatchTool("sendMail", {
        to: "to@example.test", subject: "Subject", body: "Hello", isHtml: false, skipReview: true, includeSignature,
      });
      assert.equal(result.success, true);
      assert.equal(calls.sends[0].fields.body, includeSignature === false ? "Hello" : "Hello\n\n-- \nTeam signature");
      assert.equal(calls.reviews.length, 0);
    }
  });

  it("leaves signature insertion to Thunderbird on the review path", async () => {
    const { api, calls } = loadComposeTools();
    api.buildBodyWithSignature = () => assert.fail("review must not append a second signature");
    const result = await api.dispatchTool("sendMail", {
      to: "to@example.test", subject: "Subject", body: "Hello", isHtml: false, includeSignature: true,
    });
    assert.equal(result.success, true);
    assert.equal(calls.reviews[0].composeFields.body, "Hello");
    assert.equal(calls.sends.length, 0);
  });

  it("still blocks direct sending before reading signature preferences", async () => {
    const { api, calls } = loadComposeTools({}, { blocked: true });
    api.buildBodyWithSignature = () => assert.fail("blocked send must not read the signature");
    const result = await api.dispatchTool("sendMail", { to: "to@example.test", subject: "Subject", body: "Hello", skipReview: true });
    assert.match(result.error, /blocks skipReview/);
    assert.equal(calls.sends.length, 0);
  });
});

describe("Bounded identity signature files", () => {
  it("decodes UTF-8 and closes the actual input stream", () => {
    const { api, file, calls } = loadSignatureHelpers({ bytes: "Zażółć gęślą — 👋" });
    assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file }).content, "Zażółć gęślą — 👋");
    assert.equal(calls.opens, 1);
    assert.equal(calls.closes, 1);
  });

  it("classifies HTML signature files by their extension", () => {
    const { api, file } = loadSignatureHelpers({ bytes: "<b>Jane</b>", leafName: "signature.HTML" });
    assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file }).isHtmlSig, true);
  });

  for (const fileSize of [MAX_SIGNATURE_BYTES + 1, -1, Infinity, NaN, 1.5, "12", undefined, Number.MAX_SAFE_INTEGER + 1]) {
    it(`omits a signature with invalid or excessive file size ${String(fileSize)} before opening`, () => {
      const { api, file, calls } = loadSignatureHelpers({ fileSize });
      assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file }), null);
      assert.equal(calls.opens, 0);
      assert.equal(calls.reads.length, 0);
      assert.equal(calls.warnings.length, 1);
    });
  }

  it("accepts exactly the byte limit with bounded individual reads", () => {
    const { api, file, calls, bytesRead } = loadSignatureHelpers({ bytes: "X".repeat(MAX_SIGNATURE_BYTES) });
    assert.equal(api.readSignatureFileText(file).length, MAX_SIGNATURE_BYTES);
    assert.equal(bytesRead(), MAX_SIGNATURE_BYTES);
    assert.ok(calls.reads.length > 1);
    assert.ok(calls.reads.every(count => count <= 64 * 1024));
    assert.equal(calls.closes, 1);
  });

  it("stops a growing file one byte past the limit and omits the partial signature", () => {
    const { api, file, calls, bytesRead } = loadSignatureHelpers({ bytes: "X".repeat(MAX_SIGNATURE_BYTES + 8192), fileSize: 1 });
    assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file }), null);
    assert.equal(bytesRead(), MAX_SIGNATURE_BYTES + 1);
    assert.ok(calls.reads.every(count => count <= 64 * 1024));
    assert.equal(calls.closes, 1);
    assert.equal(calls.warnings.length, 1);
  });

  for (const failure of ["initError", "streamError"]) {
    it(`closes the stream and omits an unreadable signature after ${failure}`, () => {
      const { api, file, calls } = loadSignatureHelpers({ [failure]: new Error("read failed") });
      assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file }), null);
      assert.equal(calls.closes, 1);
      assert.equal(calls.warnings.length, 1);
    });
  }

  it("does not lose a successfully read signature when closing the stream fails", () => {
    const { api, file, calls } = loadSignatureHelpers({ bytes: "Jane", closeError: new Error("already closed") });
    assert.equal(api.readSignatureFileText(file), "Jane");
    assert.equal(calls.closes, 1);
  });
});
