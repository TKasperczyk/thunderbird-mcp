"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { makeDirectSendPrefs, prefTypes, clearOriginalRawFixture, composeHeaderParser } = require("./helpers/direct-send-prefs.cjs");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
const MAX_SIGNATURE_BYTES = 1024 * 1024;
const IMAGE_SIGNATURE_BYTES = 18 * 1024 * 1024;
const DEFAULT_SIZE_WARNING_BYTES = 20 * 1024 * 1024;
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
  const calls = { stats: 0, opens: 0, closes: 0, reads: [], warnings: [], conversions: [], mimeFiles: [] };
  let cursor = 0;
  const file = {
    get fileSize() {
      calls.stats++;
      if (options.statError) throw options.statError;
      return Object.hasOwn(options, "fileSize") ? options.fileSize : bytes.length;
    },
    leafName: options.leafName || "signature.txt",
    exists: () => options.exists !== false,
    isFile: () => options.isFile !== false,
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
    ...clearOriginalRawFixture,
    TextDecoder,
    Uint8Array,
    Cr: { NS_BASE_STREAM_CLOSED: "NS_BASE_STREAM_CLOSED" },
    console: { warn: (...args) => calls.warnings.push(args), error() {} },
    NetUtil: { readInputStreamToString: (input, count) => input.read(count) },
    Ci: {
      nsIPrefBranch: prefTypes,
      nsIDocumentEncoder: { OutputFormatted: 1, OutputLFLineBreak: 2 },
      nsIMsgCompType: { New: 0, Draft: 9 },
      nsIMsgCompDeliverMode: { Now: 0, SaveAsDraft: 4 },
      nsMsgFolderFlags: { Drafts: 0x400 },
    },
    Cc: {
      "@mozilla.org/mime;1": { getService: () => ({
        getTypeFromFile(signatureFile) {
          calls.mimeFiles.push(signatureFile);
          if (options.mimeError) throw options.mimeError;
          return options.mimeType ?? "text/plain";
        },
      }) },
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
    Services: { prefs: makeDirectSendPrefs() },
    MailServices: { headerParser: composeHeaderParser, accounts: { accounts: [{ key: "account", identities: [{ key: "sender" }] }] } },
    setComposeIdentity(params) { params.identity = { key: "sender", draftsFolderURI: DRAFTS_URI, ...identity }; },
    resolveComposeFormat: (_identity, isHtml) => ({ useHtml: isHtml === true, format: 0 }),
    materializeAttachmentDescs: descs => descs,
    filePathsToAttachDescs: () => ({ descs: [], failed: [] }),
    descsToMsgAttachments: () => [],
    getAccessibleFolder: () => ({ folder }),
    findMessage: () => ({ folder, msgHdr: { folder } }),
    isSkipReviewBlocked: () => options.blocked === true,
    sendMessageDirectly: async (fields, _identity, _descs, _uri, compType, deliverMode) => {
      calls.sends.push({ fields, compType, deliverMode });
      return { success: true };
    },
    sendThroughCompose: async params => {
      calls.sends.push({ fields: params.composeFields, compType: params.type, deliverMode: api.Ci.nsIMsgCompDeliverMode.Now });
      return { success: true };
    },
  });
  Object.assign(api.Cc, {
    "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
    "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({ headers: {}, setHeader(name, value) { this.headers[name] = value; }, addAttachment() {} }) },
    "@mozilla.org/messengercompose;1": { getService: () => ({
      OpenComposeWindowWithParams(_parent, params) { calls.reviews.push(params); },
    }) },
  });
  vm.runInContext([
    snippet("DIRECT SEND ENCRYPTION GUARD"), snippet("DIRECT SEND ORIGINAL GUARD"),
    snippet("COMPOSE SIZE GUARD"), snippet("COMPOSE ACCOUNT KEY"), snippet("OUTBOUND ATTACHMENT LIMITS"), snippet("OUTBOUND MAIL TOOLS"), snippet("TOOL CALL DISPATCH"),
  ].join("\n"), api);
  return loaded;
}

describe("Direct compose identity signatures", () => {
  it("appends the configured plain signature with one separator by default", () => {
    const { api } = loadSignatureHelpers();
    assert.equal(api.buildBodyWithSignature("Hello", { htmlSigText: "Jane\nSupport" }, false, false), "Hello\n\n-- \nJane\nSupport");
    assert.equal(api.buildBodyWithSignature("Hello", { htmlSigText: "-- \nJane" }, false, false), "Hello\n\n-- \nJane");
  });

  for (const htmlSigFormat of [false, true]) {
    it(`honors suppressed separators for a ${htmlSigFormat ? "HTML" : "plain"} signature in both compose formats`, () => {
      const { api } = loadSignatureHelpers({ plainText: "Jane\n" });
      const identity = { htmlSigText: htmlSigFormat ? "<b>Jane</b>" : "Jane", htmlSigFormat, suppressSigSep: true };
      assert.equal(api.buildBodyWithSignature("Hello", identity, false, false), "Hello\n\nJane");
      assert.match(api.buildBodyWithSignature("Hello", identity, true, false),
        htmlSigFormat ? /<div class="moz-signature"><b>Jane<\/b><\/div>/ : /<div class="moz-signature">Jane<\/div>/);
    });
  }

  for (const signature of ["\n-- \nJane", "Jane\n-- \nTeam", "Jane\r\n\t--\t\r\nTeam", "Jane\n\u00a0--\u00a0\nTeam"]) {
    it(`keeps an existing separator anywhere in ${JSON.stringify(signature)}`, () => {
      const { api } = loadSignatureHelpers();
      assert.equal(api.buildBodyWithSignature("Hello", { htmlSigText: signature }, false, false), `Hello\n\n${signature}`);
    });
  }

  it("recognizes a separator in the rendered text of an HTML signature", () => {
    const signature = '<div>Jane</div><div>-- </div><div>Team</div>';
    const { api } = loadSignatureHelpers({ plainText: "\nJane\n-- \nTeam\n" });
    const body = api.buildBodyWithSignature("Hello", { htmlSigText: signature, htmlSigFormat: true }, true, false);
    assert.ok(body.includes(`<div class="moz-signature">${signature}</div>`));
    assert.equal((body.match(/-- /g) || []).length, 1);
  });

  it("does not mistake inline dashes for a separator line", () => {
    const { api } = loadSignatureHelpers();
    assert.equal(api.buildBodyWithSignature("Hello", { htmlSigText: "Jane -- Team" }, false, false), "Hello\n\n-- \nJane -- Team");
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

  for (const tool of ["sendMail", "saveDraft"]) {
    for (const isHtml of [false, true]) {
      it(`${tool} honors separator suppression in ${isHtml ? "HTML" : "plain"} production dispatch`, async () => {
        const { api, calls } = loadComposeTools({ htmlSigText: "Jane", suppressSigSep: true });
        const result = await api.dispatchTool(tool, {
          to: "to@example.test", subject: "Subject", body: "Hello", isHtml,
          ...(tool === "sendMail" ? { skipReview: true } : {}),
        });
        assert.equal(result.success, true);
        assert.equal(calls.sends.length, 1);
        assert.doesNotMatch(calls.sends[0].fields.body, /-- /);
        assert.ok(calls.sends[0].fields.body.includes("Jane"));
      });

      it(`${tool} omits image-file bytes and inline fallback in ${isHtml ? "HTML" : "plain"} production dispatch`, async () => {
        const identity = { attachSignature: true, htmlSigText: "Inline Jane" };
        const { api, calls, file } = loadComposeTools(identity, {
          bytes: Buffer.from([0x89, 0x50, 0x4e, 0x47]), leafName: "signature", mimeType: "image/png",
        });
        identity.signature = file;
        const result = await api.dispatchTool(tool, {
          to: "to@example.test", subject: "Subject", body: "Hello", isHtml,
          ...(tool === "sendMail" ? { skipReview: true } : {}),
        });
        assert.equal(result.success, true);
        assert.equal(calls.sends[0].fields.body, isHtml
          ? '<html><head><meta charset="UTF-8"></head><body>Hello</body></html>' : "Hello");
        assert.equal(calls.opens, 0);
        assert.equal(calls.reads.length, 0);
        assert.equal(calls.warnings.length, 1);
      });
    }
  }

  it("leaves signature insertion to Thunderbird on the review path", async () => {
    const { api, calls } = loadComposeTools();
    let signatureReads = 0;
    const build = api.buildSignatureFragment;
    api.buildSignatureFragment = (...args) => { signatureReads++; return build(...args); };
    const result = await api.dispatchTool("sendMail", {
      to: "to@example.test", subject: "Subject", body: "Hello", isHtml: false, includeSignature: true,
    });
    assert.equal(result.success, true);
    assert.equal(calls.reviews[0].composeFields.body, "Hello");
    assert.equal(signatureReads, 1, "estimate reads the signature, but review body leaves insertion to Thunderbird");
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

describe("Review signature size estimates", () => {
  const args = { to: "to@example.test", subject: "Subject", body: "Hello", isHtml: true, includeSignature: false };

  for (const options of [
    { mimeType: "image/png", leafName: "signature" },
    { mimeError: new Error("unknown type"), leafName: "signature.PNG" },
  ]) {
    it(`counts a large image signature using only stat (${options.leafName})`, async () => {
      const identity = { attachSignature: true };
      const { api, calls, file } = loadComposeTools(identity, { ...options, fileSize: IMAGE_SIGNATURE_BYTES });
      identity.signature = file;
      const dry = await api.dispatchTool("dryRunCompose", args);
      assert.equal(dry.wouldSucceed, true);
      assert.equal(dry.warningThresholdBytes, DEFAULT_SIZE_WARNING_BYTES);
      // 18 MiB at 72-column base64 wrapping, including the terminal CRLF.
      assert.ok(dry.estimatedEncodedBytes >= 25864876);
      assert.match(dry.warnings[0], /exceeds Thunderbird's large-message threshold/);
      assert.equal(dry.blockers.length, 0);
      assert.equal(dry.attachments.length, 0, "signature metadata is not a caller attachment");
      assert.equal(calls.stats, 1);
      assert.equal(calls.opens + calls.reads.length + calls.sends.length + calls.reviews.length, 0);

      const directArgs = { ...args, skipReview: true, includeSignature: true };
      const direct = await api.dispatchTool("dryRunCompose", directArgs);
      const unsigned = await api.dispatchTool("dryRunCompose", { ...directArgs, includeSignature: false });
      assert.equal(direct.wouldSucceed, true);
      assert.equal(direct.estimatedEncodedBytes, unsigned.estimatedEncodedBytes);
      assert.equal(direct.warnings.length, 0);
      assert.equal(calls.stats, 1, "direct sending still omits images without stat");
      assert.equal((await api.dispatchTool("sendMail", directArgs)).success, true);
      assert.equal(calls.sends.length, 1);
      assert.equal(calls.stats, 1);
      assert.equal(calls.opens + calls.reads.length, 0);

      assert.equal((await api.dispatchTool("sendMail", args)).success, true);
      assert.equal(calls.reviews.length, 1, "a review warning must not block the window");
      assert.equal(calls.reviews[0].composeFields.body, calls.sends[0].fields.body);
      assert.equal(calls.opens + calls.reads.length, 0);
    });
  }

  it("keeps a bounded small image below the threshold without reading it", async () => {
    const identity = { attachSignature: true };
    const { api, calls, file } = loadComposeTools(identity, { mimeType: "image/png", fileSize: 54 });
    identity.signature = file;
    const dry = await api.dispatchTool("dryRunCompose", args);
    assert.equal(dry.wouldSucceed, true);
    assert.ok(dry.estimatedEncodedBytes > 76 && dry.estimatedEncodedBytes < DEFAULT_SIZE_WARNING_BYTES);
    assert.equal(dry.warnings.length, 0);
    assert.equal(calls.stats, 1);
    assert.equal(calls.opens + calls.reads.length + calls.sends.length + calls.reviews.length, 0);
  });

  for (const [label, options] of [
    ["unreadable image stat", { mimeType: "image/png", statError: new Error("stat failed") }],
    ["invalid image size", { mimeType: "image/png", fileSize: -1 }],
    ["oversized text file", { fileSize: MAX_SIGNATURE_BYTES + 1 }],
    ["unreadable text file", { initError: new Error("read failed") }],
    ["HTML file with an image resource", { mimeType: "text/html", bytes: '<img src="file:///signature.png">' }],
  ]) {
    it(`reports an unknown review estimate for an ${label}`, async () => {
      const identity = { attachSignature: true };
      const { api, calls, file } = loadComposeTools(identity, options);
      identity.signature = file;
      const dry = await api.dispatchTool("dryRunCompose", args);
      assert.equal(dry.wouldSucceed, true);
      assert.equal(dry.estimatedEncodedBytes, null);
      assert.equal(dry.blockers.length, 0);
      assert.match(dry.warnings[0], /estimate excludes the identity signature/);
      assert.equal(calls.sends.length + calls.reviews.length, 0);
      if (!options.initError && !options.bytes) assert.equal(calls.opens + calls.reads.length, 0);
    });
  }

  for (const htmlSigText of ['<img src="file:///signature.png">', '<div style="background: url(https://example.test/image)">Jane</div>']) {
    it(`warns about unbounded inline signature resources: ${htmlSigText}`, async () => {
      const { api, calls } = loadComposeTools({ htmlSigText, htmlSigFormat: true });
      const dry = await api.dispatchTool("dryRunCompose", args);
      assert.equal(dry.wouldSucceed, true);
      assert.equal(dry.estimatedEncodedBytes, null);
      assert.match(dry.warnings[0], /estimate excludes the identity signature/);
      assert.equal(calls.opens + calls.reads.length + calls.sends.length + calls.reviews.length, 0);
    });
  }

  it("still bounds plain files, blank files and HTML signatures without embedded resources", async () => {
    for (const config of [
      { identity: { attachSignature: true }, bytes: "Jane" },
      { identity: { attachSignature: true }, bytes: " \n\t" },
      { identity: { htmlSigText: '<b>Jane</b> <a href="mailto:jane@example.test">Email</a>', htmlSigFormat: true } },
    ]) {
      const { api, calls, file } = loadComposeTools(config.identity, config);
      config.identity.signature = file;
      const dry = await api.dispatchTool("dryRunCompose", args);
      assert.equal(dry.wouldSucceed, true);
      assert.ok(Number.isSafeInteger(dry.estimatedEncodedBytes));
      assert.equal(dry.warnings.length, 0);
      assert.equal(calls.sends.length + calls.reviews.length, 0);
    }
  });
});

describe("Bounded identity signature files", () => {
  it("decodes UTF-8 and closes the actual input stream", () => {
    const { api, file, calls } = loadSignatureHelpers({ bytes: "Zażółć gęślą — 👋" });
    assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file }).content, "Zażółć gęślą — 👋");
    assert.equal(calls.opens, 1);
    assert.equal(calls.closes, 1);
  });

  it("classifies extensionless HTML signature files using the MIME service", () => {
    const { api, file, calls } = loadSignatureHelpers({ bytes: "<b>Jane</b>", leafName: "signature", mimeType: "text/html" });
    assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file }).isHtmlSig, true);
    assert.deepEqual(calls.mimeFiles, [file]);
  });

  it("falls back to the HTML extension when the MIME service cannot classify a file", () => {
    const { api, file } = loadSignatureHelpers({ bytes: "<b>Jane</b>", leafName: "signature.HTML", mimeError: new Error("unknown file type") });
    assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file }).isHtmlSig, true);
  });

  it("uses the MIME service result in preference to the file extension", () => {
    const { api, file } = loadSignatureHelpers({ bytes: "Jane", leafName: "signature.png", mimeType: "text/plain" });
    assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file }).content, "Jane");
  });

  for (const mimeType of ["image/png", "image/svg+xml"]) {
    it(`skips an extensionless ${mimeType} signature without opening it or using inline text`, () => {
      const { api, file, calls } = loadSignatureHelpers({ leafName: "signature", mimeType });
      assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file, htmlSigText: "Inline Jane" }), null);
      assert.deepEqual(calls.mimeFiles, [file]);
      assert.equal(calls.opens, 0);
      assert.equal(calls.reads.length, 0);
      assert.equal(calls.warnings.length, 1);
      assert.match(calls.warnings[0][0], /image signature files are omitted/);
    });
  }

  for (const options of [{ mimeError: new Error("unknown file type") }, { mimeType: "" }]) {
    it(`skips image extensions when MIME detection ${options.mimeError ? "throws" : "returns no type"}`, () => {
      const { api, file, calls } = loadSignatureHelpers({ leafName: "signature.PNG", ...options });
      assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file, htmlSigText: "Inline Jane" }), null);
      assert.equal(calls.opens, 0);
      assert.equal(calls.warnings.length, 1);
    });
  }

  for (const [label, options] of [
    ["missing", { exists: false }],
    ["directory", { isFile: false }],
    ["empty", { bytes: "" }],
    ["blank", { bytes: " \n\t" }],
    ["unreadable", { initError: new Error("read failed") }],
  ]) {
    it(`omits a ${label} signature file without falling back to inline text`, () => {
      const { api, file } = loadSignatureHelpers(options);
      assert.equal(api.getIdentitySignature({ attachSignature: true, signature: file, htmlSigText: "Inline Jane" }), null);
    });
  }

  it("does not use inline text when the selected signature file is absent", () => {
    const { api } = loadSignatureHelpers();
    assert.equal(api.getIdentitySignature({ attachSignature: true, htmlSigText: "Inline Jane" }), null);
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

describe("sendMail custom headers", () => {
  const headers = [{ name: "X-AI-Assisted", value: "Drafted with AI" }];
  const args = { to: "to@example.test", subject: "Subject", body: "Hello", headers };

  it("adds caller X- headers to the review window fields", async () => {
    const { api, calls } = loadComposeTools();
    assert.equal((await api.dispatchTool("sendMail", args)).success, true);
    assert.equal(calls.reviews[0].composeFields.headers["X-AI-Assisted"], "Drafted with AI");
  });

  it("adds caller X- headers to a direct send", async () => {
    const { api, calls } = loadComposeTools();
    assert.equal((await api.dispatchTool("sendMail", { ...args, skipReview: true })).success, true);
    assert.equal(calls.sends[0].fields.headers["X-AI-Assisted"], "Drafted with AI");
  });

  it("refuses invalid headers before opening a window, also in dryRunCompose", async () => {
    const { api, calls } = loadComposeTools();
    const bad = { ...args, headers: [{ name: "Bcc", value: "spy@example.test" }] };
    assert.match((await api.dispatchTool("sendMail", bad)).error, /must start with X-/);
    assert.match(api.dryRunCompose(bad).blockers[0], /must start with X-/);
    assert.equal(calls.reviews.length, 0);
    assert.equal(calls.sends.length, 0);
  });
});
