"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { makeDirectSendPrefs, prefTypes } = require("./helpers/direct-send-prefs.cjs");
const { makeDirectSendAttachments, assertNoAttachmentEffects } = require("./helpers/direct-send-attachments.cjs");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
// Same marker assertions and VM loading pattern as validation.test.cjs.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function loadMessageTools({ mime = { contentType: "text/plain", body: "visible" }, allowed = false, unreadable = false, raw = "Content-Type: text/plain\r\n\r\nvisible", DOMParser: Parser = null, streamError, mimeError, imap = null, remote = imap, localStream = false, hdr: headerOverrides = {} } = {}) {
  const attachments = makeDirectSendAttachments();
  const calls = {
    options: [], imports: 0, streams: 0, sends: [], reviews: 0, logs: [], order: [], timers: [],
    created: { params: 0, fields: 0, compose: 0, sender: 0 },
    attachments: attachments.calls,
    serverStreams: 0,
  };
  const folder = {
    server: remote ? { type: remote.type ?? "imap" } : {},
    getUriForMsg: () => "message-uri",
    getMsgInputStream() {
      calls.streams++;
      calls.genericStreams++;
      calls.order.push("raw");
      if (streamError) throw streamError;
      return { close() {} };
    },
    // Thunderbird 145+ bounds this stream to the message; older builds lack it.
    ...(localStream ? { getLocalMsgStream() {
      calls.streams++;
      calls.localStreams++;
      calls.order.push("raw");
      if (streamError) throw streamError;
      return { close() {} };
    } } : {}),
  };
  calls.genericStreams = 0;
  calls.localStreams = 0;
  const hdr = {
    messageId: "message-1", subject: "protected subject", author: "sender@example.test",
    recipients: "reader@example.test", folder, date: 0,
    flags: remote?.offline ? 0x80 : 0,
    ...headerOverrides,
  };
  const sandbox = {
    console: { error: (...args) => calls.logs.push(args), warn: (...args) => calls.logs.push(args) },
    DOMParser: Parser,
    TextDecoder,
    Uint8Array,
    atob,
    btoa,
    Services: { prefs: {
      ...makeDirectSendPrefs(),
      getBoolPref() { if (unreadable) throw Error("unreadable"); return allowed; },
    } },
    ChromeUtils: { importESModule: () => {
      calls.imports++;
      calls.order.push("import");
      return {
      MsgHdrToMimeMessage(msgHdr, _listener, callback, _download, options) {
        calls.options.push(options);
        calls.order.push("mime");
        if (mimeError) throw mimeError;
        callback(msgHdr, mime);
      },
      };
    } },
    findMessage: () => ({ msgHdr: hdr, folder }),
    getUserTags: () => [],
    getConfiguredGetMessagesLimit: () => 10,
    readMessageStreamFully: () => raw,
    isSkipReviewBlocked: () => false,
    isToolEnabled: () => true,
    filePathsToAttachDescs: attachments.convert,
    materializeAttachmentDescs: attachments.materialize,
    buildBodyWithSignature: body => body || "",
    MailServices: { accounts: { accounts: [{ key: "account", identities: [{ key: "sender" }] }] } },
    descsToMsgAttachments: attachments.toNative,
    Cc: {
      "@mozilla.org/timer;1": { createInstance() {
        const timer = {
          initWithCallback(callback, delay) { this.callback = callback; this.delay = delay; },
          cancel() { this.cancelled = true; },
        };
        calls.timers.push(timer);
        return timer;
      } },
      "@mozilla.org/messengercompose/composeparams;1": { createInstance() {
        calls.created.params++;
        return {};
      } },
      "@mozilla.org/messengercompose/composefields;1": { createInstance() {
        calls.created.fields++;
        return { setHeader() {} };
      } },
      "@mozilla.org/messengercompose/compose;1": { createInstance() {
        calls.created.compose++;
        return {};
      } },
      "@mozilla.org/messengercompose/send;1": { createInstance() {
        calls.created.sender++;
        return {};
      } },
    },
    Ci: { nsIPrefBranch: prefTypes, nsITimer: { TYPE_ONE_SHOT: 0 }, nsIMsgCompType: { Reply: 1, ReplyAll: 2, ForwardInline: 3 }, nsIMsgCompDeliverMode: { Now: 0 }, nsIMsgFolder: {} },
    setComposeIdentity(params) { params.identity = { key: "sender" }; },
    resolveComposeFormat: () => ({ useHtml: false, format: 0 }),
    async sendThroughCompose(params, attachments) {
      calls.sends.push({ ...params.composeFields, attachments });
      return { success: true };
    },
    async openComposeWindowWithCustomizations() { calls.reviews++; return { success: true }; },
  };
  if (remote) {
    // A remote (IMAP or add-on) message service streaming the server's copy in two chunks.
    Object.assign(sandbox.Ci, { nsMsgMessageFlags: { Offline: 0x80 } });
    sandbox.Cr = { NS_BINDING_ABORTED: 0x804b0002 };
    sandbox.Components = { isSuccessCode: status => status === 0 };
    sandbox.NetUtil = { readInputStreamToString: (stream, count) => stream.read(count) };
    sandbox.ChromeUtils.generateQI = () => () => {};
    sandbox.MailServices = {
      messageServiceFromURI: () => ({
        streamMessage(_uri, listener) {
          calls.serverStreams++;
          const data = remote.serverRaw;
          let offset = 0;
          const stream = { read(count) { const chunk = data.slice(offset, offset + count); offset += chunk.length; return chunk; } };
          listener.onStartRequest({ cancel() {} });
          const half = Math.floor(data.length / 2);
          listener.onDataAvailable(null, stream, 0, half);
          listener.onDataAvailable(null, stream, half, data.length - half);
          listener.onStopRequest(null, remote.status ?? 0);
        },
      }),
    };
  }
  vm.createContext(sandbox);
  vm.runInContext([
    source.match(/^const PREF_\w+ = .+;$/gm).join("\n"),
    snippet("PRIVACY PREFERENCE HELPERS"), snippet("MCP TEXT SANITIZATION"),
    snippet("MESSAGE TEXT CONVERSION"), snippet("RAW MIME PARSING HELPERS"),
    snippet("RAW MIME ATTACHMENT HELPERS"), snippet("INLINE IMAGE CONTENT HELPERS"),
    snippet("INLINE ATTACHMENT BASE64 HELPERS"), snippet("ENCRYPTED MESSAGE GUARD"),
    snippet("DIRECT SEND ENCRYPTION GUARD"), snippet("DIRECT SEND ORIGINAL GUARD"),
    snippet("COMPOSE SIZE GUARD"), snippet("COMPOSE ACCOUNT KEY"),
    snippet("MESSAGE READ TOOLS"), snippet("REPLY TOOL"), snippet("FORWARD TOOL"),
  ].join("\n"), sandbox);
  // The production stream reader is loaded with its helper block; serve the fixture instead.
  sandbox.readMessageStreamFully = () => raw;
  // Searched for test call-through defaults -> found the invoke lambdas below.
  // Give every direct call real path and inline inputs without changing reads,
  // review windows, or automatic drafts that omit their caller attachments.
  for (const [name, attachmentIndex, skipReviewIndex] of [
    ["replyToMessage", 9, 10],
    ["forwardMessage", 8, 9],
  ]) {
    const invoke = sandbox[name];
    sandbox[name] = (...args) => {
      if (args[skipReviewIndex] && args[attachmentIndex] === undefined) args[attachmentIndex] = attachments.entries;
      return invoke(...args);
    };
  }
  return { api: sandbox, calls };
}

// Searched for original-message refusal assertions -> found identity-only
// assertRefused in direct-send-encryption, whose no-MIME check does not fit.
// Reuse its constructor-counter and review-instruction assertions here.
function assertOriginalRefused(calls, result) {
  assert.equal(result.success, undefined);
  assert.match(result.error, /encrypted messages is blocked/);
  assert.match(result.error, /skipReview: false/);
  assert.match(result.error, /review window/);
  assert.match(result.error, /Thunderbird.*encryption/);
  assert.doesNotMatch(result.error, /enable.*read encrypted/i);
  assert.deepEqual(calls.created, { params: 0, fields: 0, compose: 0, sender: 0 });
  assert.equal(calls.sends.length, 0);
  assert.equal(calls.reviews, 0);
  assertNoAttachmentEffects(calls.attachments);
}

const encryptedTrees = [
  { label: "OpenPGP", contentType: "multipart/encrypted", parts: [] },
  { label: "nested OpenPGP", contentType: "message/rfc822", parts: [{ contentType: "multipart/encrypted", parts: [] }] },
  { label: "decrypted S/MIME", contentType: "multipart/mixed", isEncrypted: true, parts: [{ contentType: "text/plain", body: "decrypted body" }] },
  { label: "PKCS7", contentType: "application/pkcs7-mime; smime-type=enveloped-data" },
  { label: "authenticated PKCS7", contentType: 'application/pkcs7-mime; smime-type="authEnveloped-data"' },
  { label: "legacy PKCS7", contentType: "application/x-pkcs7-mime" },
  { label: "ambiguous PKCS7", contentType: "application/pkcs7-mime; smime-type=enveloped-data; smime-type=signed-data" },
  { label: "empty encrypted MIME tree", contentType: "message/rfc822", headers: { "content-type": ["application/pkcs7-mime; smime-type=enveloped-data"] }, parts: [] },
  { label: "inline OpenPGP", contentType: "text/plain", body: "-----BEGIN PGP MESSAGE-----\nciphertext" },
];

const protectedRawOriginals = [
  ["OpenPGP MIME", [
    'Content-Type: multipart/encrypted; protocol="application/pgp-encrypted"; boundary="enc"', "",
    "--enc", "Content-Type: application/pgp-encrypted", "", "Version: 1",
    "--enc", "Content-Type: application/octet-stream", "", "ciphertext", "--enc--", "",
  ].join("\r\n")],
  ["S/MIME encrypted", "Content-Type: application/pkcs7-mime; smime-type=enveloped-data\r\nContent-Transfer-Encoding: base64\r\n\r\nMIIB"],
  ["S/MIME opaque signed", "Content-Type: application/pkcs7-mime; smime-type=signed-data\r\nContent-Transfer-Encoding: base64\r\n\r\nMIIB"],
  ["legacy S/MIME", "Content-Type: application/x-pkcs7-mime\r\nContent-Transfer-Encoding: base64\r\n\r\nMIIB"],
  ["inline armor", "Content-Type: text/plain\r\n\r\n-----BEGIN PGP MESSAGE-----\r\nciphertext"],
  ["CMS filename without a reliable MIME type", "Content-Type: application/octet-stream; name=smime.p7m\r\nContent-Transfer-Encoding: base64\r\n\r\nMIIB"],
  ["unknown malformed MIME", "Content-Type: multipart/mixed\r\n\r\nno boundary"],
  ["incomplete raw MIME", "Content-Type: text/plain"],
];

describe("Encrypted message privacy", () => {
  it("getMessage keeps complete recipient and raw Cc lists for an explicit read", async () => {
    const raw = "=?UTF-8?Q?Doe=2C_John?= <j@x>, b@x, c@x, d@x";
    const decoded = "Doe, John <j@x>, b@x, c@x, d@x";
    const { api } = loadMessageTools({ hdr: {
      recipients: raw, mime2DecodedRecipients: decoded, ccList: raw,
    } });
    const result = await api.getMessage("message-1", "folder");
    assert.equal(result.recipients, decoded);
    assert.equal(result.ccList, raw);
    assert.equal(result.body, "visible");
  });

  for (const encrypted of encryptedTrees) {
    for (const rawSource of [false, true]) {
      it(`withholds ${encrypted.label}, rawSource=${rawSource}, before any content or attachment access`, async () => {
        const mime = { ...encrypted, get allUserAttachments() { return assert.fail("attachment metadata must not be read"); } };
        const { api, calls } = loadMessageTools({ mime, ...(rawSource ? { raw: protectedRawOriginals[1][1] } : {}) });
        const result = await api.getMessage("message-1", "folder", true, "html", rawSource, true);
        assert.equal(result.encryptedContentWithheld, true);
        assert.match(result.body, /Encrypted message content withheld/);
        assert.equal(result.bodyIsHtml, false);
        assert.equal(result.attachments.length, 0);
        assert.equal(result.subject, "[Encrypted message]");
        assert.equal(result.rawSource, undefined);
        assert.equal(calls.streams, 1);
        if (rawSource) assert.equal(calls.options.length, 0);
        else assert.equal(calls.options[0].examineEncryptedParts, false);
        assert.doesNotMatch(JSON.stringify(result), /decrypted body|protected subject|ciphertext/);
      });
    }
  }

  it("fails closed when the encrypted preference cannot be read", async () => {
    const { api, calls } = loadMessageTools({ mime: encryptedTrees[2], unreadable: true });
    const result = await api.getMessage("message-1", "folder");
    assert.equal(result.encryptedContentWithheld, true);
    assert.equal(calls.options[0].examineEncryptedParts, false);
    assert.equal(calls.streams, 1);
  });

  it("getMessages inherits withholding for each item", async () => {
    const { api } = loadMessageTools({ mime: encryptedTrees[2], raw: protectedRawOriginals[1][1] });
    const result = await api.getMessages([{ messageId: "message-1", folderPath: "folder" }], true, "text", true);
    assert.equal(result.messages[0].encryptedContentWithheld, true);
    assert.equal(result.succeeded, 1);
  });

  it("explicit opt-in permits decrypted content and normal attachment metadata", async () => {
    const { api, calls } = loadMessageTools({ allowed: true, mime: {
      ...encryptedTrees[2], allUserAttachments: [{ name: "document.txt", contentType: "text/plain", size: 10 }],
    } });
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.equal(result.body, "decrypted body");
    assert.notEqual(result.encryptedContentWithheld, true);
    assert.equal(result.attachments[0].name, "document.txt");
    assert.equal(calls.options[0].examineEncryptedParts, true);
  });

  it("explicit opt-in permits raw mode", async () => {
    const { api, calls } = loadMessageTools({ allowed: true, mime: encryptedTrees[2], raw: "raw data" });
    const result = await api.getMessage("message-1", "folder", false, "markdown", true);
    assert.equal(result.rawSource, "raw data");
    assert.equal(calls.streams, 1);
  });

  it("ordinary mail and multipart/signed remain readable without opting in", async () => {
    const { api, calls } = loadMessageTools({ mime: { contentType: "multipart/signed", parts: [{ contentType: "text/plain", body: "signed text" }] } });
    const result = await api.getMessage("message-1", "folder", false, "markdown");
    assert.equal(result.body, "signed text");
    assert.equal(calls.options[0].examineEncryptedParts, false);
  });

  for (const contentType of ["application/pkcs7-mime", "application/x-pkcs7-mime"]) {
    for (const isEncrypted of [false, true]) {
      for (const allowed of [false, true]) {
        it(`requires opt-in for ${contentType} signed-data, wrapper=${isEncrypted}, allowed=${allowed}`, async () => {
          const { api, calls } = loadMessageTools({ allowed, mime: {
            contentType: "message/rfc822", headers: { "content-type": [`${contentType}; SMIME-TYPE="SIGNED-DATA"`] },
            parts: [{
              contentType, headers: { "content-type": [`${contentType}; smime-type=signed-data`] },
              isEncrypted, parts: [{ contentType: "text/plain", body: "signed content" }],
            }],
          } });
          const result = await api.getMessage("message-1", "folder", false, "text");
          assert.equal(result.encryptedContentWithheld === true, !allowed);
          if (allowed) assert.equal(result.body, "signed content");
          else assert.doesNotMatch(JSON.stringify(result), /signed content/);
          assert.equal(calls.options[0].examineEncryptedParts, allowed);
          const replies = [
            await api.replyToMessage("message-1", "folder", "intro", false, false, undefined, undefined, undefined, undefined, undefined, true),
            await api.forwardMessage("message-1", "folder", "to@example.test", "intro", false, undefined, undefined, undefined, undefined, true),
          ];
          for (const reply of replies) {
            assertOriginalRefused(calls, reply);
          }
          assert.equal(calls.sends.length, 0);
        });
      }
    }
  }

  it("does not let a signed-data wrapper override encrypted descendants", async () => {
    const { api } = loadMessageTools({ mime: {
      contentType: "application/pkcs7-mime; smime-type=signed-data",
      parts: [{ contentType: "multipart/encrypted" }],
    } });
    assert.equal((await api.getMessage("message-1", "folder")).encryptedContentWithheld, true);
  });

  it("returns raw source without depending on MIME parsing", async () => {
    const { api, calls } = loadMessageTools({ mime: null });
    const result = await api.getMessage("message-1", "folder", true, "text", true);
    assert.equal(result.rawSource, "Content-Type: text/plain\r\n\r\nvisible");
    assert.equal(calls.streams, 1);
    assert.equal(calls.options.length, 0);
  });

  for (const tool of ["replyToMessage", "forwardMessage"]) {
    const invoke = (api, skipReview) => tool === "replyToMessage"
      ? api.replyToMessage("message-1", "folder", "intro", false, false, undefined, undefined, undefined, undefined, undefined, skipReview)
      : api.forwardMessage("message-1", "folder", "recipient@example.test", "intro", false, undefined, undefined, undefined, undefined, skipReview);
    for (const [label, raw] of protectedRawOriginals) {
      for (const allowed of [false, true]) {
        it(`${tool} refuses ${label} raw content before Gloda can run, opt-in=${allowed}`, async () => {
          const { api, calls } = loadMessageTools({
            allowed, raw,
            mime: { contentType: "text/plain", body: "decrypted plaintext that Gloda must never produce" },
          });
          assertOriginalRefused(calls, await invoke(api, true));
          assert.equal(calls.streams, 1);
          assert.equal(calls.imports, 0, "The MIME module is not imported for protected or unknown raw content");
          assert.equal(calls.options.length, 0, "Gloda/libmime must never receive a protected original");
          assert.deepEqual(calls.order, ["raw"]);
        });
      }
    }
    for (const encrypted of encryptedTrees) {
      for (const allowed of [false, true]) {
        it(`${tool} refuses ${encrypted.label} before any native compose creation, opt-in=${allowed}`, async () => {
          const mime = {
            ...encrypted,
            get allUserAttachments() { return assert.fail("Refused originals must not expose attachment metadata"); },
          };
          const { api, calls } = loadMessageTools({ mime, allowed });
          assertOriginalRefused(calls, await invoke(api, true));
          assert.equal(calls.streams, 1);
          assert.equal(calls.options.length, 1);
          assert.deepEqual(calls.order, ["raw", "import", "mime"], "Clear raw classification precedes the structured guard");
        });
      }
    }
    for (const [label, options] of [
      ["missing MIME result", { mime: null }],
      ["unclassified content type", { mime: { contentType: "text plain", body: "private body" } }],
      ["MIME parser failure", { mimeError: new Error("MIME parser failed") }],
      ["classification failure", { mime: { get contentType() { throw new Error("Classification failed"); } } }],
      ["body extraction failure", { mime: { contentType: "text/plain", get body() { throw new Error("Body unavailable"); } } }],
      ["missing structured body and unreadable raw", { mime: { contentType: "message/rfc822", parts: [] }, streamError: new Error("Raw unavailable") }],
      ["missing structured body and missing raw", { mime: { contentType: "message/rfc822", parts: [] }, raw: "" }],
    ]) {
      for (const allowed of [false, true]) {
        it(`${tool} refuses ${label} before any native compose creation, opt-in=${allowed}`, async () => {
          const { api, calls } = loadMessageTools({ ...options, allowed });
          assertOriginalRefused(calls, await invoke(api, true));
          assert.equal(calls.options.length, options.streamError || options.raw === "" ? 0 : 1);
          assert.equal(calls.order[0], "raw");
        });
      }
    }
    for (const allowed of [false, true]) {
      it(`${tool} allows an ordinary original with opt-in=${allowed}`, async () => {
        const { api, calls } = loadMessageTools({ allowed });
        const result = await invoke(api, true);
        assert.equal(result.success, true, result.error);
        assert.equal(calls.sends.length, 1);
        assert.match(calls.sends[0].body, /visible/);
        assert.equal(calls.options.length, 1);
        assert.equal(calls.options[0].examineEncryptedParts, false);
        assert.equal(calls.streams, 1);
        assert.deepEqual(calls.order, ["raw", "import", "mime"]);
        assert.equal(calls.attachments.tempCreates.length, 1);
        assert.equal(calls.attachments.writes.length, 1);
        assert.ok(calls.attachments.fileAccesses.length > 0);
      });
      it(`${tool} allows clear raw classification of an empty structured original, opt-in=${allowed}`, async () => {
        const { api, calls } = loadMessageTools({
          allowed, mime: { contentType: "message/rfc822", parts: [] }, raw: "Content-Type: text/plain\r\n\r\nvisible",
        });
        const result = await invoke(api, true);
        assert.equal(result.success, true, result.error);
        assert.equal(calls.sends.length, 1);
        assert.equal(calls.options.length, 1);
        assert.equal(calls.options[0].examineEncryptedParts, false);
        assert.equal(calls.streams, 1);
      });
    }
    for (const unreadable of [false, true]) {
      it(`${tool} rejects direct encrypted sends with pref off/unreadable=${unreadable}`, async () => {
        const { api, calls } = loadMessageTools({ mime: encryptedTrees[2], unreadable });
        const result = await invoke(api, true);
        assertOriginalRefused(calls, result);
      });
      it(`${tool} rejects armor from plaintext coercion with pref off/unreadable=${unreadable}`, async () => {
        const { api, calls } = loadMessageTools({
          mime: { parts: [], coerceBodyToPlaintext: () => "-----BEGIN PGP MESSAGE-----\nciphertext" }, unreadable,
        });
        const result = await invoke(api, true);
        assertOriginalRefused(calls, result);
      });
    }
    it(`${tool} refuses armor from plaintext coercion even after explicit opt-in`, async () => {
      const { api, calls } = loadMessageTools({
        mime: { parts: [], coerceBodyToPlaintext: () => "-----BEGIN PGP MESSAGE-----\nciphertext" }, allowed: true,
      });
      const result = await invoke(api, true);
      assertOriginalRefused(calls, result);
    });
    it(`${tool} refuses encrypted quoted content even after explicit opt-in`, async () => {
      const { api, calls } = loadMessageTools({ mime: encryptedTrees[2], allowed: true });
      const result = await invoke(api, true);
      assertOriginalRefused(calls, result);
    });
    it(`${tool} leaves the review path available without MIME extraction`, async () => {
      const { api, calls } = loadMessageTools({ mime: encryptedTrees[2] });
      const result = await invoke(api, false);
      assert.equal(result.success, true);
      assert.equal(calls.reviews, 1);
      assert.equal(calls.options.length, 0);
      assert.equal(calls.sends.length, 0);
    });
  }
});

describe("Encrypted message privacy for automatic reply drafts", () => {
  const invoke = api => api.replyToMessage("message-1", "folder", "intro", false, false,
    undefined, undefined, undefined, undefined, undefined, false, true);

  for (const encrypted of encryptedTrees) {
    for (const unreadable of [false, true]) {
      it(`refuses ${encrypted.label} before opening a draft window, pref unreadable=${unreadable}`, async () => {
        const mime = {
          ...encrypted,
          get allUserAttachments() { return assert.fail("attachment metadata must not be read"); },
        };
        const { api, calls } = loadMessageTools({ mime, unreadable });
        const result = await invoke(api);

        assert.match(result.error, /encrypted.*blocked/i);
        assert.equal(result.success, undefined);
        assert.equal(calls.reviews, 0);
        assert.equal(calls.sends.length, 0);
        assert.equal(calls.streams, 1);
        assert.equal(calls.options.length, 1);
        assert.equal(calls.options[0].examineEncryptedParts, false);
        assert.doesNotMatch(JSON.stringify(result), /decrypted body|protected subject|ciphertext/);
      });
    }

    it(`refuses a native ${encrypted.label} draft even after explicit opt-in`, async () => {
      const { api, calls } = loadMessageTools({ mime: encrypted, allowed: true });
      const result = await invoke(api);

      assert.match(result.error, /Automatic drafting.*blocked/);
      assert.equal(result.success, undefined);
      assert.equal(calls.reviews, 0);
      assert.equal(calls.sends.length, 0);
      assert.equal(calls.options.length, 1);
      assert.equal(calls.options[0].examineEncryptedParts, false);
    });
  }

  for (const unreadable of [false, true]) {
    it(`refuses armor produced by plaintext coercion, pref unreadable=${unreadable}`, async () => {
      const { api, calls } = loadMessageTools({
        mime: { parts: [], coerceBodyToPlaintext: () => "-----BEGIN PGP MESSAGE-----\nciphertext" },
        unreadable,
      });
      const result = await invoke(api);

      assert.match(result.error, /encrypted.*blocked/i);
      assert.equal(calls.reviews, 0);
      assert.equal(calls.sends.length, 0);
      assert.equal(calls.options[0].examineEncryptedParts, false);
      assert.doesNotMatch(JSON.stringify(result), /ciphertext|protected subject/);
    });
  }

  it("refuses drafting armor from plaintext coercion even after explicit opt-in", async () => {
    const { api, calls } = loadMessageTools({
      mime: { parts: [], coerceBodyToPlaintext: () => "-----BEGIN PGP MESSAGE-----\nciphertext" },
      allowed: true,
    });
    const result = await invoke(api);

    assert.match(result.error, /Automatic drafting.*blocked/);
    assert.equal(result.success, undefined);
    assert.equal(calls.reviews, 0);
    assert.equal(calls.sends.length, 0);
    assert.equal(calls.options[0].examineEncryptedParts, false);
  });

  for (const [label, options] of [
    ["null MIME", { mime: null }],
    ["unclassified content type", { mime: { contentType: "text plain", body: "unclassified private body" } }],
    ["parser failure", { mimeError: new Error("MIME parser failed") }],
    ["classification failure", { mime: { get contentType() { throw new Error("MIME classification failed"); } } }],
  ]) {
    for (const unreadable of [false, true]) {
      it(`refuses ${label} before native composition, pref unreadable=${unreadable}`, async () => {
        const { api, calls } = loadMessageTools({ ...options, unreadable });
        const result = await invoke(api);

        assert.equal(typeof result.error, "string");
        assert.equal(result.success, undefined);
        assert.equal(calls.reviews, 0);
        assert.equal(calls.sends.length, 0);
        assert.equal(calls.streams, 1);
        assert.equal(calls.options.length, 1);
        assert.equal(calls.options[0].examineEncryptedParts, false);
        assert.doesNotMatch(JSON.stringify(result), /unclassified private body|protected subject/);
      });
    }
  }

  for (const unreadable of [false, true]) {
    for (const mime of [
      { contentType: "text/plain", body: "visible" },
      { contentType: "multipart/signed", parts: [{ contentType: "text/plain", body: "signed content" }] },
    ]) {
      it(`permits native drafts of ${mime.contentType}, pref unreadable=${unreadable}`, async () => {
        const { api, calls } = loadMessageTools({ mime, unreadable });
        const result = await invoke(api);

        assert.equal(result.success, true);
        assert.equal(result.message, "Reply saved as draft");
        assert.equal(calls.reviews, 1);
        assert.equal(calls.sends.length, 0);
        assert.equal(calls.options.length, 1);
        assert.equal(calls.options[0].examineEncryptedParts, false);
      });
    }
  }

  it("keeps ordinary review available without privacy or MIME extraction", async () => {
    const { api, calls } = loadMessageTools({
      mime: encryptedTrees[2], unreadable: true, mimeError: new Error("MIME must not be requested for review"),
    });
    const result = await api.replyToMessage("message-1", "folder", "intro", false, false,
      undefined, undefined, undefined, undefined, undefined, false, false);

    assert.equal(result.success, true);
    assert.equal(result.message, "Reply window opened");
    assert.equal(calls.reviews, 1);
    assert.equal(calls.sends.length, 0);
    assert.equal(calls.options.length, 0);
  });
});

describe("Standalone inline PGP armor", () => {
  const header = "-----BEGIN PGP MESSAGE-----";
  for (const [label, body, encrypted] of [
    ["LF armor", `intro\n${header}\nciphertext`, true],
    ["CRLF armor with surrounding whitespace", `intro\r\n \t${header}\t \r\nciphertext`, true],
    ["header at end of input", `\t${header} `, true],
    ["prose delimiter", `The delimiter is ${header}`, false],
    ["quoted delimiter in a sentence", `The delimiter "${header}" starts a message.`, false],
    ["trailing prose", `${header} is the delimiter.`, false],
    ["quoted line", `> ${header}`, false],
  ]) {
    for (const coerced of [false, true]) {
      it(`${label} in parsed reads and direct sends, coerced=${coerced}`, async () => {
        const mime = coerced ? { parts: [], coerceBodyToPlaintext: () => body } : { contentType: "text/plain", body };
        const { api, calls } = loadMessageTools({ mime });
        const result = await api.getMessage("message-1", "folder", false, "text");
        assert.equal(result.encryptedContentWithheld === true, encrypted);
        if (!encrypted) assert.equal(result.body, body);
        const sends = [
          await api.replyToMessage("message-1", "folder", "intro", false, false, undefined, undefined, undefined, undefined, undefined, true),
          await api.forwardMessage("message-1", "folder", "to@example.test", "intro", false, undefined, undefined, undefined, undefined, true),
        ];
        for (const send of sends) {
          if (encrypted) assertOriginalRefused(calls, send);
          else assert.equal(send.success, true);
        }
        assert.equal(calls.sends.length, encrypted ? 0 : 2);
        for (const send of calls.sends) assert.ok(send.body.includes(body));
      });
    }
    for (const [encoding, raw] of [
      ["plain", `Content-Type: text/plain\r\n\r\n${body}`],
      ["base64", `Content-Type: text/plain\nContent-Transfer-Encoding: base64\n\n${Buffer.from(body).toString("base64")}`],
      ["quoted-printable", `Content-Type: text/plain\nContent-Transfer-Encoding: quoted-printable\n\n${body.replace(/-/g, "=2D")}`],
      ["UTF-16", `Content-Type: text/plain; charset=utf-16le\nContent-Transfer-Encoding: base64\n\n${Buffer.from(body, "utf16le").toString("base64")}`],
    ]) {
      for (const rawSource of [false, true]) {
        it(`${label} through ${encoding}, rawSource=${rawSource}`, async () => {
          const { api } = loadMessageTools({ raw, mime: { parts: [] } });
          const result = await api.getMessage("message-1", "folder", false, "text", rawSource);
          assert.equal(result.encryptedContentWithheld === true, encrypted);
          if (!encrypted) assert.equal(rawSource ? result.rawSource : result.body, rawSource ? raw : body);
        });
      }
    }
  }
});

describe("Mixed MIME primary bodies in reads and direct quoting", () => {
  for (const mainIsHtml of [false, true]) {
    it(`keeps the primary body ahead of an opposite-format footer, HTML main=${mainIsHtml}`, async () => {
      const mainType = mainIsHtml ? "text/html" : "text/plain";
      const html = "<p>Main discussion</p><p>Continued discussion</p>";
      const mime = { contentType: "multipart/mixed", parts: [
        { contentType: mainType, body: mainIsHtml ? "<p>Main discussion</p>" : "Main discussion\n" },
        { contentType: mainIsHtml ? "text/plain" : "text/html", body: mainIsHtml ? "Unsubscribe footer" : "<p>Unsubscribe footer</p>" },
        { contentType: mainType, body: mainIsHtml ? "<p>Continued discussion</p>" : "Continued discussion" },
      ] };
      const { api, calls } = mainIsHtml
        ? loadHtmlFixture(html, () => documentTree([
          elementNode("p", [textNode("Main discussion")]), elementNode("p", [textNode("Continued discussion")]),
        ]), { mime })
        : loadMessageTools({ mime });

      for (const format of ["text", "markdown", "html"]) {
        const result = await api.getMessage("message-1", "folder", false, format);
        assert.match(result.body, /Main discussion/);
        assert.match(result.body, /Continued discussion/);
        assert.doesNotMatch(result.body, /Unsubscribe footer/);
        assert.equal(result.bodyIsHtml, mainIsHtml && format === "html");
      }
      const reply = await api.replyToMessage("message-1", "folder", "intro", false, false,
        undefined, undefined, undefined, undefined, undefined, true);
      const forward = await api.forwardMessage("message-1", "folder", "to@example.test", "intro", false,
        undefined, undefined, undefined, undefined, true);
      assert.equal(reply.success, true);
      assert.equal(forward.success, true);
      assert.equal(calls.sends.length, 2);
      for (const fields of calls.sends) {
        assert.match(fields.body, /Main discussion/);
        assert.match(fields.body, /Continued discussion/);
        assert.doesNotMatch(fields.body, /Unsubscribe footer|<p>/);
      }
      assert.equal(calls.streams, 5);
    });
  }
});

describe("Encryption classification of joined MIME bodies", () => {
  const armor = "-----BEGIN PGP MESSAGE-----";
  const html = `<p>${armor}</p><p>ciphertext</p>`;

  for (const isHtml of [false, true]) {
    for (const allowed of [false, true]) {
      it(`classifies armor split around an attachment, HTML=${isHtml}, opt-in=${allowed}`, async () => {
        const mime = { contentType: "multipart/mixed", parts: [
          { contentType: isHtml ? "text/html" : "text/plain", partName: "1.1", body: `${isHtml ? "<p>" : ""}-----BEGIN PGP ` },
          { contentType: "application/pdf", partName: "1.2" },
          { contentType: isHtml ? "text/html" : "text/plain", partName: "1.3", body: `MESSAGE-----${isHtml ? "</p><p>ciphertext</p>" : "\nciphertext"}` },
        ], allUserAttachments: [{ partName: "1.2", name: "report.pdf", url: "mailbox-message://local/Inbox?number=1&part=1.2" }] };
        const { api, calls } = isHtml
          ? loadHtmlFixture(html, () => documentTree([
            elementNode("p", [textNode(armor)]), elementNode("p", [textNode("ciphertext")]),
          ]), { mime, allowed })
          : loadMessageTools({ mime, allowed });
        assert.equal(api.classifyMimeMessageEncryption(mime), "clear", "individual fragments have no complete armor marker");
        for (const format of ["text", "markdown", "html"]) {
          const result = await api.getMessage("message-1", "folder", false, format);
          assert.equal(result.encryptedContentWithheld === true, !allowed);
          if (allowed) assert.match(result.body, /-----BEGIN PGP MESSAGE-----/);
          else {
            assert.equal(result.subject, "[Encrypted message]");
            assert.equal(result.attachments.length, 0);
            assert.doesNotMatch(JSON.stringify(result), /ciphertext|protected subject/);
          }
        }
        const results = [
          await api.replyToMessage("message-1", "folder", "intro", false, false, undefined, undefined, undefined, undefined, undefined, true),
          await api.forwardMessage("message-1", "folder", "to@example.test", "intro", false, undefined, undefined, undefined, undefined, true),
        ];
        for (const result of results) {
          assertOriginalRefused(calls, result);
        }
        assert.equal(calls.sends.length, 0);
        assert.equal(calls.streams, 5);
      });
    }
  }

  it("classifies the full joined HTML when only the total exceeds the presentation cap", async () => {
    const cap = 2 * 1024 * 1024;
    const first = "<p>" + " ".repeat(cap - 20) + "-----BEGIN PGP ";
    const second = "MESSAGE-----</p><p>ciphertext</p>";
    const joined = first + second;
    const mime = { contentType: "multipart/mixed", parts: [
      { contentType: "text/html", body: first },
      { contentType: "text/html", body: second },
    ] };
    assert.ok(first.length < cap && second.length < cap && joined.length > cap);
    const { api, calls } = loadHtmlFixture(joined.slice(0, cap), () => documentTree([elementNode("p", [textNode("prefix")])]), { mime });
    assert.equal(api.classifyMimeMessageEncryption(mime), "clear");
    for (const format of ["text", "markdown", "html"]) {
      assert.equal((await api.getMessage("message-1", "folder", false, format)).encryptedContentWithheld, true);
    }
    assert.equal(calls.streams, 3);
  });

  it("withholds joined raw HTML if its visible text cannot be classified", async () => {
    const { api, calls } = loadMessageTools({ mime: {
      contentType: "multipart/mixed", parts: [
        { contentType: "text/html", body: "<p>-----BEGIN PGP " },
        { contentType: "text/html", body: "MESSAGE-----</p><p>ciphertext</p>" },
      ],
    }, DOMParser: class { parseFromString() { throw Error("parser failed"); } } });
    assert.equal((await api.getMessage("message-1", "folder", false, "html")).encryptedContentWithheld, true);
    assert.equal(calls.streams, 1);
  });
});

describe("Oversized HTML encryption classification", () => {
  const cap = 2 * 1024 * 1024;
  const header = "-----BEGIN PGP MESSAGE-----";
  const bodies = [
    ["armor after the cut", `<p>${" ".repeat(cap)}the delimiter is ${header}</p>`],
    ["armor split by the cut", `<p>${" ".repeat(cap - 3 - 12)}${header} is the delimiter.</p>`],
  ];
  const directSends = api => Promise.all([
    api.replyToMessage("message-1", "folder", "intro", false, false, undefined, undefined, undefined, undefined, undefined, true),
    api.forwardMessage("message-1", "folder", "to@example.test", "intro", false, undefined, undefined, undefined, undefined, true),
  ]);

  for (const [label, html] of bodies) {
    it(`withholds ${label} in structured reads and direct sends before exposing content`, async () => {
      const mime = {
        contentType: "multipart/mixed", parts: [{ contentType: "text/html", body: html }],
        get allUserAttachments() { return assert.fail("must classify before attachments"); },
      };
      const { api, calls } = loadHtmlFixture(html.slice(0, cap), () => documentTree([elementNode("p", [textNode("visible")])]), { mime });
      assert.equal(api.hasInlinePgpArmor(html), false, "the full raw source has no standalone armor line");
      for (const format of ["text", "markdown", "html"]) {
        const result = await api.getMessage("message-1", "folder", true, format);
        assert.equal(result.encryptedContentWithheld, true);
        assert.equal(result.subject, "[Encrypted message]");
      }
      assert.equal((await api.getMessage("message-1", "folder", false, "html", true)).rawSource, "Content-Type: text/plain\r\n\r\nvisible");
      for (const send of await directSends(api)) assertOriginalRefused(calls, send);
      assert.equal(calls.streams, 6);
      assert.equal(calls.sends.length, 0);
    });

    for (const [encoding, raw] of [
      ["plain", `Content-Type: text/html\r\n\r\n${html}`],
      ["base64", `Content-Type: text/html\nContent-Transfer-Encoding: base64\n\n${Buffer.from(html).toString("base64")}`],
      ["UTF-16", `Content-Type: text/html; charset=utf-16le\nContent-Transfer-Encoding: base64\n\n${Buffer.from(html, "utf16le").toString("base64")}`],
    ]) {
      it(`withholds ${label} through ${encoding} raw MIME recovery and raw output`, async () => {
        const { api, calls } = loadMessageTools({ raw, mime: { parts: [] } });
        for (const rawSource of [false, true]) {
          const result = await api.getMessage("message-1", "folder", false, "markdown", rawSource);
          assert.equal(result.encryptedContentWithheld, true);
          assert.equal(result.rawSource, undefined);
        }
        assert.equal(calls.streams, 2);
      });
    }
  }

  it("keeps presentation from inventing armor and withholds raw HTML that cannot be fully verified", async () => {
    const encodedHeader = "&#45;----BEGIN PGP MESSAGE-----";
    const prefix = `<p>${" ".repeat(cap - 3 - encodedHeader.length)}${encodedHeader}`;
    const html = `${prefix} is the delimiter.</p>`;
    assert.equal(prefix.length, cap);
    assert.equal(html.includes(header), false);
    const { api, calls } = loadHtmlFixture(prefix, () => documentTree([elementNode("p", [textNode(header)])]), {
      mime: { contentType: "text/html", body: html },
    });
    assert.equal(api.hasInlinePgpArmor(api.stripHtml(html)), true, "the cutoff does create a standalone line in presentation");
    for (const format of ["text", "markdown", "html"]) {
      const result = await api.getMessage("message-1", "folder", false, format);
      assert.notEqual(result.encryptedContentWithheld, true);
      assert.equal(result.body, format === "html" ? html : `${header}\n\n[Message body truncated at 2 MiB]`);
    }
    for (const send of await directSends(api)) assertOriginalRefused(calls, send);
    assert.equal(calls.sends.length, 0);

    const raw = `Content-Type: text/html\n\n${html}`;
    const fallback = loadHtmlFixture(prefix, () => documentTree([elementNode("p", [textNode(header)])]), { raw, mime: { parts: [] } }).api;
    assert.equal(fallback.classifyRawMessageEncryption(raw), "clear");
    const recovered = await fallback.getMessage("message-1", "folder", false, "text");
    assert.equal(recovered.encryptedContentWithheld, true);
    assert.match(recovered.body, /could not be determined/);
  });

  it("keeps standalone-line detection for HTML at and below the cap", async () => {
    for (const [text, encrypted] of [[header, true], [`${header} is the delimiter.`, false]]) {
      const html = `<p>${text}</p>`;
      const { api, calls } = loadHtmlFixture(html, () => documentTree([elementNode("p", [textNode(text)])]), {
        mime: { contentType: "text/html", body: html },
      });
      for (const format of ["text", "markdown"]) {
        const result = await api.getMessage("message-1", "folder", false, format);
        assert.equal(result.encryptedContentWithheld === true, encrypted);
      }
      for (const send of await directSends(api)) {
        if (encrypted) assertOriginalRefused(calls, send);
        else assert.equal(send.success, true);
      }
      assert.equal(calls.sends.length, encrypted ? 0 : 2);
      const exactCap = html + " ".repeat(cap - html.length);
      assert.equal(api.hasInlinePgpArmor(text, exactCap), encrypted);
    }
  });

  it("measures the encryption cap in UTF-8 bytes and honors explicit encrypted access", async () => {
    const html = `<p>${"é".repeat(cap / 2)}the delimiter is ${header}</p>`;
    assert.ok(html.length < cap);
    const { api } = loadMessageTools({ mime: { contentType: "text/html", body: html } });
    assert.equal(api.classifyMimeMessageEncryption({ contentType: "text/html", body: html }), "encrypted");
    const allowed = loadMessageTools({ allowed: true, mime: { contentType: "text/html", body: html } }).api;
    const result = await allowed.getMessage("message-1", "folder", false, "html");
    assert.notEqual(result.encryptedContentWithheld, true);
    assert.equal(result.body, html);
  });
});

describe("Encryption classification before raw output or body fallback", () => {
  const armor = "-----BEGIN PGP MESSAGE-----\nciphertext\n-----END PGP MESSAGE-----";
  const plain = "Content-Type: text/plain\n\nvisible";
  const encoded = `Content-Type: text/plain\nContent-Transfer-Encoding: base64\n\n${Buffer.from(armor).toString("base64")}`;
  const encrypted = [
    ["plain armor", `Content-Type: text/plain\n\n${armor}`],
    ["base64 armor", encoded],
    ["quoted-printable armor", `Content-Type: text/plain\nContent-Transfer-Encoding: quoted-printable\n\n${armor.replace(/-/g, "=2D")}`],
    ["UTF-16 armor", `Content-Type: text/plain; charset=utf-16le\nContent-Transfer-Encoding: base64\n\n${Buffer.from(armor, "utf16le").toString("base64")}`],
    ["nested armor", `Content-Type: multipart/mixed; boundary=b\n\n--b\n${plain}\n--b\n${encoded}\n--b--\n`],
    ["attached message armor", `Content-Type: message/rfc822\n\n${encoded}`],
  ];
  const unknown = [
    ["no header/body split", "incomplete message"],
    ["missing boundary", "Content-Type: multipart/mixed\n\nbody"],
    ["unterminated multipart", `Content-Type: multipart/mixed; boundary=b\n\n--b\n${plain}`],
    ["malformed child", `Content-Type: multipart/mixed; boundary=b\n\n--b\n${plain}\n--b\nno headers\n--b--\n`],
    ["unknown transfer encoding", "Content-Type: text/plain\nContent-Transfer-Encoding: opaque\n\nbytes"],
    ["malformed base64", "Content-Type: text/plain\nContent-Transfer-Encoding: base64\n\n%%bad%%"],
    ["truncated base64", "Content-Type: text/plain\nContent-Transfer-Encoding: base64\n\nZ"],
    ["malformed quoted-printable", "Content-Type: text/plain\nContent-Transfer-Encoding: quoted-printable\n\n=Z0"],
    ["unknown charset", "Content-Type: text/plain; charset=x-unknown\n\nbytes"],
    ["ambiguous charset", `Content-Type: text/plain; charset=utf-8; charset=utf-16le\nContent-Transfer-Encoding: base64\n\n${Buffer.from(armor, "utf16le").toString("base64")}`],
    ["extended charset", `Content-Type: text/plain; charset*=utf-8''utf-16le\nContent-Transfer-Encoding: base64\n\n${Buffer.from(armor, "utf16le").toString("base64")}`],
    ["empty charset", 'Content-Type: text/plain; charset=""\n\nbytes'],
    ["ambiguous type", "Content-Type: text/plain\nContent-Type: application/pkcs7-mime\n\nbytes"],
    ["unparsed content-type comment", "Content-Type: application/pkcs7-mime (comment); smime-type=enveloped-data\nContent-Transfer-Encoding: base64\n\nY21z"],
    ["missing type parameter separator", "Content-Type: application/pkcs7-mime smime-type=enveloped-data\n\nbytes"],
    ["empty type", "Content-Type:\n\nbytes"],
    ["ambiguous extended smime type", "Content-Type: application/pkcs7-mime; smime-type=signed-data; smime-type*=utf-8''enveloped-data\n\nbytes"],
  ];
  for (const [label, raw] of [...encrypted, ...unknown]) {
    // Read safety can recover these ordinary malformed messages. S4 keeps
    // refusing the same fixtures under its stricter send classification below.
    const recoverable = ["missing boundary", "unterminated multipart", "malformed child", "malformed base64",
      "truncated base64", "malformed quoted-printable"].includes(label);
    for (const rawSource of [false, true]) {
      it(`${recoverable ? "reads" : "withholds"} ${label} with empty structured MIME, rawSource=${rawSource}`, async () => {
        const { api, calls } = loadMessageTools({ raw, mime: {
          contentType: "message/rfc822", parts: [],
          get allUserAttachments() { return recoverable ? [] : assert.fail("must classify before attachments"); },
        } });
        const result = await api.getMessage("message-1", "folder", true, "text", rawSource, true);
        if (recoverable) {
          assert.equal(result.error, undefined);
          assert.equal(result.encryptedContentWithheld, undefined);
          if (rawSource) assert.equal(result.rawSource, raw);
          assert.equal(calls.streams, 1);
          assert.equal(calls.options.length, rawSource ? 0 : 1);
          return;
        }
        assert.equal(result.encryptedContentWithheld, true);
        assert.equal(result.attachments.length, 0);
        assert.equal(result.rawSource, undefined);
        assert.match(result.body, /withheld/);
        assert.doesNotMatch(JSON.stringify(result), /ciphertext|protected subject/);
        assert.equal(calls.streams, 1);
        assert.equal(calls.options.length, 0);
      });
    }
    for (const allowed of [false, true]) {
      it(`direct replies and forwards refuse ${label} with empty structured MIME, opt-in=${allowed}`, async () => {
        const { api, calls } = loadMessageTools({
          allowed, raw, mime: {
            contentType: "message/rfc822", parts: [],
            get allUserAttachments() { return assert.fail("Refused raw content must not expose attachment metadata"); },
          },
        });
        const results = [
          await api.replyToMessage("message-1", "folder", "intro", false, false, undefined, undefined, undefined, undefined, undefined, true),
          await api.forwardMessage("message-1", "folder", "to@example.test", "intro", false, undefined, undefined, undefined, undefined, true),
        ];
        for (const result of results) assertOriginalRefused(calls, result);
        assert.equal(calls.streams, 2);
        assert.equal(calls.options.length, 0);
        assert.equal(calls.imports, 0);
      });
    }
  }

  it("retains normal complete raw and fallback bodies without opting in", async () => {
    const { api } = loadMessageTools({ raw: plain, mime: { parts: [] } });
    assert.equal((await api.getMessage("message-1", "folder", false, "text", false)).body, "visible");
    assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).rawSource, plain);
  });

  for (const newline of ["\n", "\r\n"]) {
    for (const withHeaders of [false, true]) {
      for (const subtype of ["mixed", "alternative"]) {
        for (const rawSource of [false, true]) {
          it(`reads ordinary multipart/${subtype}, headers=${withHeaders}, newline=${JSON.stringify(newline)}, rawSource=${rawSource}`, async () => {
            const child = `${withHeaders ? `Content-Type: text/plain${newline}` : ""}${newline}Hello`;
            const raw = [
              `Content-Type: multipart/${subtype}; boundary=b`, "", "--b", child,
              "--b", "Content-Type: text/html", "", "<p>Hello</p>", "--b--", "",
            ].join(newline);
            // The raw classifier also checks the HTML part's visible text.
            const { api } = loadHtmlFixture("<p>Hello</p>", () => documentTree([elementNode("p", [textNode("Hello")])]), { raw, mime: { parts: [] } });
            const result = await api.getMessage("message-1", "folder", false, "text", rawSource);
            assert.equal(result.error, undefined);
            assert.notEqual(result.encryptedContentWithheld, true);
            assert.equal(rawSource ? result.rawSource : result.body, rawSource ? raw : "Hello");
          });
        }
      }
    }
    it(`preserves body-leading blank lines after empty headers, newline=${JSON.stringify(newline)}`, () => {
      const { api } = loadMessageTools();
      const split = api.findRawMimeHeaderBodySplit(`${newline}${newline}Hello`);
      assert.equal(split.header, "");
      assert.equal(split.body, `${newline}Hello`);
    });
  }

  it("allows raw source of deeply nested ordinary MIME within the read-safety limit", async () => {
    let raw = plain;
    for (let index = 0; index < 12; index++) raw = `Content-Type: multipart/mixed; boundary=b${index}\n\n--b${index}\n${raw}\n--b${index}--\n`;
    const { api } = loadMessageTools({ raw, mime: { parts: [] } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, raw);
  });

  it("returns a read error when the complete stream cannot be read", async () => {
    const { api } = loadMessageTools({ mime: { parts: [] }, streamError: new Error("unreadable") });
    const result = await api.getMessage("message-1", "folder", true, "text");
    assert.equal(result.error, "Failed to get message");
    assert.equal(result.encryptedContentWithheld, undefined);
  });

  it("withholds armor returned only by the MIME plaintext coercion", async () => {
    const { api } = loadMessageTools({ mime: { parts: [], coerceBodyToPlaintext: () => armor } });
    assert.equal((await api.getMessage("message-1", "folder")).encryptedContentWithheld, true);
  });

  it("allows inline armor in fallback and raw mode only after the opt-in", async () => {
    const raw = encrypted[0][1];
    const { api } = loadMessageTools({ raw, mime: { parts: [] }, allowed: true });
    assert.equal((await api.getMessage("message-1", "folder", false, "text")).body, armor);
    assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).rawSource, raw);
  });

  for (const type of [
    "signed-data", "enveloped-data", "authEnveloped-data",
    "", "signed-data; smime-type=enveloped-data", "invalid",
  ]) {
    for (const rawSource of [false, true]) {
      it(`withholds S/MIME smime-type=${type || "missing"}, rawSource=${rawSource}`, async () => {
        const raw = `Content-Type: application/pkcs7-mime${type ? `; smime-type=${type}` : ""}\nContent-Transfer-Encoding: base64\n\nY21z`;
        const { api } = loadMessageTools({ raw, mime: { contentType: "message/rfc822", parts: [] } });
        const result = await api.getMessage("message-1", "folder", false, "text", rawSource);
        assert.equal(result.encryptedContentWithheld, true);
        assert.equal(result.rawSource, undefined);
      });
    }
  }
});

// Each case pairs exact HTML input with its hand-built, browser-parsed DOM.
// Parsing/entity decoding/CSS normalization belong to Thunderbird; filtering
// and conversion still run from the production markers above.
function textNode(textContent) {
  return { nodeType: 3, textContent, childNodes: [] };
}

function commentNode(textContent) {
  return { nodeType: 8, textContent, childNodes: [] };
}

function elementNode(tagName, children = [], attributes = {}, parsedStyle = {}) {
  const node = {
    nodeType: 1,
    tagName: tagName.toUpperCase(),
    attributes,
    childNodes: [],
    // Keep the raw style attribute and its explicitly supplied CSSOM values.
    // No CSS parser or copy of the production hidden-element predicate lives here.
    style: { display: "", visibility: "", fontSize: "", opacity: "", ...parsedStyle },
    getAttribute(name) { return this.attributes[name.toLowerCase()] ?? null; },
    hasAttribute(name) { return Object.hasOwn(this.attributes, name.toLowerCase()); },
    get textContent() {
      return this.childNodes.filter(child => child.nodeType !== 8).map(child => child.textContent).join("");
    },
    get parentElement() { return this.parentNode?.nodeType === 1 ? this.parentNode : null; },
    remove() {
      if (!this.parentNode) return;
      this.parentNode.childNodes.splice(this.parentNode.childNodes.indexOf(this), 1);
      this.parentNode = null;
    },
  };
  if (node.tagName === "TEMPLATE") {
    node.content = {
      nodeType: 11,
      childNodes: children,
      get textContent() {
        return this.childNodes.filter(child => child.nodeType !== 8).map(child => child.textContent).join("");
      },
    };
  } else {
    node.childNodes = children;
  }
  for (const child of children) child.parentNode = node.content || node;
  return node;
}

function htmlDocument(body) {
  return `<!doctype html><html><head><title>head-secret</title></head><body>${body}</body></html>`;
}

function documentTree(body, head = [elementNode("title", [textNode("head-secret")])]) {
  return elementNode("html", [elementNode("head", head), elementNode("body", body)]);
}

function loadHtmlFixture(html, buildTree, options = {}) {
  return loadMessageTools({ ...options, DOMParser: class {
    parseFromString(input, mimeType) {
      assert.equal(input, html, "fixture must match the HTML passed by production");
      assert.equal(mimeType, "text/html");
      const root = buildTree(); // Fresh nodes: production removes hidden subtrees.
      return {
        documentElement: root,
        body: root.childNodes.find(node => node.tagName === "BODY"),
        querySelectorAll(selector) {
          assert.equal(selector, "*");
          const elements = [];
          function visit(node) {
            if (node.nodeType !== 1) return;
            elements.push(node);
            node.childNodes.forEach(visit); // Template content is a separate fragment.
          }
          visit(root);
          return elements;
        },
      };
    }
  } });
}

describe("Message text conversion", () => {
  it("loads encoding helpers only through the production Experiment import", () => {
    const sandbox = {};
    const globals = { DOMParser: class {}, atob, btoa, TextDecoder };
    sandbox.Cu = { importGlobalProperties(names) {
      for (const name of names) sandbox[name] = globals[name];
    } };
    vm.createContext(sandbox);
    assert.equal(vm.runInContext("typeof atob + ',' + typeof btoa + ',' + typeof TextDecoder", sandbox), "undefined,undefined,undefined");
    vm.runInContext([
      snippet("EXPERIMENT GLOBAL IMPORTS"),
      snippet("INLINE IMAGE CONTENT HELPERS"), snippet("RAW MIME PARSING HELPERS"),
    ].join("\n"), sandbox);
    assert.equal(sandbox.encodeByteStringToBase64("\x00\x80\xff"), "AID/");
    const mime = "Content-Type: text/plain; charset=utf-8\nContent-Transfer-Encoding: base64\n\nw6k=";
    assert.equal(sandbox.extractBodyPartFromRawMime(mime, "text").text, "é");
    assert.equal(sandbox.decodeRawMimeExtendedParameter("utf-8''caf%C3%A9.txt"), "café.txt");
  });

  it("warns on failed global imports and continues loading with HTML failing closed", () => {
    const failure = new Error("global import unavailable");
    const warnings = [];
    const sandbox = {
      Cu: { importGlobalProperties() { throw failure; } },
      console: { warn: (...args) => warnings.push(args) },
    };
    vm.runInNewContext([
      snippet("EXPERIMENT GLOBAL IMPORTS"),
      snippet("MCP TEXT SANITIZATION"), snippet("MESSAGE TEXT CONVERSION"),
    ].join("\n"), sandbox);
    assert.equal(warnings.length, 1);
    assert.match(warnings[0][0], /failed to import Experiment globals/);
    assert.equal(warnings[0][1], failure);
    for (const convert of [sandbox.stripHtml, sandbox.htmlToMarkdown]) {
      assert.equal(convert("<p>private</p>"), "[HTML content withheld: safe HTML parser unavailable.]");
    }
    assert.equal(sandbox.extractFormattedBody({ contentType: "text/plain", body: "ordinary message" }).body, "ordinary message");
  });

  for (const [href, expected] of [
    ["HTTPS://example.test/", "[label](HTTPS://example.test/)"],
    [" \x01\thTt\nps://example.test/", "[label](hTtps://example.test/)"],
    ["MAILTO:reader@example.test", "[label](MAILTO:reader@example.test)"],
    ["javascript:alert(1)", "label"], [" \x01JaVa\nScRiPt:alert(1)", "label"],
    ["file:///secret", "label"], ["data:text/html,secret", "label"],
    ["vbscript:secret", "label"], ["cid:secret", "label"],
    ["/relative", "label"], ["//remote.test", "label"], ["#fragment", "label"], ["", "label"],
  ]) {
    it(`filters Markdown link destination ${JSON.stringify(href)}`, () => {
      const html = `<a href="${href}">label</a>`;
      const { api } = loadHtmlFixture(html, () => documentTree([elementNode("a", [textNode("label")], { href })]));
      assert.equal(api.htmlToMarkdown(html), expected);
    });
  }
  it("escapes literal Markdown, image alt text, and destination delimiters", () => {
    const html = '<p>fixture with hostile text and attributes</p>';
    const { api } = loadHtmlFixture(html, () => documentTree([
      elementNode("a", [textNode("label")], { href: "https://safe.test/) ![pixel](https://tracker.test/p)" }),
      elementNode("img", [], { src: "https://image.test/secret", srcset: "https://image.test/secret2 2x", alt: "![alt](https://alt.test/)" }),
      textNode('<img src="https://literal.test/">'),
      elementNode("img", [], { src: "cid:secret", width: "1", height: "1" }),
    ]));
    const result = api.htmlToMarkdown(html);
    assert.match(result, /safe\.test\/%29%20!%5bpixel%5d%28https:\/\/tracker\.test\/p%29/);
    assert.ok(result.includes("\\!\\[alt\\](https://alt.test/)"));
    assert.ok(result.includes('\\<img src="https://literal.test/"\\>'));
    assert.doesNotMatch(result, /image\.test|cid:secret|(?<!\\)!\[/);
  });
  it("escapes a literal bang at generated-link boundaries across comments and elements", () => {
    const url = "https://tracker.test/p";
    const linkHtml = `<a href="${url}">x</a>`;
    const link = () => elementNode("a", [textNode("x")], { href: url });
    for (const [html, children, expected] of [
      ["!" + linkHtml, () => [textNode("!"), link()], `\\![x](${url})`],
      ["!<!--gap-->" + linkHtml, () => [textNode("!"), commentNode("gap"), link()], `\\![x](${url})`],
      ["!<span></span>" + linkHtml, () => [textNode("!"), elementNode("span"), link()], `\\![x](${url})`],
      ["!<span>" + linkHtml + "</span>", () => [textNode("!"), elementNode("span", [link()])], `\\![x](${url})`],
      ["<span>!</span>" + linkHtml, () => [elementNode("span", [textNode("!")]), link()], `\\![x](${url})`],
      ["!&#x200b;" + linkHtml, () => [textNode("!\u200b"), link()], `\\![x](${url})`],
      ["\\!" + linkHtml, () => [textNode("\\"), textNode("!"), link()], `\\\\\\![x](${url})`],
      ['<img alt="!">' + linkHtml, () => [elementNode("img", [], { alt: "!" }), link()], `\\![x](${url})`],
      ['!<img src="https://image.test/p" alt="Illustration">', () => [textNode("!"), elementNode("img", [], { src: "https://image.test/p", alt: "Illustration" })], "!Illustration"],
      ['!<img alt="[x](https://tracker.test/p)">', () => [textNode("!"), elementNode("img", [], { alt: `[x](${url})` })], `!\\[x\\](${url})`],
      ["Thanks!", () => [textNode("Thanks!")], "Thanks!"],
    ]) {
      const { api } = loadHtmlFixture(html, () => documentTree(children()));
      assert.equal(api.htmlToMarkdown(html), expected, html);
    }
  });
  it("uses code fences longer than embedded backtick runs, after invisible-character removal", () => {
    const html = "<pre>fixture</pre>";
    const value = "`\u200b``\n![pixel](https://tracker.test/)";
    const { api } = loadHtmlFixture(html, () => documentTree([elementNode("pre", [textNode(value)])]));
    assert.equal(api.htmlToMarkdown(html), "````\n```\n![pixel](https://tracker.test/)\n````");
  });
  it("keeps code containing image syntax safe inside inline, list, and table contexts", () => {
    const html = "<p>code fixture</p>";
    const image = "![pixel](https://tracker.test/)";
    const { api } = loadHtmlFixture(html, () => documentTree([
      elementNode("code", [textNode(`first\n\n${image}\n\nlast`)]),
      elementNode("ul", [elementNode("li", [elementNode("pre", [textNode(image)])])]),
      elementNode("table", [elementNode("tr", [elementNode("td", [elementNode("pre", [textNode(image)])])])]),
    ]));
    const markdown = api.htmlToMarkdown(html);
    assert.ok(markdown.includes(`\` first ${image} last \``));
    assert.ok(markdown.includes(`- \`\`\`\n  ${image}\n  \`\`\``));
    assert.ok(markdown.endsWith("\\!\\[pixel\\](https://tracker.test/)"));
  });
  it("preserves ordinary punctuation in HTML text nodes and image alt text", () => {
    const html = "<p>ordinary text fixture</p>";
    const ordinary = "some_path C# 5*3 snake_case_name _ * # | ~ ~~~ Wow! ! spaced";
    const { api } = loadHtmlFixture(html, () => documentTree([
      elementNode("p", [textNode(ordinary)]),
      elementNode("img", [], { alt: ordinary, src: "https://image.test/photo" }),
    ]));
    assert.equal(api.htmlToMarkdown(html), ordinary + "\n\n" + ordinary);
  });
  it("escapes literal link, image, and HTML delimiters without changing other text", () => {
    const html = "<p>literal syntax fixture</p>";
    const { api } = loadHtmlFixture(html, () => documentTree([
      textNode(String.raw`some\path ![x](y) [a](javascript:b) <img src=x> ! hi!`),
    ]));
    assert.equal(api.htmlToMarkdown(html), String.raw`some\\path \!\[x\](y) \[a\](javascript:b) \<img src=x\> ! hi!`);
  });
  it("escapes literal text and alt backticks beside generated inline code", () => {
    for (const content of ["![p](https://tracker.test/p)", "[a](javascript:b)", "<img src=x>"]) {
      for (const prefix of [() => textNode("`"), () => elementNode("img", [], { alt: "`" })]) {
        const html = "<p>literal-backtick fixture</p>";
        const { api } = loadHtmlFixture(html, () => documentTree([
          elementNode("p", [prefix(), elementNode("code", [textNode(content)])]),
        ]));
        assert.equal(api.htmlToMarkdown(html), "\\` ` " + content + " `");
      }
    }
  });
  it("sizes inline and block code delimiters over all nested backtick runs", () => {
    const content = "![p](https://tracker.test/p)";
    for (const [tag, expected] of [
      ["code", "```` outer ```" + content + "`` tail ````"],
      ["pre", "````\nouter ```" + content + "`` tail\n````"],
    ]) {
      const html = "<p>nested-code fixture</p>";
      const { api } = loadHtmlFixture(html, () => documentTree([
        elementNode(tag, [textNode("outer `"), elementNode("code", [textNode("``" + content + "``")]), textNode(" tail")]),
      ]));
      assert.equal(api.htmlToMarkdown(html), expected);
    }
  });
  it("keeps adjacent inline code fences from merging into unmatched delimiters", () => {
    const html = "<p>adjacent code fixture</p>";
    const { api } = loadHtmlFixture(html, () => documentTree([
      elementNode("code", [textNode("foo")]),
      elementNode("code", [textNode("``![pixel](https://tracker.test/p)")]),
    ]));
    assert.equal(api.htmlToMarkdown(html), "` foo `  ``` ``![pixel](https://tracker.test/p) ```");
  });

  it("caps HTML input at 2 MiB of UTF-8 without splitting surrogate pairs", () => {
    const { api } = loadMessageTools();
    const limit = 2 * 1024 * 1024;
    for (const unit of ["a", "é", "中", "😀"]) {
      const size = Buffer.byteLength(unit);
      const fitting = unit.repeat(Math.floor(limit / size));
      assert.equal(api.truncateHtmlForParsing(fitting).truncated, false);
      const limited = api.truncateHtmlForParsing(fitting + unit);
      assert.equal(limited.truncated, true);
      assert.equal(limited.html, fitting);
      assert.ok(Buffer.byteLength(limited.html) <= limit);
    }
    const fitting = "x".repeat(limit - 1);
    assert.equal(api.truncateHtmlForParsing(fitting + "😀tail").html, fitting);
  });
  it("parses the capped input and keeps a hidden subtree spanning the cut hidden", () => {
    const limit = 2 * 1024 * 1024;
    const prefix = '<p>Before</p><div hidden>secret'.padEnd(limit, "x");
    const html = prefix + '</div><p>After the cut</p>';
    const { api } = loadHtmlFixture(prefix, () => documentTree([
      elementNode("p", [textNode("Before")]),
      elementNode("div", [textNode("secret")], { hidden: "" }),
    ]));
    for (const format of ["text", "markdown"]) {
      const result = api.extractFormattedBody({ contentType: "text/html", body: html }, format);
      assert.equal(result.body, "Before\n\n[Message body truncated at 2 MiB]");
      assert.equal(result.bodyIsHtml, false);
    }
    assert.equal(api.extractFormattedBody({ contentType: "text/html", body: html }, "html").body, html);
    api.escapeHtml = value => value.replace(/&/g, "&amp;").replace(/</g, "&lt;");
    api.formatBodyHtml = value => value;
    vm.runInContext(snippet("COMPOSE HTML FRAGMENT"), api);
    assert.equal(api.formatBodyFragmentHtml(html, true), "Before<br><br>[Message body truncated at 2 MiB]");
  });
  it("retains readable capped text with a truncation note and bounded output", () => {
    const limit = 2 * 1024 * 1024;
    const prefix = "<p>" + "x".repeat(limit - 3);
    const { api } = loadHtmlFixture(prefix, () => documentTree([elementNode("p", [textNode(prefix.slice(3))])]));
    for (const convert of [api.stripHtml, api.htmlToMarkdown]) {
      const result = convert(prefix + "after the cut</p>");
      assert.ok(result.startsWith("x".repeat(100)));
      assert.ok(result.endsWith("\n\n[Message body truncated at 2 MiB]"));
      assert.ok(Buffer.byteLength(result) < limit + 100);
      assert.doesNotMatch(result, /after the cut|withheld/);
    }
  });
  it("does not append a note for HTML exactly at the input limit", () => {
    const html = "<p>visible</p><!--".padEnd(2 * 1024 * 1024, "x");
    const { api } = loadHtmlFixture(html, () => documentTree([elementNode("p", [textNode("visible")])]));
    assert.equal(api.stripHtml(html), "visible");
    assert.equal(api.htmlToMarkdown(html), "visible");
  });
  it("still fails closed if parsing the truncated input genuinely fails", () => {
    const inputs = [];
    const { api } = loadMessageTools({ DOMParser: class {
      parseFromString(input) { inputs.push(input); throw Error("parser failure"); }
    } });
    const html = "<p>" + "x".repeat(2 * 1024 * 1024);
    for (const convert of [api.stripHtml, api.htmlToMarkdown]) {
      assert.equal(convert(html), "[HTML content withheld: safe HTML parser unavailable.]");
    }
    assert.ok(inputs.length > 0);
    assert.ok(inputs.every(input => Buffer.byteLength(input) === 2 * 1024 * 1024));
  });
  for (const newline of ["\r", "\r\n", "\n"]) {
    it(`normalizes ${JSON.stringify(newline)} before code, list, and blockquote formatting`, () => {
      const html = "<p>carriage-return fixture</p>";
      const pixel = "![pixel](https://tracker.test/p)";
      for (const [wrap, indent, firstLine] of [
        [pre => elementNode("ul", [elementNode("li", [pre])]), "  ", "- ```"],
        [pre => elementNode("blockquote", [pre]), "> ", "> ```"],
        [pre => elementNode("blockquote", [elementNode("ul", [elementNode("li", [pre])])]), ">   ", "> - ```"],
        [pre => elementNode("ul", [elementNode("li", [elementNode("blockquote", [pre])])]), "  > ", "- > ```"],
      ]) {
        const { api } = loadHtmlFixture(html, () => documentTree([
          wrap(elementNode("pre", [textNode(`safe${newline}${newline}${pixel}`)])),
        ]));
        assert.equal(api.htmlToMarkdown(html), `${firstLine}\n${indent}safe\n${indent}\n${indent}${pixel}\n${indent}\`\`\``);
      }
    });
  }
  for (const route of ["structured", "coerced", "raw MIME"]) {
    it(`escapes only image openers in plain-text Markdown via ${route}`, async () => {
      const body = '# Heading\r\n**bold**\r![pixel](https://tracker.test/p)\n![reference][id]\n' +
        '[link](javascript:example) <img src="https://literal.test/p">\n' +
        '\\![escaped](https://tracker.test/p) \\\\![unescaped](https://tracker.test/p)';
      const expected = '# Heading\r\n**bold**\r\\![pixel](https://tracker.test/p)\n\\![reference][id]\n' +
        '[link](javascript:example) <img src="https://literal.test/p">\n' +
        '\\![escaped](https://tracker.test/p) \\\\\\![unescaped](https://tracker.test/p)';
      const mime = route === "structured" ? { contentType: "text/plain", body }
        : route === "coerced" ? { parts: [], coerceBodyToPlaintext: () => body } : { parts: [] };
      const raw = "Content-Type: text/plain; charset=utf-8\r\nContent-Transfer-Encoding: base64\r\n\r\n" + Buffer.from(body).toString("base64");
      const { api } = loadMessageTools({ mime, raw });
      for (const format of ["markdown", "text", "html", undefined]) {
        const result = await api.getMessage("message-1", "folder", false, format);
        assert.equal(result.body, !format || format === "markdown" ? expected : body);
        assert.equal(result.bodyIsHtml, false);
      }
      assert.equal((await api.getMessage("message-1", "folder", false, "markdown", true)).rawSource, raw);
    });
  }
  it("escapes image openers after the existing invisible-character removal", async () => {
    const body = "!\u200b[pixel](https://tracker.test/p)";
    for (const mime of [{ contentType: "text/plain", body }, { parts: [] }]) {
      const raw = "Content-Type: text/plain; charset=utf-8\nContent-Transfer-Encoding: base64\n\n" + Buffer.from(body).toString("base64");
      const { api } = loadMessageTools({ mime, raw });
      const result = await api.getMessage("message-1", "folder", false, "markdown");
      assert.equal(result.body, "\\![pixel](https://tracker.test/p)");
    }
  });

  for (const tag of ["script", "style", "head"]) {
    for (const closing of [`</${tag}>`, `</${tag} >`, `</${tag.toUpperCase()}\t\n >`]) {
      it(`removes the DOM subtree corresponding to ${closing}`, () => {
        // A head belongs before body; a nested head tag in body would be ignored
        // by the browser's HTML parser instead of creating a head subtree.
        const html = tag === "head"
          ? `<!doctype html><html><head><title>hidden instructions</title>${closing}<body><p>visible</p></body></html>`
          : htmlDocument(`<${tag}>hidden instructions${closing}<p>visible</p>`);
        const { api } = loadHtmlFixture(html, () => tag === "head"
          ? documentTree([elementNode("p", [textNode("visible")])], [elementNode("title", [textNode("hidden instructions")])])
          : documentTree([elementNode(tag, [textNode("hidden instructions")]), elementNode("p", [textNode("visible")])]));
        assert.equal(api.stripHtml(html), "visible");
        assert.equal(api.htmlToMarkdown(html), "visible");
      });
    }
  }
  for (const [attribute, attributes, parsedStyle] of [
    ['hidden="false"', { hidden: "false" }, {}],
    ['style="display:none"', { style: "display:none" }, { display: "none" }],
    ['style="visibility:hidden"', { style: "visibility:hidden" }, { visibility: "hidden" }],
    ['style="font-size:0px"', { style: "font-size:0px" }, { fontSize: "0px" }],
    ['style="opacity:0"', { style: "opacity:0" }, { opacity: "0" }],
    ['style="DISPLAY: NONE"', { style: "DISPLAY: NONE" }, { display: "none" }],
    ['style="color:red; display : none !important;"', { style: "color:red; display : none !important;" }, { display: "none" }],
    ['style="visibility: HIDDEN !important"', { style: "visibility: HIDDEN !important" }, { visibility: "hidden" }],
    ['style="font-size:0.0em"', { style: "font-size:0.0em" }, { fontSize: "0em" }],
    ['style="opacity:0.0"', { style: "opacity:0.0" }, { opacity: "0" }],
  ]) {
    for (const format of ["text", "markdown"]) {
      it(`drops nested ${attribute} subtree in ${format}, including inside pre/code`, () => {
        const body = htmlDocument(`<script>script-secret</script><style>style-secret</style><pre>visible<code><span ${attribute}><b>hidden-secret</b></span></code>tail</pre>`);
        const { api } = loadHtmlFixture(body, () => documentTree([
          elementNode("script", [textNode("script-secret")]),
          elementNode("style", [textNode("style-secret")]),
          elementNode("pre", [textNode("visible"), elementNode("code", [
            elementNode("span", [elementNode("b", [textNode("hidden-secret")])], attributes, parsedStyle),
          ]), textNode("tail")]),
        ]));
        const result = api.extractFormattedBody({ contentType: "text/html", body, coerceBodyToPlaintext() { assert.fail("must filter HTML before coercion"); } }, format);
        assert.equal(result.body, format === "text" ? "visibletail" : "```\nvisibletail\n```");
        assert.doesNotMatch(result.body, /secret/);
        assert.equal(result.bodyIsHtml, false);
      });
    }
  }
  for (const [hidden, buildHidden] of [
    ['<template><p>hidden-secret</p></template>', () => elementNode("template", [elementNode("p", [textNode("hidden-secret")])])],
    ['<!-- > hidden-secret -->', () => commentNode(" > hidden-secret ")],
    ['<!-- <span>hidden-secret</span> -->', () => commentNode(" <span>hidden-secret</span> ")],
    ['<pre>pre<template><b>hidden-secret</b></template><!-- > hidden-secret --></pre>', () => elementNode("pre", [
      textNode("pre"), elementNode("template", [elementNode("b", [textNode("hidden-secret")])]), commentNode(" > hidden-secret "),
    ])],
    ['<div hidden><template><p>hidden-secret</p></template></div>', () => elementNode("div", [
      elementNode("template", [elementNode("p", [textNode("hidden-secret")])]),
    ], { hidden: "" })],
  ]) {
    for (const format of ["text", "markdown"]) {
      it(`excludes inert DOM contents from ${format}: ${hidden}`, () => {
        const body = htmlDocument(`<p>visible</p>${hidden}<p>after</p>`);
        const { api } = loadHtmlFixture(body, () => documentTree([
          elementNode("p", [textNode("visible")]), buildHidden(), elementNode("p", [textNode("after")]),
        ]));
        const result = api.extractFormattedBody({ contentType: "text/html", body }, format);
        assert.match(result.body, /visible/);
        assert.match(result.body, /after/);
        assert.doesNotMatch(result.body, /hidden-secret|head-secret/);
      });
    }
  }
  it("uses text nodes without reinterpreting decoded text as markup", () => {
    const html = htmlDocument('<p>&lt;template&gt;visible&lt;/template&gt; &amp;lt;</p>');
    const { api } = loadHtmlFixture(html, () => documentTree([elementNode("p", [textNode('<template>visible</template> &lt;')])]));
    assert.equal(api.stripHtml(html), '<template>visible</template> &lt;');
  });
  it("handles uppercase tags and preserves visible siblings around a hidden subtree", () => {
    const html = htmlDocument('<DIV>before<SPAN hidden><B>secret</B></SPAN><BR>after</DIV>');
    const { api } = loadHtmlFixture(html, () => documentTree([elementNode("DIV", [
      textNode("before"), elementNode("SPAN", [elementNode("B", [textNode("secret")])], { hidden: "" }),
      elementNode("BR"), textNode("after"),
    ])]));
    assert.equal(api.stripHtml(html), "before\nafter");
    assert.equal(api.htmlToMarkdown(html), "before\nafter");
  });
  it("separates table cells and rows in message reads and direct reply/forward quotations", async () => {
    const html = htmlDocument('<table><tr><th>Field</th><th>Value</th></tr><tr><td>Account</td><td>123</td></tr><tr><td>Total</td><td>45</td></tr></table>');
    const expected = "Field Value\nAccount 123\nTotal 45";
    const { api, calls } = loadHtmlFixture(html, () => documentTree([
      elementNode("table", [elementNode("tbody", [
        elementNode("tr", [elementNode("th", [textNode("Field")]), elementNode("th", [textNode("Value")])]),
        elementNode("tr", [elementNode("td", [textNode("Account")]), elementNode("td", [textNode("123")])]),
        elementNode("tr", [elementNode("td", [textNode("Total")]), elementNode("td", [textNode("45")])]),
      ])]),
    ]), { mime: { contentType: "text/html", body: html } });
    for (const format of ["text", "markdown"]) {
      const result = await api.getMessage("message-1", "folder", false, format);
      assert.equal(result.body, expected);
      assert.notEqual(result.encryptedContentWithheld, true);
    }
    assert.equal((await api.replyToMessage("message-1", "folder", "intro", false, false, undefined, undefined, undefined, undefined, undefined, true)).success, true);
    assert.equal((await api.forwardMessage("message-1", "folder", "to@example.test", "intro", false, undefined, undefined, undefined, undefined, true)).success, true);
    assert.ok(calls.sends[0].body.includes("> Field Value\n> Account 123\n> Total 45"));
    assert.ok(calls.sends[1].body.includes(expected));
  });
  it("keeps explicitly requested HTML unchanged", () => {
    const { api } = loadMessageTools();
    const html = '<div hidden>raw\u202etext</div>';
    const result = api.extractFormattedBody({ contentType: "text/html", body: html }, "html");
    assert.equal(result.body, html);
    assert.equal(result.bodyIsHtml, true);
  });
  it("withholds HTML when no safe parser is available", () => {
    const { api } = loadMessageTools({ DOMParser: null });
    assert.doesNotMatch(api.stripHtml('<span hidden><b>secret</b></span>'), /secret/);
    assert.doesNotMatch(api.htmlToMarkdown('<div style="opacity:0">secret</div>'), /secret/);
    assert.match(api.stripHtml('<template><p>secret</p></template><!-- > secret -->'), /withheld/);
  });
  it("strips every requested Unicode class after entity decoding, preserving ZWJ and ZWNJ", () => {
    const controls = [0x200b, 0x2060, 0xfeff, ...Array.from({ length: 5 }, (_, i) => 0x202a + i), ...Array.from({ length: 4 }, (_, i) => 0x2066 + i), 0xe0000, 0xe0041, 0xe007f];
    for (const codepoint of controls) {
      const char = String.fromCodePoint(codepoint);
      const html = htmlDocument(`<p>a&#x${codepoint.toString(16)};b</p>`);
      const { api } = loadHtmlFixture(html, () => documentTree([elementNode("p", [textNode(`a${char}b`)])]));
      assert.equal(api.extractFormattedBody({ contentType: "text/plain", body: `a${char}b` }, "text").body, "ab");
      assert.equal(api.stripHtml(html), "ab");
    }
    const html = htmlDocument("a&zwj;b&zwnj;c");
    const { api } = loadHtmlFixture(html, () => documentTree([textNode("a\u200db\u200cc")]));
    assert.equal(api.stripHtml(html), "a\u200db\u200cc");
    assert.equal(api.extractFormattedBody({ contentType: "text/plain", body: "👩\u200d💻 ا\u200cب" }, "markdown").body, "👩\u200d💻 ا\u200cب");
  });
});

// Thunderbird 147/153 route multipart/signed through the OpenPGP handler, so
// Gloda flags the container as encrypted and, without examineEncryptedParts,
// returns it with no children. comm-central leaves it unflagged with children.
describe("Clear-signed mail with the encrypted-message preference off", () => {
  const PGP_SIGNATURE = "application/pgp-signature";
  const PKCS7_SIGNATURE = "application/pkcs7-signature";
  const SIGNED_TEXT = "signed text";
  const SECRET = "decrypted body";
  const signedHeaderValue = protocol => `multipart/signed; micalg=sha-256; protocol="${protocol}"; boundary="sig"`;
  const signedContainer = (parts, { protocol = PGP_SIGNATURE, isEncrypted = true, headerValue = signedHeaderValue(protocol) } = {}) => ({
    contentType: "message/rfc822", headers: { "content-type": [headerValue] },
    parts: [{ contentType: "multipart/signed", headers: { "content-type": [headerValue] }, isEncrypted, parts }],
  });
  const signatureLeaf = protocol => protocol === PGP_SIGNATURE
    ? [`Content-Type: ${PGP_SIGNATURE}; name="OpenPGP_signature.asc"`, 'Content-Disposition: attachment; filename="OpenPGP_signature.asc"', "",
      "-----BEGIN PGP SIGNATURE-----", "", "wsB5", "-----END PGP SIGNATURE-----"]
    : [`Content-Type: ${protocol}; name="smime.p7s"`, 'Content-Disposition: attachment; filename="smime.p7s"', "Content-Transfer-Encoding: base64", "", "MIIB"];
  const signedRaw = (signedPart, protocol = PGP_SIGNATURE) => [
    `Content-Type: ${signedHeaderValue(protocol)}`, "",
    "--sig", ...signedPart, "--sig", ...signatureLeaf(protocol), "--sig--", "",
  ].join("\r\n");
  const plainPart = ["Content-Type: text/plain; charset=utf-8", "", SIGNED_TEXT];
  const mixedPart = [
    'Content-Type: multipart/mixed; boundary="mix"', "",
    "--mix", ...plainPart,
    "--mix", 'Content-Type: application/pdf; name="report.pdf"', 'Content-Disposition: attachment; filename="report.pdf"',
    "Content-Transfer-Encoding: base64", "", "JVBERi0=",
    "--mix--",
  ];
  const nestedEncryptedRaws = [
    { label: "multipart/encrypted", part: [
      'Content-Type: multipart/encrypted; protocol="application/pgp-encrypted"; boundary="enc"', "",
      "--enc", "Content-Type: application/pgp-encrypted", "", "Version: 1",
      "--enc", "Content-Type: application/octet-stream", "", "ciphertext", "--enc--",
    ] },
    { label: "pkcs7-mime enveloped data", part: [
      'Content-Type: application/pkcs7-mime; smime-type=enveloped-data; name="smime.p7m"', "Content-Transfer-Encoding: base64", "", "MIIB",
    ] },
    { label: "a standalone inline PGP MESSAGE line", part: [
      "Content-Type: text/plain; charset=utf-8", "", "intro", "-----BEGIN PGP MESSAGE-----", "", "hQEM", "-----END PGP MESSAGE-----",
    ] },
  ];
  const reply = api => api.replyToMessage("message-1", "folder", "intro", false, false, undefined, undefined, undefined, undefined, undefined, true);
  const forward = api => api.forwardMessage("message-1", "folder", "to@example.test", "intro", false, undefined, undefined, undefined, undefined, true);
  const assertNoDecryption = calls => {
    for (const options of calls.options) assert.equal(options.examineEncryptedParts, false);
  };

  for (const protocol of [PGP_SIGNATURE, PKCS7_SIGNATURE, "application/x-pkcs7-signature"]) {
    it(`reads a flagged empty clear-signed container from the raw message, protocol=${protocol}`, async () => {
      const { api, calls } = loadMessageTools({ mime: signedContainer([], { protocol }), raw: signedRaw(mixedPart, protocol) });
      const result = await api.getMessage("message-1", "folder", false, "markdown");
      assert.equal(result.body, SIGNED_TEXT);
      assert.equal(result.subject, "protected subject");
      assert.notEqual(result.encryptedContentWithheld, true);
      // Attachment metadata comes from the raw message, without the detached signature.
      assert.deepEqual(JSON.parse(JSON.stringify(result.attachments)), [
        { name: "report.pdf", contentType: "application/pdf", size: 5, isInline: false },
      ]);
      assert.equal(calls.options.length, 1);
      assert.equal(calls.streams, 1);
      assertNoDecryption(calls);
    });
  }

  it("returns the raw source of flagged clear-signed mail with a single read", async () => {
    const raw = signedRaw(plainPart);
    const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw });
    const result = await api.getMessage("message-1", "folder", false, "markdown", true);
    assert.equal(result.rawSource, raw);
    assert.equal(calls.streams, 1);
    assert.equal(calls.options.length, 0);
    assertNoDecryption(calls);
  });

  it("keeps the comm-central unflagged shape in a single structured pass plus one raw check", async () => {
    const { api, calls } = loadMessageTools({
      mime: signedContainer([{ contentType: "text/plain", body: SIGNED_TEXT }], { isEncrypted: false }), raw: signedRaw(plainPart),
    });
    const result = await api.getMessage("message-1", "folder", false, "markdown");
    assert.equal(result.body, SIGNED_TEXT);
    assert.equal(calls.options.length, 1);
    // The signed-message raw checks run on every Thunderbird version.
    assert.equal(calls.streams, 1);
    assertNoDecryption(calls);
  });

  for (const allowed of [false, true]) {
    for (const protocol of [PGP_SIGNATURE, PKCS7_SIGNATURE, "application/x-pkcs7-signature"]) {
      it(`replies to and forwards flagged clear-signed mail directly, ${protocol}, opt-in=${allowed}`, async () => {
        const { api, calls } = loadMessageTools({
          allowed, mime: signedContainer([], { protocol }), raw: signedRaw(plainPart, protocol),
        });
        for (const result of [await reply(api), await forward(api)]) assert.equal(result.success, true, result.error);
        assert.equal(calls.sends.length, 2);
        for (const sent of calls.sends) assert.match(sent.body, /signed text/);
        assert.doesNotMatch(calls.sends[1].body, /PGP SIGNATURE/);
        assert.equal(calls.streams, 2);
        assert.equal(calls.options.length, 2);
        assertNoDecryption(calls);
      });
    }
  }

  it("withholds oversized signed HTML when the complete content cannot be verified", async () => {
    const cap = 2 * 1024 * 1024;
    const html = "<p>" + " ".repeat(cap) + "clear tail</p>";
    const raw = signedRaw(["Content-Type: text/html; charset=utf-8", "", html]);
    const { api, calls } = loadHtmlFixture(html.slice(0, cap),
      () => documentTree([elementNode("p", [textNode("visible prefix")])]), {
        mime: signedContainer([]), raw,
      });
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.equal(result.encryptedContentWithheld, true);
    assert.match(result.body, /could not be determined/);
    assert.equal(calls.streams, 1);
    assert.equal(calls.options.length, 0);
    assert.equal(calls.sends.length, 0);
  });

  it("refuses a direct forward that would drop the signed message's attachments", async () => {
    const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(mixedPart) });
    assert.match((await forward(api)).error, /cannot include its attachments/);
    assert.equal(calls.sends.length, 0);
  });

  for (const { label, part } of nestedEncryptedRaws) {
    it(`withholds a signed wrapper whose raw message contains ${label}`, async () => {
      const raw = signedRaw(part);
      const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw });
      for (const rawSource of [false, true]) {
        const result = await api.getMessage("message-1", "folder", false, "text", rawSource, true);
        assert.equal(result.encryptedContentWithheld, true);
        assert.equal(result.rawSource, undefined);
        assert.equal(result.attachments.length, 0);
        assert.doesNotMatch(JSON.stringify(result), /signed text|intro|ciphertext|hQEM|protected subject/);
      }
      for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
      assert.equal(calls.sends.length, 0);
      assertNoDecryption(calls);
    });
    it(`refuses a signed wrapper containing ${label} despite encrypted-read opt-in`, async () => {
      const { api, calls } = loadMessageTools({ allowed: true, mime: signedContainer([]), raw: signedRaw(part) });
      for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
      assertNoDecryption(calls);
    });
  }

  const flaggedTrees = [
    { label: "a flagged signed container that has parts", mime: signedContainer([{ contentType: "text/plain", body: SECRET }]) },
    { label: "decrypted leaves under fake containers", mime: signedContainer([{
      contentType: "multipart/fake-container", parts: [{ contentType: "multipart/fake-container", parts: [{ contentType: "text/plain", body: SECRET }] }],
    }]) },
    { label: "an unknown signature protocol", mime: signedContainer([], { protocol: "application/x-custom" }) },
    { label: "a duplicated protocol parameter", mime: signedContainer([], {
      headerValue: `multipart/signed; protocol="application/x-custom"; protocol="${PGP_SIGNATURE}"; boundary="sig"`,
    }) },
    { label: "no protocol parameter", mime: { contentType: "message/rfc822", parts: [{ contentType: "multipart/signed", isEncrypted: true, parts: [] }] } },
    { label: "a flagged non-signed container", mime: { contentType: "message/rfc822", parts: [{
      contentType: "multipart/mixed", headers: { "content-type": [signedHeaderValue(PGP_SIGNATURE)] }, isEncrypted: true, parts: [],
    }] } },
    { label: "opaque-signed S/MIME", mime: { contentType: "message/rfc822", headers: { "content-type": ["application/pkcs7-mime; smime-type=signed-data"] }, parts: [{
      contentType: "application/pkcs7-mime", headers: { "content-type": ["application/pkcs7-mime; smime-type=signed-data"] }, isEncrypted: true, parts: [],
    }] } },
  ];
  for (const { label, mime } of flaggedTrees) {
    it(`withholds ${label} in reads and after the direct-send raw check`, async () => {
      // The raw source is clear; only the structural classification may decide here.
      const { api, calls } = loadMessageTools({ mime, raw: signedRaw(plainPart) });
      for (const rawSource of [false, true]) {
        const result = await api.getMessage("message-1", "folder", true, "html", rawSource, true);
        if (rawSource) {
          assert.equal(result.rawSource, signedRaw(plainPart));
          continue;
        }
        assert.equal(result.encryptedContentWithheld, true);
        assert.doesNotMatch(JSON.stringify(result), /decrypted body|signed text|protected subject/);
      }
      for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
      assert.equal(calls.streams, 4);
      assert.equal(calls.sends.length, 0);
      assertNoDecryption(calls);
    });
    it(`refuses ${label} despite encrypted-read opt-in`, async () => {
      const { api, calls } = loadMessageTools({ allowed: true, mime, raw: signedRaw(plainPart) });
      for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
      assert.equal(calls.streams, 2);
      assertNoDecryption(calls);
    });
  }

  it("withholds the raw path when the raw message cannot be read", async () => {
    const { api, calls } = loadMessageTools({ mime: signedContainer([]), streamError: new Error("stream failed") });
    assert.equal((await api.getMessage("message-1", "folder", false, "text")).error, "Failed to get message");
    assertOriginalRefused(calls, await reply(api));
  });

  const ZWSP_ARMOR_PART = ["Content-Type: text/plain; charset=utf-8", "", "intro", "\xE2\x80\x8B-----BEGIN PGP MESSAGE-----", "", "hQEM", "-----END PGP MESSAGE-----"];
  const ARMOR_HTML = "<p>-----BEGIN PGP MESSAGE-----</p><p>hQEM</p>";
  const armorHtmlTree = () => documentTree([elementNode("p", [textNode("-----BEGIN PGP MESSAGE-----")]), elementNode("p", [textNode("hQEM")])]);
  const HTML_ARMOR_PART = ["Content-Type: text/html; charset=utf-8", "", ARMOR_HTML];
  const hiddenArmorCases = [
    { label: "a zero-width space before the armor line", load: options => loadMessageTools({ ...options, raw: signedRaw(ZWSP_ARMOR_PART) }) },
    { label: "armor split into HTML paragraphs", load: options => loadHtmlFixture(ARMOR_HTML, armorHtmlTree, { ...options, raw: signedRaw(HTML_ARMOR_PART) }) },
  ];
  for (const { label, load } of hiddenArmorCases) {
    it(`withholds signed mail with ${label} in every format and direct send`, async () => {
      const { api, calls } = load({ mime: signedContainer([]) });
      for (const format of ["text", "markdown", "html"]) {
        const result = await api.getMessage("message-1", "folder", false, format);
        assert.equal(result.encryptedContentWithheld, true, format);
        assert.doesNotMatch(JSON.stringify(result), /hQEM|intro/);
      }
      // Raw source output depends on the raw classifier alone.
      const rawResult = await api.getMessage("message-1", "folder", false, "text", true);
      assert.equal(rawResult.encryptedContentWithheld, true);
      assert.equal(rawResult.rawSource, undefined);
      for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
      assert.equal(calls.sends.length, 0);
      assertNoDecryption(calls);
    });
    it(`re-checks the extracted fallback body for ${label} even if raw classification passed`, async () => {
      const { api } = load({ mime: signedContainer([]) });
      api.classifyRawMessageEncryption = () => "clear";
      for (const format of ["text", "markdown", "html"]) {
        const result = await api.getMessage("message-1", "folder", false, format);
        assert.equal(result.encryptedContentWithheld, true, format);
        assert.doesNotMatch(JSON.stringify(result), /hQEM/);
      }
    });
  }

  it("treats bare CR line breaks around armor as encrypted on both paths", async () => {
    const armored = "intro\r-----BEGIN PGP MESSAGE-----\rhQEM\r-----END PGP MESSAGE-----";
    const { api } = loadMessageTools({ mime: { contentType: "text/plain", body: armored } });
    assert.equal(api.classifyRawMessageEncryption(`Content-Type: text/plain\r\n\r\n${armored}`), "encrypted");
    assert.equal(api.classifyRawMessageEncryption(signedRaw(["Content-Type: text/plain", "", armored])), "encrypted");
    assert.equal(api.classifyMimeMessageEncryption({ contentType: "text/plain", body: armored }), "encrypted");
    assert.equal((await api.getMessage("message-1", "folder", false, "text")).encryptedContentWithheld, true);
    assert.equal(api.hasInlinePgpBodyArmor({ contentType: "text/plain", body: "" }, armored), true);
  });

  const attachedMessagePart = [
    'Content-Type: multipart/mixed; boundary="mix"', "",
    "--mix", ...plainPart,
    "--mix", "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="earlier.eml"', "",
    "Subject: earlier", "", "earlier body",
    "--mix--",
  ];
  const PNG_BASE64 = "iVBORw0KGgo=";
  const relatedPart = (imageBase64 = PNG_BASE64) => [
    'Content-Type: multipart/related; boundary="rel"', "",
    "--rel", ...plainPart,
    "--rel", "Content-Type: image/png", "Content-ID: <img1>", "Content-Transfer-Encoding: base64", "", imageBase64,
    "--rel--",
  ];

  it("lists attached messages from the raw message like Gloda does", async () => {
    const { api } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(attachedMessagePart) });
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.equal(result.body, SIGNED_TEXT);
    assert.deepEqual(JSON.parse(JSON.stringify(result.attachments)), [
      { name: "earlier.eml", contentType: "message/rfc822", size: "Subject: earlier\r\n\r\nearlier body".length, isInline: false },
    ]);
  });

  for (const { label, part } of [
    { label: "an attached message without a disposition", part: attachedMessagePart.map(line => line.startsWith("Content-Disposition") ? "X-Note: none" : line) },
    { label: "a Content-ID image without a filename", part: relatedPart() },
  ]) {
    it(`refuses a direct forward that would drop ${label}`, async () => {
      const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(part) });
      assert.match((await forward(api)).error, /cannot include its attachments/);
      assert.equal(calls.sends.length, 0);
      assert.equal((await reply(api)).success, true);
    });
  }

  it("lists raw inline images only when the caller asks for them", async () => {
    const { api } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(relatedPart()) });
    const plain = await api.getMessage("message-1", "folder", false, "text");
    assert.equal(plain.body, SIGNED_TEXT);
    // Same rule as the Gloda path without the opt-in: related images are inline entries.
    assert.deepEqual(JSON.parse(JSON.stringify(plain.attachments)), [
      { name: "inline_1.1.2", contentType: "image/png", size: 8, isInline: true, partName: "1.1.2" },
    ]);
    assert.equal(plain.inlineImageContent, undefined);

    const withImages = await api.getMessage("message-1", "folder", false, "text", false, true);
    const [image] = withImages.attachments;
    assert.equal(image.contentType, "image/png");
    assert.equal(image.contentId, "img1");
    assert.equal(image.partName, "1.1.2");
    assert.equal(image.mcpImage.status, "included");
    assert.equal(withImages.inlineImageContent.included, 1);
    const blocks = withImages[Object.getOwnPropertySymbols(withImages)[0]];
    assert.deepEqual(JSON.parse(JSON.stringify(blocks)), [{ type: "image", data: PNG_BASE64, mimeType: "image/png" }]);
  });

  it("applies the per-image base64 limit to raw inline images", async () => {
    const big = Buffer.alloc(800 * 1024, 1).toString("base64").match(/.{1,76}/g).join("\r\n");
    const { api } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(relatedPart(big)) });
    const result = await api.getMessage("message-1", "folder", false, "text", false, true);
    assert.equal(result.attachments[0].mcpImage.status, "skipped");
    assert.match(result.attachments[0].mcpImage.reason, /limit/);
    assert.equal(result.inlineImageContent.included, 0);
  });

  it("saves raw-path attachments from the decoded bytes without a part fetch", async () => {
    const { api } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(mixedPart) });
    const written = [];
    const makeFile = filePath => ({
      path: filePath,
      append(name) { this.path += `/${name}`; },
      clone() { return makeFile(this.path); },
      create() {}, createUnique() {}, remove() {},
      exists: () => true, isDirectory: () => true,
    });
    api.Services.dirsvc = { get: () => makeFile("/tmp") };
    api.Ci.nsIFile = { DIRECTORY_TYPE: 1, NORMAL_FILE_TYPE: 0 };
    api.NetUtil = { newChannel: () => assert.fail("raw-path attachments must not be fetched") };
    api.Cc["@mozilla.org/network/file-output-stream;1"] = { createInstance: () => ({ init(file) { this.file = file; }, close() {} }) };
    api.Cc["@mozilla.org/binaryoutputstream;1"] = { createInstance: () => ({
      setOutputStream(stream) { this.stream = stream; },
      writeByteArray(bytes, length) { written.push({ path: this.stream.file.path, data: Buffer.from(bytes.slice(0, length)).toString("latin1") }); },
      close() {},
    }) };
    const result = await api.getMessage("message-1", "folder", true, "text");
    assert.equal(result.attachments[0].error, undefined);
    assert.equal(result.attachments[0].filePath, "/tmp/thunderbird-mcp/message_1/report.pdf");
    assert.deepEqual(written, [{ path: "/tmp/thunderbird-mcp/message_1/report.pdf", data: "%PDF-" }]);
  });

  it("withholds signed mail whose raw message exceeds the read cap", async () => {
    const { api, calls } = loadMessageTools({ mime: signedContainer([]) });
    const limits = [];
    api.readMessageStreamFully = (_stream, maxBytes) => {
      limits.push(maxBytes);
      throw Object.assign(new Error("message too large"), { isStreamSizeLimit: true });
    };
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.match(result.error, /message too large/);
    assertOriginalRefused(calls, await reply(api));
    assert.deepEqual(limits, [50 * 1024 * 1024, 50 * 1024 * 1024]);
    assert.equal(calls.sends.length, 0);
  });

  // libmime renders message/news, non-base64 *.eml attachments and untyped
  // multipart/digest children as embedded messages and looks inside them.
  const ENCRYPTED_INNER_MESSAGES = [
    { label: "pkcs7-mime enveloped data", lines: [
      "Subject: inner", "Content-Type: application/pkcs7-mime; smime-type=enveloped-data", "Content-Transfer-Encoding: base64", "", "MIIB",
    ] },
    { label: "multipart/encrypted with base64 armor", lines: [
      "Subject: inner", 'Content-Type: multipart/encrypted; protocol="application/pgp-encrypted"; boundary="enc"', "",
      "--enc", "Content-Type: application/pgp-encrypted", "", "Version: 1",
      "--enc", "Content-Type: application/octet-stream", "Content-Transfer-Encoding: base64", "",
      Buffer.from("-----BEGIN PGP MESSAGE-----\n\nhQEM\n-----END PGP MESSAGE-----\n").toString("base64"),
      "--enc--",
    ] },
  ];
  const CLEAR_INNER_MESSAGE = ["Subject: inner", "Content-Type: text/plain", "", "embedded text"];
  const embeddedWrappers = [
    { label: "message/news", wrap: inner => [
      'Content-Type: multipart/mixed; boundary="mix"', "",
      "--mix", ...plainPart,
      "--mix", "Content-Type: message/news", 'Content-Disposition: attachment; filename="post.eml"', "", ...inner,
      "--mix--",
    ], listed: inner => [{ name: "post.eml", contentType: "message/news", size: inner.join("\r\n").length, isInline: false }] },
    { label: "a non-base64 octet-stream named .eml", wrap: inner => [
      'Content-Type: multipart/mixed; boundary="mix"', "",
      "--mix", ...plainPart,
      "--mix", 'Content-Type: application/octet-stream; name="fwd.eml"', "Content-Transfer-Encoding: 7bit", "", ...inner,
      "--mix--",
    ], nameTyped: true },
    { label: "an untyped multipart/digest child", wrap: inner => [
      'Content-Type: multipart/mixed; boundary="mix"', "",
      "--mix", ...plainPart,
      "--mix", 'Content-Type: multipart/digest; boundary="dig"', "",
      "--dig", "", ...inner,
      "--dig--",
      "--mix--",
    ], listed: () => [] },
  ];
  for (const wrapper of embeddedWrappers) {
    for (const inner of ENCRYPTED_INNER_MESSAGES) {
      it(`withholds ${inner.label} inside ${wrapper.label} on every raw-path output`, async () => {
        const raw = signedRaw(wrapper.wrap(inner.lines));
        const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw });
        // A name-typed part is withheld as "unknown" before its content is read.
        assert.equal(api.classifyRawMessageEncryption(raw, 0, { signedRaw: true }), wrapper.nameTyped ? "unknown" : "encrypted");
        for (const [saveAttachments, rawSource] of [[false, false], [true, false], [false, true]]) {
          const result = await api.getMessage("message-1", "folder", saveAttachments, "text", rawSource, true);
          assert.equal(result.encryptedContentWithheld, true);
          assert.equal(result.rawSource, undefined);
          assert.equal(result.attachments.length, 0);
          assert.doesNotMatch(JSON.stringify(result), /signed text|MIIB|hQEM|LS0t/);
        }
        for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
        assert.equal(calls.sends.length, 0);
        assertNoDecryption(calls);
      });
    }
    if (wrapper.nameTyped) continue;
    it(`keeps a clear embedded message in ${wrapper.label} readable`, async () => {
      const raw = signedRaw(wrapper.wrap(CLEAR_INNER_MESSAGE));
      const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw });
      const result = await api.getMessage("message-1", "folder", false, "text");
      assert.equal(result.body, SIGNED_TEXT);
      assert.deepEqual(JSON.parse(JSON.stringify(result.attachments)), wrapper.listed(CLEAR_INNER_MESSAGE));
      assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).rawSource, raw);
      assert.equal((await reply(api)).success, true);
      // A direct forward cannot carry the embedded message, so it is refused.
      assert.match((await forward(api)).error, /cannot include its attachments/);
      assert.equal(calls.sends.length, 1);
    });
  }

  it("does not take an untyped digest child as the body", async () => {
    const digestOnly = ['Content-Type: multipart/digest; boundary="dig"', "", "--dig", "", ...CLEAR_INNER_MESSAGE, "--dig--"];
    const { api } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(digestOnly) });
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.notEqual(result.encryptedContentWithheld, true);
    assert.doesNotMatch(result.body, /embedded text|Subject/);
  });

  it("does not list an inline attached message without a name", async () => {
    const part = [
      'Content-Type: multipart/mixed; boundary="mix"', "",
      "--mix", ...plainPart, "--mix", "Content-Type: message/rfc822", "", ...CLEAR_INNER_MESSAGE, "--mix--",
    ];
    const { api } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(part) });
    assert.deepEqual((await api.getMessage("message-1", "folder", false, "text")).attachments.length, 0);
  });

  // libmime chooses a class from the file name for untyped or generic parts
  // (MimeHeaders_get_name order, RFC 2231 and RFC 2047 decoded), so any name that
  // could select a message or S/MIME class withholds the whole message.
  const SMIME_INNER = ["Subject: inner", "MIME-Version: 1.0", "Content-Type: application/pkcs7-mime; smime-type=enveloped-data", "Content-Transfer-Encoding: base64", "", "MIIB"];
  const mixedWith = leaf => ['Content-Type: multipart/mixed; boundary="mix"', "", "--mix", ...plainPart, "--mix", ...leaf, "--mix--"];
  const OCTET = "Content-Type: application/octet-stream";
  const nameTypedLeaves = [
    { label: "an octet-stream named .p7m with base64 CMS", leaf: [`${OCTET}; name="secret.p7m"`, "Content-Transfer-Encoding: base64", "", "MIIB"] },
    { label: "an octet-stream named .mail wrapping S/MIME", leaf: [`${OCTET}; name="fwd.mail"`, "Content-Transfer-Encoding: 7bit", "", ...SMIME_INNER] },
    { label: "an octet-stream named .art wrapping S/MIME", leaf: [`${OCTET}; name="fwd.art"`, "", ...SMIME_INNER] },
    { label: "an RFC 2231 continued .eml filename", leaf: [OCTET, 'Content-Disposition: attachment; filename*0="fwd."; filename*1="eml"', "", ...SMIME_INNER] },
    { label: "an RFC 2047 encoded .eml name", leaf: [`${OCTET}; name="=?utf-8?Q?fwd.eml?="`, "", ...SMIME_INNER] },
    { label: "a Content-Name .eml", leaf: [OCTET, "Content-Name: fwd.eml", "", ...SMIME_INNER] },
    { label: "an X-Sun-Data-Name .eml", leaf: [OCTET, "X-Sun-Data-Name: fwd.eml", "", ...SMIME_INNER] },
    { label: "an untyped part named .pgp", leaf: ['Content-Disposition: attachment; filename="a.pgp"', "", "binary"] },
    { label: "an x-unknown-content-type part named .gpg", leaf: ['Content-Type: application/x-unknown-content-type; name="a.GPG"', "", "binary"] },
    { label: "an untyped embedded message root with a .p7m filename", leaf: [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="earlier.eml"', "",
      "Subject: inner", 'Content-Disposition: attachment; filename="x.p7m"', "Content-Transfer-Encoding: base64", "", "MIIB",
    ] },
    { label: "an untyped embedded message root with any name", leaf: [
      "Content-Type: message/rfc822", "", "Subject: inner", "Content-Name: notes.txt", "", "embedded text",
    ] },
    { label: "a message/rfc822 part named .p7m over an untyped base64 root", leaf: [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="x.p7m"', "",
      "From: x@example.invalid", "Subject: inner", "Content-Transfer-Encoding: base64", "",
      "MIAGCSqGSIb3DQEHA6CAMIACAQAxggEwMIIBLAIBADCBlDCBjjELMAkGA1UEBhMC",
    ] },
    { label: "a message/rfc822 part with a protected name over a typed root", leaf: [
      "Content-Type: message/rfc822; name=\"x.pgp\"", "", "Subject: inner", "MIME-Version: 1.0", "Content-Type: text/plain", "", "embedded text",
    ] },
    { label: "an untyped embedded root with a base64 transfer encoding", leaf: [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="earlier.eml"', "",
      "From: x@example.invalid", "Subject: inner", "Content-Transfer-Encoding: base64", "", "ZW1iZWRkZWQ=",
    ] },
    { label: "an untyped embedded root with a quoted-printable transfer encoding", leaf: [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="earlier.eml"', "",
      "Subject: inner", "Content-Transfer-Encoding: quoted-printable", "", "embedded text",
    ] },
    // Mozilla's parameter parser reads these as x.p7m; ours would not.
    { label: "a backslash inside a quoted name", leaf: [`${OCTET}; name="x.p7\\m"`, "", "data"] },
    { label: "a backslash inside a quoted filename", leaf: [OCTET, 'Content-Disposition: attachment; filename="x.\\p7m"', "", "data"] },
    { label: "text after a closing quote", leaf: [OCTET, 'Content-Disposition: attachment; filename="x.p7m"junk', "", "data"] },
    { label: "a comma before the first parameter", leaf: [OCTET, "Content-Disposition: attachment, filename=x.p7m", "", "data"] },
    { label: "a disposition without a token", leaf: [OCTET, "Content-Disposition: filename=x.p7m", "", "data"] },
    { label: "a NUL inside a filename", leaf: [OCTET, "Content-Disposition: attachment; filename=x.p7m\0.bin", "", "data"] },
    { label: "a backslash in Content-Name", leaf: [OCTET, "Content-Name: x.p7\\m", "", "data"] },
    { label: "a message part with a backslash in its filename", leaf: [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="x.\\p7m"', "",
      "Subject: inner", "MIME-Version: 1.0", "Content-Type: text/plain", "", "embedded text",
    ] },
    // libmime ends headers at the first empty line under CR, LF or CRLF.
    { label: "headers ended by LF then CRLF", leaf: ["X-Note: 1\n", "Content-Type: text/plain", "", "data"] },
    { label: "headers ended by CRLF then CR", leaf: ["X-Note: 1", "\rContent-Type: text/plain", "", "data"] },
    { label: "an unknown filename charset", leaf: [OCTET, "Content-Disposition: attachment; filename*=x-no-such-charset''report.pdf", "", "data"] },
    { label: "a gap in RFC 2231 continuations", leaf: [OCTET, 'Content-Disposition: attachment; filename*0="report"; filename*2=".pdf"', "", "data"] },
    { label: "a malformed RFC 2231 percent escape", leaf: [OCTET, "Content-Disposition: attachment; filename*=utf-8''report%G1.pdf", "", "data"] },
    { label: "a malformed RFC 2047 encoded word", leaf: [`${OCTET}; name="=?utf-8?B?***?="`, "", "data"] },
    { label: "a uuencoded .p7m in an untyped embedded body", leaf: [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="earlier.eml"', "",
      "Subject: inner", "", "intro", "begin 644 x.p7m", "M04)#", "end",
    ] },
    { label: "a BinHex block in an untyped embedded body", leaf: [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="earlier.eml"', "",
      "Subject: inner", "", "(This file must be converted with BinHex 4.0)", ":data:",
    ] },
    { label: "a yEnc .pgp in an embedded text body", leaf: [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="earlier.eml"', "",
      "Subject: inner", "Content-Type: text/plain", "", "=ybegin line=128 size=4 name=x.pgp", "data", "=yend size=4",
    ] },
  ];
  for (const { label, leaf } of nameTypedLeaves) {
    it(`withholds signed mail with ${label} on every raw-path output`, async () => {
      const raw = signedRaw(mixedWith(leaf));
      const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw });
      assert.equal(api.classifyRawMessageEncryption(raw, 0, { signedRaw: true }), "unknown");
      for (const [saveAttachments, rawSource] of [[false, false], [true, false], [false, true]]) {
        const result = await api.getMessage("message-1", "folder", saveAttachments, "text", rawSource, true);
        assert.equal(result.encryptedContentWithheld, true);
        assert.equal(result.rawSource, undefined);
        assert.equal(result.attachments.length, 0);
        assert.doesNotMatch(JSON.stringify(result), /signed text|MIIB|protected subject/);
      }
      for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
      assert.equal(calls.sends.length, 0);
      assertNoDecryption(calls);
    });
  }

  it("lists RFC 2231 and RFC 2047 attachment names decoded and keeps the message readable", async () => {
    const encoded = Buffer.from("été.pdf").toString("base64");
    const part = [
      'Content-Type: multipart/mixed; boundary="mix"', "",
      "--mix", ...plainPart,
      "--mix", "Content-Type: application/pdf",
      "Content-Disposition: attachment; filename*0*=utf-8''r%C3%A9; filename*1*=sum%C3%A9; filename*2=\".pdf\"",
      "Content-Transfer-Encoding: base64", "", "JVBERi0=",
      "--mix", `Content-Type: application/pdf; name="=?utf-8?B?${encoded}?="`, "Content-Disposition: attachment",
      "Content-Transfer-Encoding: base64", "", "JVBERi0=",
      "--mix--",
    ];
    const raw = signedRaw(part);
    const { api } = loadMessageTools({ mime: signedContainer([]), raw });
    assert.equal(api.classifyRawMessageEncryption(raw, 0, { signedRaw: true }), "clear");
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.equal(result.body, SIGNED_TEXT);
    assert.deepEqual(Array.from(result.attachments, attachment => attachment.name), ["résumé.pdf", "été.pdf"]);
    assert.equal((await reply(api)).success, true);
  });

  // Strict well-formedness gate on the signed raw path: any construct a MIME
  // parser could read differently withholds the message.
  const PKCS7_LEAF = ['Content-Type: application/pkcs7-mime; smime-type=enveloped-data; name="smime.p7m"', "Content-Transfer-Encoding: base64", "", "MIIB"];
  const TEXT_LEAF = ["Content-Type: text/plain", "", "visible"];
  // Our parser follows boundary "A"; libmime follows "B" or treats the
  // separator differently and so sees the PKCS7 part.
  const hiddenPartBody = (seen, hidden) => ["", `--${hidden}`, ...PKCS7_LEAF, `--${hidden}--`, `--${seen}`, ...TEXT_LEAF, `--${seen}--`];
  const malformedLeaves = [
    { label: "boundary plus an RFC 2231 extended boundary", leaf: ["Content-Type: multipart/mixed; boundary=\"A\"; boundary*=''B", ...hiddenPartBody("A", "B")] },
    { label: "boundary plus a continued boundary", leaf: ['Content-Type: multipart/mixed; boundary="A"; boundary*0="B"', ...hiddenPartBody("A", "B")] },
    { label: "a backslash inside the boundary", leaf: ['Content-Type: multipart/mixed; boundary="A\\B"', ...hiddenPartBody("A\\B", "AB")] },
    { label: "text after the quoted boundary", leaf: ['Content-Type: multipart/mixed; boundary="AB"junk', ...hiddenPartBody("AB", "ABjunk")] },
    { label: "a form-feed separator", leaf: ['Content-Type: multipart/mixed; boundary="A"', "", "--A", ...TEXT_LEAF, "--A\f", ...PKCS7_LEAF, "--A--"] },
    { label: "a vertical-tab separator", leaf: ['Content-Type: multipart/mixed; boundary="A"', "", "--A", ...TEXT_LEAF, "--A\v", ...PKCS7_LEAF, "--A--"] },
    { label: "a duplicate parameter", leaf: ["Content-Type: text/plain; charset=utf-8; charset=us-ascii", "", "visible"] },
    { label: "an RFC 2231 charset", leaf: ["Content-Type: text/plain; charset*=utf-8''utf-8", "", "visible"] },
    { label: "a backslash in a quoted charset", leaf: ['Content-Type: text/plain; charset="utf\\-8"', "", "visible"] },
    { label: "text after a quoted disposition parameter", leaf: ['Content-Type: application/pdf', 'Content-Disposition: attachment; filename="a.pdf" x', "", "data"] },
    { label: "an unknown transfer encoding", leaf: ["Content-Type: text/plain", "Content-Transfer-Encoding: x-uuencode", "", "visible"] },
    { label: "a duplicate transfer encoding", leaf: ["Content-Type: text/plain", "Content-Transfer-Encoding: 7bit", "Content-Transfer-Encoding: base64", "", "dmlzaWJsZQ=="] },
    { label: "a line prefixed with an enclosing boundary", leaf: ["Content-Type: text/plain", "", "visible", "--sig-not-a-separator"] },
    { label: "a header name with invalid characters", leaf: ["X Bad: 1", "Content-Type: text/plain", "", "visible"] },
    { label: "a NUL in a header", leaf: ["X-Note: a\0b", "Content-Type: text/plain", "", "visible"] },
    { label: "a bare CR inside a header line", leaf: ["X-Note: a\rContent-Type: application/pkcs7-mime", "Content-Type: text/plain", "", "visible"] },
    { label: "a nested multipart without a closing delimiter", leaf: ['Content-Type: multipart/mixed; boundary="A"', "", "--A", ...TEXT_LEAF] },
    { label: "a boundary longer than 70 characters", leaf: [`Content-Type: multipart/mixed; boundary="${"b".repeat(71)}"`, "", `--${"b".repeat(71)}`, ...TEXT_LEAF, `--${"b".repeat(71)}--`] },
    { label: "a plain filename that disagrees with its continuation", leaf: ["Content-Type: application/pdf", 'Content-Disposition: attachment; filename="a.pdf"; filename*0="x.p7"; filename*1="m"', "", "data"] },
    // Thunderbird before about 152 does not decode encoded embedded messages,
    // so a quoted-printable soft break makes the header blocks differ.
    { label: "a quoted-printable embedded message hiding pkcs7-mime", leaf: [
      "Content-Type: message/rfc822", "Content-Transfer-Encoding: quoted-printable", "",
      "From: a@example.test", "Subject: fwd", "MIME-Version: 1.0",
      "X-Foo: a=", "Content-Type: application/pkcs7-mime", "X-Bar: b=", "Content-Transfer-Encoding: base64", "", "MIAGCSqGSIb3DQEHA6CAMIACAQAx",
    ] },
    { label: "a quoted-printable embedded message hiding multipart/encrypted", leaf: [
      "Content-Type: message/rfc822", "Content-Transfer-Encoding: quoted-printable", "",
      "From: a@example.test", "Subject: fwd", "MIME-Version: 1.0",
      "X-Foo: a=", "Content-Type: multipart/encrypted", "X-Bar: b=", "Content-Transfer-Encoding: 7bit", "", "MIAGCSqGSIb3DQEHA6CAMIACAQAx",
    ] },
    { label: "a base64 embedded message", leaf: [
      "Content-Type: message/rfc822", "Content-Transfer-Encoding: base64", "",
      Buffer.from("Subject: inner\r\nContent-Type: text/plain\r\n\r\nembedded text").toString("base64"),
    ] },
    { label: "a multipart with a base64 transfer encoding", leaf: ['Content-Type: multipart/mixed; boundary="A"', "Content-Transfer-Encoding: base64", "", "--A", ...TEXT_LEAF, "--A--"] },
  ];
  for (const { label, leaf } of malformedLeaves) {
    it(`withholds signed mail with ${label} through the well-formedness gate`, async () => {
      const raw = signedRaw(mixedWith(leaf));
      const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw });
      assert.equal(api.isStrictlyWellFormedRawMime(raw), false);
      assert.equal(api.classifyRawMessageEncryption(raw, 0, { signedRaw: true }), "unknown");
      for (const [saveAttachments, rawSource] of [[false, false], [true, false], [false, true]]) {
        const result = await api.getMessage("message-1", "folder", saveAttachments, "text", rawSource, true);
        assert.equal(result.encryptedContentWithheld, true);
        assert.equal(result.rawSource, undefined);
        assert.equal(result.attachments.length, 0);
        assert.doesNotMatch(JSON.stringify(result), /signed text|visible|MIIB|protected subject/);
      }
      for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
      assert.equal(calls.sends.length, 0);
      assertNoDecryption(calls);
    });
  }

  const realisticSigned = [
    { label: "Thunderbird PGP/MIME with protected headers, alternative body and an RFC 2231 attachment name",
      header: 'multipart/signed; micalg=pgp-sha256; protocol="application/pgp-signature"; boundary="------------JWnu0yhmPHsQ5Ue2FfXcrbJq"',
      attachments: ["résumé.pdf"],
      html: "<p>signed text</p>\r\n",
      lines: [
        "Content-Type: multipart/signed; micalg=pgp-sha256;",
        ' protocol="application/pgp-signature";',
        ' boundary="------------JWnu0yhmPHsQ5Ue2FfXcrbJq"',
        "", "This is an OpenPGP/MIME signed message (RFC 4880 and 3156)",
        "--------------JWnu0yhmPHsQ5Ue2FfXcrbJq",
        'Content-Type: multipart/mixed; boundary="------------0LmkWlLRnn3AQsF9ZpWKFn7f";',
        ' protected-headers="v1"',
        "From: Tom <tom@example.invalid>", "To: reader@example.invalid", "Message-ID: <a@example.invalid>", "Subject: Hello",
        "", "--------------0LmkWlLRnn3AQsF9ZpWKFn7f",
        "Content-Type: multipart/alternative;", ' boundary="------------q3sJ6Ijb2QGgM5HfK0Ip4pXz"',
        "", "--------------q3sJ6Ijb2QGgM5HfK0Ip4pXz",
        "Content-Type: text/plain; charset=UTF-8; format=flowed", "Content-Transfer-Encoding: base64",
        "", Buffer.from("signed text\r\n").toString("base64"), "",
        "--------------q3sJ6Ijb2QGgM5HfK0Ip4pXz",
        "Content-Type: text/html; charset=UTF-8", "Content-Transfer-Encoding: 7bit",
        "", "<p>signed text</p>", "",
        "--------------q3sJ6Ijb2QGgM5HfK0Ip4pXz--", "",
        "--------------0LmkWlLRnn3AQsF9ZpWKFn7f",
        `Content-Type: application/pdf; name="=?UTF-8?B?${Buffer.from("résumé.pdf").toString("base64")}?="`,
        "Content-Disposition: attachment; filename*0*=UTF-8''r%C3%A9sum%C3%A9;", " filename*1*=.pdf",
        "Content-Transfer-Encoding: base64", "", "JVBERi0=", "",
        "--------------0LmkWlLRnn3AQsF9ZpWKFn7f--", "",
        "--------------JWnu0yhmPHsQ5Ue2FfXcrbJq",
        'Content-Type: application/pgp-signature; name="OpenPGP_signature.asc"',
        "Content-Description: OpenPGP digital signature",
        'Content-Disposition: attachment; filename="OpenPGP_signature.asc"',
        "", "-----BEGIN PGP SIGNATURE-----", "", "wsB5BAABCAAjFiEE", "-----END PGP SIGNATURE-----", "",
        "--------------JWnu0yhmPHsQ5Ue2FfXcrbJq--", "",
      ] },
    { label: "Outlook S/MIME with folded parameters",
      header: 'multipart/signed; protocol="application/x-pkcs7-signature"; micalg=SHA1; boundary="----=_NextPart_000_0007_01DA1234.56789ABC"',
      attachments: [],
      lines: [
        "Content-Type: multipart/signed;", '\tprotocol="application/x-pkcs7-signature";', "\tmicalg=SHA1;",
        '\tboundary="----=_NextPart_000_0007_01DA1234.56789ABC"', "MIME-Version: 1.0",
        "", "This is a multi-part message in MIME format.", "",
        "------=_NextPart_000_0007_01DA1234.56789ABC",
        "Content-Type: multipart/alternative;", '\tboundary="----=_NextPart_001_0008_01DA1234.56789ABC"',
        "", "", "------=_NextPart_001_0008_01DA1234.56789ABC",
        "Content-Type: text/plain;", '\tcharset="us-ascii"', "Content-Transfer-Encoding: 7bit",
        "", "signed text", "",
        "------=_NextPart_001_0008_01DA1234.56789ABC",
        "Content-Type: text/html;", '\tcharset="us-ascii"', "Content-Transfer-Encoding: quoted-printable",
        "", "<p>signed text</p>",
        "------=_NextPart_001_0008_01DA1234.56789ABC--", "",
        "------=_NextPart_000_0007_01DA1234.56789ABC",
        "Content-Type: application/x-pkcs7-signature;", '\tname="smime.p7s"', "Content-Transfer-Encoding: base64",
        "Content-Disposition: attachment;", '\tfilename="smime.p7s"',
        "", "MIIB", "", "------=_NextPart_000_0007_01DA1234.56789ABC--", "",
      ] },
    { label: "Apple Mail PGP/MIME",
      header: 'multipart/signed; boundary="Apple-Mail=_5A1B2C3D-0000-4000-8000-ABCDEFABCDEF"; protocol="application/pgp-signature"; micalg=pgp-sha512',
      attachments: [],
      lines: [
        "Content-Type: multipart/signed;", '\tboundary="Apple-Mail=_5A1B2C3D-0000-4000-8000-ABCDEFABCDEF";',
        '\tprotocol="application/pgp-signature";', "\tmicalg=pgp-sha512",
        "Mime-Version: 1.0 (Mac OS X Mail 16.0 \\(3731.500.231\\))",
        "", "", "--Apple-Mail=_5A1B2C3D-0000-4000-8000-ABCDEFABCDEF",
        "Content-Transfer-Encoding: 7bit", "Content-Type: text/plain;", "\tcharset=us-ascii",
        "", "signed text", "",
        "--Apple-Mail=_5A1B2C3D-0000-4000-8000-ABCDEFABCDEF",
        "Content-Transfer-Encoding: 7bit", "Content-Disposition: attachment;", "\tfilename=signature.asc",
        "Content-Type: application/pgp-signature;", "\tname=signature.asc", "Content-Description: Message signed with OpenPGP",
        "", "-----BEGIN PGP SIGNATURE-----", "", "iQEz", "-----END PGP SIGNATURE-----", "",
        "--Apple-Mail=_5A1B2C3D-0000-4000-8000-ABCDEFABCDEF--", "",
      ] },
  ];
  for (const { label, header, attachments, lines, html = "<p>signed text</p>" } of realisticSigned) {
    it(`passes the gate and reads realistic ${label}`, async () => {
      const raw = lines.join("\r\n");
      const { api, calls } = loadHtmlFixture(html, () => documentTree([elementNode("p", [textNode(SIGNED_TEXT)])]), {
        mime: signedContainer([], { headerValue: header }), raw,
      });
      assert.equal(api.isStrictlyWellFormedRawMime(raw), true);
      assert.equal(api.classifyRawMessageEncryption(raw, 0, { signedRaw: true }), "clear");
      const result = await api.getMessage("message-1", "folder", false, "text");
      assert.notEqual(result.encryptedContentWithheld, true);
      assert.match(result.body, /signed text/);
      assert.deepEqual(Array.from(result.attachments, attachment => attachment.name), attachments);
      assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).rawSource, raw);
      assert.equal((await reply(api)).success, true);
      assert.match(calls.sends[0].body, /signed text/);
      assertNoDecryption(calls);
    });
  }

  // Every occurrence of a delimiter must be an exact line, as Thunderbird's own
  // multipart/signed splitter finds them with indexOf.
  const delimiterCases = [
    { label: "a padded opening delimiter", raw: () => signedRaw(plainPart).replace("\r\n\r\n--sig\r\n", "\r\n\r\n--sig \r\n") },
    { label: "a padded separator", raw: () => signedRaw(plainPart).replace("\r\n--sig\r\nContent-Type: application/pgp-signature", "\r\n--sig\t\r\nContent-Type: application/pgp-signature") },
    { label: "a delimiter at the end of an LF preamble line", raw: () => [
      `Content-Type: ${signedHeaderValue(PGP_SIGNATURE)}`, "", "junk--sig", "--sig", ...plainPart, "--sig", ...signatureLeaf(PGP_SIGNATURE), "--sig--", "",
    ].join("\n") },
    { label: "a delimiter in the middle of a line in part 1", raw: () => signedRaw(["Content-Type: text/plain", "", "signed text --sig", "more"]) },
    { label: "a closing delimiter in the middle of a line", raw: () => signedRaw(["Content-Type: text/plain", "", "signed text --sig--", "more"]) },
    { label: "a lone CR in a body", raw: () => signedRaw(["Content-Type: text/plain", "", "signed\rtext"]) },
  ];
  // Boundary derivation: mimeVerify strips matching quotes, and parsers
  // normalise folded whitespace differently, so only a plain boundary is accepted.
  const SIG_LEAF_LINES = [`Content-Type: ${PGP_SIGNATURE}; name="OpenPGP_signature.asc"`, "", "-----BEGIN PGP SIGNATURE-----", "", "wsB5", "-----END PGP SIGNATURE-----"];
  const quotedBoundaryRaw = boundaryParam => [
    `Content-Type: multipart/signed; micalg=pgp-sha256; protocol="${PGP_SIGNATURE}"; boundary=${boundaryParam}`, "MIME-Version: 1.0", "",
    "--'sig'", "Content-Type: text/plain; charset=utf-8", "", "visible",
    "--sig", ...PKCS7_LEAF, "--sig", "Content-Type: application/pgp-signature", "", "x", "--sig--", "",
    "--'sig'", ...SIG_LEAF_LINES, "--'sig'--", "",
  ].join("\r\n");
  const FOLDED_BOUNDARY = "abc   def";
  delimiterCases.push(
    { label: `a double-quoted boundary "'sig'"`, raw: () => quotedBoundaryRaw(`"'sig'"`) },
    { label: "a single-quoted boundary 'sig'", raw: () => quotedBoundaryRaw("'sig'") },
    { label: "an empty quoted boundary ''", raw: () => quotedBoundaryRaw("''").replace(/--'sig'/g, "--''") },
    { label: "a folded quoted boundary with repeated spaces", raw: () => [
      `Content-Type: multipart/signed; micalg=pgp-sha256; protocol="${PGP_SIGNATURE}"; boundary="abc`, '   def"', "MIME-Version: 1.0", "",
      "--abc def", "Content-Type: text/plain; charset=utf-8", "", "visible",
      `--${FOLDED_BOUNDARY}`, ...PKCS7_LEAF,
      `--${FOLDED_BOUNDARY}`, ...SIG_LEAF_LINES, `--${FOLDED_BOUNDARY}--`, "--abc def--", "",
    ].join("\r\n") },
  );
  // Thunderbird decodes an ESC sequence in a name with the message charset
  // (ISO-2022-JP), turning "smime.p7<ESC>(Bm" into "smime.p7m".
  const ESC = "\x1B";
  const iso2022Raw = ({ topParams = "", sunCharset = false, leafHeaders }) => [
    `Content-Type: ${signedHeaderValue(PGP_SIGNATURE)}${topParams}`, ...(sunCharset ? ["X-Sun-Charset: ISO-2022-JP"] : []), "MIME-Version: 1.0", "",
    "--sig", 'Content-Type: multipart/mixed; boundary="mix"', "",
    "--mix", "Content-Type: text/plain; charset=utf-8", "", "visible",
    "--mix", ...leafHeaders, "Content-Transfer-Encoding: base64", "", "MIAGCSqGSIb3DQEHA6CAMIACAQAx",
    "--mix--", "",
    "--sig", ...SIG_LEAF_LINES, "--sig--", "",
  ].join("\r\n");
  delimiterCases.push(
    { label: "an ISO-2022-JP escape in a Content-Type name", raw: () => iso2022Raw({
      topParams: "; charset=ISO-2022-JP", leafHeaders: [`Content-Type: application/octet-stream; name="smime.p7${ESC}(Bm"`],
    }) },
    { label: "an ISO-2022-JP escape in a Content-Disposition filename", raw: () => iso2022Raw({
      topParams: "; charset=ISO-2022-JP", leafHeaders: ["Content-Type: application/octet-stream", `Content-Disposition: attachment; filename="smime.p7${ESC}(Bm"`],
    }) },
    { label: "an escape in a name with an X-Sun-Charset header", raw: () => iso2022Raw({
      sunCharset: true, leafHeaders: [`Content-Type: application/octet-stream; name="smime.p7${ESC}(Bm"`],
    }) },
    { label: "a DEL in an unrelated header", raw: () => signedRaw(["X-Note: a\x7Fb", ...plainPart]) },
  );
  for (const { label, raw: build } of delimiterCases) {
    it(`withholds signed mail with ${label} through the well-formedness gate`, async () => {
      const raw = build();
      const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw });
      assert.equal(api.isStrictlyWellFormedRawMime(raw), false);
      assert.equal(api.classifyRawMessageEncryption(raw, 0, { signedRaw: true }), "unknown");
      for (const [saveAttachments, rawSource] of [[false, false], [true, false], [false, true]]) {
        const result = await api.getMessage("message-1", "folder", saveAttachments, "text", rawSource, true);
        assert.equal(result.encryptedContentWithheld, true);
        assert.equal(result.rawSource, undefined);
        assert.equal(result.attachments.length, 0);
        assert.doesNotMatch(JSON.stringify(result), /signed|visible|MIIB|protected subject/);
      }
      for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
      assert.equal(calls.sends.length, 0);
      assertNoDecryption(calls);
    });
  }

  for (const newline of ["\r\n", "\n"]) {
    it(`reads realistic Thunderbird signed mail with its preamble, newline=${JSON.stringify(newline)}`, async () => {
      const thunderbird = realisticSigned[0];
      const raw = thunderbird.lines.join(newline);
      const { api } = loadHtmlFixture(`<p>signed text</p>${newline}`, () => documentTree([elementNode("p", [textNode(SIGNED_TEXT)])]), {
        mime: signedContainer([], { headerValue: thunderbird.header }), raw,
      });
      assert.match(raw, /This is an OpenPGP\/MIME signed message/);
      assert.equal(api.isStrictlyWellFormedRawMime(raw), true);
      const result = await api.getMessage("message-1", "folder", false, "text");
      assert.notEqual(result.encryptedContentWithheld, true);
      assert.match(result.body, /signed text/);
      assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).rawSource, raw);
    });
  }

  for (const boundary of [
    "------------0LmkWlLRnn3AQsF9ZpWKFn7f", "_000_DB9PR01MB1234ABCDEF_", "Apple-Mail=_5A1B2C3D-0000-4000-8000-ABCDEFABCDEF",
    "Apple-Mail-2--123456789", "000000000000a1b2c3d4e5f6a7b8", "----K-9.FairEmail+0123", "=-AbCdEf0123456789==", "nextPart1234567.abcdEFGH",
  ]) {
    it(`accepts the real-world boundary style ${boundary}`, async () => {
      const header = `multipart/signed; micalg=pgp-sha256; protocol="${PGP_SIGNATURE}"; boundary="${boundary}"`;
      const raw = [`Content-Type: ${header}`, "", `--${boundary}`, ...plainPart, `--${boundary}`, ...signatureLeaf(PGP_SIGNATURE), `--${boundary}--`, ""].join("\r\n");
      const { api } = loadMessageTools({ mime: signedContainer([], { headerValue: header }), raw });
      assert.equal(api.isStrictlyWellFormedRawMime(raw), true);
      const result = await api.getMessage("message-1", "folder", false, "text");
      assert.equal(result.body, SIGNED_TEXT);
      assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).rawSource, raw);
    });
  }

  it("accepts RFC 5322 header names such as server-added X-Spam_score", async () => {
    const apple = realisticSigned.find(fixture => fixture.label.startsWith("Apple"));
    const raw = [apple.lines[0], ...apple.lines.slice(1, 4), "X-Spam_score: -0.1", "X-Spam_bar: /", ...apple.lines.slice(4)].join("\r\n");
    const { api } = loadMessageTools({ mime: signedContainer([], { headerValue: apple.header }), raw });
    assert.equal(api.isStrictlyWellFormedRawMime(raw), true);
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.notEqual(result.encryptedContentWithheld, true);
    assert.match(result.body, /signed text/);
  });

  // Thunderbird 156 (comm-central) shows a signed container unflagged with its
  // children, so the message never takes the "raw" path. The signed-message
  // checks must still apply.
  const UNTYPED_INNER_LEAF = [
    "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="x.p7m"', "",
    "From: x@example.invalid", "Subject: inner", "Content-Transfer-Encoding: base64", "",
    "MIAGCSqGSIb3DQEHA6CAMIACAQAxggEwMIIBLAIBADCBlDCBjjELMAkGA1UEBhMC",
  ];
  const commCentralTree = () => signedContainer([{ contentType: "multipart/mixed", parts: [
    { contentType: "text/plain", body: SIGNED_TEXT },
    { contentType: "message/rfc822", parts: [{ contentType: "text/plain", body: "inner" }] },
  ] }], { isEncrypted: false });

  it("withholds the untyped-inner input in the comm-central shape on every output", async () => {
    const { api, calls } = loadMessageTools({ mime: commCentralTree(), raw: signedRaw(mixedWith(UNTYPED_INNER_LEAF)) });
    for (const [saveAttachments, rawSource] of [[false, false], [true, false], [false, true]]) {
      const result = await api.getMessage("message-1", "folder", saveAttachments, "text", rawSource, true);
      assert.equal(result.encryptedContentWithheld, true);
      assert.equal(result.rawSource, undefined);
      assert.equal(result.attachments.length, 0);
      assert.doesNotMatch(JSON.stringify(result), /signed text|MIAGCS|protected subject/);
    }
    for (const result of [await reply(api), await forward(api)]) assertOriginalRefused(calls, result);
    assert.equal(calls.sends.length, 0);
    assertNoDecryption(calls);
  });

  it("applies the signed-message checks to rawSource and the body fallback from the raw headers alone", async () => {
    // No Gloda headers: only the stored message says it is multipart/signed.
    const { api } = loadMessageTools({ mime: { contentType: "message/rfc822", parts: [] }, raw: signedRaw(mixedWith(UNTYPED_INNER_LEAF)) });
    for (const rawSource of [false, true]) {
      const result = await api.getMessage("message-1", "folder", false, "text", rawSource);
      assert.equal(result.encryptedContentWithheld, true);
      assert.equal(result.rawSource, undefined);
    }
  });

  it("reads realistic clear-signed mail in the comm-central shape", async () => {
    const thunderbird = realisticSigned[0];
    const raw = thunderbird.lines.join("\r\n");
    const { api, calls } = loadHtmlFixture("<p>signed text</p>\r\n", () => documentTree([elementNode("p", [textNode(SIGNED_TEXT)])]), {
      mime: signedContainer([{ contentType: "text/plain", body: SIGNED_TEXT }], { isEncrypted: false, headerValue: thunderbird.header }), raw,
    });
    assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).rawSource, raw);
    assert.equal((await api.getMessage("message-1", "folder", false, "text")).body, SIGNED_TEXT);
    assert.equal((await reply(api)).success, true);
    assertNoDecryption(calls);
  });

  it("withholds ambiguous untyped-inner input in unsigned mail before parsing", async () => {
    const raw = [...mixedWith(UNTYPED_INNER_LEAF), ""].join("\r\n");
    const { api, calls } = loadMessageTools({ mime: { contentType: "message/rfc822", parts: [] }, raw });
    assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).encryptedContentWithheld, true);
    assert.equal((await api.getMessage("message-1", "folder", false, "text")).encryptedContentWithheld, true);
    assert.equal(calls.streams, 2);
    assert.equal(calls.options.length, 0);
  });

  it("checks embedded MIME ambiguity in unsigned mail before parsing", async () => {
    for (const name of ["OpenPGP_0x1234.asc", "Meeting.msg", "notes.mht"]) {
      const raw = [
        'Content-Type: multipart/mixed; boundary="mix"', "",
        "--mix", ...plainPart,
        "--mix", `${OCTET}; name="${name}"`, `Content-Disposition: attachment; filename="${name}"`, "", "data",
        "--mix", "Content-Type: message/rfc822", "", "Subject: inner", "Content-Transfer-Encoding: base64", "", "ZW1iZWRkZWQ=",
        "--mix--", "",
      ].join("\r\n");
      const { api } = loadMessageTools({ raw, mime: { parts: [] } });
      assert.equal(api.classifyRawMessageEncryption(raw), "clear", name);
      assert.equal(api.classifyRawMessageEncryption(raw, 0, { signedRaw: true }), "unknown", name);
      assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).encryptedContentWithheld, true, name);
      assert.equal((await api.getMessage("message-1", "folder", false, "text")).encryptedContentWithheld, true, name);
    }
  });

  it("keeps a classic pre-MIME embedded message readable", async () => {
    const leaf = [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="earlier.eml"', "",
      "From: x@example.invalid", "Subject: inner", "", "plain pre-MIME text",
    ];
    const raw = signedRaw(mixedWith(leaf));
    const { api } = loadMessageTools({ mime: signedContainer([]), raw });
    assert.equal(api.classifyRawMessageEncryption(raw, 0, { signedRaw: true }), "clear");
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.equal(result.body, SIGNED_TEXT);
    assert.deepEqual(Array.from(result.attachments, attachment => attachment.name), ["earlier.eml"]);
    assert.equal((await api.getMessage("message-1", "folder", false, "text", true)).rawSource, raw);
    assert.equal((await reply(api)).success, true);
  });

  it("keeps untyped embedded bodies readable when encoded files have ordinary names", async () => {
    const leaf = [
      "Content-Type: message/rfc822", 'Content-Disposition: attachment; filename="earlier.eml"', "",
      "Subject: inner", "", "begin 644 photo.jpg", "M04)#", "end",
    ];
    const { api } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(mixedWith(leaf)) });
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.equal(result.body, SIGNED_TEXT);
    assert.deepEqual(Array.from(result.attachments, attachment => attachment.name), ["earlier.eml"]);
  });

  for (const header of ["Content-Name", "X-Sun-Data-Name"]) {
    it(`refuses a direct forward that would drop a part named only by ${header}`, async () => {
      const leaf = ["Content-Type: application/pdf", `${header}: report.pdf`, "Content-Transfer-Encoding: base64", "", "JVBERi0="];
      const { api, calls } = loadMessageTools({ mime: signedContainer([]), raw: signedRaw(mixedWith(leaf)) });
      assert.deepEqual(Array.from((await api.getMessage("message-1", "folder", false, "text")).attachments, attachment => attachment.name), ["report.pdf"]);
      assert.match((await forward(api)).error, /cannot include its attachments/);
      assert.equal(calls.sends.length, 0);
    });
  }

  it("never requests decryption from Gloda while the preference is off", async () => {
    const { api, calls } = loadMessageTools();
    await api.getMessage("message-1", "folder");
    await reply(api);
    await forward(api);
    assert.equal(calls.options.length, 3);
    assertNoDecryption(calls);
  });
});

describe("Attachment and forward handling of external attachment links", () => {
  function fileStub(filePath, created) {
    return {
      path: filePath,
      append(name) { this.path += `/${name}`; },
      clone() { return fileStub(this.path, created); },
      create() {},
      createUnique() { created.push(this.path); },
      remove() {},
      exists: () => true, isDirectory: () => true,
    };
  }

  function loadSaver(attachments, extraMime = {}) {
    const mime = { contentType: "text/plain", body: "visible", allUserAttachments: attachments, ...extraMime };
    const { api, calls } = loadMessageTools({ mime });
    const created = [];
    const channels = [];
    api.Services.dirsvc = { get: () => fileStub("/tmp", created) };
    api.Ci.nsIFile = { DIRECTORY_TYPE: 1, NORMAL_FILE_TYPE: 0 };
    api.NetUtil = {
      newChannel: options => { channels.push(options); return {}; },
      asyncFetch(_channel, callback) { callback({ data: "report pdf" }, 0); },
      asyncCopy(inputStream, outputStream, callback) {
        outputStream.file.fileSize = inputStream.data.length;
        callback(0);
      },
    };
    api.Cc["@mozilla.org/network/file-output-stream;1"] = { createInstance: () => ({ init(file) { this.file = file; } }) };
    return { api, calls, created, channels };
  }

  const attachment = (url, extra = {}) => ({ name: "report.pdf", contentType: "application/pdf", size: 10, url, ...extra });
  const EXTERNAL = /^External attachment link refused/;
  const MAIL_STORE_URLS = [
    "x-moz-ews://user@exchange.example.test/inbox/AAMk?part=1.2",
    "x-moz-graph://user@graph.example.test/inbox/AAMk?part=1.2",
    "imap://user@mail.example.test/fetch%3EUID%3E/INBOX%3E42?part=1.2",
    "mailbox-message://nobody@Local%20Folders/Inbox?number=1&part=1.2",
  ];
  const REFUSED_ATTACHMENT_URLS = [
    "file:///etc/passwd", "http://127.0.0.1:8080/admin", "https://example.test/report.pdf",
    "ftp://example.test/report.pdf", "data:text/plain,secret", "blob:https://example.test/uuid",
    "jar:file:///tmp/archive.jar!/report.pdf", "resource://gre/modules/Services.sys.mjs",
    "chrome://browser/content/browser.xhtml", "moz-extension://test/report.pdf", "about:config",
    "view-source:https://example.test/report.pdf", "FILE:///etc/passwd", "  file:///etc/passwd",
  ];

  for (const url of [
    "file:///etc/passwd", "file:///C:/Users/alice/.ssh/id_rsa", "http://127.0.0.1:8080/admin", "https://example.test/report.pdf",
  ]) {
    it(`refuses the external attachment ${url} on save without a file or channel`, async () => {
      const { api, created, channels } = loadSaver([attachment(url, { isExternal: true })]);
      const result = await api.getMessage("message-1", "folder", true, "text");
      assert.match(result.attachments[0].error, EXTERNAL);
      assert.equal(result.attachments[0].filePath, undefined);
      assert.equal(JSON.stringify(result).includes("alice"), false);
      assert.equal(JSON.stringify(result).includes(url), false);
      assert.deepEqual(created, []);
      assert.deepEqual(channels, []);
    });
  }

  it("refuses only the external attachment and still saves the mail-store one", async () => {
    const { api, created, channels } = loadSaver([
      attachment("file:///etc/passwd", { name: "evil.txt", isExternal: true }),
      attachment(MAIL_STORE_URLS[0], { name: "good.pdf" }),
    ]);
    const result = await api.getMessage("message-1", "folder", true, "text");
    assert.match(result.attachments[0].error, EXTERNAL);
    assert.equal(result.attachments[0].filePath, undefined);
    assert.equal(result.attachments[1].error, undefined);
    assert.equal(result.attachments[1].filePath, created[0]);
    assert.deepEqual(channels.map(options => options.uri), [MAIL_STORE_URLS[0]]);
    assert.equal(created.length, 1);
    assert.match(created[0], /good\.pdf$/);
  });

  for (const url of MAIL_STORE_URLS) {
    it(`opens a system-principal channel for the non-external attachment ${url}`, async () => {
      const { api, created, channels } = loadSaver([attachment(url)]);
      const result = await api.getMessage("message-1", "folder", true, "text");
      assert.equal(result.attachments[0].error, undefined);
      assert.equal(result.attachments[0].filePath, created[0]);
      assert.equal(channels.length, 1);
      assert.equal(channels[0].uri, url);
      assert.equal(channels[0].loadUsingSystemPrincipal, true);
      assert.equal(created.length, 1);
    });
  }

  for (const url of REFUSED_ATTACHMENT_URLS) {
    it(`still refuses ${JSON.stringify(url)} by scheme when the attachment is not flagged external`, async () => {
      const { api, created, channels } = loadSaver([attachment(url)]);
      const result = await api.getMessage("message-1", "folder", true, "text");
      assert.match(result.attachments[0].error, /^Refusing to fetch a [a-z-]+: URL/);
      assert.deepEqual(created, []);
      assert.deepEqual(channels, []);
    });
  }

  it("lists an external attachment when saving is not requested and fetches nothing", async () => {
    const { api, channels } = loadSaver([attachment("file:///etc/passwd", { isExternal: true })]);
    const result = await api.getMessage("message-1", "folder", false, "text");
    assert.equal(result.attachments[0].name, "report.pdf");
    assert.equal(result.attachments[0].error, undefined);
    assert.deepEqual(channels, []);
  });

  describe("inline image fallback", () => {
    const inlineMime = (url, extra) => ({
      contentType: "multipart/mixed", partName: "",
      parts: [{
        contentType: "multipart/related", partName: "1",
        parts: [
          { contentType: "text/plain", partName: "1.1", body: "visible" },
          { contentType: "image/png", partName: "1.2", size: 8, name: "pic.png", headers: { "content-id": ["<img1>"] } },
        ],
      }],
      allUserAttachments: [attachment(url, { name: "pic.png", contentType: "image/png", size: 8, partName: "1.2", ...extra })],
    });

    for (const [label, url, extra, expected] of [
      ["an external attachment", REFUSED_ATTACHMENT_URLS[0], { isExternal: true }, EXTERNAL],
      ...REFUSED_ATTACHMENT_URLS.map(url => [JSON.stringify(url), url, {}, /^Refusing to fetch/]),
    ]) {
      it(`does not fetch the fallback URL of ${label}`, async () => {
        const { api, channels } = loadSaver([], inlineMime(url, extra));
        const result = await api.getMessage("message-1", "folder", false, "text", false, true);
        const image = result.attachments.find(info => info.isInline);
        assert.ok(image, "the inline image is listed");
        assert.equal(image.mcpImage.status, "skipped");
        assert.match(image.mcpImage.reason, expected);
        assert.equal(JSON.stringify(result).includes("passwd"), false);
        assert.deepEqual(channels, []);
      });
    }
  });

  describe("direct forward", () => {
    const forward = api => api.forwardMessage("message-1", "folder", "to@example.test", "intro", false, undefined, undefined, undefined, undefined, true);

    it("refuses the whole send when any original attachment is external", async () => {
      const { api, calls } = loadSaver([
        attachment(MAIL_STORE_URLS[0], { name: "good.pdf" }),
        attachment("file:///home/alice/.ssh/id_rsa", { name: "id_rsa", isExternal: true }),
      ]);
      const result = await forward(api);
      assert.match(result.error, /external attachment link/);
      assert.match(result.error, /skipReview: false/);
      assert.equal(JSON.stringify(result).includes("alice"), false);
      assert.deepEqual(calls.sends, []);
    });

    for (const url of REFUSED_ATTACHMENT_URLS) {
      it(`refuses ${JSON.stringify(url)} when the original attachment is not flagged external`, async () => {
        const { api, calls } = loadSaver([
          attachment(MAIL_STORE_URLS[0], { name: "good.pdf" }),
          attachment(url),
        ]);
        const result = await forward(api);
        assert.match(result.error, /Refusing to fetch/);
        assert.match(result.error, /skipReview: false/);
        assert.equal(result.error.includes(url.trim()), false);
        assert.deepEqual(calls.sends, []);
      });
    }

    it("sends ordinary attachments, including Exchange URLs, unchanged", async () => {
      const { api, calls } = loadSaver(MAIL_STORE_URLS.map((url, i) => attachment(url, { name: `file${i}.pdf` })));
      const result = await forward(api);
      assert.equal(result.success, true);
      assert.deepEqual(Array.from(calls.sends[0].attachments.slice(0, MAIL_STORE_URLS.length), desc => desc.url), MAIL_STORE_URLS);
      assert.equal(calls.sends[0].attachments.length, MAIL_STORE_URLS.length + 2);
    });
  });
});

describe("getMessage rawSource reads the complete message", () => {
  const FULL = "Content-Type: multipart/mixed; boundary=b\r\n\r\n--b\r\nContent-Type: text/plain\r\n\r\nhello\r\n--b--\r\n";
  const PREFIX = FULL.slice(0, 60);

  it("streams an IMAP message that is not fully offline from the server, not its partial local copy", async () => {
    const { api, calls } = loadMessageTools({ raw: PREFIX, imap: { offline: false, serverRaw: FULL } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, FULL);
    assert.equal(calls.serverStreams, 1);
    assert.equal(calls.streams, 0, "the local store is not read for an incomplete copy");
  });

  it("reads the local copy of an IMAP message flagged Offline", async () => {
    const { api, calls } = loadMessageTools({ raw: FULL, imap: { offline: true, serverRaw: "unused" } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, FULL);
    assert.equal(calls.serverStreams, 0);
    assert.equal(calls.streams, 1);
  });

  it("reports a failed server fetch as an error, never as a partial source", async () => {
    const { api } = loadMessageTools({ raw: PREFIX, imap: { offline: false, serverRaw: FULL.slice(0, 30), status: 0x80004005 } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, undefined);
    assert.equal(result.error, "Failed to read raw source");
  });
});

// Add-on providers (Owl for Exchange) keep the IMAP offline-store contract: the
// Offline flag marks a complete local copy, and the message service streams the
// rest. Owl also writes "From - <date>" plus X-Mozilla-* headers in front of
// every cached message; the mbox store escapes that line to ">From ".
describe("getMessage reads add-on provider (Owl) messages like IMAP", () => {
  const FULL = "Content-Type: multipart/mixed; boundary=b\r\n\r\n--b\r\nContent-Type: text/plain\r\n\r\nhello\r\n--b--\r\n";
  const ENVELOPES = [">From - Sat Oct 24 2026 04:51:38\r\n", ">>From - escaped twice\r\n", "From - Fri Oct 23 17:51:38 2026\n"];

  it("streams an Owl message without an offline copy from its message service", async () => {
    const { api, calls } = loadMessageTools({ streamError: new Error("not available offline"), remote: { type: "owl", offline: false, serverRaw: FULL } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, FULL);
    assert.equal(calls.serverStreams, 1);
    assert.equal(calls.streams, 0, "no local read without an offline copy");
  });

  it("reads the offline copy of an Owl message without touching the server", async () => {
    const { api, calls } = loadMessageTools({ raw: FULL, remote: { type: "owl", offline: true, serverRaw: "unused" } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, FULL);
    assert.equal(calls.serverStreams, 0);
    assert.equal(calls.streams, 1);
  });

  it("opens the stored copy through getLocalMsgStream when Thunderbird offers it", async () => {
    const { api, calls } = loadMessageTools({ raw: FULL, localStream: true, remote: { type: "owl", offline: true, serverRaw: "unused" } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, FULL);
    assert.equal(calls.localStreams, 1);
    assert.equal(calls.genericStreams, 0, "the unbounded folder stream is not opened");
  });

  it("falls back to getMsgInputStream on builds without getLocalMsgStream", async () => {
    const { api, calls } = loadMessageTools({ raw: FULL });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, FULL);
    assert.equal(calls.genericStreams, 1);
    assert.equal(calls.localStreams, 0);
  });

  it("streams from the message service when a stale Offline flag hides a missing copy", async () => {
    const { api, calls } = loadMessageTools({ streamError: new Error("missing"), remote: { type: "owl", offline: true, serverRaw: FULL } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, FULL);
    assert.equal(calls.serverStreams, 1);
  });

  for (const envelope of ENVELOPES) {
    it(`drops the stored envelope line ${JSON.stringify(envelope)} from local and streamed copies`, async () => {
      const local = await loadMessageTools({ raw: envelope + FULL, remote: { type: "owl", offline: true, serverRaw: "unused" } })
        .api.getMessage("message-1", "folder", false, "text", true);
      assert.equal(local.rawSource, FULL);
      const streamed = await loadMessageTools({ remote: { type: "owl", offline: false, serverRaw: envelope + FULL } })
        .api.getMessage("message-1", "folder", false, "text", true);
      assert.equal(streamed.rawSource, FULL);
    });
  }

  it("completes a provider slice that ends inside the final CRLF", async () => {
    const short = ENVELOPES[0] + FULL.slice(0, -1);
    assert.ok(short.endsWith("\r"));
    const local = await loadMessageTools({ raw: short, remote: { type: "owl", offline: true, serverRaw: "unused" } })
      .api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(local.rawSource, FULL);
    const streamed = await loadMessageTools({ remote: { type: "owl", offline: false, serverRaw: short } })
      .api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(streamed.rawSource, FULL);
    // Only a final bare CR is completed; one inside the message stays a hazard.
    const inner = await loadMessageTools({ raw: "Content-Type: text/plain\r\n\rvisible", remote: { type: "owl", offline: true, serverRaw: "unused" } })
      .api.getMessage("message-1", "folder", false, "text", false);
    assert.equal(inner.encryptedContentWithheld, true);
  });

  it("reads a cached message behind an envelope line without the encrypted-message opt-in", async () => {
    const { api } = loadMessageTools({ raw: ENVELOPES[0] + "X-Mozilla-Status: 0001\r\nContent-Type: text/plain\r\n\r\nvisible" });
    const result = await api.getMessage("message-1", "folder", false, "text", false);
    assert.equal(result.encryptedContentWithheld, undefined);
    assert.equal(result.body, "visible");
  });

  it("keeps an envelope-looking header-less line inside the body", async () => {
    const raw = "Content-Type: text/plain\r\n\r\n>From - Sat Oct 24 2026 04:51:38\r\nquoted";
    const { api } = loadMessageTools({ raw, remote: { type: "owl", offline: true, serverRaw: "unused" } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.rawSource, raw);
  });

  it("does not download for a local store, or for an empty server type", async () => {
    for (const type of ["none", "pop3", ""]) {
      const { api, calls } = loadMessageTools({ streamError: new Error("missing"), remote: { type, offline: false, serverRaw: FULL } });
      const result = await api.getMessage("message-1", "folder", false, "text", true);
      assert.equal(result.error, "Failed to read raw source", type);
      assert.equal(calls.serverStreams, 0, type);
    }
  });
});
