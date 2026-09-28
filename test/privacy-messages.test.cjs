"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
// Same marker assertions and VM loading pattern as validation.test.cjs.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function loadMessageTools({ mime = { contentType: "text/plain", body: "visible" }, allowed = false, unreadable = false, raw = "raw MIME", DOMParser: Parser = null, streamError } = {}) {
  const calls = { options: [], streams: 0, sends: [], reviews: 0, logs: [] };
  const folder = {
    server: {},
    getUriForMsg: () => "message-uri",
    getMsgInputStream() {
      calls.streams++;
      if (streamError) throw streamError;
      return { close() {} };
    },
  };
  const hdr = {
    messageId: "message-1", subject: "protected subject", author: "sender@example.test",
    recipients: "reader@example.test", folder, date: 0,
  };
  const sandbox = {
    console: { error: (...args) => calls.logs.push(args), warn: (...args) => calls.logs.push(args) },
    DOMParser: Parser,
    TextDecoder,
    Uint8Array,
    atob,
    Services: { prefs: { getBoolPref() { if (unreadable) throw Error("unreadable"); return allowed; } } },
    ChromeUtils: { importESModule: () => ({
      MsgHdrToMimeMessage(msgHdr, _listener, callback, _download, options) {
        calls.options.push(options);
        callback(msgHdr, mime);
      },
    }) },
    findMessage: () => ({ msgHdr: hdr, folder }),
    getUserTags: () => [],
    getConfiguredGetMessagesLimit: () => 10,
    readMessageStreamFully: () => raw,
    isSkipReviewBlocked: () => false,
    filePathsToAttachDescs: () => ({ descs: [], failed: [] }),
    Cc: {
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({ setHeader() {} }) },
    },
    Ci: { nsIMsgCompType: { Reply: 1, ReplyAll: 2, ForwardInline: 3 }, nsIMsgCompDeliverMode: { Now: 0 }, nsIMsgFolder: {} },
    setComposeIdentity(params) { params.identity = {}; },
    resolveComposeFormat: () => ({ useHtml: false, format: 0 }),
    markMessageDispositionState() {},
    async sendMessageDirectly(fields, _identity, attachments) {
      calls.sends.push({ ...fields, attachments });
      return { success: true };
    },
    async openComposeWindowWithCustomizations() { calls.reviews++; return { success: true }; },
  };
  vm.createContext(sandbox);
  vm.runInContext([
    source.match(/^const PREF_\w+ = .+;$/gm).join("\n"),
    snippet("PRIVACY PREFERENCE HELPERS"), snippet("MCP TEXT SANITIZATION"),
    snippet("MESSAGE TEXT CONVERSION"), snippet("RAW MIME PARSING HELPERS"),
    snippet("INLINE ATTACHMENT BASE64 HELPERS"), snippet("ENCRYPTED MESSAGE GUARD"),
    snippet("MESSAGE READ TOOLS"), snippet("REPLY TOOL"), snippet("FORWARD TOOL"),
  ].join("\n"), sandbox);
  return { api: sandbox, calls };
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

describe("Encrypted message privacy", () => {
  for (const encrypted of encryptedTrees) {
    for (const rawSource of [false, true]) {
      it(`withholds ${encrypted.label}, rawSource=${rawSource}, before any content or attachment access`, async () => {
        const mime = { ...encrypted, get allUserAttachments() { return assert.fail("attachment metadata must not be read"); } };
        const { api, calls } = loadMessageTools({ mime });
        const result = await api.getMessage("message-1", "folder", true, "html", rawSource, true);
        assert.equal(result.encryptedContentWithheld, true);
        assert.match(result.body, /Encrypted message content withheld/);
        assert.equal(result.bodyIsHtml, false);
        assert.equal(result.attachments.length, 0);
        assert.equal(result.subject, "[Encrypted message]");
        assert.equal(result.rawSource, undefined);
        assert.equal(calls.streams, 0);
        assert.equal(calls.options[0].examineEncryptedParts, false);
        assert.doesNotMatch(JSON.stringify(result), /decrypted body|protected subject|ciphertext/);
      });
    }
  }

  it("fails closed when the encrypted preference cannot be read", async () => {
    const { api, calls } = loadMessageTools({ mime: encryptedTrees[2], unreadable: true });
    const result = await api.getMessage("message-1", "folder");
    assert.equal(result.encryptedContentWithheld, true);
    assert.equal(calls.options[0].examineEncryptedParts, false);
    assert.equal(calls.streams, 0);
  });

  it("getMessages inherits withholding for each item", async () => {
    const { api } = loadMessageTools({ mime: encryptedTrees[2] });
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
            if (allowed) assert.equal(reply.success, true);
            else assert.match(reply.error, /encrypted messages is blocked/);
          }
          assert.equal(calls.sends.length, allowed ? 2 : 0);
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

  it("does not fall back to raw data when MIME parsing fails", async () => {
    const { api, calls } = loadMessageTools({ mime: null });
    const result = await api.getMessage("message-1", "folder", true, "text", true);
    assert.match(result.error, /parse message/);
    assert.equal(calls.streams, 0);
  });

  for (const tool of ["replyToMessage", "forwardMessage"]) {
    const invoke = (api, skipReview) => tool === "replyToMessage"
      ? api.replyToMessage("message-1", "folder", "intro", false, false, undefined, undefined, undefined, undefined, undefined, skipReview)
      : api.forwardMessage("message-1", "folder", "recipient@example.test", "intro", false, undefined, undefined, undefined, undefined, skipReview);
    for (const unreadable of [false, true]) {
      it(`${tool} rejects direct encrypted sends with pref off/unreadable=${unreadable}`, async () => {
        const { api, calls } = loadMessageTools({ mime: encryptedTrees[2], unreadable });
        const result = await invoke(api, true);
        assert.match(result.error, /encrypted messages is blocked/);
        assert.match(result.error, /skipReview: false/);
        assert.equal(calls.sends.length, 0);
        assert.equal(calls.options[0].examineEncryptedParts, false);
      });
      it(`${tool} rejects armor from plaintext coercion with pref off/unreadable=${unreadable}`, async () => {
        const { api, calls } = loadMessageTools({
          mime: { parts: [], coerceBodyToPlaintext: () => "-----BEGIN PGP MESSAGE-----\nciphertext" }, unreadable,
        });
        const result = await invoke(api, true);
        assert.match(result.error, /encrypted messages is blocked/);
        assert.equal(calls.sends.length, 0);
      });
    }
    it(`${tool} permits armor from plaintext coercion after explicit opt-in`, async () => {
      const { api, calls } = loadMessageTools({
        mime: { parts: [], coerceBodyToPlaintext: () => "-----BEGIN PGP MESSAGE-----\nciphertext" }, allowed: true,
      });
      const result = await invoke(api, true);
      assert.equal(result.success, true);
      assert.match(calls.sends[0].body, /-----BEGIN PGP MESSAGE-----/);
    });
    it(`${tool} sends encrypted quoted content after explicit opt-in`, async () => {
      const { api, calls } = loadMessageTools({ mime: encryptedTrees[2], allowed: true });
      const result = await invoke(api, true);
      assert.equal(result.success, true);
      assert.match(calls.sends[0].body, /decrypted body/);
      assert.equal(calls.options[0].examineEncryptedParts, true);
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
          if (encrypted) assert.match(send.error, /encrypted messages is blocked/);
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
      assert.equal((await api.getMessage("message-1", "folder", false, "html", true)).encryptedContentWithheld, true);
      for (const send of await directSends(api)) assert.match(send.error, /encrypted messages is blocked/);
      assert.equal(calls.streams, 0);
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

  it("never classifies a standalone armor line created by the presentation cut", async () => {
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
    for (const send of await directSends(api)) assert.equal(send.success, true);
    assert.equal(calls.sends.length, 2);

    const raw = `Content-Type: text/html\n\n${html}`;
    const fallback = loadHtmlFixture(prefix, () => documentTree([elementNode("p", [textNode(header)])]), { raw, mime: { parts: [] } }).api;
    assert.equal(fallback.classifyRawMessageEncryption(raw), "clear");
    const recovered = await fallback.getMessage("message-1", "folder", false, "text");
    assert.notEqual(recovered.encryptedContentWithheld, true);
    assert.equal(recovered.body, `${header}\n\n[Message body truncated at 2 MiB]`);
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
        if (encrypted) assert.match(send.error, /encrypted messages is blocked/);
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
    assert.equal(api.isEncryptedMimeMessage({ contentType: "text/html", body: html }), true);
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
    for (const rawSource of [false, true]) {
      it(`withholds ${label} with empty structured MIME, rawSource=${rawSource}`, async () => {
        const { api, calls } = loadMessageTools({ raw, mime: {
          contentType: "message/rfc822", parts: [],
          get allUserAttachments() { return assert.fail("must classify before attachments"); },
        } });
        const result = await api.getMessage("message-1", "folder", true, "text", rawSource, true);
        assert.equal(result.encryptedContentWithheld, true);
        assert.equal(result.attachments.length, 0);
        assert.equal(result.rawSource, undefined);
        assert.match(result.body, /withheld/);
        assert.doesNotMatch(JSON.stringify(result), /ciphertext|protected subject/);
        assert.equal(calls.streams, 1);
        assert.equal(calls.options[0].examineEncryptedParts, false);
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
            const { api } = loadMessageTools({ raw, mime: { parts: [] } });
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

  it("does not guess when the MIME depth limit is reached", async () => {
    let raw = plain;
    for (let index = 0; index < 12; index++) raw = `Content-Type: multipart/mixed; boundary=b${index}\n\n--b${index}\n${raw}\n--b${index}--\n`;
    const { api } = loadMessageTools({ raw, mime: { parts: [] } });
    const result = await api.getMessage("message-1", "folder", false, "text", true);
    assert.equal(result.encryptedContentWithheld, true);
    assert.match(result.body, /could not be determined/);
  });

  it("withholds all content when fallback stream reading fails", async () => {
    const { api } = loadMessageTools({ mime: { parts: [] }, streamError: new Error("unreadable") });
    const result = await api.getMessage("message-1", "folder", true, "text");
    assert.equal(result.encryptedContentWithheld, true);
    assert.equal(result.attachments.length, 0);
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
