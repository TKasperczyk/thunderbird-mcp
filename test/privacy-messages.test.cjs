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
