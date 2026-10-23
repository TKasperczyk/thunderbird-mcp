"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

// Searched test/{raw-mime-body,raw-mime-attachments,privacy-messages,
// direct-send-preflight}.test.cjs and test/helpers for VM loaders and MIME
// builders. Their loaders are private or include message-tool/native stubs.
// Follow their production marker + vm pattern; keep this corpus entirely inline.
const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
const source = fs.readFileSync(apiPath, "utf8");
const production = [
  "INLINE ATTACHMENT BASE64 HELPERS", "MCP TEXT SANITIZATION",
  "MESSAGE TEXT CONVERSION", "RAW MIME PARSING HELPERS", "ENCRYPTED MESSAGE GUARD",
].map(name => {
  const begin = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, begin);
  assert.ok(begin >= 0 && end > begin, `Missing production markers: ${name}`);
  return source.slice(begin, end);
}).join("\n");

function loadReadSafety() {
  const sandbox = { TextDecoder, Uint8Array, atob, btoa };
  vm.createContext(sandbox);
  vm.runInContext(production, sandbox, { filename: apiPath });
  return sandbox;
}

// Smaller versions of the local MIME builders found by the search above; no
// shared helper is exported and no fixtures are written to disk.
function entity(headers, body = "ordinary content") {
  return `${headers.join("\r\n")}\r\n\r\n${body}`;
}

function multipart(parts, { boundary = "read:(safe)?/,+='", padding = " \t", subtype = "mixed", preamble = "", epilogue = "", headers = [] } = {}) {
  return entity([`Content-Type: multipart/${subtype}; boundary="${boundary}"`, ...headers], [
    preamble,
    ...parts.map(part => `--${boundary}${padding}\r\n${part}`),
    `--${boundary}--${padding}`, epilogue,
  ].join("\r\n"));
}

const plain = entity(["Content-Type: text/plain; charset=utf-8"], "readable ordinary mail");
const unusual = multipart([plain]);
const api = loadReadSafety();

function assertFallback(raw, expected, diagnostic = "") {
  assert.equal(api.classifyRawMessageEncryption(raw, 0, { signedRaw: true }), "unknown", `must exercise tier 2: ${diagnostic}`);
  assert.equal(api.classifyRawMessageReadSafety(raw), expected, diagnostic);
}

describe("raw message read safety: strict classification is authoritative", () => {
  const signed = multipart([
    plain,
    entity(["Content-Type: application/pgp-signature"], "detached signature"),
  ], { boundary: "signature", padding: "", subtype: 'signed; protocol="application/pgp-signature"' });

  for (const [name, raw, expected] of [
    ["plain mail", plain, "clear"],
    ["UTF-8 mail", entity(["Content-Type: text/plain; charset=utf-8"], "hello"), "clear"],
    ["interpretable UTF-16 mail", entity([
      "Content-Type: text/plain; charset=utf-16le", "Content-Transfer-Encoding: base64",
    ], Buffer.from("ordinary content", "utf16le").toString("base64")), "clear"],
    ["interpretable stateful-charset mail", entity([
      "Content-Type: text/plain; charset=iso-2022-jp", "Content-Transfer-Encoding: base64",
    ], Buffer.from("ordinary content").toString("base64")), "clear"],
    ["detached signed mail", signed, "clear"],
    ["an interpretable encoded ordinary filename", entity([
      "Content-Type: application/pdf", 'Content-Disposition: attachment; filename="=?utf-8?Q?report.pdf?="',
    ]), "clear"],
    ["S/MIME", entity(["Content-Type: application/pkcs7-mime"], "MIIB"), "encrypted"],
    ["the S/MIME x- alias", entity(["Content-Type: application/x-pkcs7-mime"], "MIIB"), "encrypted"],
    ["PGP/MIME", multipart([plain], { boundary: "crypt", padding: "", subtype: "encrypted" }), "encrypted"],
    ["inline armor", entity(["Content-Type: text/plain"], "-----BEGIN PGP MESSAGE-----\r\nopaque"), "encrypted"],
    ["base64 inline armor", entity([
      "Content-Type: text/plain", "Content-Transfer-Encoding: base64",
    ], Buffer.from("-----BEGIN PGP MESSAGE-----\r\nopaque").toString("base64")), "encrypted"],
  ]) {
    it(`retains strict ${expected} for ${name} without scanning tier 2`, () => {
      const h = loadReadSafety();
      const strict = h.classifyRawMessageEncryption;
      assert.equal(strict(raw, 0, { signedRaw: true }), expected);
      const calls = [];
      h.classifyRawMessageEncryption = (...args) => {
        calls.push(args);
        return strict(...args);
      };
      h.hasRawMessageReadIndicators = () => { throw Error("tier 2 must not run"); };
      assert.equal(h.classifyRawMessageReadSafety(raw), expected);
      assert.equal(calls[0][0], raw);
      assert.equal(calls[0][1], 0);
      assert.deepEqual({ ...calls[0][2] }, { signedRaw: true });
    });
  }

  it("keeps the S4 strict decision unknown after a read becomes clear", () => {
    assertFallback(unusual, "clear");
    assert.equal(api.classifyRawMessageEncryption(unusual, 0, { signedRaw: true, rejectTruncatedHtml: true }), "unknown");
    assert.equal(api.checkSignedRawMessage(null, { raw: unusual }).state, "unknown");
  });

  it("uses no MIME tree or header parser during tier 2", () => {
    const h = loadReadSafety();
    const strict = h.classifyRawMessageEncryption;
    h.classifyRawMessageEncryption = (...args) => {
      const result = strict(...args);
      assert.equal(result, "unknown");
      for (const name of ["parseRawMimeEntity", "parseRawMimeHeaders", "parseRawMimeHeaderValue", "splitRawMimeMultipartBody", "findRawMimeHeaderBodySplit"]) {
        h[name] = () => { throw Error(`unexpected parser dependency: ${name}`); };
      }
      return result;
    };
    assert.equal(h.classifyRawMessageReadSafety(unusual), "clear");
  });

  it("fails closed if the indicator scan throws", () => {
    const h = loadReadSafety();
    h.normalizeRawReadMimeValue = () => { throw Error("uninterpretable value"); };
    assert.equal(h.classifyRawMessageReadSafety(unusual), "unknown");
  });
});

describe("raw message read safety: ordinary unusual MIME remains readable", () => {
  const cases = [
    ["RFC boundary punctuation and transport padding", unusual],
    ["spaces within the quoted boundary", multipart([plain], { boundary: "ordinary boundary" })],
    ["a folded boundary parameter", unusual.replace('; boundary=', ';\r\n\tboundary=')],
    ["a folded quoted boundary value", multipart([plain], { boundary: "ordinary boundary" }).replace('="ordinary boundary"', '="ordinary\r\n boundary"')],
    ["opening and closing delimiter padding", multipart([plain], { boundary: "simple", padding: " \t " })],
    ["ordinary boundary-like prose", multipart([entity(["Content-Type: text/plain"], 'text --simple marker; --simple-- in a sentence\r\n--simple-more-text')], { boundary: "simple" })],
    ["unbalanced prose quotes and parentheses", multipart([entity(["Content-Type: text/plain"], 'a "quote and an (unfinished thought')])],
    ["header-like ordinary prose", multipart([entity(["Content-Type: text/plain"], 'About Content-Type: text/plain\r\nContent-Disposition: inline\r\ntext follows')])],
    ["MIME fields mentioned in a sentence", entity(["Content-Type: text/plain (comment)"], 'I saw Content-Type: text/plain in an example.\r\nThe Content-Disposition: field needs a "quote here.')],
    ["ordinary marker prefixes", multipart([entity(["Content-Type: text/plain"], '-----BEGIN ORDINARY NOTE-----\r\nBEGIN PGP is a phrase, not armor')])],
    ["an indented body", entity(["Content-Type: text/plain (ordinary)"], "  indented line\r\n\tmore text")],
    ["non-ASCII body and Subject", multipart([entity(["Subject: Zażółć gęślą", "Content-Type: text/plain; charset=utf-8"], "Zażółć gęślą jaźń")])],
    ["encoded unrelated header", multipart([entity(["Subject: =?utf-8?Q?ordinary?=", "Content-Type: text/plain"], "hello")])],
    ["comments around a media type", entity(["Content-Type: (note) text (note) / (note) plain (note)"])],
    ["nested comments with quotes inside them", entity(['Content-Type: text/plain (outer (inner) "a quote); charset=utf-8'])],
    ["comments around parameters", entity(['Content-Type: text/plain; (note) charset (note) = (note) "utf-8" (note)'])],
    ["a comment inside a media token", entity(["Content-Type: te(note)xt/plain"])],
    ["ordinary name comments", multipart([entity(['Content-Type: application/pdf; name=report.(note)pdf', 'Content-Disposition: attachment; filename="annual report.pdf" (note)'])])],
    ["parentheses and semicolons inside quoted ordinary names", multipart([entity(['Content-Type: application/pdf; name="report (2026); final.pdf"'])])],
    ["identity encoding with comments", entity(["Content-Type: text/plain", "Content-Transfer-Encoding: (note) 8bit (note)"])],
    ["ordinary base64 attachment", multipart([plain, entity([
      'Content-Type: application/pdf; name="report.pdf"', "Content-Transfer-Encoding: base64",
    ], "JVBERi0xLjQK")])],
    ["ordinary quoted-printable leaf", multipart([entity([
      "Content-Type: text/plain; charset=utf-8", "Content-Transfer-Encoding: quoted-printable",
    ], "ordinary=20mail")])],
    ["ordinary uuencode", multipart([entity([], "begin 644 report.pdf\r\n$9&%T80``\r\n`\r\nend")])],
    ["ordinary yEnc", multipart([entity([], "=ybegin line=128 size=4 name=report.pdf\r\ndata\r\n=yend size=4")])],
    ["clear message/rfc822", multipart([entity(["Content-Type: message/rfc822"], plain)])],
    ["clear message/global", multipart([entity(["Content-Type: message/global"], plain)])],
    ["clear digest children", multipart([entity([], plain)], { subtype: "digest" })],
    ["clear name-inferred embedded mail", multipart([entity(['Content-Type: application/octet-stream; name="forward.eml"'], plain)])],
    ["all identity encodings in embedded mail", multipart([entity([
      "Content-Type: message/rfc822", "Content-Transfer-Encoding: binary",
    ], entity(["Content-Type: text/plain", "Content-Transfer-Encoding: 8bit"], "hello")), entity([
      "Content-Type: text/plain", "Content-Transfer-Encoding: 7bit",
    ])])],
    ["a benign preamble and epilogue", multipart([plain], { preamble: "reader preamble", epilogue: 'footer with --read:(safe)?/,+=\' marker text' })],
    ["LF-only message", unusual.replace(/\r\n/g, "\n")],
    ["missing final newline", unusual.trimEnd()],
    ["deep ordinary nesting beyond the strict tree limit", Array.from({ length: 15 }).reduce((raw, _, i) => multipart([raw], { boundary: `level-${i}`, padding: "" }), plain)],
  ];
  for (const [name, raw] of cases) {
    it(`reads ${name}`, () => assertFallback(raw, "clear", name));
  }
  it("accepts byte snapshots as well as strings", () => {
    assertFallback(new Uint8Array(Buffer.from(unusual, "latin1")), "clear");
    assertFallback(Buffer.from(unusual, "latin1"), "clear");
  });
});

describe("raw message read safety: every S/MIME and PGP type/protocol indication withholds", () => {
  const types = [
    "application/pkcs7-mime", "application/x-pkcs7-mime", "application/pkcs7-signature",
    "application/x-pkcs7-signature", "application/pkcs7", "application/smime", "application/s-mime",
    "application/pgp", "application/x-pgp", "application/pgp-encrypted", "application/pgp-signature",
    "application/pgp-keys", "application/x-pgp-message", "application/openpgp", "text/pgp",
    "multipart/encrypted",
  ];
  const variants = [
    ["literal", type => type],
    ["mixed case", type => type.toUpperCase()],
    ["folded", type => type.replace("/", " /\r\n\t")],
    ["comments", type => type.replace("/", " (outer (inner)) / (note) ")],
    ["in-token comment", type => type.slice(0, 3) + "(note)" + type.slice(3)],
    ["quotes", type => type.split("/").map(word => `"${word}"`).join("/")],
    ["split tokens", type => type.replace(/./g, "$& ")],
  ];
  for (const type of types) {
    for (const [variant, spell] of variants) {
      it(`withholds ${type}, ${variant}, as type and as protocol`, () => {
        const value = spell(type);
        for (const header of [`Content-Type: ${value}`, `Content-Type: multipart/signed; boundary=x; protocol="${value}"`]) {
          assertFallback(multipart([plain, entity([header])]), "unknown", header);
        }
      });
    }
  }
  it("treats smime-type alone as an indication without trusting its value", () => {
    for (const value of ["signed-data", "enveloped-data", "certs-only", "ordinary"]) {
      assertFallback(multipart([entity([`Content-Type: application/octet-stream; smime-type=${value}`])]), "unknown", value);
    }
  });
});

describe("raw message read safety: native filename routes and encodings withhold", () => {
  const extensions = ["p7m", "p7c", "p7z", "p7s", "p7b", "p7r", "p10", "p12", "pfx", "pem", "pkcs7", "pgp", "gpg", "asc", "sig"];
  for (const extension of extensions) {
    it(`withholds .${extension} through every native filename source`, () => {
      for (const filename of [`payload.${extension}`, `PAYLOAD.${extension.toUpperCase()}`, `payload. ${extension.split("").join(" ")}`]) {
        for (const headers of [
          [`Content-Type: application/octet-stream; name="${filename}"`],
          ["Content-Type: application/pdf", `Content-Disposition: attachment; filename="${filename}"`],
          ["Content-Type: text/plain", `Content-Name: ${filename}`],
          ["Content-Type: application/pdf", `X-Sun-Data-Name: "${filename}"`],
        ]) assertFallback(multipart([entity(headers)]), "unknown", headers.join(" | "));
      }
    });
  }
  for (const [name, headers] of [
    ["folded filename", ['Content-Disposition: attachment; filename="payload.\r\n p7m"']],
    ["comment in unquoted filename", ["Content-Disposition: attachment; filename=payload.p7(note)m"]],
    ["lower-priority Content-Type name", ['Content-Disposition: attachment; filename="report.pdf"', 'Content-Type: application/pdf; name="payload.p7m"']],
    ["lower-priority legacy name", ['Content-Disposition: attachment; filename="report.pdf"', "Content-Name: payload.asc"]],
    ["lowest-priority Sun name", ["Content-Name: report.pdf", "X-Sun-Data-Name: payload.gpg"]],
  ]) it(`withholds ${name}`, () => assertFallback(multipart([entity(headers)]), "unknown", name));

  for (const [name, body] of [
    ["uuencoded protected data", "begin 644 payload.p7m\r\nopaque\r\nend"],
    ["uuencoded signature", "begin 0644 payload.p7s\r\nopaque\r\nend"],
    ["yEnc protected data", "=ybegin line=128 size=4 name=payload.pgp\r\nopaque"],
    ["yEnc signature", "=ybegin line=128 size=4 name=payload.sig\r\nopaque"],
    ["BinHex opaque filename", "(This file must be converted with BinHex 4.0)\r\nopaque"],
    ["encoded file with an uninterpretable name", "begin 644 =?utf-8?Q?report.pdf?=\r\nopaque\r\nend"],
  ]) it(`withholds ${name}`, () => assertFallback(multipart([entity([], body)]), "unknown", name));
});

describe("raw message read safety: uninterpretable routing metadata withholds", () => {
  const cases = [
    ["RFC 2047 Q name", 'Content-Disposition: attachment; filename="=?utf-8?Q?payload=2Ep7m?="'],
    ["RFC 2047 B name", 'Content-Type: application/octet-stream; name="=?utf-8?B?cGF5bG9hZC5wN20=?="'],
    ["RFC 2047 ordinary name", 'Content-Disposition: attachment; filename="=?utf-8?Q?report.pdf?="'],
    ["RFC 2047 media type", "Content-Type: =?utf-8?Q?application/pkcs7-mime?="],
    ["RFC 2047 legacy name", "Content-Name: =?utf-8?Q?report.pdf?="],
    ["RFC 2047 Sun name", "X-Sun-Data-Name: =?utf-8?Q?report.pdf?="],
    ["malformed encoded word", 'Content-Disposition: attachment; filename="=?bad"'],
    ["spaced encoded word", 'Content-Disposition: attachment; filename="= ?utf-8?Q?report.pdf?="'],
    ["comment-split encoded word", 'Content-Type: application/pdf; name==(note)?utf-8?Q?report.pdf?='],
    ["RFC 2231 extended name", "Content-Type: application/pdf; name*=utf-8''report.pdf"],
    ["RFC 2231 extended filename", "Content-Disposition: attachment; filename*=utf-8''payload%2Ep7m"],
    ["RFC 2231 continued name", 'Content-Type: application/pdf; name*0="payload.p7"; name*1="m"'],
    ["RFC 2231 extended continuation", "Content-Disposition: attachment; filename*0*=utf-8''payload; filename*1*=%2Ep7m"],
    ["RFC 2231 boundary", "Content-Type: multipart/mixed; boundary*=utf-8''ordinary"],
    ["RFC 2231 protocol", "Content-Type: multipart/signed; protocol*=utf-8''application%2Fpgp-encrypted"],
    ["RFC 2231 charset", "Content-Type: text/plain; charset*=utf-8''utf-8"],
    ["backslash media type", "Content-Type: application/pkcs7\\-mime"],
    ["backslash filename", 'Content-Disposition: attachment; filename="payload.p7\\m"'],
    ["backslash ordinary boundary", 'Content-Type: multipart/mixed; boundary="ordinary\\boundary"'],
    ["escaped quote", 'Content-Disposition: attachment; filename="report\\".pdf"'],
    ["percent escape", 'Content-Disposition: attachment; filename="payload%2Ep7m"'],
    ["spaced percent escape", 'Content-Disposition: attachment; filename="payload% 2Ep7m"'],
    ["comment-split percent escape", 'Content-Disposition: attachment; filename=payload%(note)2Ep7m'],
    ["non-ASCII media type", "Content-Type: application/pkcs7-mimé"],
    ["non-ASCII disposition", "Content-Disposition: attachmént"],
    ["non-ASCII type parameter", 'Content-Type: application/pdf; name="résumé.pdf"'],
    ["non-ASCII disposition parameter", 'Content-Disposition: attachment; filename="résumé.pdf"'],
    ["non-ASCII legacy name", "Content-Name: résumé.pdf"],
    ["non-ASCII Sun name", "X-Sun-Data-Name: résumé.pdf"],
    ["non-ASCII in comments", "Content-Type: text/plain (résumé)"],
    ["unterminated comment", "Content-Type: text/plain (note"],
    ["unmatched closing comment", "Content-Type: text/plain note)"],
    ["unterminated quote", 'Content-Type: text/plain; charset="utf-8'],
    ["quoted parameter suffix", 'Content-Disposition: attachment; filename="report.pdf"junk'],
    ["duplicate parameter", "Content-Type: text/plain; charset=utf-8; charset=us-ascii"],
    ["bare parameter", "Content-Type: text/plain; mystery"],
    ["malformed parameter assignment", "Content-Disposition: attachment; filename==report.pdf"],
    ["empty media type", "Content-Type:"],
    ["invalid media type", "Content-Type: ordinary"],
    ["comma-separated media types", "Content-Type: text/plain, application/pdf"],
    ["unknown transfer encoding", "Content-Transfer-Encoding: x-opaque"],
    ["uuencode transfer encoding", "Content-Transfer-Encoding: x-uuencode"],
    ["non-ASCII transfer encoding", "Content-Transfer-Encoding: bäsé64"],
    ["empty transfer encoding", "Content-Transfer-Encoding:"],
    ["control-byte escape", "Content-Disposition: attachment; filename=report\x1B(B.pdf"],
    ["NUL", "Content-Type: text/plain\0"],
    ["DEL", "Content-Disposition: attachment; filename=report\x7F.pdf"],
    ["bare CR", "Content-Type: text/plain\rContent-Type: application/pkcs7-mime"],
    ["folded field name", "Content-\r\n Type: text/plain"],
  ];
  for (const [name, header] of cases) {
    it(`withholds ${name} at nested depth`, () => {
      assertFallback(multipart([multipart([entity([header])])]), "unknown", name);
    });
  }
  for (const header of ["Content-Type: text/plain", "Content-Disposition: inline", "Content-Transfer-Encoding: 7bit", "Content-Name: report.pdf", "X-Sun-Data-Name: report.pdf"]) {
    it(`withholds duplicate ${header.split(":")[0]} in one header run`, () => {
      assertFallback(multipart([entity([header, header])]), "unknown", header);
    });
  }
  for (const raw of [undefined, null, {}, [], 42, "", "Content-Type: text/plain", "Subject: plain\n", "Content-Type: text/plain (comment)\r\n", "invalid header\r\n\r\ntext"]) {
    it(`withholds an unreadable or incomplete snapshot: ${JSON.stringify(raw)}`, () => {
      assert.equal(api.classifyRawMessageReadSafety(raw), "unknown");
    });
  }
});

describe("raw message read safety: embedded decoding is global, not a guessed tree relationship", () => {
  const embeddedTypes = ["message/rfc822", "message/news", "message/global", "message/external-body", "message/unknown", "multipart/digest"];
  for (const type of embeddedTypes) {
    for (const encoding of ["base64", "quoted-printable", "(comment) BASE64 (comment)", '"quoted-printable"']) {
      it(`withholds ${type} with ${encoding} anywhere in the snapshot`, () => {
        const embeddedPart = entity([`Content-Type: ${type}`], plain);
        const encodedPart = entity(["Content-Type: application/pdf", `Content-Transfer-Encoding: ${encoding}`], "ZGF0YQ==");
        for (const raw of [
          multipart([embeddedPart, encodedPart]),
          multipart([encodedPart, embeddedPart]),
          multipart([multipart([embeddedPart]), multipart([encodedPart])]),
          multipart([entity([`Content-Type: ${type}`, `Content-Transfer-Encoding: ${encoding}`], "ZGF0YQ==")]),
          multipart([embeddedPart], { epilogue: `Content-Transfer-Encoding: ${encoding}\r\n` }),
        ]) assertFallback(raw, "unknown", `${type}, ${encoding}`);
      });
    }
  }
  for (const extension of ["eml", "mail", "art", "nws", "mht", "mhtml", "msg"]) {
    it(`treats .${extension} filenames as possible embedded messages`, () => {
      assertFallback(multipart([
        entity([`Content-Type: application/octet-stream; name="forward.${extension}"`], plain),
        entity(["Content-Type: application/pdf", "Content-Transfer-Encoding: base64"], "ZGF0YQ=="),
      ]), "unknown");
    });
  }
  it("withholds encoded inner headers that contain no visible crypto spelling", () => {
    const inner = entity(["Content-Type: application/pkcs7-mime"], "MIIB");
    for (const [encoding, body] of [
      ["base64", Buffer.from(inner).toString("base64")],
      ["quoted-printable", Array.from(Buffer.from(inner), b => `=${b.toString(16).padStart(2, "0")}`).join("")],
    ]) assertFallback(multipart([entity(["Content-Type: message/rfc822", `Content-Transfer-Encoding: ${encoding}`], body)]), "unknown");
  });
});

describe("raw message read safety: whole-snapshot indications and armor", () => {
  const hazard = entity(["Content-Type: application/pkcs7-mime"], "MIIB");
  for (const [name, raw] of [
    ["preamble", multipart([plain], { preamble: hazard })],
    ["epilogue", multipart([plain], { epilogue: hazard })],
    ["body-looking text", multipart([entity(["Content-Type: text/plain"], hazard)])],
    ["header after a mismatched boundary", multipart([plain], { epilogue: `--different\r\n${hazard}\r\n--different--` })],
    ["deeply nested content", Array.from({ length: 15 }).reduce((raw, _, i) => multipart([raw], { boundary: `deep-${i}`, padding: "" }), hazard)],
    ["after the HTML display limit", multipart([plain], { epilogue: "a".repeat(2 * 1024 * 1024 + 64) + `\r\n${hazard}` })],
  ]) it(`withholds a type indication in the ${name}`, () => assertFallback(raw, "unknown", name));

  for (const marker of [
    "-----BEGIN PGP MESSAGE-----", "-----END PGP MESSAGE-----",
    "-----BEGIN PGP SIGNED MESSAGE-----", "-----BEGIN PGP SIGNATURE-----",
    "-----BEGIN PGP PUBLIC KEY BLOCK-----", "-----BEGIN PGP PRIVATE KEY BLOCK-----",
    "-----begin pgp message-----", "-----BEGIN\r\n PGP\tMESSAGE-----",
    "-----B E G I N P G P MESSAGE-----", '-----"BEGIN" "PGP" MESSAGE-----',
    "-----BE\u200BGIN PG\u2060P MESSAGE-----",
  ]) {
    it(`withholds armor ${JSON.stringify(marker)} throughout the snapshot`, () => {
      for (const raw of [
        multipart([entity(["Content-Type: text/plain"], marker)]),
        multipart([plain], { preamble: marker }),
        multipart([plain], { epilogue: `ordinary text quoting ${marker} inline` }),
        multipart([entity(["Content-Type: text/html"], `<p>${marker}</p>`)]),
      ]) assertFallback(raw, "unknown", marker);
    });
  }
  it("scans armor beyond presentation limits", () => {
    assertFallback(multipart([entity(["Content-Type: text/html"], "a".repeat(2 * 1024 * 1024 + 64) + "-----BEGIN PGP MESSAGE-----")]), "unknown");
  });
  for (const encoding of ["base64", "quoted-printable"]) {
    for (const variant of ["ordinary", "junk", "chunk boundary", "UTF-16"]) {
      it(`finds ${encoding} armor with ${variant} across all raw quartet alignments`, () => {
        const marker = "-----BEGIN PGP MESSAGE-----";
        const prefix = variant === "chunk boundary" ? "x".repeat(4093) : "";
        const bytes = Buffer.from(prefix + marker, variant === "UTF-16" ? "utf16le" : "utf8");
        let encoded = encoding === "base64" ? bytes.toString("base64")
          : Array.from(bytes, byte => `=${byte.toString(16).padStart(2, "0")}`).join("");
        if (variant === "junk") encoded = encoding === "base64" ? encoded.split("").join("!\r\n")
          : encoded.replace(/(.{6})/g, "$1=\r\n");
        for (let padding = 0; padding < 4; padding++) {
          const raw = multipart([entity([
            "X-Ordinary: " + "a".repeat(padding),
            "Content-Type: text/plain", `Content-Transfer-Encoding: ${encoding}`,
          ], encoded)]);
          assertFallback(raw, "unknown", `${variant}, ${encoding}, alignment ${padding}`);
        }
      });
    }
  }
  it("retains QP escapes and soft line breaks at scanner chunk boundaries", () => {
    for (const boundaryOffset of [4094, 4095, 4096]) {
      const lead = multipart([entity(["Content-Type: text/plain", "Content-Transfer-Encoding: quoted-printable"], "")]);
      const start = lead.indexOf("\r\n\r\n", lead.indexOf("Content-Transfer-Encoding")) + 4;
      for (const escape of ["=2D", "=\r\n-"]) {
        const raw = lead.slice(0, start) + "x".repeat(boundaryOffset - start) + escape + "----BEGIN PGP MESSAGE-----" + lead.slice(start);
        assertFallback(raw, "unknown", `${escape}, ${boundaryOffset}`);
      }
    }
  });
  it("handles long ordinary slash runs in a filename", () => {
    assertFallback(entity(['Content-Type: application/octet-stream (comment); name="' + "/".repeat(200000) + 'ordinary.txt"']), "clear");
  });
});

describe("raw message read safety: HTML and charset transformations", () => {
  // No literal END marker: each fixture must be caught by its transformed
  // opening marker, not by a second, already visible armor indication.
  const armor = "-----BEGIN PGP MESSAGE-----\r\nopaque";
  const htmlCases = [
    ["decimal hyphen references", armor.replaceAll("-", "&#45;")],
    ["hexadecimal hyphen references", armor.replaceAll("-", "&#X002D;")],
    ["semicolonless references", "&#45-&#x2d--BEGIN PGP MESSAGE-----"],
    ["an entity in BEGIN", armor.replace("BEGIN", "B&#69;GIN")],
    ["an entity in PGP", armor.replace("PGP", "P&#71;P")],
    ["named invisible references", armor.replace("BEGIN", "BE&NegativeThinSpace;GIN")],
    ["markup between every opening dash", "-<b></b>-<i></i>-<span></span>--BEGIN PGP MESSAGE-----"],
    ["markup inside BEGIN", armor.replace("BEGIN", "BE<b>GIN</b>")],
    ["quoted greater-than in a tag", armor.replace("BEGIN", 'BE<span title=">">GIN</span>')],
    ["comments inside BEGIN", armor.replace("BEGIN", "BE<!-- hidden > markup -->GIN")],
    ["malformed markup inside BEGIN", armor.replace("BEGIN", "BE<b a='unterminated>GIN")],
    ["a long tag inside BEGIN", armor.replace("BEGIN", `BE<span title="${"x".repeat(8192)}">GIN</span>`)],
    ["charset escape selected by HTML metadata", '<meta http-equiv="content-type" content="text/html; charset=iso-2022-jp">' + armor.replaceAll("PGP", "P\x1b(BGP")],
    ["charset escape inside a numeric reference", armor.replaceAll("-", "&#4\x1b(B5;")],
    ["charset escape inside a reference introducer", armor.replaceAll("-", "&\x1b(B#45;")],
  ];
  for (const [name, html] of htmlCases) {
    for (const encoding of ["8bit", "base64", "quoted-printable"]) {
      it(`withholds ${name} after ${encoding}, without a DOM or MIME tree`, () => {
        const bytes = Buffer.from(`<pre>${html}</pre>`, "utf8");
        // libmime tolerates non-alphabet base64 bytes and QP soft breaks.
        // The unusual container forces Tier 2 for valid encodings as well.
        const body = encoding === "base64" ? bytes.toString("base64").split("").join("!\r\n")
          : encoding === "quoted-printable" ? Array.from(bytes, byte => `=${byte.toString(16).padStart(2, "0")}`).join("=\r\n")
            : bytes.toString("latin1");
        for (let alignment = 0; alignment < 4; alignment++) {
          assertFallback(multipart([entity([
            "X-Ordinary: " + "a".repeat(alignment), "Content-Type: text/html; charset=utf-8",
            `Content-Transfer-Encoding: ${encoding}`,
          ], body)]), "unknown", `${name}, ${encoding}, alignment ${alignment}`);
        }
      });
    }
  }

  it("retains transformation indicators across every decoded chunk split", () => {
    for (const text of ["&#45;", "&#X0002d", "-----BE<span>GIN", "-----BE&Tab;GIN", "-\x1b(B", "&#4\x1b(B5;", "&\x1b(B#45;"]) {
      for (let split = 0; split <= text.length; split++) {
        assert.equal(api.hasRawReadPgpArmor([text.slice(0, split), text.slice(split)]), true, `${JSON.stringify(text)}, split ${split}`);
      }
    }
  });

  it("keeps numeric reference state bounded without a reference-length cutoff", () => {
    assertFallback(multipart([entity(["Content-Type: text/html"], `&#${"0".repeat(32768)}45;BEGIN PGP`)]), "unknown");
    assertFallback(multipart([entity(["Content-Type: text/html"], `&#${"9".repeat(32768)};ordinary text`)]), "clear");
    assertFallback(multipart([entity(["Content-Type: text/html"], "x".repeat(2 * 1024 * 1024 + 64) + "&#45;BEGIN PGP")]), "unknown");
  });

  it("withholds the original top-level malformed-base64 HTML reproduction", () => {
    const raw = entity(["Content-Type: text/html", "Content-Transfer-Encoding: base64"],
      Buffer.from(`<pre>${armor.replaceAll("-", "&#45;")}</pre>`).toString("base64").split("").join("!"));
    assertFallback(raw, "unknown");
  });

  for (const charset of ["", "x-unknown", "iso-2022-jp", "csiso2022jp", "utf-16", "utf-16le", "utf-16be", "utf-32", "replacement", "gbk", "gb18030", "big5", "euc-jp", "shift_jis", "euc-kr"]) {
    it(`withholds uninterpretable charset ${JSON.stringify(charset)} in every declaration source`, () => {
      for (const headers of [
        [`Content-Type: text/plain; charset="${charset}"`],
        ["Content-Type: text/plain", `X-Sun-Charset: "${charset}"`],
      ]) {
        assertFallback(multipart([entity(headers)]), "unknown");
        assertFallback(multipart([plain], { epilogue: entity(headers) }), "unknown");
      }
    });
  }

  it("withholds multibyte trail bytes that obscure a converted armor marker", () => {
    const bytes = Buffer.concat([
      Buffer.from("-----BE"), Buffer.from("8136a437", "hex"),
      Buffer.from("GIN PGP MESSAGE-----\r\nopaque"),
    ]);
    assert.equal(new TextDecoder("gb18030").decode(bytes), "-----BE\u200bGIN PGP MESSAGE-----\r\nopaque");
    for (const encoding of ["8bit", "base64", "quoted-printable"]) {
      const body = encoding === "base64" ? bytes.toString("base64")
        : encoding === "quoted-printable" ? Array.from(bytes, byte => `=${byte.toString(16).padStart(2, "0")}`).join("")
          : bytes.toString("latin1");
      assertFallback(multipart([entity([
        "Content-Type: text/plain; charset=gb18030", `Content-Transfer-Encoding: ${encoding}`,
      ], body)]), "unknown", encoding);
    }
  });

  it("withholds charset-transformed armor after tolerant transfer decoding", () => {
    const stateful = armor.replaceAll("PGP", "P\x1b(BGP");
    assert.equal(new TextDecoder("iso-2022-jp").decode(Buffer.from(stateful)), armor);
    for (const body of [stateful, armor.replaceAll("-", "-\x1b(B")]) {
      for (const charset of [null, "iso-2022-jp"]) {
        for (const encoding of ["base64", "quoted-printable"]) {
          const bytes = Buffer.from(body);
          const encoded = encoding === "base64" ? bytes.toString("base64").split("").join("!")
            : Array.from(bytes, byte => `=${byte.toString(16).padStart(2, "0")}`).join("=\r\n");
          assertFallback(multipart([entity([
            "Content-Type: text/plain" + (charset ? `; charset=${charset}` : ""),
            `Content-Transfer-Encoding: ${encoding}`,
          ], encoded)]), "unknown", `${charset}, ${encoding}`);
        }
      }
    }
  });

  const ordinaryHtml = '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<!-- ordinary comment --><p>Well-known &amp; ordinary &#65; &#160; &#xA0; mail.</p>' +
    '<a href="https://example.invalid/?a=1&amp;b=2">Details</a><b title=">">Thanks</b></body></html>';
  const htmlPart = entity(["Content-Type: text/html; charset=utf-8"], ordinaryHtml);
  for (const [name, raw] of [
    ["alternative HTML", multipart([plain, htmlPart], { subtype: "alternative" })],
    ["mixed HTML and PDF", multipart([htmlPart, entity([
      "Content-Type: application/pdf", "Content-Transfer-Encoding: base64",
    ], "JVBERi0xLjQK")])],
    ["calendar and HTML", multipart([htmlPart, entity([
      "Content-Type: text/calendar; method=REQUEST; charset=utf-8",
    ], "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR")], { subtype: "alternative" })],
    ["forwarded HTML", multipart([entity(["Content-Type: message/rfc822"], htmlPart)])],
    ["tolerant base64 HTML", multipart([entity([
      "Content-Type: text/html; charset=utf-8", "Content-Transfer-Encoding: base64",
    ], Buffer.from(ordinaryHtml).toString("base64").split("").join("!"))])],
    ["quoted-printable HTML", multipart([entity([
      "Content-Type: text/html; charset=utf-8", "Content-Transfer-Encoding: quoted-printable",
    ], ordinaryHtml.replaceAll("=", "=3D"))])],
  ]) it(`still reads ordinary unusual ${name}`, () => assertFallback(raw, "clear", name));

  for (const charset of ["utf-8", "unicode-1-1-utf-8", "us-ascii", "ISO-8859-1", "windows-1252", "latin2", "koi8-r"]) {
    it(`accepts byte-scannable charset ${charset} and its aliases`, () => {
      assertFallback(multipart([entity([`Content-Type: text/plain; charset="${charset}"`])]), "clear");
      assertFallback(multipart([entity(["Content-Type: text/plain", `X-Sun-Charset: ${charset}`])]), "clear");
    });
  }
});

describe("raw message read safety: HTML comment delimiters are not armor prefixes", () => {
  const htmlPart = (html, encoding = null) => entity([
    "Content-Type: text/html; charset=utf-8", ...(encoding ? [`Content-Transfer-Encoding: ${encoding}`] : []),
  ], encoding === "base64" ? Buffer.from(html).toString("base64").split("").join("!")
    : encoding === "quoted-printable" ? html.replaceAll("=", "=3D") : html);

  for (const [name, html] of [
    ["Outlook conditional comments", '<html><body><!--[if !mso]><!--><p>ordinary mail</p><!--<![endif]--><!-- tracking --></body></html>'],
    ["a comment delimiter cut by an entity", "<p>x</p><!--&nbsp;-->"],
    ["a bogus comment", "<p>x</p><!-<![endif]>"],
  ]) {
    it(`still reads ${name}`, () => {
      for (const encoding of [null, "base64", "quoted-printable"]) {
        assertFallback(multipart([htmlPart(html, encoding)]), "clear", `${name}, ${encoding}`);
      }
    });
  }

  for (const [name, html] of [
    ["armor cut by a comment", "<p>-----<!-- -->BEGIN PGP MESSAGE-----</p>"],
    ["armor after a comment delimiter", "<!--<![endif]-->-----<i>BEGIN PGP MESSAGE-----</i>"],
    ["a dash run split by tags", "<p>-<b>-</b>-<b>-</b>-BEGIN PGP MESSAGE-----</p>"],
    ["armor inside a comment opener's dashes", "<!-----BEGIN PGP MESSAGE-----"],
    ["a comment-looking dash run longer than the tail window", `<!--${"-".repeat(20)}<b>BEGIN PGP MESSAGE-----</b>`],
    ["a dash run after a closed comment", "<!-- x -->--<!-- y -->---BEGIN PGP MESSAGE-----"],
  ]) it(`withholds ${name}`, () => assertFallback(multipart([htmlPart(html)]), "unknown", name));
});
