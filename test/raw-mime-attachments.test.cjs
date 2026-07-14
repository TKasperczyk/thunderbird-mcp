"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadRawMimeHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  function markedSnippet(startMarker, endMarker) {
    const start = source.indexOf(startMarker);
    const end = source.indexOf(endMarker);
    assert.ok(start >= 0, `${startMarker} missing`);
    assert.ok(end > start, `${endMarker} missing`);
    return source.slice(start, end);
  }

  const snippet = [
    markedSnippet("// BEGIN RAW MIME PARSING HELPERS", "// END RAW MIME PARSING HELPERS"),
    markedSnippet("// BEGIN RAW MIME ATTACHMENT HELPERS", "// END RAW MIME ATTACHMENT HELPERS"),
  ].join("\n");
  const sandbox = {
    Uint8Array,
    TextDecoder,
    atob: globalThis.atob,
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.parseAttachmentPartsFromRawMime = parseAttachmentPartsFromRawMime;
this.attachmentSaveLooksWrong = attachmentSaveLooksWrong;`,
    sandbox
  );
  return {
    parseAttachmentPartsFromRawMime: sandbox.parseAttachmentPartsFromRawMime,
    attachmentSaveLooksWrong: sandbox.attachmentSaveLooksWrong,
  };
}

const {
  parseAttachmentPartsFromRawMime,
  attachmentSaveLooksWrong,
} = loadRawMimeHelpers();

function utf8(bytes) {
  return Buffer.from(bytes).toString("utf8");
}

function latin1Bytes(bytes) {
  return Array.from(bytes);
}

function byFilename(parts) {
  return Object.fromEntries(parts.map(part => [part.filename, part]));
}

function multipart(boundary, parts, headers = "") {
  return `${headers || `Content-Type: multipart/mixed; boundary="${boundary}"`}
MIME-Version: 1.0

${parts.map(part => `--${boundary}
${part}`).join("\n")}
--${boundary}--
`;
}

function attachmentPart({ filename, contentType = "application/octet-stream", cte = "base64", body, disposition = "attachment" }) {
  return `Content-Type: ${contentType}
Content-Transfer-Encoding: ${cte}
Content-Disposition: ${disposition}; filename="${filename}"

${body}`;
}

describe("raw MIME attachment parser", () => {
  it("extracts multipart/alternative sibling attachments", () => {
    const raw = `Content-Type: multipart/alternative;
 boundary="===============9006705364210621469=="
MIME-Version: 1.0
From: test-sender@example.com
To: test-recipient@example.com
Subject: Repro: attachments hidden in multipart/alternative
Message-ID: <repro-mpa-attachments@example.com>
Date: Sun, 24 May 2026 12:00:00 +0200

--===============9006705364210621469==
Content-Type: text/html; charset="utf-8"
MIME-Version: 1.0
Content-Transfer-Encoding: base64

PGh0bWw+PGJvZHk+PHA+Qm9keSB0ZXh0LiBUd28gYXR0YWNobWVudHMgc2hvdWxkIGJlIHZpc2li
bGUuPC9wPjwvYm9keT48L2h0bWw+

--===============9006705364210621469==
Content-Type: application/octet-stream
MIME-Version: 1.0
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="test.xml"

PD94bWwgdmVyc2lvbj0nMS4wJz8+Cjxub3RlPlRoaXMgaXMgYSBkdW1teSBYTUwgYXR0YWNobWVu
dC48L25vdGU+Cg==

--===============9006705364210621469==
Content-Type: application/octet-stream
MIME-Version: 1.0
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="test.pdf"

JVBERi0xLjQKJWR1bW15LXBkZi1ieXRlcwplbmRvYmoKc3RhcnR4cmVmCjAKJSVFT0YK

--===============9006705364210621469==--
`;

    const parts = byFilename(parseAttachmentPartsFromRawMime(raw));
    assert.deepEqual(Object.keys(parts).sort(), ["test.pdf", "test.xml"]);
    assert.equal(utf8(parts["test.xml"].bytes), "<?xml version='1.0'?>\n<note>This is a dummy XML attachment.</note>\n");
    assert.equal(utf8(parts["test.pdf"].bytes), "%PDF-1.4\n%dummy-pdf-bytes\nendobj\nstartxref\n0\n%%EOF\n");
  });

  it("extracts a standard nested multipart attachment", () => {
    const alternative = `Content-Type: multipart/alternative; boundary="alt"

--alt
Content-Type: text/plain

Plain body
--alt
Content-Type: text/html

<p>HTML body</p>
--alt--`;
    const pdf = attachmentPart({
      filename: "nested.pdf",
      contentType: "application/pdf",
      body: "bmVzdGVkLXBkZg==",
    });

    const parts = parseAttachmentPartsFromRawMime(multipart("outer", [alternative, pdf]));
    assert.equal(parts.length, 1);
    assert.equal(parts[0].filename, "nested.pdf");
    assert.equal(parts[0].contentType, "application/pdf");
    assert.equal(utf8(parts[0].bytes), "nested-pdf");
  });

  it("opt-in parsing identifies CID images without treating regular image attachments as inline", () => {
    const regularAttachment = attachmentPart({
      filename: "photo.png",
      contentType: "image/png",
      body: "cmVndWxhci1hdHRhY2htZW50",
    });
    const related = `Content-Type: multipart/related; boundary="related"

--related
Content-Type: text/html; charset=utf-8

<p>See <img src="cid:diagram@example.test"></p>
--related
Content-Type: image/png
Content-ID: <diagram@example.test>
Content-Transfer-Encoding: base64

aW5saW5lLXBpeGVscw==
--related
${regularAttachment}
--related--`;
    const raw = multipart("outer", [related]);

    const defaultParts = parseAttachmentPartsFromRawMime(raw);
    assert.equal(defaultParts.length, 1);
    assert.equal(defaultParts[0].filename, "photo.png");
    assert.equal(defaultParts[0].isInline, false);

    const optInParts = parseAttachmentPartsFromRawMime(raw, { includeInlineImages: true });
    const inlinePart = optInParts.find(part => part.isInline);
    const regularPart = optInParts.find(part => part.filename === "photo.png");

    assert.ok(inlinePart);
    assert.equal(inlinePart.partName, "1.1.2");
    assert.equal(inlinePart.contentId, "diagram@example.test");
    assert.equal(inlinePart.contentType, "image/png");
    assert.equal(utf8(inlinePart.bytes), "inline-pixels");
    assert.ok(regularPart);
    assert.equal(regularPart.isInline, false);
  });

  it("decodes supported content-transfer-encoding variants", () => {
    const binary = String.fromCharCode(0, 1, 2, 255);
    const raw8bit = String.fromCharCode(0x63, 0x61, 0x66, 0xE9);
    const raw = multipart("cte", [
      attachmentPart({ filename: "base64.txt", cte: "base64", body: "aGVsbG8gYmFzZTY0" }),
      attachmentPart({ filename: "qp.txt", cte: "quoted-printable", body: "hello=20quoted=2Dprintable=0A" }),
      attachmentPart({ filename: "plain.txt", cte: "7bit", body: "plain ascii" }),
      attachmentPart({ filename: "eight.bin", cte: "8bit", body: raw8bit }),
      attachmentPart({ filename: "binary.bin", cte: "binary", body: binary }),
    ]);

    const parts = byFilename(parseAttachmentPartsFromRawMime(raw));
    assert.equal(utf8(parts["base64.txt"].bytes), "hello base64");
    assert.equal(utf8(parts["qp.txt"].bytes), "hello quoted-printable\n");
    assert.equal(utf8(parts["plain.txt"].bytes), "plain ascii");
    assert.deepEqual(latin1Bytes(parts["eight.bin"].bytes), [0x63, 0x61, 0x66, 0xE9]);
    assert.deepEqual(latin1Bytes(parts["binary.bin"].bytes), [0, 1, 2, 255]);
  });

  it("handles folded headers", () => {
    const raw = `Content-Type: multipart/mixed; boundary="fold"

--fold
Content-Type: application/octet-stream
Content-Transfer-Encoding: base64
Content-Disposition: attachment;
 filename="folded.txt"

Zm9sZGVk
--fold--
`;
    const parts = parseAttachmentPartsFromRawMime(raw);
    assert.equal(parts.length, 1);
    assert.equal(parts[0].filename, "folded.txt");
    assert.equal(utf8(parts[0].bytes), "folded");
  });

  it("handles CRLF, LF, and bare-CR line endings", () => {
    const lf = `Content-Type: multipart/mixed; boundary="lines"

--lines
Content-Type: application/octet-stream
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="lines.txt"

bGluZXM=
--lines--
`;
    for (const lineEnding of ["\r\n", "\n", "\r"]) {
      const raw = lf.replace(/\n/g, lineEnding);
      const parts = parseAttachmentPartsFromRawMime(raw);
      assert.equal(parts.length, 1);
      assert.equal(parts[0].filename, "lines.txt");
      assert.equal(utf8(parts[0].bytes), "lines");
    }
  });

  it("does not false-match boundary-prefix lines inside part bodies", () => {
    const raw = `Content-Type: multipart/mixed; boundary="strict"

--strict
Content-Type: application/octet-stream
Content-Transfer-Encoding: 7bit
Content-Disposition: attachment; filename="strict.txt"

alpha\r
--strict-extra\r
omega
--strict--
`;
    const parts = parseAttachmentPartsFromRawMime(raw);
    assert.equal(parts.length, 1);
    assert.equal(utf8(parts[0].bytes), "alpha\r\n--strict-extra\r\nomega");
  });

  it("returns empty for top-level single-part text/plain", () => {
    const raw = `Content-Type: text/plain; charset=utf-8
Content-Transfer-Encoding: 7bit

single part body
`;
    assert.equal(parseAttachmentPartsFromRawMime(raw).length, 0);
  });

  it("does not descend into message/rfc822 parts", () => {
    const forwarded = `Content-Type: message/rfc822

Content-Type: multipart/mixed; boundary="inner"

--inner
Content-Type: application/octet-stream
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="hidden.txt"

aGlkZGVu
--inner--
`;
    assert.equal(parseAttachmentPartsFromRawMime(multipart("outer", [forwarded])).length, 0);
  });

  it("fails gracefully on malformed MIME", () => {
    assert.equal(parseAttachmentPartsFromRawMime(`Content-Type: multipart/mixed

body`).length, 0);

    const missingClosing = `Content-Type: multipart/mixed; boundary="broken"

--broken
Content-Type: application/octet-stream
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="partial.txt"

cGFydGlhbA==
`;
    const partial = parseAttachmentPartsFromRawMime(missingClosing);
    assert.equal(partial.length, 1);
    assert.equal(utf8(partial[0].bytes), "partial");

    const unknownCte = multipart("badcte", [
      attachmentPart({ filename: "unknown.bin", cte: "x-unknown", body: "ignored" }),
    ]);
    assert.equal(parseAttachmentPartsFromRawMime(unknownCte).length, 0);
  });

  it("returns gracefully for deeply nested multipart input", () => {
    let part = attachmentPart({ filename: "too-deep.txt", body: "ZGVlcA==" });
    for (let i = 35; i >= 1; i--) {
      part = `Content-Type: multipart/mixed; boundary="deep${i}"

--deep${i}
${part}
--deep${i}--`;
    }
    const raw = multipart("root", [part]);
    let parts = null;
    assert.doesNotThrow(() => {
      parts = parseAttachmentPartsFromRawMime(raw);
    });
    assert.equal(parts.length, 0);
  });
});

describe("attachment save size heuristic", () => {
  it("allows exact empty and small tolerated variance", () => {
    assert.equal(attachmentSaveLooksWrong(0, 0), false);
    assert.equal(attachmentSaveLooksWrong(null, 42), false);
    assert.equal(attachmentSaveLooksWrong(100, 95), false);
    assert.equal(attachmentSaveLooksWrong(100, 99), false);
  });

  it("rejects zero-byte and significant mismatches", () => {
    assert.equal(attachmentSaveLooksWrong(null, 0), true);
    assert.equal(attachmentSaveLooksWrong(100, 0), true);
    assert.equal(attachmentSaveLooksWrong(100, 91), true);
    assert.equal(attachmentSaveLooksWrong(1000, 940), true);
  });
});
