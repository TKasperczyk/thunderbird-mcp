"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadRawMimeBodyHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN RAW MIME PARSING HELPERS";
  const endMarker = "// END RAW MIME PARSING HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "raw MIME parsing helper start marker missing");
  assert.ok(end > start, "raw MIME parsing helper end marker missing");

  const sandbox = {
    Uint8Array,
    TextDecoder,
    atob: globalThis.atob,
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(start, end)}
this.extractBodyPartFromRawMime = extractBodyPartFromRawMime;`,
    sandbox
  );
  return sandbox.extractBodyPartFromRawMime;
}

const extractBodyPartFromRawMime = loadRawMimeBodyHelpers();

function mimePart({
  body,
  contentType = "text/plain",
  charset,
  cte = "7bit",
  disposition,
  extraHeaders = [],
}) {
  const headers = [
    `Content-Type: ${contentType}${charset ? `; charset="${charset}"` : ""}`,
    `Content-Transfer-Encoding: ${cte}`,
    ...extraHeaders,
  ];
  if (disposition) headers.push(`Content-Disposition: ${disposition}`);
  return `${headers.join("\n")}\n\n${body}`;
}

function multipart({
  parts,
  subtype = "mixed",
  boundary = "boundary",
  terminal = true,
  preamble = "",
  epilogue = "",
  lineEnding = "\n",
  foldedBoundary = false,
}) {
  const contentType = foldedBoundary
    ? `Content-Type: multipart/${subtype};\n boundary="${boundary}"`
    : `Content-Type: multipart/${subtype}; boundary="${boundary}"`;
  let raw = `${contentType}\nMIME-Version: 1.0\n\n`;
  if (preamble) raw += `${preamble}\n`;
  for (const part of parts) raw += `--${boundary}\n${part}\n`;
  if (terminal) raw += `--${boundary}--\n${epilogue}`;
  else raw = raw.slice(0, -1);
  return lineEnding === "\n" ? raw : raw.replace(/\n/g, lineEnding);
}

function bytesToQuotedPrintable(bytes) {
  return Array.from(bytes, byte => `=${byte.toString(16).padStart(2, "0").toUpperCase()}`).join("");
}

function utf8Bytes(text) {
  return Buffer.from(text, "utf8");
}

function hexBytes(hex) {
  return Buffer.from(hex, "hex");
}

function assertExtracted(raw, bodyFormat, expectedText, expectedIsHtml = false) {
  const result = extractBodyPartFromRawMime(raw, bodyFormat);
  assert.ok(result, "expected a body part");
  assert.equal(result.text, expectedText);
  assert.equal(result.isHtml, expectedIsHtml);
  return result;
}

describe("raw MIME body extraction", () => {
  it("chooses multipart/alternative by bodyFormat preference", () => {
    const raw = multipart({
      subtype: "alternative",
      boundary: "alt;quoted",
      foldedBoundary: true,
      parts: [
        mimePart({ body: "Plain version" }),
        mimePart({ body: "<p>HTML version</p>", contentType: "text/html" }),
      ],
    });

    assertExtracted(raw, "text", "Plain version");
    assertExtracted(raw, "html", "<p>HTML version</p>", true);
    assertExtracted(raw, "markdown", "<p>HTML version</p>", true);
  });

  it("falls back to the available alternative representation", () => {
    const htmlOnly = multipart({
      subtype: "alternative",
      parts: [mimePart({ body: "<b>HTML only</b>", contentType: "text/html" })],
    });
    const plainOnly = multipart({
      subtype: "alternative",
      parts: [mimePart({ body: "Plain only" })],
    });

    assertExtracted(htmlOnly, "text", "<b>HTML only</b>", true);
    assertExtracted(plainOnly, "html", "Plain only");
  });

  it("walks mixed and related parts depth-first while skipping text attachments", () => {
    const alternative = multipart({
      subtype: "alternative",
      boundary: "nested-alt",
      parts: [
        mimePart({ body: "Nested plain" }),
        mimePart({ body: "<p>Nested HTML</p>", contentType: "text/html" }),
      ],
    });
    const related = multipart({
      subtype: "related",
      boundary: "related",
      parts: [
        alternative,
        mimePart({ body: "aW1hZ2U=", contentType: "image/png", cte: "base64" }),
      ],
    });
    const raw = multipart({
      subtype: "mixed",
      boundary: "outer",
      parts: [
        mimePart({ body: "Ignore attachment", disposition: "attachment; filename=notes.txt" }),
        related,
        mimePart({ body: "Later body" }),
      ],
    });

    assertExtracted(raw, "text", "Nested plain");
    assertExtracted(raw, "html", "<p>Nested HTML</p>", true);

    const directMixed = multipart({
      parts: [
        mimePart({ body: "First plain body" }),
        mimePart({ body: "<p>Later HTML body</p>", contentType: "text/html" }),
      ],
    });
    assertExtracted(directMixed, "html", "First plain body");

    const htmlWithInlineImage = multipart({
      subtype: "related",
      parts: [
        mimePart({ body: "<p>HTML with CID image</p>", contentType: "text/html" }),
        mimePart({ body: "aW1hZ2U=", contentType: "image/png", cte: "base64" }),
      ],
    });
    assertExtracted(htmlWithInlineImage, "markdown", "<p>HTML with CID image</p>", true);
  });

  it("decodes each text part with its own charset and transfer encoding", () => {
    const fixtures = [
      {
        charset: "utf-8",
        cte: "quoted-printable",
        body: bytesToQuotedPrintable(utf8Bytes("Zażółć gęślą jaźń")),
        expected: "Zażółć gęślą jaźń",
      },
      {
        charset: "iso-8859-2",
        cte: "base64",
        body: hexBytes("5a61bff3b3e62067eab66cb1206a61bcf1").toString("base64"),
        expected: "Zażółć gęślą jaźń",
      },
      {
        charset: "windows-1250",
        cte: "8bit",
        body: hexBytes("50f8ed6c699a209e6c759d6f75e86bfd206bf9f22080").toString("latin1"),
        expected: "Příliš žluťoučký kůň €",
      },
    ];

    for (const [index, fixture] of fixtures.entries()) {
      const raw = multipart({
        boundary: `charset-${index}`,
        parts: [mimePart(fixture)],
      });
      assertExtracted(raw, "text", fixture.expected);
    }
  });

  it("handles CRLF, bare LF/CR, preamble, epilogue, and quoted boundaries", () => {
    for (const lineEnding of ["\r\n", "\n", "\r"]) {
      const raw = multipart({
        boundary: "line-endings;1",
        foldedBoundary: true,
        lineEnding,
        preamble: "This preamble is not a body.",
        epilogue: "This epilogue is not a body.",
        parts: [mimePart({ body: "Selected body" })],
      });
      assertExtracted(raw, "text", "Selected body");
    }
  });

  it("keeps the final part when the terminal boundary is missing", () => {
    const raw = multipart({
      boundary: "unterminated",
      terminal: false,
      parts: [
        mimePart({ body: "First attachment", disposition: "attachment" }),
        mimePart({ body: "Recovered final body" }),
      ],
    });
    assertExtracted(raw, "text", "Recovered final body");
  });

  it("skips message/rfc822 and malformed candidates before a valid body", () => {
    const forwarded = `Content-Type: message/rfc822\n\n${multipart({
      boundary: "forwarded",
      parts: [mimePart({ body: "Hidden forwarded body" })],
    })}`;
    const raw = multipart({
      parts: [
        forwarded,
        mimePart({ body: "not valid base64!", cte: "base64" }),
        mimePart({ body: "Visible body" }),
      ],
    });
    assertExtracted(raw, "text", "Visible body");
  });

  it("caps traversal at ten nested MIME levels", () => {
    function nested(levels) {
      let raw = mimePart({ body: "Deep body" });
      for (let i = 0; i < levels; i++) {
        raw = multipart({ boundary: `depth-${i}`, parts: [raw] });
      }
      return raw;
    }

    assertExtracted(nested(10), "text", "Deep body");
    assert.equal(extractBodyPartFromRawMime(nested(11), "text"), null);
  });

  it("preserves the singlepart decoder behavior", () => {
    const utf8Qp = mimePart({
      body: `${bytesToQuotedPrintable(utf8Bytes("soft break: "))}=\r\n${bytesToQuotedPrintable(utf8Bytes("żółć"))}`,
      charset: "utf-8",
      cte: "quoted-printable",
    });
    assertExtracted(utf8Qp, "text", "soft break: żółć");

    const htmlBytes = utf8Bytes("<p>raw HTML ✓</p>");
    const base64Html = mimePart({
      body: `${htmlBytes.toString("base64").slice(0, 8)}\r\n${htmlBytes.toString("base64").slice(8)}`,
      contentType: "text/html",
      charset: "utf-8",
      cte: "base64",
    });
    assertExtracted(base64Html, "markdown", "<p>raw HTML ✓</p>", true);

    const unknownCharset = mimePart({
      body: utf8Bytes("UTF-8 fallback ✓").toString("latin1"),
      charset: "x-unknown-charset",
      cte: "8bit",
    });
    const fallback = assertExtracted(unknownCharset, "text", "UTF-8 fallback ✓");
    assert.equal(fallback.charsetFallback, true);

    assertExtracted(
      mimePart({ body: "", contentType: "text/html" }),
      "html",
      "",
      true
    );
    assertExtracted("Content-Type: ; charset=utf-8\n\nMalformed plain default", "text", "Malformed plain default");
  });

  it("returns null for unsupported top-level content", () => {
    const raw = mimePart({ body: "JVBERi0=", contentType: "application/pdf", cte: "base64" });
    assert.equal(extractBodyPartFromRawMime(raw, "text"), null);
  });
});
