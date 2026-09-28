"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadExtractFormattedBody() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const start = source.indexOf("function extractBodyContent(");
  const end = source.indexOf("function formatBodyHtml(", start);
  const sanitizeStart = source.indexOf("// BEGIN MCP TEXT SANITIZATION");
  const sanitizeEnd = source.indexOf("// END MCP TEXT SANITIZATION");
  assert.ok(sanitizeStart >= 0 && sanitizeEnd > sanitizeStart, "text sanitization markers missing");
  assert.ok(start >= 0, "extractBodyContent start marker missing");
  assert.ok(end > start, "extractFormattedBody end marker missing");

  const sandbox = {
    htmlToMarkdown: html => `markdown:${html}`,
    stripHtml: html => html.replace(/<[^>]*>/g, ""),
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(sanitizeStart, sanitizeEnd)}
${source.slice(start, end)}
this.extractFormattedBody = extractFormattedBody;`,
    sandbox
  );
  return sandbox.extractFormattedBody;
}

const extractFormattedBody = loadExtractFormattedBody();

function assertBody(result, body, bodyIsHtml) {
  assert.equal(result.body, body);
  assert.equal(result.bodyIsHtml, bodyIsHtml);
}

describe("structured MIME body extraction", () => {
  it("selects the requested multipart/alternative representation", () => {
    const message = {
      contentType: "multipart/alternative",
      parts: [
        { contentType: "text/plain; charset=utf-8", body: "Plain version" },
        { contentType: "text/html; charset=utf-8", body: "<p>HTML version</p>" },
      ],
    };

    assertBody(extractFormattedBody(message, "text"), "Plain version", false);
    assertBody(extractFormattedBody(message, "html"), "<p>HTML version</p>", true);
    assertBody(
      extractFormattedBody(message, "markdown"),
      "markdown:<p>HTML version</p>",
      false
    );
  });

  it("joins only the primary body's type in message order outside multipart/alternative", () => {
    const message = {
      contentType: "multipart/mixed",
      parts: [
        { contentType: "text/html", body: "<p>First body</p>" },
        { contentType: "text/plain", body: "Later body" },
        { contentType: "text/html", body: "<p>Last body</p>" },
        { contentType: "multipart/mixed", parts: [
          { contentType: "text/html", body: "<p>Nested HTML</p>" },
          { contentType: "text/plain", body: "\nNested plain" },
        ] },
      ],
    };

    assertBody(extractFormattedBody(message, "text"), "First bodyLast bodyNested HTML", false);
    assertBody(extractFormattedBody(message, "html"), "<p>First body</p><p>Last body</p><p>Nested HTML</p>", true);
    assertBody(extractFormattedBody(message, "markdown"), "markdown:<p>First body</p><p>Last body</p><p>Nested HTML</p>", false);
  });

  it("keeps both Apple Mail HTML fragments around an inline PDF", () => {
    const html = "<p>Before the PDF</p><p>After the PDF</p>";
    const message = {
      contentType: "message/rfc822", partName: "",
      allUserAttachments: [{ partName: "1.2.2", name: "report.pdf" }],
      parts: [{
        contentType: "multipart/alternative", partName: "1",
        parts: [
          { contentType: "text/plain", partName: "1.1", body: "Plain alternative" },
          { contentType: "multipart/mixed", partName: "1.2", parts: [
            { contentType: "text/html; charset=utf-8", partName: "1.2.1", body: "<p>Before the PDF</p>" },
            { contentType: "application/pdf", partName: "1.2.2", body: "PDF bytes" },
            { contentType: "text/html", partName: "1.2.3", body: "<p>After the PDF</p>" },
          ] },
        ],
      }],
    };
    assertBody(extractFormattedBody(message, "text"), "Plain alternative", false);
    assertBody(extractFormattedBody(message, "html"), html, true);
    assertBody(extractFormattedBody(message, "markdown"), `markdown:${html}`, false);
  });

  it("keeps a plain primary body when HTML fragments occur later", () => {
    const message = { contentType: "multipart/mixed", parts: [
      { contentType: "text/plain", body: "First plain\n" },
      { contentType: "multipart/related", parts: [
        { contentType: "text/html", body: "<p>First HTML</p>" },
        { contentType: "image/png", body: "image bytes" },
      ] },
      { contentType: "text/plain", body: "Last plain" },
      { contentType: "text/html", body: "<p>Last HTML</p>" },
    ] };
    assertBody(extractFormattedBody(message, "html"), "First plain\nLast plain", false);
    assertBody(extractFormattedBody(message, "markdown"), "First plain\nLast plain", false);
    assertBody(extractFormattedBody(message, "text"), "First plain\nLast plain", false);
  });

  it("does not replace an HTML main body with a plain unsubscribe footer", () => {
    const message = { contentType: "multipart/mixed", parts: [
      { contentType: "text/html", body: "<p>Main discussion</p>" },
      { contentType: "text/plain", body: "Unsubscribe footer" },
    ] };
    assertBody(extractFormattedBody(message, "text"), "Main discussion", false);
    assertBody(extractFormattedBody(message, "markdown"), "markdown:<p>Main discussion</p>", false);
    assertBody(extractFormattedBody(message, "html"), "<p>Main discussion</p>", true);
  });

  it("does not replace a plain main body with an HTML unsubscribe footer", () => {
    const message = { contentType: "multipart/mixed", parts: [
      { contentType: "text/plain", body: "Main discussion" },
      { contentType: "text/html", body: "<p>Unsubscribe footer</p>" },
    ] };
    for (const format of ["text", "markdown", "html"]) {
      assertBody(extractFormattedBody(message, format), "Main discussion", false);
    }
  });

  it("selects Outlook's related HTML while ignoring its image", () => {
    const message = {
      contentType: "multipart/alternative",
      allUserAttachments: [{ partName: "1.2.2", name: "image.png" }],
      parts: [
        { contentType: "text/plain", partName: "1.1", body: "Plain Outlook message" },
        { contentType: "multipart/related", partName: "1.2", parts: [
          { contentType: "text/html", partName: "1.2.1", body: '<p>Outlook message<img src="cid:image"></p>' },
          { contentType: "image/png", partName: "1.2.2", body: "image bytes" },
        ] },
      ],
    };
    assertBody(extractFormattedBody(message, "text"), "Plain Outlook message", false);
    assertBody(extractFormattedBody(message, "html"), '<p>Outlook message<img src="cid:image"></p>', true);
  });

  it("includes a mailing-list footer in message order without duplicating alternatives", () => {
    const message = {
      contentType: "multipart/mixed",
      parts: [
        { contentType: "multipart/alternative", parts: [
          { contentType: "text/plain", body: "Discussion\n" },
          { contentType: "text/html", body: "<p>Discussion</p>" },
        ] },
        { contentType: "text/plain", body: "\nList footer\n" },
        { contentType: "multipart/mixed", parts: [{ contentType: "text/plain", body: "Unsubscribe details" }] },
      ],
    };
    assertBody(extractFormattedBody(message, "text"), "Discussion\n\nList footer\nUnsubscribe details", false);
    assertBody(extractFormattedBody(message, "html"), "<p>Discussion</p>", true);
  });

  it("excludes text attachments by partName, skipping attached messages and null parts", () => {
    const message = {
      contentType: "message/rfc822",
      allUserAttachments: [null, { partName: "1.1", name: "notes.txt" }, { partName: "1.4", name: "page.html" }],
      parts: [{ contentType: "multipart/mixed", parts: [
        null,
        { contentType: "text/plain", partName: "1.1", body: "attached text secret" },
        { contentType: "text/plain", partName: "1.2", body: "Body\n" },
        { contentType: "message/rfc822", partName: "1.3", parts: [{ contentType: "text/plain", body: "attached message secret" }] },
        { contentType: "text/html", partName: "1.4", body: "<p>attached HTML secret</p>" },
        { contentType: "text/plain", partName: "1.5", body: "Footer" },
        null,
      ] }],
    };
    for (const format of ["text", "html", "markdown"]) {
      assertBody(extractFormattedBody(message, format), "Body\nFooter", false);
    }
    assertBody(extractFormattedBody(null, "text"), "", false);
  });

  for (const brokenList of ["getter", "partial iteration"]) {
    it(`falls back to no attachment exclusion on ${brokenList} failure`, () => {
      const message = {
        contentType: "multipart/mixed",
        parts: [
          { contentType: "text/plain", partName: "1.1", body: "First\n" },
          { contentType: "text/plain", partName: "1.2", body: "Second" },
        ],
        get allUserAttachments() {
          if (brokenList === "getter") throw Error("unavailable");
          return [{ partName: "1.1" }, { get partName() { throw Error("unreadable record"); } }];
        },
      };
      assertBody(extractFormattedBody(message, "text"), "First\nSecond", false);
    });
  }

  it("preserves fallback and first-choice selection in alternatives, including nulls", () => {
    const message = { contentType: "multipart/alternative", parts: [
      null,
      { contentType: "multipart/mixed", parts: [
        { contentType: "text/html", body: "<p>A</p>" },
        null,
        { contentType: "text/html", body: "<p>B</p>" },
      ] },
      { contentType: "text/html", body: "<p>Unused second HTML choice</p>" },
    ] };
    assertBody(extractFormattedBody(message, "text"), "AB", false);
    assertBody(extractFormattedBody(message, "html"), "<p>A</p><p>B</p>", true);
  });
});
