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
  assert.ok(start >= 0, "extractBodyContent start marker missing");
  assert.ok(end > start, "extractFormattedBody end marker missing");

  const sandbox = {
    htmlToMarkdown: html => `markdown:${html}`,
    stripHtml: html => html.replace(/<[^>]*>/g, ""),
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(start, end)}
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

  it("preserves message order outside multipart/alternative", () => {
    const message = {
      contentType: "multipart/mixed",
      parts: [
        { contentType: "text/html", body: "<p>First body</p>" },
        { contentType: "text/plain", body: "Later body" },
      ],
    };

    assertBody(extractFormattedBody(message, "text"), "First body", false);
  });
});
