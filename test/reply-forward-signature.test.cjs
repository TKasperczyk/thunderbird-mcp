"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const apiSource = fs.readFileSync(
  path.join(__dirname, "..", "extension", "mcp_server", "api.js"),
  "utf8"
);

// Mirrors the pure body assembly helper expected in api.js. The Thunderbird
// compose APIs are unavailable in Node, so behavior and integration are
// checked separately.
function buildBodyWithSignatureAndBlock(body, sigFragment, block, useHtml) {
  if (useHtml) {
    return `<html><head><meta charset="UTF-8"></head><body>${body}${sigFragment}${block}</body></html>`;
  }
  return `${body || ""}${sigFragment}${sigFragment ? "\n\n" : "\n\n"}${block}`;
}

function getFunctionSource(functionName, nextFunctionName) {
  const start = apiSource.indexOf(`function ${functionName}(`);
  const end = apiSource.indexOf(`function ${nextFunctionName}(`, start + 1);
  assert.notEqual(start, -1, `${functionName} should exist in api.js`);
  assert.notEqual(end, -1, `${nextFunctionName} should follow ${functionName}`);
  return apiSource.slice(start, end);
}

describe("Reply and forward signatures", () => {
  it("places the signature between reply text and quote in HTML", () => {
    assert.equal(
      buildBodyWithSignatureAndBlock("Hello", "<br><div>Sig</div>", "<br><blockquote>Quote</blockquote>", true),
      "<html><head><meta charset=\"UTF-8\"></head><body>Hello<br><div>Sig</div><br><blockquote>Quote</blockquote></body></html>"
    );
  });

  it("places the signature between reply text and quote in plain text", () => {
    assert.equal(
      buildBodyWithSignatureAndBlock("Hello", "\n\n-- \nSig", "On date wrote:\n> Quote", false),
      "Hello\n\n-- \nSig\n\nOn date wrote:\n> Quote"
    );
  });

  it("leaves the body and quote unchanged when there is no signature", () => {
    assert.equal(
      buildBodyWithSignatureAndBlock("Hello", "", "On date wrote:\n> Quote", false),
      "Hello\n\nOn date wrote:\n> Quote"
    );
  });

  it("places a forward signature before the forwarded block", () => {
    assert.equal(
      buildBodyWithSignatureAndBlock("Intro<br><br>", "<br><div>Sig</div>", "<blockquote>Forward</blockquote>", true),
      "<html><head><meta charset=\"UTF-8\"></head><body>Intro<br><br><br><div>Sig</div><blockquote>Forward</blockquote></body></html>"
    );
  });

  it("builds the identity signature in the reply skipReview branch", () => {
    const replySource = getFunctionSource("replyToMessage", "forwardMessage");
    const skipReviewStart = replySource.indexOf("if (skipReview) {");
    const directSend = replySource.indexOf("sendMessageDirectly(", skipReviewStart);
    assert.notEqual(skipReviewStart, -1);
    assert.notEqual(directSend, -1);
    assert.match(replySource.slice(skipReviewStart, directSend), /buildSignatureFragment\(/);
    assert.match(replySource.slice(skipReviewStart, directSend), /buildBodyWithSignatureAndBlock\(/);
  });

  it("builds the identity signature in the forward skipReview branch", () => {
    const forwardSource = getFunctionSource("forwardMessage", "getRecentMessages");
    const skipReviewStart = forwardSource.indexOf("if (skipReview) {");
    const directSend = forwardSource.indexOf("sendMessageDirectly(", skipReviewStart);
    assert.notEqual(skipReviewStart, -1);
    assert.notEqual(directSend, -1);
    assert.match(forwardSource.slice(skipReviewStart, directSend), /buildSignatureFragment\(/);
    assert.match(forwardSource.slice(skipReviewStart, directSend), /buildBodyWithSignatureAndBlock\(/);
  });
});
