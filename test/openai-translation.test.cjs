"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { describe, it } = require("node:test");

const root = path.join(__dirname, "..");
const apiSource = fs.readFileSync(
  path.join(root, "extension", "mcp_server", "api.js"),
  "utf8"
);
const schemaSource = fs.readFileSync(
  path.join(root, "extension", "mcp_server", "schema.json"),
  "utf8"
);
const optionsHtml = fs.readFileSync(
  path.join(root, "extension", "options.html"),
  "utf8"
);
const optionsJs = fs.readFileSync(
  path.join(root, "extension", "options.js"),
  "utf8"
);

describe("OpenAI high-quality inline translation", () => {
  it("uses gpt-4o-mini with stateless structured output", () => {
    assert.match(apiSource, /OPENAI_TRANSLATION_MODEL = "gpt-4o-mini"/);
    assert.match(apiSource, /store: false/);
    assert.match(apiSource, /type: "json_schema"/);
    assert.match(apiSource, /name: "email_translation"/);
    assert.match(apiSource, /temperature: 0/);
  });

  it("handles incomplete, refused, and fenced Responses API output", () => {
    assert.match(apiSource, /response\?\.status === "incomplete"/);
    assert.match(apiSource, /reason === "max_output_tokens"/);
    assert.match(apiSource, /content\?\.type === "refusal"/);
    assert.match(apiSource, /replace\(\/\^```\(\?:json\)\?/);
    assert.match(apiSource, /Math\.max\(2000,/);
  });

  it("stores the API key in Thunderbird's password manager", () => {
    assert.match(apiSource, /Services\.logins\.searchLoginsAsync/);
    assert.match(apiSource, /Services\.logins\.addLoginAsync/);
    assert.match(apiSource, /Services\.logins\.removeLoginAsync/);
    assert.doesNotMatch(apiSource, /Services\.logins\s*\.findLogins/);
    assert.doesNotMatch(apiSource, /setStringPref\([^\n]*apiKey/i);
    assert.match(optionsHtml, /type="password" id="openAIApiKey"/);
  });

  it("sends only bounded visible text segments", () => {
    assert.match(apiSource, /createTreeWalker/);
    assert.match(apiSource, /SHOW_TEXT/);
    assert.match(apiSource, /script,style,noscript,template,svg,code,pre/);
    assert.match(apiSource, /OPENAI_MAX_VISIBLE_CHARACTERS/);
    assert.match(apiSource, /OPENAI_MAX_TEXT_SEGMENTS/);
  });

  it("checks protected tokens and a monthly hard limit", () => {
    assert.match(apiSource, /preservesProtectedTranslationTokens/);
    assert.match(apiSource, /PREF_OPENAI_MONTHLY_LIMIT_CENTS/);
    assert.match(apiSource, /usage\.microdollars \+ estimatedMicrodollars > limitMicrodollars/);
    assert.match(apiSource, /addOpenAITranslationUsage/);
  });

  it("falls back to Mozilla and caches successful translations", () => {
    assert.match(apiSource, /__tbMcpOpenAITranslationCache/);
    assert.match(apiSource, /Mozilla yerel çevirisine geçiliyor/);
    assert.match(apiSource, /provider: "mozilla-local"/);
  });

  it("exposes safe settings without returning the secret", () => {
    assert.match(schemaSource, /getOpenAITranslationConfig/);
    assert.match(schemaSource, /setOpenAITranslationConfig/);
    assert.match(optionsJs, /getOpenAITranslationConfig/);
    assert.match(optionsJs, /setOpenAITranslationConfig/);
    assert.match(apiSource, /keyConfigured: Boolean\(await getOpenAITranslationApiKey\(\)\)/);
  });
});
