"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const manifest = require("../extension/manifest.json");
const schema = require("../extension/mcp_server/schema.json");
const backgroundSource = fs.readFileSync(
  path.resolve(__dirname, "../extension/background.js"),
  "utf8"
);
const apiSource = fs.readFileSync(
  path.resolve(__dirname, "../extension/mcp_server/api.js"),
  "utf8"
);
const popupSource = fs.readFileSync(
  path.resolve(__dirname, "../extension/translation-popup.js"),
  "utf8"
);
const popupHtml = fs.readFileSync(
  path.resolve(__dirname, "../extension/translation-popup.html"),
  "utf8"
);

describe("Thunderbird inline translation button", () => {
  it("uses a message toolbar action with compact translation choices", () => {
    assert.equal(manifest.message_display_action.default_title, "AI / Basit Çeviri");
    assert.equal(manifest.message_display_action.default_popup, "translation-popup.html");
  });

  it("dispatches AI and local popup buttons to the inline translation API", () => {
    assert.match(popupSource, /translateDisplayedMessageInline\(/);
    assert.match(popupSource, /translate\("openai-fast"\)/);
    assert.match(popupSource, /translate\("openai-quality"\)/);
    assert.match(popupSource, /translate\("local"\)/);
    assert.match(popupSource, /AI çevirisi gösteriliyor/);
    assert.match(popupSource, /Yerel çeviri gösteriliyor/);
    assert.match(popupSource, /"restore"/);
    assert.match(popupHtml, /id="showOriginal"/);
    assert.match(popupHtml, /id="qualityTranslate"/);
    assert.match(popupHtml, /⚡ Hızlı AI/);
    assert.match(popupHtml, /Kaliteli AI/);
    assert.match(popupHtml, />Orijinal</);
    assert.doesNotMatch(backgroundSource, /messageDisplayAction\.onClicked\.addListener/);
  });

  it("declares HTML-preserving and inline experiment API methods", () => {
    const functions = schema[0].functions;
    const enableLocalTranslations = functions.find(
      item => item.name === "enableLocalTranslations"
    );
    const detectLanguage = functions.find(item => item.name === "detectLanguage");
    const translateText = functions.find(item => item.name === "translateText");
    const translateInline = functions.find(item => item.name === "translateDisplayedMessageInline");
    assert.ok(enableLocalTranslations);
    assert.ok(detectLanguage);
    assert.ok(translateText);
    assert.ok(translateInline);
    assert.ok(translateText.parameters.some(param => param.name === "isHtml"));
    assert.deepEqual(
      translateInline.parameters.map(param => param.name),
      ["tabId", "targetLanguage", "sourceLanguage", "provider"]
    );
  });

  it("translates the rendered HTML and subject while caching both originals", () => {
    assert.match(apiSource, /translator\.translate\(text, Boolean\(isHtml\)\)/);
    assert.match(apiSource, /const originalHtml = display\.body\.innerHTML/);
    assert.match(
      apiSource,
      /const originalSubject =\s+display\.canonicalSubject \|\| readSubject\(display\.subjectElement\)/
    );
    assert.match(apiSource, /display\.body\.innerHTML = result\.text/);
    assert.match(apiSource, /writeSubject\(display\.subjectElement, subjectResult\.text\)/);
    assert.match(apiSource, /subjectTitle: originalSubjectTitle/);
    assert.match(apiSource, /data-thunderbird-mcp-translated/);
    assert.match(apiSource, /LanguageDetector\.detectLanguage/);
    assert.match(popupSource, /"auto"/);
  });

  it("restores the canonical database subject even after an extension reload", () => {
    assert.match(apiSource, /displayedHeader\?\.mime2DecodedSubject/);
    assert.match(apiSource, /canonicalSubject: display\.canonicalSubject/);
    assert.match(
      apiSource,
      /original\.canonicalSubject \|\| display\.canonicalSubject \|\| original\.subject/
    );
    assert.match(apiSource, /if \(display\.canonicalSubject\) \{/);
  });

  it("enables Thunderbird's local translation actor during startup and keeps it enabled", () => {
    assert.match(apiSource, /browser\.translations\.enable/);
    assert.match(apiSource, /await ensureLocalTranslationsEnabled\(\)/);
    assert.match(backgroundSource, /await browser\.mcpServer\.enableLocalTranslations\(\)/);
    assert.doesNotMatch(apiSource, /restoreLocalTranslationsPreference/);
  });

  it("guards against a missing translation engine port with a restart hint", () => {
    assert.match(apiSource, /if \(port\) return port/);
    assert.match(apiSource, /Thunderbird'ü tamamen kapatıp yeniden açın/);
  });

  it("prevents duplicate translations and stale message updates", () => {
    assert.match(apiSource, /__tbMcpInlineTranslationsInFlight/);
    assert.match(apiSource, /Bu ileti zaten Türkçeye çevriliyor/);
    assert.match(apiSource, /currentDisplay\.messageURI !== display\.messageURI/);
  });

  it("detects the source language before inserting a Turkish progress notice", () => {
    assert.match(apiSource, /const originalTextContent = display\.body\.textContent \|\| ""/);
    assert.match(apiSource, /const detectionText = `\$\{originalSubject\}\\n\$\{originalTextContent\}`/);
  });

  it("keeps the AI button on OpenAI and does not silently fall back", () => {
    assert.match(apiSource, /provider must be auto, openai, openai-fast, openai-quality, local, or restore/);
    assert.match(apiSource, /providerMode === "restore"/);
    assert.match(apiSource, /const explicitOpenAI = providerMode\.startsWith\("openai"\)/);
    assert.match(apiSource, /explicitOpenAI && !segmentPackage/);
    assert.match(apiSource, /requestedProvider: providerMode/);
  });

  it("does not block detection when the model support catalogue is temporarily unavailable", () => {
    assert.match(apiSource, /supportStatus = "unverified"/);
    assert.match(apiSource, /supportedLanguage = detectedLanguage/);
    assert.match(apiSource, /supportLookupError/);
  });

  it("uses only official Mozilla translation catalogues with verified attachments", () => {
    assert.match(apiSource, /https:\/\/firefox\.settings\.services\.mozilla\.com\/v1\//);
    assert.match(apiSource, /translations-models-v2/);
    assert.match(apiSource, /translations-wasm-v2/);
    assert.match(apiSource, /credentials: "omit"/);
    assert.match(apiSource, /referrerPolicy: "no-referrer"/);
    assert.match(apiSource, /NetUtil\.newChannel/);
    assert.match(apiSource, /loadUsingSystemPrincipal: true/);
    assert.match(apiSource, /FIREFOX_TRANSLATION_DOWNLOAD_HOSTS\.has\(url\.hostname\)/);
    assert.match(apiSource, /Ci\.nsIRequest\.LOAD_ANONYMOUS/);
    assert.match(apiSource, /hiddenDOMWindow\.URL/);
    assert.match(apiSource, /hiddenDOMWindow\.Blob/);
    assert.match(apiSource, /actualHash !== cacheKey/);
    assert.match(apiSource, /attachment hash verification failed/);
    assert.match(apiSource, /useMockedTranslator: false/);
    assert.match(apiSource, /ensureOfficialMozillaTranslationClients\(TranslationsParent\)/);
  });
});
