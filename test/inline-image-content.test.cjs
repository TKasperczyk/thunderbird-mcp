"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadInlineImageContentHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN INLINE IMAGE CONTENT HELPERS";
  const endMarker = "// END INLINE IMAGE CONTENT HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "inline image helper start marker missing");
  assert.ok(end > start, "inline image helper end marker missing");

  const snippet = source.slice(start, end);
  const sandbox = { btoa: globalThis.btoa };
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.helpers = {
  MAX_INLINE_IMAGE_BASE64_BYTES,
  MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES,
  normalizeInlineImageMimeType,
  normalizeInlineImageContentId,
  correlateInlineImageRecords,
  getInlineImageContentIdReferences,
  orderInlineImageRecordsForBody,
  getBase64EncodedSize,
  getInlineImageSkipReason,
  encodeByteStringToBase64,
  setExtraMcpContentBlocks,
  buildToolResultContent,
};`,
    sandbox
  );
  return sandbox.helpers;
}

const {
  MAX_INLINE_IMAGE_BASE64_BYTES,
  MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES,
  normalizeInlineImageMimeType,
  normalizeInlineImageContentId,
  correlateInlineImageRecords,
  getInlineImageContentIdReferences,
  orderInlineImageRecordsForBody,
  getBase64EncodedSize,
  getInlineImageSkipReason,
  encodeByteStringToBase64,
  setExtraMcpContentBlocks,
  buildToolResultContent,
} = loadInlineImageContentHelpers();

describe("inline image MCP content helpers", () => {
  it("keeps the default single text block byte-identical", () => {
    const toolResult = {
      id: "message-1",
      body: "![diagram](cid:diagram@example.test)",
      attachments: [],
    };
    const expectedText = JSON.stringify(toolResult, null, 2);

    const content = buildToolResultContent(toolResult);

    assert.equal(content.length, 1);
    assert.equal(content[0].type, "text");
    assert.equal(content[0].text, expectedText);
  });

  it("places image blocks after text without serializing their base64 into text", () => {
    const toolResult = {
      attachments: [{
        name: "diagram.png",
        contentId: "diagram@example.test",
        isInline: true,
      }],
    };
    const imageBlock = {
      type: "image",
      data: "iVBORw0KGgo=",
      mimeType: "image/png",
    };

    setExtraMcpContentBlocks(toolResult, [imageBlock]);
    const content = buildToolResultContent(toolResult);

    assert.equal(content.length, 2);
    assert.equal(content[0].type, "text");
    assert.strictEqual(content[1], imageBlock);
    assert.equal(content[0].text.includes(imageBlock.data), false);
    assert.deepEqual(JSON.parse(content[0].text), toolResult);
  });

  it("calculates encoded sizes at base64 boundaries", () => {
    assert.equal(getBase64EncodedSize(0), 0);
    assert.equal(getBase64EncodedSize(1), 4);
    assert.equal(getBase64EncodedSize(3), 4);
    assert.equal(getBase64EncodedSize(4), 8);

    const maxRawBytes = Math.floor(MAX_INLINE_IMAGE_BASE64_BYTES / 4) * 3;
    assert.equal(getBase64EncodedSize(maxRawBytes), MAX_INLINE_IMAGE_BASE64_BYTES);
    assert.equal(getBase64EncodedSize(maxRawBytes + 1), MAX_INLINE_IMAGE_BASE64_BYTES + 4);
  });

  it("normalizes supported MIME types and rejects unsupported image formats", () => {
    assert.equal(normalizeInlineImageMimeType(" IMAGE/JPEG; name=photo.jpg "), "image/jpeg");
    assert.equal(getInlineImageSkipReason("image/png", 4, 0), "");
    assert.equal(getInlineImageSkipReason("image/jpeg", 4, 0), "");
    assert.equal(getInlineImageSkipReason("image/gif", 4, 0), "");
    assert.equal(getInlineImageSkipReason("image/webp", 4, 0), "");
    assert.match(getInlineImageSkipReason("image/svg+xml", 4, 0), /Unsupported MIME type/);
    assert.match(getInlineImageSkipReason("image/bmp", 4, 0), /Unsupported MIME type/);
  });

  it("keeps one metadata entry and one block candidate for a named attachment with Content-ID", () => {
    const namedAttachment = {
      info: { name: "signature-logo.png" },
      contentId: " <SIGNATURE.Logo@Example.Test> ",
      partName: "1.3",
    };
    const inlineImage = {
      contentId: "signature.logo@example.test",
      partName: "1.4",
    };

    const result = correlateInlineImageRecords([namedAttachment], [inlineImage]);

    assert.equal(result.metadataEntries.length, 1);
    assert.equal(result.inlineImageEntries.length, 1);
    assert.strictEqual(result.inlineImageEntries[0].metadataRecord, namedAttachment);
    assert.strictEqual(result.inlineImageEntries[0].inlineImage, inlineImage);
    assert.equal(result.inlineImageEntries[0].matchedKnownAttachment, true);
  });

  it("keeps one metadata entry and one block candidate when Gloda stripped inline headers", () => {
    const namedAttachment = {
      info: { name: "signature-logo.png" },
      contentId: "",
      partName: "1.3",
    };
    const strippedInlineImage = {
      contentId: "",
      partName: "1.3",
    };

    const result = correlateInlineImageRecords(
      [namedAttachment],
      [strippedInlineImage]
    );

    assert.equal(result.metadataEntries.length, 1);
    assert.equal(result.inlineImageEntries.length, 1);
    assert.strictEqual(result.inlineImageEntries[0].metadataRecord, namedAttachment);
    assert.strictEqual(result.inlineImageEntries[0].inlineImage, strippedInlineImage);
    assert.equal(result.inlineImageEntries[0].matchedKnownAttachment, true);
  });

  it("orders rendered CID references before unreferenced inline images", () => {
    const sources = [
      { id: "unreferenced-1", info: { contentId: "unused-1@example.test" } },
      { id: "second", info: { contentId: "second@example.test" } },
      { id: "first", info: { contentId: "FIRST@EXAMPLE.TEST" } },
      { id: "unreferenced-2", info: { contentId: "unused-2@example.test" } },
    ];
    const body = [
      "![first](cid:%66irst%40example.test)",
      '<img src="CID:<SECOND@EXAMPLE.TEST>">',
      "![first again](cid:first@example.test)",
    ].join("\n");

    assert.deepEqual(
      [...getInlineImageContentIdReferences(body)],
      ["first@example.test", "second@example.test"]
    );
    assert.deepEqual(
      [...orderInlineImageRecordsForBody(sources, body)].map(source => source.id),
      ["first", "second", "unreferenced-1", "unreferenced-2"]
    );
    assert.equal(normalizeInlineImageContentId("CID:<FIRST@EXAMPLE.TEST>"), "first@example.test");
  });

  it("allows the exact per-image budget and skips one byte over it", () => {
    assert.equal(
      getInlineImageSkipReason("image/png", MAX_INLINE_IMAGE_BASE64_BYTES, 0),
      ""
    );
    assert.match(
      getInlineImageSkipReason("image/png", MAX_INLINE_IMAGE_BASE64_BYTES + 1, 0),
      /per-image base64 limit/
    );
  });

  it("allows the exact total budget and skips content that would exceed it", () => {
    const finalImageSize = MAX_INLINE_IMAGE_BASE64_BYTES;
    const exactTotalBefore = MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES - finalImageSize;
    assert.equal(
      getInlineImageSkipReason("image/webp", finalImageSize, exactTotalBefore),
      ""
    );
    assert.match(
      getInlineImageSkipReason("image/webp", finalImageSize, exactTotalBefore + 1),
      /total base64 limit/
    );
  });

  it("encodes arbitrary MIME bytes without UTF-8 conversion", () => {
    const byteString = String.fromCharCode(0, 1, 2, 0x7f, 0x80, 0xff);
    const encoded = encodeByteStringToBase64(byteString);
    assert.deepEqual([...Buffer.from(encoded, "base64")], [0, 1, 2, 0x7f, 0x80, 0xff]);
  });
});
