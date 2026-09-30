"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

// The MIME assembly itself runs inside Thunderbird (MessageSend/MimeMessage).
// What can be checked here: the pure planning/body-rewriting helpers, and that
// inlineImages is declared and forwarded for every tool that supports it.
const API_PATH = path.join(__dirname, "..", "extension", "mcp_server", "api.js");
const source = fs.readFileSync(API_PATH, "utf8");

function limitsSnippet() {
  const start = source.indexOf("// BEGIN OUTBOUND ATTACHMENT LIMITS");
  const end = source.indexOf("// END OUTBOUND ATTACHMENT LIMITS");
  assert.ok(start >= 0 && end > start, "outbound attachment limit markers missing");
  return source.slice(start, end);
}

function loadHelpers() {
  const startMarker = "// BEGIN OUTBOUND INLINE IMAGE HELPERS";
  const endMarker = "// END OUTBOUND INLINE IMAGE HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0 && end > start, "outbound inline image helper markers missing");
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(
    `${limitsSnippet()}
${source.slice(start, end)}
this.helpers = {
  normalizeOutboundContentId,
  guessOutboundImageContentType,
  planOutboundInlineImages,
  findUnreferencedContentIds,
  replaceContentIdReferences,
  OUTBOUND_INLINE_IMAGES_SCHEMA,
};`,
    sandbox
  );
  return sandbox.helpers;
}

const {
  normalizeOutboundContentId,
  guessOutboundImageContentType,
  planOutboundInlineImages,
  findUnreferencedContentIds,
  replaceContentIdReferences,
  OUTBOUND_INLINE_IMAGES_SCHEMA,
} = loadHelpers();

describe("normalizeOutboundContentId", () => {
  it("accepts plain ids and strips cid: and angle brackets", () => {
    assert.equal(normalizeOutboundContentId("screen1.png"), "screen1.png");
    assert.equal(normalizeOutboundContentId("cid:screen1.png"), "screen1.png");
    assert.equal(normalizeOutboundContentId("<part1.abc@example.com>"), "part1.abc@example.com");
  });

  it("rejects ids that could break out of an attribute or header", () => {
    for (const bad of ["", "a b.png", 'x".png', "x>.png", "<>", "a\r\nX-Evil: 1"]) {
      assert.equal(normalizeOutboundContentId(bad), null, JSON.stringify(bad));
    }
  });
});

describe("guessOutboundImageContentType", () => {
  it("maps common image extensions case-insensitively", () => {
    assert.equal(guessOutboundImageContentType("a.PNG"), "image/png");
    assert.equal(guessOutboundImageContentType("C:\\x\\shot.jpeg"), "image/jpeg");
    assert.equal(guessOutboundImageContentType("noext"), "");
    assert.equal(guessOutboundImageContentType("doc.pdf"), "");
  });
});

describe("planOutboundInlineImages", () => {
  it("returns an empty plan for a missing argument", () => {
    const { plan, errors } = planOutboundInlineImages(undefined);
    assert.equal(plan.length, 0);
    assert.equal(errors.length, 0);
  });

  it("defaults cid and name from the file name of path", () => {
    const { plan, errors } = planOutboundInlineImages([{ path: "C:\\shots\\ekran 1.png" }, { path: "/tmp/b.jpg", cid: "b" }]);
    assert.deepEqual([...errors], ["inlineImages[0]: invalid cid \"ekran 1.png\" (letters, digits and . @ - _ + = etc.; no spaces, quotes or angle brackets)"]);
    assert.equal(plan.length, 1);
    assert.equal(plan[0].cid, "b");
    assert.equal(plan[0].name, "b.jpg");
    assert.equal(plan[0].contentType, "image/jpeg");
    assert.equal(plan[0].entry, "/tmp/b.jpg");
  });

  it("builds a base64 attachment entry with a content type", () => {
    const { plan, errors } = planOutboundInlineImages([{ cid: "logo", name: "logo.gif", base64: "R0lGOA==" }]);
    assert.equal(errors.length, 0);
    assert.deepEqual({ ...plan[0].entry }, { name: "logo.gif", contentType: "image/gif", base64: "R0lGOA==" });
  });

  it("requires contentType when nothing can be guessed, and only image/*", () => {
    assert.match(planOutboundInlineImages([{ cid: "x", base64: "QQ==" }]).errors[0], /contentType is required/);
    assert.match(planOutboundInlineImages([{ cid: "x", contentType: "text/html", base64: "QQ==" }]).errors[0], /must be image/);
  });

  it("rejects duplicate cids, missing data and both sources at once", () => {
    const { errors } = planOutboundInlineImages([
      { cid: "a.png", base64: "QQ==" },
      { cid: "A.PNG", base64: "QQ==" },
      { cid: "c.png" },
      { cid: "d.png", base64: "QQ==", path: "/tmp/d.png" },
    ]);
    assert.equal(errors.length, 3);
    assert.match(errors[0], /duplicate cid/);
    assert.match(errors[1], /one of base64 or path/);
    assert.match(errors[2], /either base64 or path/);
  });
});

describe("cid references in the body", () => {
  const body = '<p>Hi</p><img src="cid:a.png"><img src=\'cid:b\' alt=x><img src="cid:a.png2">';

  it("finds unreferenced cids by whole URL, not prefix", () => {
    assert.deepEqual([...findUnreferencedContentIds(body, ["a.png", "b", "c"])], ["c"]);
    assert.deepEqual([...findUnreferencedContentIds('<img src="cid:a.png2">', ["a.png"])], ["a.png"]);
  });

  it("replaces every whole reference and leaves longer ids alone", () => {
    const out = replaceContentIdReferences(body + '<img src="cid:a.png">', { "a.png": "data:image/png;base64,QQ==" });
    assert.equal(out.split("data:image/png;base64,QQ==").length - 1, 2);
    assert.ok(out.includes("cid:a.png2"));
    assert.ok(out.includes("cid:b"));
  });

  it("treats regex metacharacters in a cid literally", () => {
    assert.deepEqual([...findUnreferencedContentIds('<img src="cid:aXpng">', ["a.png"])], ["a.png"]);
    const out = replaceContentIdReferences('<img src="cid:a+b$1.png">', { "a+b$1.png": "U" });
    assert.equal(out, '<img src="U">');
  });
});

describe("inlineImages wiring", () => {
  const tools = {
    sendMail: "composeMail",
    saveDraft: "saveDraft",
    replyToMessage: "replyToMessage",
  };

  function toolDefinition(name) {
    const start = source.indexOf(`name: "${name}",`);
    assert.notEqual(start, -1, `${name} definition not found`);
    const next = source.indexOf("        name: \"", start + 20);
    return source.slice(start, next === -1 ? source.length : next);
  }

  for (const [tool, impl] of Object.entries(tools)) {
    it(`${tool} declares inlineImages and forwards it to ${impl}`, () => {
      assert.ok(toolDefinition(tool).includes("inlineImages: OUTBOUND_INLINE_IMAGES_SCHEMA"), `${tool} schema is missing inlineImages`);
      const marker = `case "${tool}":`;
      const at = source.indexOf(marker);
      const line = source.slice(at, source.indexOf(";", at + marker.length));
      assert.ok(line.includes("args.inlineImages"), `${tool} dispatch drops inlineImages`);
      const sig = source.match(new RegExp(`function ${impl}\\(([^)]*)\\)`));
      assert.ok(sig && sig[1].trim().endsWith("inlineImages"), `${impl} does not take inlineImages last`);
    });
  }

  it("does not claim support for forwardMessage", () => {
    assert.ok(!toolDefinition("forwardMessage").includes("inlineImages"));
  });

  it("schema items require a source and forbid unknown keys", () => {
    assert.equal(OUTBOUND_INLINE_IMAGES_SCHEMA.items.additionalProperties, false);
    assert.equal(OUTBOUND_INLINE_IMAGES_SCHEMA.items.anyOf.length, 3);
  });
});
