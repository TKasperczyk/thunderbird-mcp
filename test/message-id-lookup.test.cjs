"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const SAMPLE = "8a0e908686e3471f962ec8fa93849104@contentai.ru";
const REAL_ID = "mebrt6w2q7f6_stmc_t0tpupj1dr_pi8b_fbu~n+brg8-u0jmx2j4gqf7lb26@alfastrahmail.ru";

function extractFunction(source, name) {
  const marker = `function ${name}(`;
  const start = source.indexOf(marker);
  assert.ok(start >= 0, `${name} missing`);
  const brace = source.indexOf("{", start);
  let depth = 0;
  for (let i = brace; i < source.length; i++) {
    if (source[i] === "{") depth++;
    else if (source[i] === "}") {
      depth--;
      if (depth === 0) return source.slice(start, i + 1);
    }
  }
  assert.fail(`unterminated ${name}`);
}

function loadHelpers() {
  const source = fs.readFileSync(
    path.join(__dirname, "..", "extension", "mcp_server", "api.js"),
    "utf8"
  );
  const flagStart = source.indexOf("const FOLDER_FLAG_VIRTUAL");
  const flagEnd = source.indexOf("function normalizeRfcMessageId");
  assert.ok(flagStart >= 0 && flagEnd > flagStart, "folder flag constants missing");
  const names = [
    "normalizeRfcMessageId",
    "messageIdForHeaderLookup",
    "folderLookupPriorityFromFlags",
    "lookupHeaderInDatabase",
  ];
  const script = [
    source.slice(flagStart, flagEnd),
    ...names.map(name => extractFunction(source, name)),
    "this.helpers = { normalizeRfcMessageId, messageIdForHeaderLookup, folderLookupPriorityFromFlags, lookupHeaderInDatabase };",
  ].join("\n");
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(script, sandbox);
  return sandbox.helpers;
}

const {
  normalizeRfcMessageId,
  messageIdForHeaderLookup,
  folderLookupPriorityFromFlags,
  lookupHeaderInDatabase,
} = loadHelpers();

const INBOX = 0x00001000;
const SENT = 0x00000200;
const DRAFTS = 0x00000400;
const ARCHIVE = 0x00004000;
const ALLMAIL = 0x00008000;
const JUNK = 0x40000000;
const TRASH = 0x00000100;

describe("normalizeRfcMessageId", () => {
  it("strips mid:, brackets, and one level of percent-encoding", () => {
    assert.equal(normalizeRfcMessageId(`mid:${SAMPLE}`).messageId, SAMPLE);
    assert.equal(normalizeRfcMessageId(`<${SAMPLE}>`).messageId, SAMPLE);
    assert.equal(normalizeRfcMessageId(`mid:<${SAMPLE}>`).messageId, SAMPLE);
    assert.equal(
      normalizeRfcMessageId("8a0e908686e3471f962ec8fa93849104%40contentai.ru").messageId,
      SAMPLE
    );
    assert.equal(normalizeRfcMessageId(`%3Cmid:${SAMPLE.replace("@", "%40")}%3E`).messageId, SAMPLE);
    assert.equal(normalizeRfcMessageId(`  MID:<${SAMPLE}>  `).messageId, SAMPLE);
  });

  it("preserves local-part case and leaves + and ~ untouched", () => {
    assert.equal(normalizeRfcMessageId(`mid:${REAL_ID}`).messageId, REAL_ID);
    assert.equal(normalizeRfcMessageId("mid:Case.ID@Example.COM").messageId, "Case.ID@Example.COM");
  });

  it("rejects values that are not an RFC Message-ID", () => {
    assert.match(normalizeRfcMessageId("mid:not-an-id").error, /Not an RFC Message-ID/);
    assert.match(normalizeRfcMessageId("").error, /non-empty/);
    assert.match(normalizeRfcMessageId("   ").error, /non-empty/);
    assert.match(normalizeRfcMessageId(null).error, /non-empty/);
  });

  it("keeps a malformed percent-encoded id that already contains @", () => {
    assert.equal(normalizeRfcMessageId("foo%zz@bar.com").messageId, "foo%zz@bar.com");
  });

  it("does not rewrite a stored id that already contains a literal %HH sequence", () => {
    assert.equal(normalizeRfcMessageId("foo%2Bbar@example.com").messageId, "foo%2Bbar@example.com");
    assert.equal(normalizeRfcMessageId("foo%40bar@example.com").messageId, "foo%40bar@example.com");
    assert.equal(normalizeRfcMessageId("<foo%2Bbar@example.com>").messageId, "foo%2Bbar@example.com");
  });

  it("decodes a mid: link once after stripping the prefix", () => {
    assert.equal(normalizeRfcMessageId("mid:foo%2Bbar@baz.com").messageId, "foo+bar@baz.com");
    assert.equal(normalizeRfcMessageId("MID:foo%2Bbar@baz.com").messageId, "foo+bar@baz.com");
  });

  it("rejects ids that must not start a folder scan", () => {
    for (const value of [
      "foo bar@example.com",
      "imap-message://user@host/INBOX#1",
      "mailto:user@example.com",
      "MAILTO:user@example.com",
      "@",
      "@domain",
      "local@",
    ]) {
      assert.match(normalizeRfcMessageId(value).error, /Not an RFC Message-ID/, value);
    }
  });
});

describe("messageIdForHeaderLookup", () => {
  it("uses the normalized id when it is an RFC Message-ID", () => {
    assert.equal(messageIdForHeaderLookup(` mid:<${SAMPLE}> `), SAMPLE);
  });

  it("keeps an unusual id that has no @", () => {
    assert.equal(messageIdForHeaderLookup("  numeric-key  "), "numeric-key");
  });

  it("leaves a stored id containing %HH unchanged", () => {
    assert.equal(messageIdForHeaderLookup("foo%2Bbar@example.com"), "foo%2Bbar@example.com");
    assert.equal(messageIdForHeaderLookup("foo%40bar@example.com"), "foo%40bar@example.com");
    assert.equal(messageIdForHeaderLookup("mid:foo%2Bbar@baz.com"), "foo+bar@baz.com");
  });

  it("returns null for blank or non-string values", () => {
    assert.equal(messageIdForHeaderLookup(""), null);
    assert.equal(messageIdForHeaderLookup("  "), null);
    assert.equal(messageIdForHeaderLookup(12), null);
  });
});

describe("folderLookupPriorityFromFlags", () => {
  it("orders inbox, sent, drafts, and archive ahead of other folders", () => {
    assert.equal(folderLookupPriorityFromFlags([INBOX]), 0);
    assert.equal(folderLookupPriorityFromFlags([SENT]), 1);
    assert.equal(folderLookupPriorityFromFlags([DRAFTS]), 2);
    assert.equal(folderLookupPriorityFromFlags([ARCHIVE]), 3);
    assert.equal(folderLookupPriorityFromFlags([0]), 4);
  });

  it("ranks All Mail after ordinary folders even when Archive is also set", () => {
    assert.equal(folderLookupPriorityFromFlags([ALLMAIL]), 5);
    assert.equal(folderLookupPriorityFromFlags([ALLMAIL | ARCHIVE]), 5);
    assert.equal(folderLookupPriorityFromFlags([0, ALLMAIL]), 5);
    assert.equal(folderLookupPriorityFromFlags([INBOX, ALLMAIL]), 0);
    assert.equal(folderLookupPriorityFromFlags([INBOX | ALLMAIL]), 0);
    assert.equal(folderLookupPriorityFromFlags([SENT, ALLMAIL]), 1);
    assert.equal(folderLookupPriorityFromFlags([DRAFTS, ALLMAIL]), 2);
  });

  it("treats a folder inside junk or trash as that folder", () => {
    assert.equal(folderLookupPriorityFromFlags([0, JUNK]), 6);
    assert.equal(folderLookupPriorityFromFlags([0, 0, TRASH]), 7);
  });

  it("lets an inbox flag outrank a trash ancestor", () => {
    assert.equal(folderLookupPriorityFromFlags([INBOX, TRASH]), 0);
  });

  it("sorts a stable folder list by that priority", () => {
    const folders = [
      { index: 0, priority: folderLookupPriorityFromFlags([TRASH]) },
      { index: 1, priority: folderLookupPriorityFromFlags([0]) },
      { index: 2, priority: folderLookupPriorityFromFlags([INBOX]) },
      { index: 3, priority: folderLookupPriorityFromFlags([0, JUNK]) },
      { index: 4, priority: folderLookupPriorityFromFlags([SENT]) },
      { index: 5, priority: folderLookupPriorityFromFlags([0]) },
      { index: 6, priority: folderLookupPriorityFromFlags([ALLMAIL | ARCHIVE]) },
      { index: 7, priority: folderLookupPriorityFromFlags([ARCHIVE]) },
      { index: 8, priority: folderLookupPriorityFromFlags([0, ALLMAIL]) },
    ];
    folders.sort((a, b) => a.priority - b.priority || a.index - b.index);
    assert.deepEqual(folders.map(folder => folder.index), [2, 4, 7, 1, 5, 6, 8, 3, 0]);
  });
});

describe("lookupHeaderInDatabase", () => {
  it("returns a direct hit without enumerating", () => {
    let enumerated = false;
    const hdr = { messageId: SAMPLE };
    const found = lookupHeaderInDatabase({
      getMsgHdrForMessageID(id) {
        return id === SAMPLE ? hdr : null;
      },
      enumerateMessages() {
        enumerated = true;
        return [];
      },
    }, SAMPLE);
    assert.equal(found, hdr);
    assert.equal(enumerated, false);
  });

  it("does not enumerate after a direct miss", () => {
    let enumerated = false;
    const found = lookupHeaderInDatabase({
      getMsgHdrForMessageID() {
        return null;
      },
      enumerateMessages() {
        enumerated = true;
        return [{ messageId: SAMPLE }];
      },
    }, SAMPLE);
    assert.equal(found, null);
    assert.equal(enumerated, false);
  });

  it("does not enumerate when the database has no direct lookup", () => {
    let enumerated = false;
    const found = lookupHeaderInDatabase({
      enumerateMessages() {
        enumerated = true;
        return [{ messageId: SAMPLE }];
      },
    }, SAMPLE);
    assert.equal(found, null);
    assert.equal(enumerated, false);
  });

  it("drops a header whose id does not match and does not scan", () => {
    let enumerated = false;
    const found = lookupHeaderInDatabase({
      getMsgHdrForMessageID() {
        return { messageId: "different@example.com" };
      },
      enumerateMessages() {
        enumerated = true;
        return [{ messageId: SAMPLE }];
      },
    }, SAMPLE);
    assert.equal(found, null);
    assert.equal(enumerated, false);
  });

  it("returns null when the direct lookup throws", () => {
    const found = lookupHeaderInDatabase({
      getMsgHdrForMessageID() {
        throw new Error("db closed");
      },
    }, SAMPLE);
    assert.equal(found, null);
  });
});
