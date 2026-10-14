"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");

function snippet(startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0 && end > start, `Production markers missing: ${startMarker}`);
  return source.slice(start, end);
}

function constant(name) {
  const match = source.match(new RegExp(`^const ${name} = [^\\n]+;`, "m"));
  assert.ok(match, `Production constant missing: ${name}`);
  return match[0];
}

// The real production schemas of the two header tools, evaluated in isolation.
function productionHeaderTools() {
  const start = source.indexOf('      {\n        name: "getMessageHeaders"');
  const startCrLf = source.indexOf('      {\r\n        name: "getMessageHeaders"');
  const from = start >= 0 ? start : startCrLf;
  const end = source.indexOf('name: "sendMail"', from);
  const to = source.lastIndexOf("      {", end);
  assert.ok(from >= 0 && to > from, "production header tool schemas missing");
  const sandbox = {};
  vm.runInNewContext(`this.tools = [${source.slice(from, to)}];`, sandbox);
  return sandbox.tools;
}

const ENCRYPTED_RAW = 'From: sender@example.test\r\nSubject: =?utf-8?q?Encrypted_Message?=\r\nContent-Type: multipart/encrypted; protocol="application/pgp-encrypted"; boundary="b"\r\n\r\n--b--\r\n';
const DECRYPTION_OKAY = 0x00020000;

function makeHarness({ rows, direct = true, allowEncrypted = false, msPerYield = 0 } = {}) {
  const stats = { yields: 0, listAllKeys: 0, directLookups: 0, headerReads: 0 };
  let generation = 0;
  let now = 0;
  let accessible = true;
  let onYield = null;
  let databaseHook = null;
  const headers = new Map();
  const byMessageId = new Map();

  const folder = {
    URI: "imap://account1/INBOX",
    prettyName: "INBOX",
    hasSubFolders: false,
    server: null,
    updateFolder() {},
    getMsgInputStream(msgHdr) {
      let offset = 0;
      const raw = msgHdr.raw;
      return {
        available() { return Math.min(raw.length - offset, 1000); },
        read(count) { const chunk = raw.slice(offset, offset + count); offset += chunk.length; return chunk; },
        close() {},
      };
    },
    get msgDatabase() {
      if (databaseHook) databaseHook();
      const openedGeneration = generation;
      const assertLive = () => assert.equal(openedGeneration, generation, "Database retained across a yield");
      const db = {
        listAllKeys() { assertLive(); stats.listAllKeys++; return Array.from(headers.keys()); },
        enumerateMessages() { assert.fail("A header lookup must not open a native message enumerator"); },
        containsKey(key) { assertLive(); return headers.has(key); },
        getMsgHdrForKey(key) { assertLive(); stats.headerReads++; return headers.get(key); },
      };
      if (direct) {
        db.getMsgHdrForMessageID = id => { stats.directLookups++; return byMessageId.get(id) || null; };
      }
      return db;
    },
  };

  rows.forEach((row, index) => {
    const properties = { keywords: "custom junk $label1", ...row.properties };
    const header = {
      messageKey: index + 1,
      messageId: `m-${index}@example.com`,
      threadId: 100 + index,
      date: (index + 1) * 1000000,
      subject: "Raw subject",
      author: "Raw author",
      recipients: "Raw recipient",
      mime2DecodedSubject: "Decoded subject",
      mime2DecodedAuthor: "Alice Smith",
      mime2DecodedRecipients: "Bob Jones",
      ccList: "plain@example.com",
      flags: 0,
      isRead: true,
      isFlagged: false,
      messageSize: 321,
      folder,
      getStringProperty: name => properties[name] || "",
      getUint32Property: name => (name === "enigmail" ? (row.enigmail || 0) : 0),
      ...row.hdr,
    };
    for (const field of row.throwOn || []) {
      Object.defineProperty(header, field, { get() { throw new Error(`unreadable ${field}`); } });
    }
    headers.set(header.messageKey, header);
    // Rows flagged indirectOnly miss the direct lookup but are still in the DB.
    if (!row.indirectOnly) byMessageId.set(header.messageId, header);
  });

  const account = { key: "account1", incomingServer: { type: "imap", rootFolder: folder } };
  folder.server = account.incomingServer;

  class FakeDate extends Date {
    static now() { return now; }
  }
  const runtime = vm.createContext({
    Date: FakeDate,
    Services: {
      prefs: { getBoolPref: (name, fallback) => (name.endsWith(".allowEncryptedMessages") ? allowEncrypted : fallback) },
      tm: {
        dispatchToMainThread(resolve) {
          setImmediate(() => {
            stats.yields++;
            now += msPerYield;
            generation++;
            if (onYield) onYield(stats.yields);
            resolve();
          });
        },
      },
    },
    Ci: { nsMsgMessageFlags: { Expunged: 0x8 }, nsITimer: { TYPE_ONE_SHOT: 0 } },
    NetUtil: { readInputStreamToString: (stream, count) => stream.read(count) },
    MailServices: {
      mimeConverter: {
        decodeMimeHeader: value => value.replace(/=\?utf-8\?q\?([^?]*)\?=/gi, (_, text) => text.replace(/_/g, " ")),
      },
      folderLookup: { getFolderForURL: uri => (uri === folder.URI ? folder : null) },
      accounts: { accounts: [account], findAccountForServer: () => (accessible ? account : null) },
    },
    isAccountAllowed: () => true,
    getAllowedAccountIds: () => [],
    console,
  });
  const keywordStart = source.indexOf("const INTERNAL_KEYWORDS = new Set([");
  const keywordEnd = source.indexOf("]);", keywordStart);
  const tools = productionHeaderTools();
  runtime.buildTools = () => tools;
  vm.runInContext([
    ...["SEARCH_YIELD_EVERY", "SEARCH_TIME_BUDGET_MS", "MAX_HEADER_BATCH_IDS", "PREF_ALLOW_ENCRYPTED_MESSAGES"].map(constant),
    snippet("// BEGIN PRIVACY PREFERENCE HELPERS", "// END PRIVACY PREFERENCE HELPERS"),
    snippet("// BEGIN RAW MIME PARSING HELPERS", "// END RAW MIME PARSING HELPERS"),
    snippet("// BEGIN PROTECTED SUBJECT HELPERS", "// END PROTECTED SUBJECT HELPERS"),
    source.slice(keywordStart, keywordEnd + 3),
    snippet("// BEGIN SEARCH RESULT HELPERS", "// END SEARCH RESULT HELPERS"),
    snippet("function isFolderAccessible(", "function toColumnarTable("),
    snippet("function getUserTags(", "// BEGIN OUTBOUND ATTACHMENT CONVERSION"),
    snippet("function openFolder(", "function findTrashFolder("),
    snippet("// BEGIN MESSAGE HEADER READ TOOLS", "// END MESSAGE HEADER READ TOOLS"),
    snippet("// BEGIN INLINE ATTACHMENT BASE64 HELPERS", "// END INLINE ATTACHMENT BASE64 HELPERS"),
    snippet("// BEGIN TOOL SCHEMA VALIDATOR", "// END TOOL SCHEMA VALIDATOR"),
    snippet("// BEGIN MCP TEXT SANITIZATION", "// END MCP TEXT SANITIZATION"),
    snippet("// BEGIN TOOL CALL DISPATCH", "// END TOOL CALL DISPATCH"),
  ].join("\n"), runtime);

  return {
    runtime,
    folder,
    stats,
    headers,
    advance(ms) { now += ms; },
    revokeAccess() { accessible = false; },
    setOnYield(callback) { onYield = callback; },
    setDatabaseHook(callback) { databaseHook = callback; },
    async call(name, args) {
      const result = await runtime.callTool(name, args);
      return JSON.parse(JSON.stringify(result));
    },
  };
}

const rowsOf = count => Array.from({ length: count }, () => ({}));
const ids = (count, start = 0) => Array.from({ length: count }, (_, index) => `m-${start + index}@example.com`);

describe("header object shape", () => {
  it("uses the field names and types of the other read tools", async () => {
    const h = makeHarness({ rows: rowsOf(2) });
    const result = await h.call("getMessageHeaders", { messageId: "m-1@example.com", folderPath: h.folder.URI });
    assert.deepEqual(Object.keys(result).sort(), [
      "author", "ccList", "date", "flagged", "id", "inReplyTo", "read", "recipients",
      "references", "size", "subject", "tags", "threadId",
    ]);
    assert.equal(result.id, "m-1@example.com");
    assert.equal(result.threadId, 101);
    assert.equal(typeof result.threadId, "number");
    assert.equal(result.read, true);
    assert.equal(result.flagged, false);
    assert.equal(result.subject, "Decoded subject");
    assert.equal(result.author, "Alice Smith");
    assert.equal(result.recipients, "Bob Jones");
    assert.deepEqual(result.tags, ["custom", "$label1"]);
    assert.equal(result.size, 321);
    assert.equal(result.date, new Date(2000000 / 1000).toISOString());
  });

  it("RFC 2047-decodes ccList itself", async () => {
    const h = makeHarness({
      rows: [{ hdr: { ccList: "=?utf-8?q?Jose_Garcia?= <jose@example.com>, plain@example.com" } }, { hdr: { ccList: "" } }],
    });
    const [first, second] = [
      await h.call("getMessageHeaders", { messageId: "m-0@example.com", folderPath: h.folder.URI }),
      await h.call("getMessageHeaders", { messageId: "m-1@example.com", folderPath: h.folder.URI }),
    ];
    assert.equal(first.ccList, "Jose Garcia <jose@example.com>, plain@example.com");
    assert.equal(second.ccList, "");
  });

  it("falls back to the raw ccList when decoding throws", async () => {
    const h = makeHarness({ rows: [{ hdr: { ccList: "a@example.com" } }] });
    h.runtime.MailServices.mimeConverter.decodeMimeHeader = () => { throw new Error("bad charset"); };
    const result = await h.call("getMessageHeaders", { messageId: "m-0@example.com", folderPath: h.folder.URI });
    assert.equal(result.ccList, "a@example.com");
  });

  it("derives inReplyTo from the references property, not a separate property", async () => {
    const h = makeHarness({
      rows: [
        // References present: the last entry is the direct parent.
        { properties: { references: "<root@example.com> <mid@example.com> <parent@example.com>" } },
        // Thunderbird stores a lone In-Reply-To in references.
        { properties: { references: "<only@example.com>" } },
        // A stale in-reply-to property must not be consulted.
        { properties: { "in-reply-to": "<ignored@example.com>" } },
        // Bare, unbracketed ids and odd spacing.
        { properties: { references: "  a@example.com\n  b@example.com " } },
      ],
    });
    const read = async index => h.call("getMessageHeaders", { messageId: `m-${index}@example.com`, folderPath: h.folder.URI });
    const [a, b, c, d] = [await read(0), await read(1), await read(2), await read(3)];
    assert.equal(a.inReplyTo, "parent@example.com");
    assert.equal(a.references, "<root@example.com> <mid@example.com> <parent@example.com>");
    assert.equal(b.inReplyTo, "only@example.com");
    assert.equal(c.inReplyTo, "");
    assert.equal(c.references, "");
    assert.equal(d.inReplyTo, "b@example.com");
  });

  it("withholds subjects of recorded-encrypted messages unless allowed", async () => {
    const rows = [{ enigmail: DECRYPTION_OKAY, hdr: { raw: ENCRYPTED_RAW, subject: "Cached", mime2DecodedSubject: "Cached decrypted subject" } }];
    const blocked = makeHarness({ rows });
    const single = await blocked.call("getMessageHeaders", { messageId: "m-0@example.com", folderPath: blocked.folder.URI });
    assert.equal(single.subject, "Encrypted Message");
    assert.equal(single.encryptedContentWithheld, true);
    const batch = await blocked.call("batchGetMessageHeaders", { messageIds: ["m-0@example.com"], folderPath: blocked.folder.URI });
    assert.equal(batch.headers["m-0@example.com"].subject, "Encrypted Message");
    assert.doesNotMatch(JSON.stringify(batch), /Cached decrypted/);

    const allowed = makeHarness({ rows, allowEncrypted: true });
    const open = await allowed.call("getMessageHeaders", { messageId: "m-0@example.com", folderPath: allowed.folder.URI });
    assert.equal(open.subject, "Cached decrypted subject");
    assert.equal(open.encryptedContentWithheld, undefined);
  });
});

describe("header lookup and cooperative fallback scan", () => {
  it("answers direct hits without scanning or yielding", async () => {
    const h = makeHarness({ rows: rowsOf(1000) });
    const result = await h.call("batchGetMessageHeaders", { messageIds: ids(10), folderPath: h.folder.URI });
    assert.equal(result.failed, 0);
    assert.equal(result.total, 10);
    assert.equal(Object.keys(result.headers).length, 10);
    assert.equal(h.stats.listAllKeys, 0);
    assert.equal(h.stats.yields, 0);
    assert.equal(result.truncated, undefined);
    assert.equal(result.notScanned, undefined);
  });

  it("finds ids that miss the direct lookup with one yielding scan, past the old 50,000 cap", async () => {
    const rows = rowsOf(60000);
    rows[59990].indirectOnly = true;
    rows[100].indirectOnly = true;
    const h = makeHarness({ rows });
    const result = await h.call("batchGetMessageHeaders", {
      messageIds: ["m-59990@example.com", "m-100@example.com", "m-5@example.com"],
      folderPath: h.folder.URI,
    });
    assert.equal(result.failed, 0);
    assert.equal(result.headers["m-59990@example.com"].id, "m-59990@example.com");
    assert.equal(result.headers["m-100@example.com"].threadId, 200);
    assert.equal(h.stats.listAllKeys, 1, "a single shared scan serves all missed ids");
    assert.ok(h.stats.yields >= Math.floor(59990 / 250) - 1, `expected a yield every 250 rows, saw ${h.stats.yields}`);
  });

  it("stops scanning once every missed id is found", async () => {
    const rows = rowsOf(5000);
    rows[10].indirectOnly = true;
    const h = makeHarness({ rows });
    const result = await h.call("batchGetMessageHeaders", { messageIds: ["m-10@example.com"], folderPath: h.folder.URI });
    assert.equal(result.failed, 0);
    assert.equal(h.stats.yields, 0, "found inside the first chunk");
    assert.ok(h.stats.headerReads <= 250);
  });

  it("reports definitively absent ids as not found after a complete scan", async () => {
    const h = makeHarness({ rows: rowsOf(600) });
    const result = await h.call("batchGetMessageHeaders", {
      messageIds: ["m-1@example.com", "absent@example.com"],
      folderPath: h.folder.URI,
    });
    assert.equal(result.failed, 1);
    assert.deepEqual(result.headers["absent@example.com"], { error: "not found in this folder" });
    assert.equal(result.truncated, undefined);
    assert.equal(result.notScanned, undefined);
    assert.ok(h.stats.yields >= 1, "the scan yielded");
  });

  it("distinguishes not scanned from not found when the time budget expires", async () => {
    const rows = rowsOf(5000);
    rows[4000].indirectOnly = true;
    const h = makeHarness({ rows, msPerYield: 6000 });
    const result = await h.call("batchGetMessageHeaders", {
      messageIds: ["m-1@example.com", "m-4000@example.com", "absent@example.com"],
      folderPath: h.folder.URI,
    });
    assert.equal(result.headers["m-1@example.com"].id, "m-1@example.com");
    for (const id of ["m-4000@example.com", "absent@example.com"]) {
      assert.equal(result.headers[id].notScanned, true);
      assert.match(result.headers[id].error, /not scanned/);
    }
    assert.deepEqual(result.notScanned.sort(), ["absent@example.com", "m-4000@example.com"]);
    assert.equal(result.truncated, true);
    assert.match(result.message, /20-second time budget/);
    assert.equal(result.failed, 2);
    // Bounded by the budget: 20s at 6s per yield is four chunks, not the folder.
    assert.ok(h.stats.headerReads < 5000);
  });

  it("reports an unscanned single message with notScanned instead of plain not found", async () => {
    const rows = rowsOf(2000);
    const h = makeHarness({ rows, msPerYield: 25000 });
    const result = await h.call("getMessageHeaders", { messageId: "absent@example.com", folderPath: h.folder.URI });
    assert.equal(result.notScanned, true);
    assert.match(result.error, /Message not found: absent@example.com/);
    assert.match(result.error, /time budget/);

    const complete = makeHarness({ rows: rowsOf(10) });
    const absent = await complete.call("getMessageHeaders", { messageId: "absent@example.com", folderPath: complete.folder.URI });
    assert.deepEqual(absent, { error: "Message not found: absent@example.com" });
  });

  it("treats unreadable rows as not scanned rather than not found", async () => {
    const rows = rowsOf(20);
    rows[3].throwOn = ["flags"];
    const h = makeHarness({ rows });
    const result = await h.call("batchGetMessageHeaders", { messageIds: ["absent@example.com"], folderPath: h.folder.URI });
    assert.deepEqual(result.notScanned, ["absent@example.com"]);
    assert.equal(result.truncated, true);
    assert.doesNotMatch(result.message, /time budget/);
  });

  it("falls back to the scan when direct lookup is unavailable", async () => {
    const h = makeHarness({ rows: rowsOf(30), direct: false });
    const result = await h.call("batchGetMessageHeaders", { messageIds: ["m-29@example.com"], folderPath: h.folder.URI });
    assert.equal(result.failed, 0);
    assert.equal(result.headers["m-29@example.com"].id, "m-29@example.com");
  });

  it("returns an error and no rows when access is revoked during the scan", async () => {
    const h = makeHarness({ rows: rowsOf(2000) });
    h.setOnYield(() => h.revokeAccess());
    const result = await h.call("batchGetMessageHeaders", {
      messageIds: ["m-1@example.com", "absent@example.com"],
      folderPath: h.folder.URI,
    });
    assert.match(result.error, /not accessible/);
    assert.equal(result.headers, undefined);
  });

  it("counts duplicate ids once in failed but fully in total", async () => {
    const h = makeHarness({ rows: rowsOf(3) });
    const result = await h.call("batchGetMessageHeaders", {
      messageIds: ["m-0@example.com", "m-0@example.com", "absent@example.com", "absent@example.com"],
      folderPath: h.folder.URI,
    });
    assert.equal(result.total, 4);
    assert.equal(result.failed, 1);
  });
});

describe("handler guards and error surfacing", () => {
  it("returns the empty fast path without opening the folder", async () => {
    const h = makeHarness({ rows: rowsOf(1) });
    h.runtime.openFolder = () => assert.fail("must not open the folder");
    const result = await h.call("batchGetMessageHeaders", { messageIds: [], folderPath: h.folder.URI });
    assert.deepEqual(result, { headers: {}, total: 0, failed: 0 });
  });

  it("rejects an over-cap array at the schema, before the handler sees it", async () => {
    const h = makeHarness({ rows: rowsOf(1) });
    const args = { messageIds: Array.from({ length: 100000 }, () => 42), folderPath: h.folder.URI };
    const errors = h.runtime.validateToolArgs("batchGetMessageHeaders", args);
    assert.equal(errors.length, 1);
    assert.match(errors[0], /at most 200 item/);
    assert.equal(h.runtime.validateToolArgs("batchGetMessageHeaders", { messageIds: ids(200), folderPath: h.folder.URI }).length, 0);
  });

  it("keeps a handler-side cap and rejects non-string ids for direct callers", async () => {
    const h = makeHarness({ rows: rowsOf(1) });
    const tooMany = await h.runtime.batchGetMessageHeaders(ids(201), h.folder.URI);
    assert.match(tooMany.error, /Too many ids: 201/);
    const bad = await h.runtime.batchGetMessageHeaders(["m-0@example.com", null], h.folder.URI);
    assert.match(bad.error, /non-empty strings/);
    const notArray = await h.runtime.batchGetMessageHeaders("m-0@example.com", h.folder.URI);
    assert.match(notArray.error, /must be an array/);
  });

  it("surfaces unexpected exceptions as { error }, not as a thrown rejection", async () => {
    const h = makeHarness({ rows: rowsOf(1) });
    h.runtime.openFolder = () => { throw new Error("boom"); };
    const single = await h.runtime.getMessageHeaders("m-0@example.com", h.folder.URI);
    assert.match(single.error, /boom/);
    const batch = await h.runtime.batchGetMessageHeaders(["m-0@example.com"], h.folder.URI);
    assert.match(batch.error, /boom/);
  });

  it("returns the folder error for unknown folders", async () => {
    const h = makeHarness({ rows: rowsOf(1) });
    const result = await h.call("batchGetMessageHeaders", { messageIds: ["m-0@example.com"], folderPath: "imap://account1/Missing" });
    assert.match(result.error, /Folder not found/);
  });

  it("keeps a header conversion failure local to its id", async () => {
    const h = makeHarness({ rows: [{ throwOn: ["isRead"] }, {}] });
    const result = await h.call("batchGetMessageHeaders", { messageIds: ["m-0@example.com", "m-1@example.com"], folderPath: h.folder.URI });
    assert.deepEqual(result.headers["m-0@example.com"], { error: "Could not read message header" });
    assert.equal(result.headers["m-1@example.com"].id, "m-1@example.com");
    assert.equal(result.failed, 1);
  });
});

describe("bracketed and untrimmed ids", () => {
  it("matches bracketed ids on the direct path and keys results by the caller's id", async () => {
    const h = makeHarness({ rows: rowsOf(50) });
    const result = await h.call("batchGetMessageHeaders", {
      messageIds: ["<m-1@example.com>", " m-2@example.com ", "m-1@example.com"],
      folderPath: h.folder.URI,
    });
    assert.equal(result.failed, 0);
    assert.equal(result.total, 3);
    assert.deepEqual(Object.keys(result.headers).sort(), [" m-2@example.com ", "<m-1@example.com>", "m-1@example.com"]);
    assert.equal(result.headers["<m-1@example.com>"].id, "m-1@example.com");
    assert.equal(result.headers[" m-2@example.com "].id, "m-2@example.com");
    assert.equal(h.stats.listAllKeys, 0, "no scan for a resolvable id");
    assert.equal(h.stats.yields, 0);
    const single = await h.call("getMessageHeaders", { messageId: "<m-3@example.com>", folderPath: h.folder.URI });
    assert.equal(single.id, "m-3@example.com");
  });

  it("matches bracketed ids on the scan path", async () => {
    const rows = rowsOf(800);
    rows[700].indirectOnly = true;
    const h = makeHarness({ rows });
    const result = await h.call("batchGetMessageHeaders", { messageIds: ["<m-700@example.com>"], folderPath: h.folder.URI });
    assert.equal(result.failed, 0);
    assert.equal(result.headers["<m-700@example.com>"].id, "m-700@example.com");
    assert.equal(h.stats.listAllKeys, 1);
    const single = await h.call("getMessageHeaders", { messageId: "<m-700@example.com>", folderPath: h.folder.URI });
    assert.equal(single.id, "m-700@example.com");
  });

  it("reports an id that normalises to nothing as not found without scanning", async () => {
    const h = makeHarness({ rows: [{ hdr: { messageId: "" } }, ...rowsOf(500)] });
    const result = await h.call("batchGetMessageHeaders", { messageIds: ["<>", "m-1@example.com"], folderPath: h.folder.URI });
    assert.deepEqual(result.headers["<>"], { error: "not found in this folder" });
    assert.equal(result.headers["m-1@example.com"].id, "m-1@example.com");
    assert.equal(h.stats.listAllKeys, 0);
  });
});

describe("scan robustness", () => {
  it("keeps resolved headers and reports the rest as notScanned when the database getter throws after a yield", async () => {
    const rows = rowsOf(2000);
    rows[1500].indirectOnly = true;
    const h = makeHarness({ rows });
    let reads = 0;
    // 1 openFolder, 2 listAllKeys, 3 first chunk, 4 second chunk (throws).
    h.setDatabaseHook(() => { if (++reads >= 4) throw new Error("database closed"); });
    const result = await h.call("batchGetMessageHeaders", {
      messageIds: ["m-1@example.com", "m-1500@example.com"],
      folderPath: h.folder.URI,
    });
    assert.equal(result.error, undefined);
    assert.equal(result.headers["m-1@example.com"].id, "m-1@example.com");
    assert.equal(result.headers["m-1500@example.com"].notScanned, true);
    assert.deepEqual(result.notScanned, ["m-1500@example.com"]);
    assert.equal(result.truncated, true);
    assert.equal(h.stats.yields, 1);
  });

  it("treats an expunged direct-lookup hit as gone", async () => {
    const h = makeHarness({ rows: [{ hdr: { flags: 0x8 } }, {}] });
    const result = await h.call("batchGetMessageHeaders", {
      messageIds: ["m-0@example.com", "m-1@example.com"],
      folderPath: h.folder.URI,
    });
    assert.deepEqual(result.headers["m-0@example.com"], { error: "not found in this folder" });
    assert.equal(result.headers["m-1@example.com"].id, "m-1@example.com");
    assert.equal(result.truncated, undefined);
    const single = await h.call("getMessageHeaders", { messageId: "m-0@example.com", folderPath: h.folder.URI });
    assert.deepEqual(single, { error: "Message not found: m-0@example.com" });
  });
});

describe("inReplyTo from parsed references", () => {
  it("prefers Thunderbird's parsed references and falls back to the stored string", async () => {
    const parsed = ["root@example.com", "parent@example.com"];
    const h = makeHarness({
      rows: [
        { properties: { references: "<other@example.com>" }, hdr: { numReferences: 2, getStringReference: index => parsed[index] } },
        { properties: { references: "<a@example.com> <b@example.com>" }, hdr: { numReferences: 0, getStringReference: () => assert.fail("not used") } },
        { properties: { references: "<c@example.com>" }, throwOn: ["numReferences"] },
      ],
    });
    const read = async index => h.call("getMessageHeaders", { messageId: `m-${index}@example.com`, folderPath: h.folder.URI });
    assert.equal((await read(0)).inReplyTo, "parent@example.com");
    assert.equal((await read(0)).references, "<other@example.com>");
    assert.equal((await read(1)).inReplyTo, "b@example.com");
    assert.equal((await read(2)).inReplyTo, "c@example.com");
  });
});

describe("scanFolderCooperatively", () => {
  const scanOf = (deadline = 1e9) => ({ deadline, truncated: false });
  async function run(h, visit, options, scan = scanOf()) {
    const outcome = await h.runtime.scanFolderCooperatively(h.folder, scan, visit, options);
    return { ...outcome };
  }

  it("visits live headers in key order, yielding every SEARCH_YIELD_EVERY rows", async () => {
    const h = makeHarness({ rows: rowsOf(600) });
    const seen = [];
    const outcome = await run(h, hdr => { seen.push(hdr.messageId); });
    assert.deepEqual(outcome, { complete: true, stopped: false, scanned: 600 });
    assert.deepEqual(seen, ids(600));
    assert.equal(h.stats.listAllKeys, 1);
    assert.equal(h.stats.yields, 2, "three chunks of 250 yield twice");
  });

  it("walks the snapshot from the end with newestFirst, across chunks", async () => {
    const h = makeHarness({ rows: rowsOf(600) });
    const seen = [];
    const outcome = await run(h, hdr => { seen.push(hdr.messageId); }, { newestFirst: true });
    assert.equal(outcome.complete, true);
    assert.equal(outcome.scanned, 600);
    assert.deepEqual(seen, ids(600).reverse());
  });

  it("stops when visit returns true, without yielding or visiting further", async () => {
    const h = makeHarness({ rows: rowsOf(1000) });
    const seen = [];
    const outcome = await run(h, hdr => { seen.push(hdr.messageId); return seen.length === 3; }, { newestFirst: true });
    assert.deepEqual(outcome, { complete: true, stopped: true, scanned: 3 });
    assert.deepEqual(seen, ["m-999@example.com", "m-998@example.com", "m-997@example.com"]);
    assert.equal(h.stats.yields, 0);
  });

  it("only a strict true from visit stops the scan", async () => {
    const h = makeHarness({ rows: rowsOf(10) });
    const outcome = await run(h, () => 1);
    assert.deepEqual(outcome, { complete: true, stopped: false, scanned: 10 });
  });

  it("skips expunged rows without counting them", async () => {
    const h = makeHarness({ rows: [{ hdr: { flags: 0x8 } }, {}, {}] });
    const seen = [];
    const outcome = await run(h, hdr => { seen.push(hdr.messageId); });
    assert.deepEqual(outcome, { complete: true, stopped: false, scanned: 2 });
    assert.deepEqual(seen, ["m-1@example.com", "m-2@example.com"]);
  });

  it("marks the scan incomplete on unreadable rows and a throwing visitor, but continues", async () => {
    const rows = rowsOf(6);
    rows[2].throwOn = ["flags"];
    const h = makeHarness({ rows });
    const seen = [];
    const outcome = await run(h, hdr => {
      seen.push(hdr.messageId);
      if (hdr.messageId === "m-4@example.com") throw new Error("visitor failed");
    });
    assert.equal(outcome.complete, false);
    assert.equal(outcome.stopped, false);
    assert.deepEqual(seen, ["m-0@example.com", "m-1@example.com", "m-3@example.com", "m-4@example.com", "m-5@example.com"]);
  });

  it("marks rows that vanish after the snapshot as incomplete", async () => {
    const h = makeHarness({ rows: rowsOf(600) });
    h.setOnYield(() => h.headers.delete(400));
    const seen = [];
    const outcome = await run(h, hdr => { seen.push(hdr.messageId); });
    assert.equal(outcome.complete, false);
    assert.equal(outcome.scanned, 599);
    assert.ok(!seen.includes("m-399@example.com"));
  });

  it("reports an expired time budget as incomplete and flags the scan", async () => {
    const h = makeHarness({ rows: rowsOf(10) });
    const scan = scanOf(-1);
    const seen = [];
    const outcome = await run(h, hdr => { seen.push(hdr); }, undefined, scan);
    assert.deepEqual(outcome, { complete: false, stopped: false, scanned: 0 });
    assert.equal(seen.length, 0);
    assert.equal(scan.timedOut, true);
  });

  it("stops between chunks when the budget runs out", async () => {
    const h = makeHarness({ rows: rowsOf(2000), msPerYield: 6000 });
    const scan = scanOf(20000);
    const outcome = await run(h, () => {}, undefined, scan);
    assert.equal(outcome.complete, false);
    assert.equal(scan.timedOut, true);
    assert.ok(outcome.scanned < 2000 && outcome.scanned % 250 === 0);
  });

  it("is incomplete when access is lost, the database getter throws or keys cannot be listed", async () => {
    const revoked = makeHarness({ rows: rowsOf(1000) });
    revoked.setOnYield(() => revoked.revokeAccess());
    const lost = await run(revoked, () => {});
    assert.equal(lost.complete, false);
    assert.equal(lost.scanned, 250);

    const throwing = makeHarness({ rows: rowsOf(1000) });
    let reads = 0;
    throwing.setDatabaseHook(() => { if (++reads >= 3) throw new Error("closed"); });
    const closed = await run(throwing, () => {});
    assert.equal(closed.complete, false);
    assert.equal(closed.scanned, 250);

    const noKeys = makeHarness({ rows: rowsOf(10) });
    noKeys.setDatabaseHook(() => { throw new Error("no database"); });
    assert.deepEqual(await run(noKeys, () => assert.fail("must not visit")), { complete: false, stopped: false, scanned: 0 });
  });

  it("retains no database across yields", async () => {
    // The harness fails any database call made after the generation changed,
    // and every yield bumps the generation.
    const h = makeHarness({ rows: rowsOf(800) });
    const outcome = await run(h, () => {}, { newestFirst: true });
    assert.equal(outcome.complete, true);
    assert.ok(h.stats.yields >= 3);
  });
});
