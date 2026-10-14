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

// The real production schemas of the header and query tools, evaluated in isolation.
function productionTools() {
  const startLf = source.indexOf('      {\n        name: "getMessageHeaders"');
  const from = startLf >= 0 ? startLf : source.indexOf('      {\r\n        name: "getMessageHeaders"');
  const end = source.indexOf('name: "sendMail"', from);
  const to = source.lastIndexOf("      {", end);
  assert.ok(from >= 0 && to > from, "production tool schemas missing");
  const sandbox = {};
  vm.runInNewContext(`this.tools = [${source.slice(from, to)}];`, sandbox);
  return sandbox.tools;
}

const ENCRYPTED_RAW = 'From: sender@example.test\r\nSubject: =?utf-8?q?Encrypted_Message?=\r\nContent-Type: multipart/encrypted; protocol="application/pgp-encrypted"; boundary="b"\r\n\r\n--b--\r\n';
const DECRYPTION_OKAY = 0x00020000;
const INBOX_FLAG = 0x1000;
const NO_THREAD = 0xffffffff;

// Minimal stand-in for MailServices.headerParser.parseEncodedHeader.
let parseCalls = 0;
function parseEncodedHeader(raw) {
  parseCalls++;
  return (raw.match(/(?:[^,"]|"[^"]*")+/g) || []).map(part => {
    const angle = part.match(/<([^>]+)>/);
    const email = (angle ? angle[1] : part).trim();
    return { email, name: angle ? part.replace(angle[0], "").trim().replace(/^"|"$/g, "") : "" };
  });
}

/**
 * folders: [{ uri, rows, noInbox, threadApi, direct }]. Every folder is its own
 * account whose server root (isServer, no database) holds the folder as its
 * Inbox unless noInbox. Row fields: id, key (message key, default index + 1),
 * thread, author, date (ms; 0 = undated), hdr, properties, enigmail, throwOn.
 * Message keys follow insertion order, like arrival order, not date order.
 */
function makeHarness({ folders, allowEncrypted = false, msPerYield = 0 } = {}) {
  const stats = { yields: 0, listAllKeys: 0, headerReads: 0, hydrated: 0 };
  let onDirectLookup = null;
  parseCalls = 0;
  let generation = 0;
  let now = 1700000000000;
  const revoked = new Set();
  let onYield = null;
  const byUri = new Map();
  const accounts = [];

  for (const [folderIndex, spec] of folders.entries()) {
    const headers = new Map();
    const byMessageId = new Map();
    const folder = {
      URI: spec.uri,
      prettyName: "INBOX",
      isServer: false,
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
        const openedGeneration = generation;
        const assertLive = () => assert.equal(openedGeneration, generation, "Database retained across a yield");
        const db = {
          listAllKeys() { assertLive(); stats.listAllKeys++; return Array.from(headers.keys()); },
          enumerateMessages() { assert.fail("A folder query must not open a native message enumerator"); },
          containsKey(key) { assertLive(); return headers.has(key); },
          getMsgHdrForKey(key) { assertLive(); stats.headerReads++; return headers.get(key); },
        };
        if (spec.direct !== false) {
          db.getMsgHdrForMessageID = id => {
            assertLive();
            const found = byMessageId.get(id) || null;
            if (found && onDirectLookup) onDirectLookup(found);
            return found;
          };
        }
        if (spec.threadApi !== false) {
          db.getThreadContainingMsgHdr = hdr => {
            assertLive();
            const keys = [...headers.values()].filter(h => h.threadId === hdr.threadId).map(h => h.messageKey);
            return { numChildren: keys.length, getChildKeyAt: i => keys[i] };
          };
        }
        return db;
      },
    };
    spec.rows.forEach((row, index) => {
      const properties = { keywords: "custom", ...row.properties };
      const header = {
        messageKey: row.key === undefined ? index + 1 : row.key,
        messageId: row.id || `f${folderIndex}-m${index}@example.com`,
        threadId: row.thread === undefined ? 100 + index : row.thread,
        date: row.date === undefined ? (index + 1) * 1000000 : row.date * 1000,
        subject: "Raw subject",
        author: row.author || "Alice Smith <alice@example.com>",
        recipients: "Bob Jones",
        mime2DecodedSubject: "Decoded subject",
        mime2DecodedAuthor: row.author || "Alice Smith <alice@example.com>",
        mime2DecodedRecipients: "Bob Jones",
        ccList: "",
        flags: 0,
        isRead: true,
        isFlagged: false,
        messageSize: 321,
        folder,
        getStringProperty: name => {
          if (name === "references") stats.hydrated++;
          return properties[name] || "";
        },
        getUint32Property: name => (name === "enigmail" ? (row.enigmail || 0) : 0),
        ...row.hdr,
      };
      for (const field of row.throwOn || []) {
        Object.defineProperty(header, field, { get() { throw new Error(`unreadable ${field}`); } });
      }
      headers.set(header.messageKey, header);
      byMessageId.set(header.messageId, header);
    });
    const root = {
      URI: `imap://account${folderIndex}`,
      isServer: true,
      getFolderWithFlags: flag => (flag === INBOX_FLAG && !spec.noInbox ? folder : null),
      get msgDatabase() { throw new Error("a server root has no database"); },
    };
    const server = { type: "imap", rootFolder: root };
    folder.server = server;
    root.server = server;
    const account = { key: `account${folderIndex}`, incomingServer: server };
    accounts.push(account);
    byUri.set(spec.uri, { folder, account, headers });
  }

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
    Ci: { nsMsgMessageFlags: { Expunged: 0x8 }, nsMsgFolderFlags: { Inbox: INBOX_FLAG }, nsITimer: { TYPE_ONE_SHOT: 0 } },
    NetUtil: { readInputStreamToString: (stream, count) => stream.read(count) },
    MailServices: {
      mimeConverter: {
        decodeMimeHeader: value => value.replace(/=\?utf-8\?q\?([^?]*)\?=/gi, (_, text) => text.replace(/_/g, " ")),
      },
      headerParser: { parseEncodedHeader },
      folderLookup: { getFolderForURL: uri => (byUri.has(uri) ? byUri.get(uri).folder : null) },
      accounts: {
        accounts,
        findAccountForServer: server => accounts.find(a => a.incomingServer === server) || null,
      },
    },
    isAccountAllowed: key => !revoked.has(key),
    getAllowedAccountIds: () => [],
    console,
  });
  const keywordStart = source.indexOf("const INTERNAL_KEYWORDS = new Set([");
  const keywordEnd = source.indexOf("]);", keywordStart);
  const tools = productionTools();
  runtime.buildTools = () => tools;
  const calendarPattern = source.match(/^\s*const CALENDAR_EMAIL_PATTERN = [^\r\n]+;/m);
  assert.ok(calendarPattern, "Production constant missing: CALENDAR_EMAIL_PATTERN");
  vm.runInContext([
    calendarPattern[0],
    ...[
      "SEARCH_YIELD_EVERY", "SEARCH_TIME_BUDGET_MS", "MAX_HEADER_BATCH_IDS", "PREF_ALLOW_ENCRYPTED_MESSAGES",
      "DEFAULT_THREAD_RESULTS", "MAX_THREAD_RESULTS", "DEFAULT_SENDER_RESULTS", "MAX_SENDER_RESULTS",
      "DEFAULT_SENDER_SCAN_CAP", "MAX_SENDER_SCAN_CAP", "NSMSGKEY_NONE",
    ].map(constant),
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
    stats,
    now: () => now,
    parses: () => parseCalls,
    setOnDirectLookup(callback) { onDirectLookup = callback; },
    removeRow(uri, key) {
      const { headers } = byUri.get(uri);
      const header = headers.get(key);
      headers.delete(key);
      return header;
    },
    revoke(uri) { revoked.add(byUri.get(uri).account.key); },
    setOnYield(callback) { onYield = callback; },
    async call(name, args) {
      const result = await runtime.callTool(name, args);
      return JSON.parse(JSON.stringify(result));
    },
  };
}

const INBOX = "imap://account0/INBOX";
const rowsOf = (count, make = () => ({})) => Array.from({ length: count }, (_, i) => make(i));
const oneFolder = (rows, extra = {}) => makeHarness({ folders: [{ uri: INBOX, rows, ...extra }] });
const idsOf = list => list.map(m => m.id);

describe("query tool schemas", () => {
  const tools = Object.fromEntries(productionTools().map(t => [t.name, t.inputSchema]));

  it("declare the bounds the handlers clamp to", () => {
    const thread = tools.searchByThread.properties;
    assert.equal(thread.maxResults.minimum, 1);
    assert.equal(thread.maxResults.maximum, 500);
    assert.equal(thread.messageId.minLength, 1);
    assert.equal(thread.folderPath.minLength, 1);
    const sender = tools.getSenderHistory.properties;
    assert.equal(sender.maxResults.maximum, 200);
    assert.equal(sender.scanCap.maximum, 50000);
    assert.equal(sender.scanCap.minimum, 1);
    assert.equal(sender.sinceDays.minimum, 1);
    assert.equal(sender.email.minLength, 1);
  });

  it("keep the schema maximums in sync with the handler constants", () => {
    const h = oneFolder([]);
    assert.equal(vm.runInContext("MAX_THREAD_RESULTS", h.runtime), tools.searchByThread.properties.maxResults.maximum);
    assert.equal(vm.runInContext("MAX_SENDER_RESULTS", h.runtime), tools.getSenderHistory.properties.maxResults.maximum);
    assert.equal(vm.runInContext("MAX_SENDER_SCAN_CAP", h.runtime), tools.getSenderHistory.properties.scanCap.maximum);
  });

  it("reject empty strings through the production validator", () => {
    const h = oneFolder([]);
    assert.match(h.runtime.validateToolArgs("searchByThread", { messageId: "", folderPath: INBOX })[0], /at least 1 character/);
    assert.match(h.runtime.validateToolArgs("getSenderHistory", { email: "" })[0], /at least 1 character/);
  });
});

describe("searchByThread", () => {
  it("returns the shared header shape newest-first from the thread object, without scanning the folder", async () => {
    // Message keys are arrival order and deliberately not in date order.
    const h = oneFolder([
      { id: "a@x", thread: 7, date: 1000 },
      { id: "b@x", thread: 7, date: 3000 },
      { id: "other@x", thread: 8, date: 5000 },
      { id: "c@x", thread: 7, date: 2000 },
    ].concat(rowsOf(2000, i => ({ thread: 1000 + i }))));
    const result = await h.call("searchByThread", { messageId: "a@x", folderPath: INBOX });
    assert.equal(result.threadId, 7);
    assert.equal(typeof result.threadId, "number");
    assert.equal(result.anchorId, "a@x");
    assert.deepEqual(idsOf(result.thread), ["b@x", "c@x", "a@x"]);
    assert.equal(result.total, 3);
    assert.equal(result.truncated, false);
    assert.equal(result.notScanned, undefined);
    assert.deepEqual(Object.keys(result.thread[0]).sort(), [
      "author", "ccList", "date", "flagged", "id", "inReplyTo", "read", "recipients",
      "references", "size", "subject", "tags", "threadId",
    ]);
    assert.equal(h.stats.listAllKeys, 0, "no folder enumeration");
    assert.equal(h.stats.yields, 0);
    assert.ok(h.stats.headerReads <= 10, `bounded by the thread size, read ${h.stats.headerReads}`);
  });

  it("accepts a bracketed anchor id", async () => {
    const h = oneFolder([{ id: "a@x", thread: 7 }]);
    const result = await h.call("searchByThread", { messageId: "<a@x>", folderPath: INBOX });
    assert.equal(result.anchorId, "a@x");
  });

  it("treats thread id 0 and message key 0 as real, and nsMsgKey_None as no thread", async () => {
    const zero = oneFolder([
      { id: "a@x", key: 0, thread: 0, date: 1000 },
      { id: "b@x", key: 5, thread: 0, date: 2000 },
      { id: "c@x", key: 6, thread: 9, date: 3000 },
    ]);
    const result = await zero.call("searchByThread", { messageId: "a@x", folderPath: INBOX });
    assert.equal(result.threadId, 0);
    assert.equal(result.anchored, undefined);
    assert.deepEqual(idsOf(result.thread), ["b@x", "a@x"]);

    const none = oneFolder([{ id: "a@x", thread: NO_THREAD }, { id: "b@x", thread: NO_THREAD }]);
    const alone = await none.call("searchByThread", { messageId: "a@x", folderPath: INBOX });
    assert.equal(alone.threadId, null);
    assert.equal(alone.anchored, true);
    assert.deepEqual(idsOf(alone.thread), ["a@x"]);
    assert.equal(alone.truncated, false);
  });

  it("finds an anchor that misses the direct lookup and shares one budget with the thread read", async () => {
    const rows = rowsOf(3000, i => ({ thread: i % 700 === 0 ? 4 : 1000 + i }));
    const h = oneFolder(rows, { direct: false });
    const result = await h.call("searchByThread", { messageId: "f0-m2801@example.com", folderPath: INBOX });
    assert.equal(result.threadId, 1000 + 2801);
    assert.equal(result.total, 1);

    const shared = makeHarness({ folders: [{ uri: INBOX, rows: rowsOf(3000, () => ({ thread: 4 })), direct: false }], msPerYield: 4000 });
    const budget = await shared.call("searchByThread", { messageId: "f0-m600@example.com", folderPath: INBOX });
    assert.equal(budget.notScanned, true, "the anchor scan consumed the budget the thread read needed");
    assert.equal(budget.anchorId, "f0-m600@example.com");
    assert.ok(idsOf(budget.thread).includes("f0-m600@example.com"));
    assert.match(budget.message, /20-second time budget/);
  });

  it("falls back to a cooperative folder scan when no thread object is available", async () => {
    const rows = rowsOf(1000, i => ({ thread: i === 3 || i === 900 ? 4 : 1000 + i }));
    const h = oneFolder(rows, { threadApi: false });
    const result = await h.call("searchByThread", { messageId: "f0-m3@example.com", folderPath: INBOX });
    assert.equal(result.total, 2);
    assert.equal(result.truncated, false);
    assert.ok(h.stats.yields >= 3);
  });

  it("reports the result cap as truncated, keeps the newest rows and always the anchor", async () => {
    const h = oneFolder(rowsOf(30, i => ({ thread: 1, date: i + 1 })));
    const result = await h.call("searchByThread", { messageId: "f0-m0@example.com", folderPath: INBOX, maxResults: 5 });
    assert.equal(result.thread.length, 5);
    assert.equal(result.total, 30);
    assert.equal(result.truncated, true);
    assert.equal(result.notScanned, undefined);
    assert.equal(result.thread[0].id, "f0-m29@example.com");
    assert.ok(idsOf(result.thread).includes("f0-m0@example.com"), "the oldest anchor is kept");
  });

  it("clamps an over-cap maxResults to 500", async () => {
    const h = oneFolder(rowsOf(600, () => ({ thread: 1 })));
    const result = await h.call("searchByThread", { messageId: "f0-m0@example.com", folderPath: INBOX, maxResults: 99999 });
    assert.equal(result.thread.length, 500);
    assert.equal(result.total, 600);
    assert.equal(result.truncated, true);
  });

  it("flags an unfinished thread read as notScanned with the anchor still present", async () => {
    const h = makeHarness({ folders: [{ uri: INBOX, rows: rowsOf(5000, () => ({ thread: 1 })) }], msPerYield: 6000 });
    const result = await h.call("searchByThread", { messageId: "f0-m0@example.com", folderPath: INBOX });
    assert.equal(result.truncated, true);
    assert.equal(result.notScanned, true);
    assert.match(result.message, /20-second time budget/);
    assert.ok(idsOf(result.thread).includes("f0-m0@example.com"));
    assert.ok(h.stats.headerReads < 5000, "bounded by the time budget");
  });

  it("treats unreadable members as an incomplete read", async () => {
    const rows = rowsOf(20, () => ({ thread: 1 }));
    rows[5].throwOn = ["flags"];
    const result = await oneFolder(rows).call("searchByThread", { messageId: "f0-m0@example.com", folderPath: INBOX });
    assert.equal(result.notScanned, true);
    assert.doesNotMatch(result.message, /time budget/);
  });

  it("distinguishes an absent anchor from an unscanned one", async () => {
    const absent = await oneFolder(rowsOf(10)).call("searchByThread", { messageId: "nope@x", folderPath: INBOX });
    assert.deepEqual(absent, { error: "Message not found: nope@x" });

    const slow = makeHarness({ folders: [{ uri: INBOX, rows: rowsOf(2000) }], msPerYield: 25000 });
    const unknown = await slow.call("searchByThread", { messageId: "nope@x", folderPath: INBOX });
    assert.equal(unknown.notScanned, true);
    assert.match(unknown.error, /time budget/);
  });

  it("re-reads the anchor for the final page and flags it when it vanished during the read", async () => {
    const h = oneFolder(rowsOf(600, () => ({ thread: 1 })));
    h.setOnYield(count => { if (count === 1) h.removeRow(INBOX, 1); });
    const result = await h.call("searchByThread", { messageId: "f0-m0@example.com", folderPath: INBOX });
    assert.equal(result.anchorId, "f0-m0@example.com");
    assert.ok(!idsOf(result.thread).includes("f0-m0@example.com"), "a stale anchor is not returned");
    assert.equal(result.notScanned, true);
    assert.equal(result.truncated, true);
    assert.match(result.message, /disappeared/);
  });

  it("reports a vanished anchor instead of scanning the folder", async () => {
    const h = oneFolder(rowsOf(2000, i => ({ thread: i < 3 ? 1 : 1000 + i })));
    h.setOnDirectLookup(hdr => h.removeRow(INBOX, hdr.messageKey));
    const result = await h.call("searchByThread", { messageId: "f0-m0@example.com", folderPath: INBOX });
    assert.deepEqual(result, { error: "Message not found: f0-m0@example.com" });
    assert.equal(h.stats.listAllKeys, 0, "no fallback folder scan");
  });

  it("withholds encrypted subjects like the base header tools", async () => {
    const rows = [{ id: "e@x", thread: 1, enigmail: DECRYPTION_OKAY, hdr: { raw: ENCRYPTED_RAW, mime2DecodedSubject: "Cached decrypted subject" } }];
    const blocked = await oneFolder(rows).call("searchByThread", { messageId: "e@x", folderPath: INBOX });
    assert.equal(blocked.thread[0].subject, "Encrypted Message");
    assert.doesNotMatch(JSON.stringify(blocked), /Cached decrypted/);
    const open = await makeHarness({ folders: [{ uri: INBOX, rows }], allowEncrypted: true })
      .call("searchByThread", { messageId: "e@x", folderPath: INBOX });
    assert.equal(open.thread[0].subject, "Cached decrypted subject");
  });

  it("returns an error and no rows when access is revoked mid-read", async () => {
    const h = oneFolder(rowsOf(2000, () => ({ thread: 1 })));
    h.setOnYield(() => h.revoke(INBOX));
    const result = await h.call("searchByThread", { messageId: "f0-m0@example.com", folderPath: INBOX });
    assert.match(result.error, /not accessible/);
    assert.equal(result.thread, undefined);
  });

  it("returns { error } for guards, unknown folders and unexpected exceptions", async () => {
    const h = oneFolder(rowsOf(2));
    assert.match((await h.runtime.searchByThread("", INBOX)).error, /non-empty string/);
    assert.match((await h.call("searchByThread", { messageId: "x", folderPath: "imap://account0/Missing" })).error, /Folder not found/);
    h.runtime.openFolder = () => { throw new Error("boom"); };
    assert.match((await h.runtime.searchByThread("f0-m0@example.com", INBOX)).error, /boom/);
  });
});

describe("getSenderHistory", () => {
  const DAY = 86400 * 1000;

  it("matches the parsed address exactly and case-insensitively, never a substring or a display name", async () => {
    const h = oneFolder([
      { id: "hit1@x", author: "Ann <ANN@example.com>", date: 1000 },
      { id: "hit2@x", author: "ann@example.com", date: 3000 },
      { id: "joann@x", author: "Jo <joann@example.com>", date: 4000 },
      { id: "suffix@x", author: "ann@example.com.evil.test", date: 5000 },
      { id: "spoof@x", author: '"ann@example.com" <attacker@evil.test>', date: 6000 },
      { id: "comma@x", author: '"Doe, Ann" <ann@example.com>', date: 2000 },
      { id: "other@x", author: "Bob <bob@example.com>", date: 9000 },
    ]);
    for (const query of ["Ann@Example.com", "mailto:ann@example.com", "MAILTO:ANN@example.com"]) {
      const result = await h.call("getSenderHistory", { email: query });
      assert.deepEqual(idsOf(result.messages), ["hit2@x", "comma@x", "hit1@x"], query);
      assert.equal(result.truncated, false);
    }
  });

  it("rejects a query that is not one plain address, instead of returning a confident empty result", async () => {
    const h = oneFolder(rowsOf(2));
    for (const query of [
      "alice", "a@x.test, b@x.test", "   ", "Ann <ann@example.com>",
      "mailto:alice@example.com?subject=x", "mailto:alice@example.com?cc=b@x.test",
      "http://alice@example.com", "alice@example.com\n", "ali ce@example.com",
    ]) {
      const result = await h.call("getSenderHistory", { email: query });
      assert.match(result.error, /single email address|non-empty/, JSON.stringify(query));
      assert.equal(result.messages, undefined, JSON.stringify(query));
    }
  });

  it("skips the header parser for rows that cannot match", async () => {
    const h = oneFolder([
      ...rowsOf(500, () => ({ author: "Bob <bob@example.com>" })),
      { id: "hit@x", author: "=?utf-8?q?Ann?= <ann@example.com>" },
    ]);
    const result = await h.call("getSenderHistory", { email: "ann@example.com" });
    assert.deepEqual(idsOf(result.messages), ["hit@x"]);
    assert.ok(h.parses() <= 3, `parser ran ${h.parses()} times for 501 rows`);
  });

  it("returns the shared shape plus folderPath, newest by date despite key order", async () => {
    const h = oneFolder([
      { id: "old@x", date: 1000, key: 9 },
      { id: "new@x", date: 9000, key: 2 },
      { id: "mid@x", date: 5000, key: 4 },
    ]);
    const result = await h.call("getSenderHistory", { email: "alice@example.com" });
    assert.deepEqual(idsOf(result.messages), ["new@x", "mid@x", "old@x"]);
    assert.equal(result.messages[0].folderPath, INBOX);
    assert.equal(result.messages[0].read, true);
    assert.equal(typeof result.messages[0].threadId, "number");
    assert.equal(result.totalScanned, 3);
    assert.equal(result.truncated, false);
    assert.equal(result.notScanned, undefined);
    assert.equal(result.notScannedFolders, undefined);
  });

  it("reads each Inbox newest-first so scanCap examines the newest keys, not the oldest", async () => {
    const h = oneFolder(rowsOf(50, i => ({ id: `m${i}@x`, date: i + 1 })));
    const result = await h.call("getSenderHistory", { email: "alice@example.com", scanCap: 10, maxResults: 200 });
    assert.equal(result.messages.length, 10);
    assert.deepEqual(idsOf(result.messages), Array.from({ length: 10 }, (_, i) => `m${49 - i}@x`));
    assert.equal(result.truncated, true);
    assert.equal(result.notScanned, true);
    assert.deepEqual(result.notScannedFolders, [INBOX]);
    assert.match(result.message, /scanCap of 10/);
    assert.match(result.message, /newest-first[^.]*older/);

    const exact = await oneFolder(rowsOf(10)).call("getSenderHistory", { email: "alice@example.com", scanCap: 10 });
    assert.equal(exact.truncated, false, "a folder exactly at the cap was fully examined");
  });

  it("does not let messages outside sinceDays consume scanCap, and keeps undated ones", async () => {
    const now = oneFolder([]).now();
    const rows = [
      { id: "recent-low@x", date: now - 2 * DAY },
      { id: "recent-mid@x", date: now - 3 * DAY },
      { id: "undated@x", date: 0 },
      ...rowsOf(40, i => ({ id: `stale${i}@x`, date: now - (100 + i) * DAY })),
    ];
    const h = makeHarness({ folders: [{ uri: INBOX, rows }] });
    const result = await h.call("getSenderHistory", { email: "alice@example.com", sinceDays: 30, scanCap: 3 });
    assert.deepEqual(new Set(idsOf(result.messages)), new Set(["recent-low@x", "recent-mid@x", "undated@x"]));
    assert.equal(result.truncated, false, "stale rows were skipped without using the cap");
  });

  it("scans every account so newest-first holds across them", async () => {
    const h = makeHarness({
      folders: [
        { uri: "imap://account0/INBOX", rows: rowsOf(5, i => ({ id: `old${i}@x`, date: 1000 + i })) },
        { uri: "imap://account1/INBOX", rows: rowsOf(3, i => ({ id: `new${i}@x`, date: 50000 + i })) },
      ],
    });
    const result = await h.call("getSenderHistory", { email: "alice@example.com", maxResults: 4 });
    assert.deepEqual(idsOf(result.messages), ["new2@x", "new1@x", "new0@x", "old4@x"]);
    assert.equal(result.truncated, true, "matches exceeded maxResults");
    assert.equal(result.notScanned, undefined);
  });

  it("skips accounts without an Inbox and server roots without reporting truncation", async () => {
    const h = makeHarness({
      folders: [
        { uri: "imap://account0/INBOX", rows: [{ id: "ok@x" }] },
        { uri: "imap://account1/Other", rows: [{ id: "elsewhere@x" }], noInbox: true },
      ],
    });
    const result = await h.call("getSenderHistory", { email: "alice@example.com" });
    assert.deepEqual(idsOf(result.messages), ["ok@x"]);
    assert.equal(result.truncated, false);
    assert.equal(result.notScanned, undefined);
  });

  it("skips accounts that are not accessible", async () => {
    const h = makeHarness({
      folders: [
        { uri: "imap://account0/INBOX", rows: [{ id: "ok@x" }] },
        { uri: "imap://account1/INBOX", rows: [{ id: "hidden@x" }] },
      ],
    });
    h.revoke("imap://account1/INBOX");
    const result = await h.call("getSenderHistory", { email: "alice@example.com" });
    assert.deepEqual(idsOf(result.messages), ["ok@x"]);
  });

  it("scans past 50,000 rows cooperatively and clamps scanCap", async () => {
    const h = oneFolder(rowsOf(60000, i => ({ author: i === 0 ? "target@example.com" : "other@example.com" })));
    const result = await h.call("getSenderHistory", { email: "target@example.com", scanCap: 99999999 });
    assert.equal(result.messages.length, 0, "the oldest key is beyond the 50,000 newest examined");
    assert.deepEqual(result.notScannedFolders, [INBOX]);
    assert.ok(result.totalScanned >= 50000 && result.totalScanned <= 50001);
    assert.ok(h.stats.yields >= 50000 / 250 - 1);
  });

  it("marks later accounts notScannedFolders when the shared time budget runs out", async () => {
    const h = makeHarness({
      folders: [
        { uri: "imap://account0/INBOX", rows: rowsOf(3000) },
        { uri: "imap://account1/INBOX", rows: rowsOf(10) },
      ],
      msPerYield: 6000,
    });
    const result = await h.call("getSenderHistory", { email: "alice@example.com", maxResults: 200 });
    assert.equal(result.truncated, true);
    assert.equal(result.notScanned, true);
    assert.deepEqual(result.notScannedFolders.sort(), ["imap://account0/INBOX", "imap://account1/INBOX"]);
    assert.match(result.message, /20-second time budget/);
    assert.ok(h.stats.headerReads < 3010);
  });

  it("reports every truncation cause when they mix", async () => {
    const now = oneFolder([]).now();
    const h = makeHarness({
      folders: [
        { uri: "imap://account0/INBOX", rows: rowsOf(100, () => ({ date: now - DAY })) },
        { uri: "imap://account1/INBOX", rows: rowsOf(3000, () => ({ date: now - 100 * DAY })) },
      ],
      msPerYield: 8000,
    });
    const result = await h.call("getSenderHistory", { email: "alice@example.com", scanCap: 50, sinceDays: 30 });
    assert.match(result.message, /scanCap of 50/);
    assert.match(result.message, /20-second time budget/);
  });

  it("treats unreadable rows as an incomplete scan", async () => {
    const rows = rowsOf(10);
    rows[2].throwOn = ["flags"];
    const result = await oneFolder(rows).call("getSenderHistory", { email: "alice@example.com" });
    assert.deepEqual(result.notScannedFolders, [INBOX]);
    assert.equal(result.truncated, true);
    assert.match(result.message, /unreadable/);
  });

  it("hydrates only the final page", async () => {
    const h = oneFolder(rowsOf(100, i => ({ date: i + 1 })));
    const result = await h.call("getSenderHistory", { email: "alice@example.com", maxResults: 3 });
    assert.equal(result.messages.length, 3);
    assert.equal(h.stats.hydrated, 3, "header objects were built for three rows, not for every match");
    assert.equal(result.truncated, true);
  });

  it("withholds encrypted subjects like the base header tools", async () => {
    const rows = [{ id: "e@x", enigmail: DECRYPTION_OKAY, hdr: { raw: ENCRYPTED_RAW, mime2DecodedSubject: "Cached decrypted subject" } }];
    const result = await oneFolder(rows).call("getSenderHistory", { email: "alice@example.com" });
    assert.equal(result.messages[0].subject, "Encrypted Message");
    assert.doesNotMatch(JSON.stringify(result), /Cached decrypted/);
  });

  it("exposes nothing about an account revoked mid-scan, and keeps totals consistent", async () => {
    const h = makeHarness({
      folders: [
        { uri: "imap://account0/INBOX", rows: rowsOf(600) },
        { uri: "imap://account1/INBOX", rows: rowsOf(5, i => ({ id: `kept${i}@x` })) },
      ],
    });
    h.setOnYield(() => h.revoke("imap://account0/INBOX"));
    const result = await h.call("getSenderHistory", { email: "alice@example.com", maxResults: 200 });
    assert.deepEqual(idsOf(result.messages).sort(), Array.from({ length: 5 }, (_, i) => `kept${i}@x`));
    assert.ok(result.messages.every(m => m.folderPath === "imap://account1/INBOX"));
    assert.doesNotMatch(JSON.stringify(result), /account0/);
    assert.equal(result.totalScanned, 5);
    assert.equal(result.truncated, false);
  });

  it("clamps maxResults to 200 and returns { error } for bad input or exceptions", async () => {
    const h = oneFolder(rowsOf(300));
    const result = await h.call("getSenderHistory", { email: "alice@example.com", maxResults: 99999 });
    assert.equal(result.messages.length, 200);
    assert.equal(result.truncated, true);
    h.runtime.getAccessibleAccounts = () => { throw new Error("boom"); };
    assert.match((await h.runtime.getSenderHistory("alice@example.com")).error, /boom/);
  });
});
