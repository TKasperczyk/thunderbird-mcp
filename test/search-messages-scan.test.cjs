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

function makeFolder(name, { account = "account1", count = 0, messages, children = [], isServer = false, legacy = false } = {}) {
  const stats = { refreshes: 0, databaseReads: 0, snapshots: 0, headerReads: 0, contentReads: 0, propertyReads: 0 };
  const hooks = {};
  let generation = 0;
  const headers = new Map();
  const folder = {
    URI: `imap://${account}/${name}`,
    prettyName: name,
    account,
    isServer,
    subFolders: children,
    get hasSubFolders() { return children.length > 0; },
    updateFolder() { stats.refreshes++; },
    invalidateDatabase() { generation++; },
    headers,
    stats,
    hooks,
    get msgDatabase() {
      stats.databaseReads++;
      if (hooks.database) hooks.database();
      const openedGeneration = generation;
      const assertLive = () => assert.equal(openedGeneration, generation, "Database retained across a yield");
      const database = {
        listAllKeys() {
          assertLive();
          stats.snapshots++;
          if (hooks.snapshot) hooks.snapshot();
          return Array.from(headers.keys());
        },
        enumerateMessages() { assert.fail("A search must not open a native message enumerator"); },
      };
      database[legacy ? "ContainsKey" : "containsKey"] = key => {
        assertLive();
        return headers.has(key);
      };
      database[legacy ? "GetMsgHdrForKey" : "getMsgHdrForKey"] = key => {
        assertLive();
        stats.headerReads++;
        if (hooks.header) hooks.header(key);
        return headers.get(key);
      };
      return database;
    },
  };
  const rows = messages || Array.from({ length: count }, (_, index) => ({ messageId: `m-${index}@example.com`, date: (index + 1) * 1000000 }));
  rows.forEach((row, index) => {
    const properties = { preview: "A preview", keywords: "custom junk $label1", ...row.properties };
    const header = {
      messageKey: index + 1,
      messageId: `m-${index}@example.com`,
      threadId: index,
      date: (index + 1) * 1000000,
      subject: "Raw subject",
      author: "Raw author",
      recipients: "Raw recipient",
      ccList: "copy@example.com",
      flags: 0,
      isRead: false,
      isFlagged: true,
      folder,
      getStringProperty(name) { stats.propertyReads++; return properties[name] || ""; },
      ...row,
    };
    for (const [field, fallback] of [["mime2DecodedSubject", "Decoded subject"], ["mime2DecodedAuthor", "Alice Smith"], ["mime2DecodedRecipients", "Bob Jones"]]) {
      Object.defineProperty(header, field, {
        configurable: true,
        get() { stats.contentReads++; return row[field] ?? fallback; },
      });
    }
    headers.set(header.messageKey, header);
  });
  return folder;
}

function makeHarness({ roots = [makeFolder("INBOX", { count: 3 })], onYield, glodaItems, glodaLimit = 20000, waitForGloda = false } = {}) {
  const folders = new Map();
  function addFolder(folder) {
    folders.set(folder.URI, folder);
    for (const child of folder.subFolders) addFolder(child);
  }
  roots.forEach(addFolder);
  const accounts = roots.map(root => ({ key: root.account, incomingServer: { type: "imap", rootFolder: root } }));
  for (const folder of folders.values()) folder.server = accounts.find(account => account.key === folder.account).incomingServer;
  const allowedAccounts = new Set(accounts.map(account => account.key));
  const timers = [];
  const searchers = [];
  let now = 0;
  let yields = 0;
  let harness;
  class SearchDate extends Date {
    static now() { return now; }
  }
  class GlodaMsgSearcher {
    constructor(listener) {
      this.listener = listener;
      this.retrievalLimit = glodaLimit;
      searchers.push(this);
    }
    getCollection() {
      if (!waitForGloda) queueMicrotask(() => this.listener?.onQueryCompleted({ items: glodaItems || [] }));
    }
  }
  const runtime = vm.createContext({
    Date: SearchDate,
    GlodaMsgSearcher,
    Services: {
      tm: {
        dispatchToMainThread(resolve) {
          setImmediate(() => {
            yields++;
            for (const folder of folders.values()) folder.invalidateDatabase();
            if (onYield) onYield(harness);
            resolve();
          });
        },
      },
    },
    Cc: {
      "@mozilla.org/timer;1": {
        createInstance() {
          const timer = {
            canceled: false,
            initWithCallback(callback, interval) { this.callback = callback; this.interval = interval; },
            cancel() { this.canceled = true; },
            fire() { if (!this.canceled) this.callback(); },
          };
          timers.push(timer);
          return timer;
        },
      },
    },
    Ci: { nsMsgMessageFlags: { Expunged: 0x8 }, nsITimer: { TYPE_ONE_SHOT: 0 } },
    MailServices: {
      folderLookup: { getFolderForURL: uri => folders.get(uri) },
      accounts: { accounts, findAccountForServer: server => accounts.find(account => account.incomingServer === server) },
    },
    isAccountAllowed: key => allowedAccounts.has(key),
    getAllowedAccountIds: () => [],
    buildTools: () => [{ name: "searchMessages", group: "messages" }],
  });
  const keywordStart = source.indexOf("const INTERNAL_KEYWORDS = new Set([");
  const keywordEnd = source.indexOf("]);", keywordStart);
  assert.ok(keywordStart >= 0 && keywordEnd > keywordStart);
  vm.runInContext([
    ...["DEFAULT_MAX_RESULTS", "MAX_SEARCH_RESULTS_CAP", "SEARCH_YIELD_EVERY", "SEARCH_TIME_BUDGET_MS"].map(constant),
    source.slice(keywordStart, keywordEnd + 3),
    snippet("// BEGIN SEARCH RESULT HELPERS", "// END SEARCH RESULT HELPERS"),
    snippet("function isFolderAccessible(", "function toColumnarTable("),
    snippet("function getUserTags(", "// BEGIN OUTBOUND ATTACHMENT CONVERSION"),
    snippet("// BEGIN MESSAGE SEARCH", "// END MESSAGE SEARCH"),
    snippet("// BEGIN MCP TEXT SANITIZATION", "// END MCP TEXT SANITIZATION"),
    snippet("// BEGIN TOOL DISPATCH", "// END TOOL DISPATCH"),
  ].join("\n"), runtime);
  harness = {
    folders, roots, allowedAccounts, timers, searchers, runtime,
    advance(milliseconds) { now += milliseconds; },
    get yields() { return yields; },
    async search(args = {}) {
      const result = await runtime.callTool("searchMessages", { query: "", folderPath: roots[0].URI, ...args });
      return JSON.parse(JSON.stringify(result));
    },
  };
  return harness;
}

describe("searchMessages production scan and pagination", () => {
  it("sorts all matches before pagination and reaches both ends beyond 10000", async () => {
    const folder = makeFolder("Archive", { count: 10023 });
    const h = makeHarness({ roots: [folder] });
    const oldest = await h.search({ offset: 0, maxResults: 2, sortOrder: "asc" });
    const newest = await h.search({ offset: 0, maxResults: 2, sortOrder: "desc" });
    const deep = await h.search({ offset: 10020, maxResults: 5, sortOrder: "asc" });
    assert.deepEqual(oldest.messages.map(row => row.id), ["m-0@example.com", "m-1@example.com"]);
    assert.deepEqual(newest.messages.map(row => row.id), ["m-10022@example.com", "m-10021@example.com"]);
    assert.deepEqual(deep.messages.map(row => row.id), ["m-10020@example.com", "m-10021@example.com", "m-10022@example.com"]);
    for (const page of [oldest, newest, deep]) {
      assert.equal(page.totalMatches, 10023);
      assert.equal(page.truncated, undefined);
      assert.ok(page.messages.every(row => !("_dateTs" in row) && !("_messageKey" in row)));
    }
    assert.equal(oldest.hasMore, true);
    assert.equal(deep.hasMore, false);
    assert.ok(h.yields > 0);
  });

  it("returns exact counts past the old collection cap without hydrating content", async () => {
    const folder = makeFolder("Archive", { count: 10011 });
    const h = makeHarness({ roots: [folder] });
    assert.deepEqual(await h.search({ countOnly: true }), { count: 10011 });
    assert.equal(folder.stats.contentReads, 0);
    assert.equal(folder.stats.propertyReads, 0);
  });

  it("hydrates only the requested page while preserving output fields and tag filtering", async () => {
    const folder = makeFolder("Archive", { count: 1200 });
    const h = makeHarness({ roots: [folder] });
    const page = await h.search({ offset: 1000, maxResults: 2, sortOrder: "asc" });
    assert.equal(folder.stats.contentReads, 6);
    assert.equal(folder.stats.propertyReads, 4);
    assert.deepEqual(page.messages[0], {
      id: "m-1000@example.com", threadId: 1000,
      subject: "Decoded subject", author: "Alice Smith", recipients: "Bob Jones", ccList: "copy@example.com",
      date: new Date(1001000).toISOString(), folder: "Archive", folderPath: folder.URI,
      read: false, flagged: true, tags: ["custom", "$label1"], preview: "A preview",
    });
  });

  it("keeps the bare-array shape for complete searches without offset", async () => {
    const h = makeHarness();
    const result = await h.search({ maxResults: 1 });
    assert.ok(Array.isArray(result));
    assert.equal(result[0].id, "m-2@example.com");
  });

  it("yields on nonmatching headers and reacquires databases without enumerators", async () => {
    const folder = makeFolder("Archive", { count: 751, legacy: true });
    const h = makeHarness({ roots: [folder] });
    assert.deepEqual(await h.search({ query: "no such content" }), []);
    assert.equal(h.yields, 3);
    assert.equal(folder.stats.snapshots, 1);
    assert.equal(folder.stats.headerReads, 751);
    assert.equal(folder.stats.databaseReads, 5);
  });

  it("excludes expunged rows returned by listAllKeys without claiming truncation", async () => {
    const folder = makeFolder("Archive", { messages: [{ flags: 0x8 }, {}, { flags: 0x8 }, {}] });
    const h = makeHarness({ roots: [folder] });
    assert.deepEqual(await h.search({ countOnly: true }), { count: 2 });
    assert.deepEqual((await h.search()).map(row => row.id), ["m-3@example.com", "m-1@example.com"]);
  });

  it("deduplicates across all folders before deep pagination", async () => {
    const inbox = makeFolder("Inbox", { count: 10003 });
    const allMail = makeFolder("All Mail", { count: 10003 });
    const root = makeFolder("Root", { isServer: true, children: [inbox, allMail] });
    const h = makeHarness({ roots: [root] });
    const page = await h.search({ offset: 10000, maxResults: 5, sortOrder: "asc" });
    assert.equal(page.totalMatches, 10003);
    assert.equal(page.hasMore, false);
    assert.deepEqual(page.messages.map(row => row.id), ["m-10000@example.com", "m-10001@example.com", "m-10002@example.com"]);
    assert.ok(page.messages.every(row => row.folderPath === inbox.URI));
    assert.ok(page.messages.every(row => row.dupLocations.length === 1 && row.dupLocations[0] === allMail.URI));
    assert.deepEqual(await h.search({ countOnly: true, dedupByMessageId: false }), { count: 20006 });
  });

  it("preserves query operators, AND tokens, decoded headers, preview and cheap filters", async () => {
    const folder = makeFolder("Inbox", { messages: [
      { mime2DecodedSubject: "Projekt Żółw", mime2DecodedAuthor: "Alice Smith", date: Date.parse("2025-01-02T12:00:00Z") * 1000,
        properties: { preview: "Budget review", keywords: "project=20alpha" } },
      { mime2DecodedSubject: "Budget", mime2DecodedAuthor: "Bob Jones", isRead: true, isFlagged: false,
        date: Date.parse("2025-01-03T12:00:00Z") * 1000 },
    ] });
    const h = makeHarness({ roots: [folder] });
    for (const query of ["alice budget", "from:alice smith", "subject:projekt żółw", "to:bob jones", "cc:copy@example.com"]) {
      const rows = await h.search({ query, unreadOnly: true, flaggedOnly: true, tag: "project=20alpha", startDate: "2025-01-02", endDate: "2025-01-02" });
      assert.deepEqual(rows.map(row => row.id), ["m-0@example.com"]);
    }
    for (const query of ["subject:budget", "alice missing", "   ", "from:"]) {
      assert.deepEqual(await h.search({ query, unreadOnly: true }), []);
    }
  });
});

describe("searchMessages production budget, mutations and access", () => {
  for (const mode of ["array", "page", "count"]) {
    it(`preserves the ${mode} response contract when time-limited`, async () => {
      const h = makeHarness({ roots: [makeFolder("Archive", { count: 700 })], onYield(runtime) { runtime.advance(20000); } });
      const result = await h.search({ maxResults: 2, ...(mode === "page" ? { offset: 999 } : {}), ...(mode === "count" ? { countOnly: true } : {}) });
      if (mode === "array") {
        assert.ok(Array.isArray(result));
        assert.equal(result.length, 2);
        assert.deepEqual(Object.keys(result), ["0", "1"]);
        return;
      }
      assert.equal(result.truncated, true);
      assert.match(result.message, /20-second.*partial/);
      if (mode === "count") assert.equal(result.count, 250);
      else {
        assert.equal(result.hasMore, false);
        assert.equal(result.totalMatches, 250);
        assert.equal(result.messages.length, 0);
      }
    });
  }

  it("terminates pagination through time-limited header results independently of truncation", async () => {
    const h = makeHarness({ roots: [makeFolder("Archive", { count: 700 })], onYield(runtime) { runtime.advance(20000); } });
    let offset = 0;
    let pages = 0;
    let result;
    do {
      result = await h.search({ offset, maxResults: 100, sortOrder: "asc" });
      assert.equal(result.truncated, true);
      assert.equal(result.totalMatches, 250);
      assert.equal(result.hasMore, offset + result.messages.length < result.totalMatches);
      offset += result.messages.length;
      assert.ok(++pages <= 3, "Pagination must stop at the collected total");
    } while (result.hasMore);
    assert.equal(offset, 250);
    assert.equal(pages, 3);
    const beyond = await h.search({ offset, maxResults: 100 });
    assert.deepEqual(beyond.messages, []);
    assert.equal(beyond.hasMore, false);
    assert.equal(beyond.truncated, true);
  });

  it("checks the deadline within a chunk, including nonmatches", async () => {
    const folder = makeFolder("Slow", { count: 200 });
    const h = makeHarness({ roots: [folder] });
    folder.hooks.header = () => h.advance(2000);
    const result = await h.search({ query: "never matches", countOnly: true });
    assert.equal(result.count, 0);
    assert.equal(result.truncated, true);
    assert.equal(folder.stats.headerReads, 10);
  });

  it("reports a native key snapshot that consumed the budget before scanning headers", async () => {
    const folder = makeFolder("Slow", { count: 20 });
    const h = makeHarness({ roots: [folder] });
    folder.hooks.snapshot = () => h.advance(25000);
    const result = await h.search({ countOnly: true });
    assert.equal(result.count, 0);
    assert.equal(result.truncated, true);
    assert.equal(folder.stats.headerReads, 0);
  });

  it("continues past removed or failing keys and marks incomplete results", async () => {
    const folder = makeFolder("Archive", { count: 300 });
    folder.hooks.header = key => { if (key === 270) throw new Error("Unreadable header"); };
    const h = makeHarness({ roots: [folder], onYield() { folder.headers.delete(251); } });
    const result = await h.search({ offset: 0, maxResults: 200 });
    assert.equal(result.totalMatches, 298);
    assert.equal(result.truncated, true);
    assert.equal(result.messages[0].id, "m-299@example.com");
    assert.ok(result.messages.every(row => !["m-250@example.com", "m-269@example.com"].includes(row.id)));
  });

  it("does not hydrate a key reused for a different message during a yield", async () => {
    const folder = makeFolder("Archive", { count: 251 });
    const h = makeHarness({ roots: [folder], onYield() { folder.headers.get(1).messageId = "replacement@example.com"; } });
    const result = await h.search({ offset: 0, maxResults: 1, sortOrder: "asc" });
    assert.equal(result.truncated, true);
    assert.deepEqual(result.messages, []);
    assert.equal(result.hasMore, false, "An empty page must not keep the caller's loop running");
  });

  it("continues into sibling folders after a database failure", async () => {
    const broken = makeFolder("Broken", { count: 1 });
    broken.hooks.database = () => { throw new Error("Database unavailable"); };
    const inbox = makeFolder("Inbox", { count: 1 });
    const root = makeFolder("Root", { isServer: true, children: [broken, inbox] });
    const h = makeHarness({ roots: [root] });
    const result = await h.search({ offset: 0 });
    assert.equal(result.truncated, true);
    assert.deepEqual(result.messages.map(row => row.folderPath), [inbox.URI]);
  });

  for (const countOnly of [false, true]) {
    it(`rechecks revoked accounts before returning ${countOnly ? "counts" : "messages and duplicate locations"}`, async () => {
      const revoked = makeFolder("Private", { account: "private", count: 2 });
      const safe = makeFolder("Safe", { account: "safe", count: 1 });
      const h = makeHarness({ roots: [revoked, safe], onYield(runtime) { runtime.allowedAccounts.delete("private"); } });
      const result = await h.search({ folderPath: null, offset: 0, countOnly });
      assert.equal(result.truncated, true);
      if (countOnly) assert.equal(result.count, 1);
      else {
        assert.equal(result.messages.length, 1);
        assert.equal(result.messages[0].folderPath, safe.URI);
        assert.equal(result.messages[0].dupLocations, undefined);
        assert.ok(!JSON.stringify(result).includes(revoked.URI));
      }
    });
  }

  it("stops the current folder when its account is revoked during a chunk yield", async () => {
    const folder = makeFolder("Private", { count: 500 });
    const h = makeHarness({ roots: [folder], onYield(runtime) { runtime.allowedAccounts.clear(); } });
    const result = await h.search({ countOnly: true });
    assert.equal(result.truncated, true);
    assert.equal(result.count, 0);
    assert.equal(folder.stats.headerReads, 250);
  });

  it("denies an inaccessible explicit folder before refresh or database access", async () => {
    const folder = makeFolder("Private", { count: 1 });
    const h = makeHarness({ roots: [folder] });
    h.allowedAccounts.clear();
    assert.match((await h.search()).error, /Account not accessible/);
    assert.equal(folder.stats.databaseReads, 0);
    assert.equal(folder.stats.refreshes, 0);
  });

  it("preserves IMAP refresh for unscoped recursive searches and honors includeSubfolders", async () => {
    const child = makeFolder("Nested", { count: 1 });
    const inbox = makeFolder("Inbox", { count: 1, children: [child] });
    const root = makeFolder("Root", { isServer: true, children: [inbox] });
    const h = makeHarness({ roots: [root] });
    await h.search({ folderPath: null });
    for (const folder of [root, inbox, child]) assert.equal(folder.stats.refreshes, 1);
    await h.search({ folderPath: inbox.URI, includeSubfolders: false });
    assert.equal(inbox.stats.refreshes, 2);
    assert.equal(child.stats.refreshes, 1);
  });

  it("keeps final untrusted-text sanitization while preserving identifiers", async () => {
    const folder = makeFolder("In\u200bbox", { messages: [{ messageId: "id\u200b@example.com", mime2DecodedSubject: "Hello\u200bthere", properties: { preview: "A\u202eB" } }] });
    const h = makeHarness({ roots: [folder] });
    const [row] = await h.search();
    assert.equal(row.subject, "Hellothere");
    assert.equal(row.preview, "AB");
    assert.equal(row.folder, "Inbox");
    assert.equal(row.folderPath, folder.URI);
    assert.equal(row.id, "id\u200b@example.com");
  });
});

describe("searchMessages production Gloda collection", () => {
  for (const offset of [undefined, null]) {
    it(`keeps plain arrays without completeness fields when offset is ${offset}`, async () => {
      const folder = makeFolder("Archive", { count: 2 });
      const glodaItems = Array.from(folder.headers.values(), folderMessage => ({ folderMessage }));
      const h = makeHarness({ roots: [folder], glodaItems, glodaLimit: 2 });
      const result = await h.runtime.callTool("searchMessages", { query: "body query", searchBody: true, folderPath: folder.URI, offset });
      assert.ok(Array.isArray(result));
      assert.deepEqual(Array.from(result, row => row.id), ["m-1@example.com", "m-0@example.com"]);
      assert.deepEqual(Object.keys(result), ["0", "1"]);
      assert.equal(result.truncated, undefined);
      assert.equal(result.message, undefined);
      assert.equal(result.hasMore, undefined);
    });
  }

  it("keeps an empty array when an unpaginated Gloda search times out", async () => {
    const h = makeHarness({ waitForGloda: true });
    const pending = h.search({ query: "body query", searchBody: true });
    h.advance(20000);
    h.timers[0].fire();
    assert.deepEqual(await pending, []);
  });

  it("terminates pagination through Gloda candidates while keeping truncation separate", async () => {
    const folder = makeFolder("Archive", { count: 5 });
    const glodaItems = Array.from(folder.headers.values(), folderMessage => ({ folderMessage }));
    const h = makeHarness({ roots: [folder], glodaItems, glodaLimit: 5 });
    const ids = [];
    let offset = 0;
    let pages = 0;
    let result;
    do {
      result = await h.search({ query: "body query", searchBody: true, offset, maxResults: 2, sortOrder: "asc" });
      assert.equal(result.truncated, true);
      assert.equal(result.totalMatches, 5);
      assert.equal(result.hasMore, offset + result.messages.length < result.totalMatches);
      ids.push(...result.messages.map(row => row.id));
      offset += result.messages.length;
      assert.ok(++pages <= 3, "Pagination must stop at the collected total");
    } while (result.hasMore);
    assert.deepEqual(ids, Array.from({ length: 5 }, (_, index) => `m-${index}@example.com`));
    for (const beyondOffset of [5, 10]) {
      const beyond = await h.search({ query: "body query", searchBody: true, offset: beyondOffset });
      assert.deepEqual(beyond.messages, []);
      assert.equal(beyond.hasMore, false);
      assert.equal(beyond.truncated, true);
    }
  });

  it("sorts and paginates beyond 10000 Gloda matches without applying a first-N cap", async () => {
    const folder = makeFolder("Archive", { count: 10003 });
    const glodaItems = Array.from(folder.headers.values(), folderMessage => ({ folderMessage }));
    const h = makeHarness({ roots: [folder], glodaItems });
    const page = await h.search({ query: "body query", searchBody: true, offset: 10000, sortOrder: "asc", maxResults: 5 });
    assert.equal(page.totalMatches, 10003);
    assert.equal(page.truncated, true);
    assert.equal(page.hasMore, false);
    assert.deepEqual(page.messages.map(row => row.id), ["m-10000@example.com", "m-10001@example.com", "m-10002@example.com"]);
    assert.ok(h.yields > 0);
    assert.equal(folder.stats.contentReads, 9);
    assert.equal(h.timers[0].canceled, true);
  });

  for (const countOnly of [false, true]) {
    it(`reports Gloda's retrieval ceiling before filtering ${countOnly ? "counts" : "rows"}`, async () => {
      const folder = makeFolder("Archive", { messages: [{ isRead: true }, {}, { isRead: true }] });
      const glodaItems = Array.from(folder.headers.values(), folderMessage => ({ folderMessage }));
      const h = makeHarness({ roots: [folder], glodaItems, glodaLimit: 3 });
      const result = await h.search({ query: "body query", searchBody: true, unreadOnly: true, offset: 0, countOnly });
      assert.equal(result.truncated, true);
      assert.match(result.message, /relevance-limited candidate set/);
      if (countOnly) assert.equal(result.count, 1);
      else { assert.equal(result.messages.length, 1); assert.equal(result.hasMore, false); }
    });
  }

  it("does not claim completeness for a collection below Gloda's retrieval limit", async () => {
    const folder = makeFolder("Archive", { count: 2 });
    const glodaItems = Array.from(folder.headers.values(), folderMessage => ({ folderMessage }));
    const h = makeHarness({ roots: [folder], glodaItems, glodaLimit: 3 });
    const result = await h.search({ query: "body query", searchBody: true, offset: 0 });
    assert.equal(result.messages.length, 2);
    assert.equal(result.truncated, true);
    assert.equal(result.hasMore, false);
    assert.match(result.message, /completeness cannot be verified/);
  });

  it("returns partial results when Gloda never completes and detaches late callbacks", async () => {
    const h = makeHarness({ waitForGloda: true });
    const pending = h.search({ query: "body query", searchBody: true, countOnly: true });
    assert.equal(h.timers[0].interval, 20000);
    h.advance(20000);
    h.timers[0].fire();
    const result = await pending;
    assert.equal(result.count, 0);
    assert.equal(result.truncated, true);
    assert.equal(h.searchers[0].listener, null);
    assert.equal(h.timers[0].canceled, true);
  });

  it("handles timeout during cooperative Gloda processing without a second completion", async () => {
    const folder = makeFolder("Archive", { count: 600 });
    const glodaItems = Array.from(folder.headers.values(), folderMessage => ({ folderMessage }));
    const h = makeHarness({ roots: [folder], glodaItems, onYield(runtime) { runtime.advance(20000); runtime.timers[0].fire(); } });
    const result = await h.search({ query: "body query", searchBody: true, countOnly: true });
    assert.equal(result.truncated, true);
    assert.equal(result.count, 250);
    assert.equal(h.timers[0].canceled, true);
  });

  it("filters access again after Gloda processing yields", async () => {
    const folder = makeFolder("Private", { count: 300 });
    const glodaItems = Array.from(folder.headers.values(), folderMessage => ({ folderMessage }));
    const h = makeHarness({ roots: [folder], glodaItems, onYield(runtime) { runtime.allowedAccounts.clear(); } });
    const result = await h.search({ query: "body query", searchBody: true, offset: 0 });
    assert.equal(result.truncated, true);
    assert.deepEqual(result.messages, []);
  });
});
