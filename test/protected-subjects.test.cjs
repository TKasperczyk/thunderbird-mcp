"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
// Same marker assertions and VM loading pattern as privacy-messages.test.cjs.
function snippet(startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${startMarker}`);
  return source.slice(start, end);
}

const DECRYPTION_OKAY = 0x00020000;
const GOOD_SIGNATURE = 0x00000001;
const HEADER_LIMIT = 16 * 1024;
const CACHED_SUBJECT = "Cached decrypted subject";
const OUTER_SUBJECT = "Encrypted Message";
const ENCRYPTED_RAW = `From: sender@example.test\r\nSubject: =?utf-8?q?Encrypted_Message?=\r\nContent-Type: multipart/encrypted; protocol="application/pgp-encrypted"; boundary="b"\r\n\r\n--b--\r\n`;
const CLEAR_RAW = "From: sender@example.test\r\nSubject: Plain\r\nContent-Type: text/plain\r\n\r\nbody\r\n";

function makeStream(raw, stats) {
  let offset = 0;
  return {
    // Small chunks prove the header read stops at the header terminator.
    available() { return Math.min(raw.length - offset, 1000); },
    read(count) {
      const chunk = raw.slice(offset, offset + count);
      offset += chunk.length;
      stats.bytesRead += chunk.length;
      return chunk;
    },
    close() { stats.closed++; },
  };
}

function makeHeader(folder, { id, enigmail = 0, raw = CLEAR_RAW, date = 1, subject = CACHED_SUBJECT } = {}) {
  return {
    messageId: id, threadId: 1, folder, raw, date, flags: 0, isRead: false, isFlagged: false,
    subject, mime2DecodedSubject: subject, author: "sender@example.test", recipients: "reader@example.test",
    ccList: "",
    getStringProperty: (name) => (name === "preview" ? "Cached decrypted preview" : ""),
    getUint32Property: (name) => (name === "enigmail" ? enigmail : 0),
  };
}

function loadTools({ allowed = false, streamError = null } = {}) {
  const stats = { streams: 0, bytesRead: 0, closed: 0, displayed: 0 };
  const folder = {
    URI: "imap://account/INBOX", prettyName: "INBOX", hasSubFolders: false,
    getMsgInputStream(msgHdr) {
      stats.streams++;
      if (streamError) throw streamError;
      return makeStream(msgHdr.raw, stats);
    },
  };
  const headers = [];
  folder.msgDatabase = { enumerateMessages: () => headers };
  const sandbox = {
    console,
    Services: { prefs: { getBoolPref: (name, fallback) => (name.endsWith(".allowEncryptedMessages") ? allowed : fallback) } },
    NetUtil: { readInputStreamToString: (stream, count) => stream.read(count) },
    MailServices: {
      mimeConverter: { decodeMimeHeader: (value) => value.replace(/=\?utf-8\?q\?([^?]*)\?=/gi, (_, text) => text.replace(/_/g, " ")) },
    },
    ChromeUtils: { importESModule: () => ({ MailUtils: { displayMessageInFolderTab() { stats.displayed++; } } }) },
    findMessage: (messageId) => {
      const msgHdr = headers.find((header) => header.messageId === messageId);
      return msgHdr ? { msgHdr, folder } : { error: "not found" };
    },
    openFolder: () => ({ folder, db: folder.msgDatabase }),
    getAccessibleAccounts: () => [],
    getUserTags: () => [],
    DEFAULT_MAX_RESULTS: 50,
    MAX_SEARCH_RESULTS_CAP: 200,
    SEARCH_COLLECTION_CAP: 10000,
  };
  vm.createContext(sandbox);
  vm.runInContext([
    source.match(/^const PREF_\w+ = .+;$/gm).join("\n"),
    snippet("// BEGIN PRIVACY PREFERENCE HELPERS", "// END PRIVACY PREFERENCE HELPERS"),
    snippet("// BEGIN RAW MIME PARSING HELPERS", "// END RAW MIME PARSING HELPERS"),
    snippet("// BEGIN ENCRYPTED MESSAGE GUARD", "// END ENCRYPTED MESSAGE GUARD"),
    snippet("// BEGIN SEARCH RESULT HELPERS", "// END SEARCH RESULT HELPERS"),
    snippet("function displayMessage(", "function isTrashOrDescendant("),
  ].join("\n"), sandbox);
  return {
    api: sandbox, stats, folder,
    add(options) {
      const header = makeHeader(folder, options);
      headers.push(header);
      return header;
    },
  };
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

describe("getRecentMessages protects subjects of encrypted mail", () => {
  it("returns the transmitted subject without a preview for recorded encrypted messages", () => {
    const h = loadTools();
    h.add({ id: "encrypted", enigmail: DECRYPTION_OKAY, raw: ENCRYPTED_RAW, date: 2 });
    h.add({ id: "clear", date: 1 });
    const rows = plain(h.api.getRecentMessages(h.folder.URI, 36500));
    assert.equal(rows.length, 2);
    assert.equal(rows[0].id, "encrypted");
    assert.equal(rows[0].subject, OUTER_SUBJECT);
    assert.equal(rows[0].preview, undefined);
    assert.equal(rows[0].encryptedContentWithheld, true);
    assert.equal(rows[1].subject, CACHED_SUBJECT);
    assert.equal(rows[1].preview, "Cached decrypted preview");
    assert.equal(rows[1].encryptedContentWithheld, undefined);
    assert.doesNotMatch(JSON.stringify(rows[0]), /Cached decrypted/);
    assert.equal(h.stats.streams, 1, "only the protected row is read");
    assert.equal(h.stats.closed, 1);
  });

  it("reads stored messages only for protected rows on the returned page", () => {
    const h = loadTools();
    for (let index = 0; index < 30; index++) {
      h.add({ id: `m-${index}`, enigmail: DECRYPTION_OKAY, raw: ENCRYPTED_RAW, date: index + 1 });
    }
    const page = plain(h.api.getRecentMessages(h.folder.URI, 36500, 3, 0));
    assert.equal(page.messages.length, 3);
    assert.equal(page.totalMatches, 30);
    for (const row of page.messages) {
      assert.equal(row.subject, OUTER_SUBJECT);
      assert.equal(row.encryptedContentWithheld, true);
      assert.equal(Object.hasOwn(row, "_protectedHdr"), false);
    }
    assert.equal(h.stats.streams, 3);
  });

  it("returns cached subjects and previews after explicit opt-in", () => {
    const h = loadTools({ allowed: true });
    h.add({ id: "encrypted", enigmail: DECRYPTION_OKAY, raw: ENCRYPTED_RAW });
    const [row] = plain(h.api.getRecentMessages(h.folder.URI, 36500));
    assert.equal(row.subject, CACHED_SUBJECT);
    assert.equal(row.preview, "Cached decrypted preview");
    assert.equal(row.encryptedContentWithheld, undefined);
    assert.equal(h.stats.streams, 0);
  });

  it("leaves signed-only messages unchanged", () => {
    const h = loadTools();
    h.add({ id: "signed", enigmail: GOOD_SIGNATURE });
    const [row] = plain(h.api.getRecentMessages(h.folder.URI, 36500));
    assert.equal(row.subject, CACHED_SUBJECT);
    assert.equal(h.stats.streams, 0);
  });

  it("uses a placeholder when the transmitted subject cannot be read within the header bound", () => {
    const oversized = `X-Padding: ${"a".repeat(HEADER_LIMIT)}\r\nSubject: Outer\r\n\r\nbody`;
    const duplicate = "Subject: One\r\nSubject: Two\r\n\r\nbody";
    for (const [label, options] of [
      ["oversized headers", { raw: oversized }],
      ["duplicate subjects", { raw: duplicate }],
      ["unreadable stream", { streamError: new Error("not available offline") }],
    ]) {
      const h = loadTools({ streamError: options.streamError });
      h.add({ id: "encrypted", enigmail: DECRYPTION_OKAY, raw: options.raw || ENCRYPTED_RAW });
      const [row] = plain(h.api.getRecentMessages(h.folder.URI, 36500));
      assert.equal(row.subject, "[Encrypted message]", label);
      assert.equal(row.encryptedContentWithheld, true, label);
      assert.ok(h.stats.bytesRead <= HEADER_LIMIT, `${label}: read ${h.stats.bytesRead} bytes`);
    }
  });
});

describe("displayMessage while encrypted message access is off", () => {
  it("refuses messages that are encrypted or recorded as encrypted, before opening them", () => {
    for (const options of [
      { raw: ENCRYPTED_RAW },
      { raw: "Subject: S/MIME\r\nContent-Type: application/pkcs7-mime; smime-type=enveloped-data\r\n\r\nx" },
      { raw: "Subject: Ambiguous\r\nContent-Type: text/plain\r\nContent-Type: multipart/encrypted\r\n\r\nx" },
      { raw: CLEAR_RAW, enigmail: DECRYPTION_OKAY },
    ]) {
      const h = loadTools();
      h.add({ id: "message", ...options });
      const result = h.api.displayMessage("message", h.folder.URI);
      assert.match(result.error, /Encrypted messages cannot be displayed while encrypted message access is off/);
      assert.equal(result.subject, undefined);
      assert.equal(h.stats.displayed, 0);
      assert.ok(h.stats.bytesRead <= HEADER_LIMIT);
    }
  });

  it("displays clear messages and, after opt-in, encrypted ones", () => {
    for (const [allowed, raw] of [[false, CLEAR_RAW], [true, ENCRYPTED_RAW]]) {
      const h = loadTools({ allowed });
      h.add({ id: "message", raw });
      const result = h.api.displayMessage("message", h.folder.URI);
      assert.equal(result.success, true, result.error);
      assert.equal(h.stats.displayed, 1);
    }
  });
});
