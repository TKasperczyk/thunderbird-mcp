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
  assert.ok(start >= 0, `Missing api.js marker: ${startMarker}`);
  assert.ok(end > start, `Missing api.js marker: ${endMarker}`);
  return source.slice(start, end);
}

// Runs the production parser, header reader, access checks and handler.
function load({ allowEncrypted = false, accountAllowed = true } = {}) {
  const reads = { streams: 0 };
  const sandbox = {
    console,
    Services: { prefs: { getBoolPref: (_name, fallback) => (allowEncrypted ? true : fallback) } },
    NetUtil: { readInputStreamToString: (stream, count) => stream.read(count) },
    isAccountAllowed: () => accountAllowed,
    MailServices: null,
  };
  sandbox.PREF_ALLOW_ENCRYPTED_MESSAGES = "extensions.thunderbird-mcp.allowEncryptedMessages";
  const runtime = vm.createContext(sandbox);
  vm.runInContext([
    snippet("function isPrivacyOptInEnabled(", "// BEGIN CONTACT WRITE GATE"),
    snippet("function isFolderAccessible(", "function toColumnarTable("),
    snippet("function openFolder(", "function findMessage("),
    snippet("function findMessage(", "// BEGIN MESSAGE SEARCH"),
    snippet("function findRawMimeHeaderBodySplit(", "function splitRawMimeHeaderParameters("),
    snippet("const OPENPGP_ENCRYPTION_STATUS_MASK", "// The Subject header as transmitted"),
    snippet("// BEGIN UNSUBSCRIBE HEADER PARSER", "// END UNSUBSCRIBE INFO TOOL"),
    "this.parseListUnsubscribe = parseListUnsubscribe; this.getUnsubscribeInfo = getUnsubscribeInfo;",
  ].join("\n"), runtime);

  function addMessage(rawHeaders, { enigmail = 0, raw } = {}) {
    const text = raw ?? `${rawHeaders}\r\n\r\nbody`;
    const msgHdr = {
      messageId: "msg-1",
      getUint32Property: () => enigmail,
      folder: {
        getMsgInputStream() {
          reads.streams++;
          let offset = 0;
          return {
            available: () => text.length - offset,
            read(count) { const chunk = text.slice(offset, offset + count); offset += chunk.length; return chunk; },
            close() {},
          };
        },
      },
    };
    const server = { type: "imap" };
    const folder = {
      URI: "imap://one/INBOX", server, updateFolder() {},
      msgDatabase: { getMsgHdrForMessageID: id => (id === "msg-1" ? msgHdr : null), enumerateMessages: () => [msgHdr] },
    };
    runtime.MailServices = {
      accounts: { findAccountForServer: () => ({ key: "one" }) },
      folderLookup: { getFolderForURL: uri => (uri === folder.URI ? folder : null) },
    };
    return folder;
  }
  return { runtime, addMessage, reads, parse: (list, post) => JSON.parse(JSON.stringify(runtime.parseListUnsubscribe(list, post))) };
}

describe("parseListUnsubscribe", () => {
  it("returns https and mailto targets with one-click support (RFC 2369 / 8058)", () => {
    const { parse } = load();
    const result = parse(
      ["<mailto:unsub@example.com?subject=unsubscribe%20me>, <https://example.com/u?id=1>"],
      ["List-Unsubscribe=One-Click"],
    );
    assert.deepEqual(result.mailto, {
      uri: "mailto:unsub@example.com?subject=unsubscribe%20me", address: "unsub@example.com",
      subject: "unsubscribe me", body: null, ignoredFields: [],
    });
    assert.deepEqual(result.https, [{ uri: "https://example.com/u?id=1", host: "example.com" }]);
    assert.equal(result.oneClick, true);
    assert.deepEqual(result.rejected, []);
  });

  it("does not advertise one-click without a post header, with a different one, or without an https target", () => {
    const { parse } = load();
    assert.equal(parse(["<https://a.example.com/u>"], []).oneClick, false);
    assert.equal(parse(["<https://a.example.com/u>"], ["List-Unsubscribe=Other"]).oneClick, false);
    assert.equal(parse(["<https://a.example.com/u>"], ["List-Unsubscribe=One-Click", "List-Unsubscribe=One-Click"]).oneClick, false);
    const noHttps = parse(["<mailto:u@example.com>"], ["List-Unsubscribe=One-Click"]);
    assert.equal(noHttps.oneClick, false);
    assert.match(noHttps.warnings.join(" "), /https/);
  });

  it("treats http and unknown schemes as rejected, never as targets", () => {
    const { parse } = load();
    const result = parse(["<http://a.example.com/u>, <javascript:alert(1)>, <file:///etc/passwd>, <ftp://a.example.com/>"], []);
    assert.deepEqual(result.https, []);
    assert.equal(result.mailto, null);
    assert.deepEqual(result.rejected.map(item => item.reason), ["insecure-scheme", "unsupported-scheme", "unsupported-scheme", "unsupported-scheme"]);
  });

  it("rejects https URLs with userinfo, spaces, control characters or backslashes", () => {
    const { parse } = load();
    for (const uri of [
      "https://trusted.example@evil.example.com/u",
      "https://evil.example.com\\@trusted.example.com/u",
      "https://a.example.com/ u",
      "https://a.example.com/\u0000x",
      "https://a.example.com/‮x",
      "https://",
    ]) {
      const result = parse([`<${uri}>`], ["List-Unsubscribe=One-Click"]);
      assert.deepEqual(result.https, [], JSON.stringify(uri));
      assert.equal(result.oneClick, false, JSON.stringify(uri));
      assert.equal(result.rejected.length, 1, JSON.stringify(uri));
    }
  });

  it("applies a strict host policy, so every advertised https target is safe to request", () => {
    const { parse } = load();
    const hostile = [
      "https://127.0.0.1/u", "https://169.254.169.254/latest", "https://localhost:8443/u", "https://[::1]/u",
      "https://2130706433/u", "https://0x7f.0.0.1/u", "https://intranet/u", "https://printer.local/u",
      "https://news.example.com:8443/u", "https://news.example.com./u", "https://app.internal/u",
    ];
    for (const uri of hostile) {
      const result = parse([`<${uri}>`], ["List-Unsubscribe=One-Click"]);
      assert.deepEqual(result.https, [], uri);
      assert.equal(result.oneClick, false, uri);
      assert.equal(result.rejected.length, 1, uri);
    }
    const allowed = parse(["<https://News.Example.com:443/u>, <https://xn--bcher-kva.example.org/u>"], ["List-Unsubscribe=One-Click"]);
    assert.deepEqual(allowed.https.map(item => item.host), ["news.example.com", "xn--bcher-kva.example.org"]);
    assert.equal(allowed.oneClick, true);
  });

  it("rejects control, bidi and oversized values after percent-decoding in mailto fields", () => {
    const { parse } = load();
    for (const uri of [
      "mailto:a@example.com?subject=hi%00there", "mailto:a@example.com?subject=%E2%80%AEevil", "mailto:a@example.com?body=x%E2%80%A8y",
      "mailto:a@example.com?body=x%07y", "mailto:a%00@example.com", `mailto:a@example.com?subject=${"s".repeat(201)}`,
      `mailto:a@example.com?body=${"b".repeat(1001)}`,
    ]) {
      const result = parse([`<${uri}>`], []);
      assert.equal(result.mailto, null, uri.slice(0, 60));
      assert.equal(result.rejected.length, 1, uri.slice(0, 60));
    }
    const ok = parse(["<mailto:a@example.com?subject=unsubscribe&body=line1%0D%0Aline2%09tab>"], []);
    assert.equal(ok.mailto.body, "line1\r\nline2\ttab");
  });

  it("bounds the ignored mailto field names it reports", () => {
    const { parse } = load();
    const fields = Array.from({ length: 12 }, (_, i) => `x${i}=1`).join("&") + "&" + "a".repeat(30) + "=1&bad%20name=1";
    const result = parse([`<mailto:a@example.com?${fields}>`], []);
    assert.equal(result.mailto.ignoredFields.length, 5);
    assert.ok(result.mailto.ignoredFields.every(name => /^[a-z0-9-]{1,20}$/.test(name)));
  });

  it("rejects mailto targets with several recipients or header injection", () => {
    const { parse } = load();
    for (const uri of [
      "mailto:a@example.com,b@example.com",
      "mailto:a@example.com%2Cb@example.com",
      "mailto:a@example.com%0d%0aBcc:%20evil@example.com",
      "mailto:",
      "mailto:not-an-address",
      "mailto:a@example.com?subject=x%0d%0aBcc:%20evil@example.com",
      "mailto:a@example.com?subject=%E0%A4%A",
    ]) {
      const result = parse([`<${uri}>`], []);
      assert.equal(result.mailto, null, uri);
      assert.equal(result.rejected.length, 1, uri);
    }
  });

  it("keeps only the first mailto and reports cc/bcc style fields as ignored", () => {
    const { parse } = load();
    const result = parse(["<mailto:a@example.com?cc=x@example.com&bcc=y@example.com>, <mailto:b@example.com>"], []);
    assert.equal(result.mailto.address, "a@example.com");
    assert.deepEqual(result.mailto.ignoredFields, ["cc", "bcc"]);
    assert.deepEqual(result.rejected.map(item => item.reason), ["additional-mailto"]);
  });

  it("is bounded: oversized headers, entries and URIs", () => {
    const { parse } = load();
    const many = Array.from({ length: 30 }, (_, i) => `<https://h${i}.example.com/u>`).join(",");
    const capped = parse([many], []);
    assert.equal(capped.https.length, 10);
    assert.match(capped.warnings.join(" "), /More than 10/);

    const long = parse([`<https://a.example.com/${"x".repeat(3000)}>`], []);
    assert.deepEqual(long.rejected.map(item => item.reason), ["too-long"]);
    assert.ok(long.rejected[0].uri.length <= 203, "echoed text is truncated");

    const huge = parse([`<https://a.example.com/u>${" ".repeat(9000)}<https://b.example.com/u>`], []);
    assert.deepEqual(huge.https.map(item => item.host), ["a.example.com"]);
    assert.match(huge.warnings.join(" "), /longer than 8192/);
  });

  it("ignores garbage outside angle brackets, comments, and unterminated entries", () => {
    const { parse } = load();
    const garbage = parse(["https://a.example.com/u, (comment) <https://b.example.com/u> junk <https://c.example.com/u"], []);
    assert.deepEqual(garbage.https.map(item => item.host), ["b.example.com"]);
    assert.match(garbage.warnings.join(" "), /outside <URI>/);
  });

  it("uses the first of several List-Unsubscribe headers and says so", () => {
    const { parse } = load();
    const result = parse(["<https://a.example.com/u>", "<https://evil.example.com/u>"], []);
    assert.deepEqual(result.https.map(item => item.host), ["a.example.com"]);
    assert.match(result.warnings.join(" "), /Multiple/);
  });

  it("reports nothing for a message without the header", () => {
    const { parse } = load();
    const result = parse(undefined, undefined);
    assert.deepEqual([result.mailto, result.https, result.rejected, result.oneClick], [null, [], [], false]);
  });
});

describe("production getUnsubscribeInfo", () => {
  const HEADERS = [
    "From: News <news@example.com>",
    "List-Unsubscribe: <mailto:u@example.com>,",
    "  <https://example.com/u?id=1>",
    "List-Unsubscribe-Post: List-Unsubscribe=One-Click",
    "Subject: Weekly",
  ].join("\r\n");

  it("reads the stored header block, unfolds it, and reports the targets without side effects", () => {
    const h = load();
    const folder = h.addMessage(HEADERS);
    const result = JSON.parse(JSON.stringify(h.runtime.getUnsubscribeInfo("msg-1", folder.URI)));
    assert.equal(result.id, "msg-1");
    assert.equal(result.hasUnsubscribe, true);
    assert.equal(result.mailto.address, "u@example.com");
    assert.equal(result.https[0].host, "example.com");
    assert.equal(result.oneClick, true);
  });

  it("reports hasUnsubscribe:false when only rejected targets exist", () => {
    const h = load();
    const folder = h.addMessage("List-Unsubscribe: <http://example.com/u>");
    const result = JSON.parse(JSON.stringify(h.runtime.getUnsubscribeInfo("msg-1", folder.URI)));
    assert.equal(result.hasUnsubscribe, false);
    assert.equal(result.rejected.length, 1);
  });

  it("does not look into a header-less body: a List-Unsubscribe line in the body is ignored", () => {
    const h = load();
    const folder = h.addMessage("", { raw: "Subject: x\r\n\r\nList-Unsubscribe: <https://evil.example.com/u>\r\n" });
    const result = JSON.parse(JSON.stringify(h.runtime.getUnsubscribeInfo("msg-1", folder.URI)));
    assert.equal(result.hasUnsubscribe, false);
  });

  it("withholds recorded-encrypted messages without reading the message", () => {
    const h = load();
    const folder = h.addMessage(HEADERS, { enigmail: 0x00000100 });
    const result = JSON.parse(JSON.stringify(h.runtime.getUnsubscribeInfo("msg-1", folder.URI)));
    assert.equal(result.encryptedContentWithheld, true);
    assert.equal(result.https, undefined);
    assert.equal(h.reads.streams, 0);
  });

  it("returns the targets of an encrypted message when the user allowed encrypted access", () => {
    const h = load({ allowEncrypted: true });
    const folder = h.addMessage(HEADERS, { enigmail: 0x00000100 });
    const result = JSON.parse(JSON.stringify(h.runtime.getUnsubscribeInfo("msg-1", folder.URI)));
    assert.equal(result.hasUnsubscribe, true);
  });

  it("honors account restrictions", () => {
    const h = load({ accountAllowed: false });
    const folder = h.addMessage(HEADERS);
    const result = JSON.parse(JSON.stringify(h.runtime.getUnsubscribeInfo("msg-1", folder.URI)));
    assert.match(result.error, /not accessible/);
    assert.equal(h.reads.streams, 0);
  });

  it("reports an unreadable header block instead of guessing", () => {
    const h = load();
    const folder = h.addMessage("", { raw: "" });
    const result = JSON.parse(JSON.stringify(h.runtime.getUnsubscribeInfo("msg-1", folder.URI)));
    assert.match(result.error, /header block could not be read/);
  });

  it("validates its arguments and unknown messages", () => {
    const h = load();
    const folder = h.addMessage(HEADERS);
    assert.match(h.runtime.getUnsubscribeInfo("", folder.URI).error, /messageId/);
    assert.match(h.runtime.getUnsubscribeInfo("msg-1", "").error, /folderPath/);
    assert.match(h.runtime.getUnsubscribeInfo("nope", folder.URI).error, /Message not found/);
  });
});
