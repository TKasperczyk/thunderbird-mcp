"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { makeDirectSendPrefs, prefTypes } = require("./helpers/direct-send-prefs.cjs");
const { makeDirectSendAttachments, assertNoAttachmentEffects } = require("./helpers/direct-send-attachments.cjs");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
const RAW_LIMIT = 50 * 1024 * 1024;
const HTML_LIMIT = 2 * 1024 * 1024;
const TIME_LIMIT = 15000;
const RAW_HEADER = "Content-Type: text/plain\r\n\r\n";
const SIZE_WARNING_PREF = "mailnews.message_warning_size";
const DEFAULT_SIZE_WARNING_BYTES = 20 * 1024 * 1024;
const SMALL_SIZE_WARNING_BYTES = 500000;
const tools = ["replyToMessage", "forwardMessage"];

// Searched for shared marker loaders -> found local snippet functions in
// direct-send-encryption and privacy-messages. Reuse their extraction pattern.
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

// Searched for raw-stream, timer, clock, and preflight runtimes -> found local
// read-message-stream and direct-send-outcome mocks, but no shared combination.
// Retain their stream/native timer patterns and reuse real attachment conversion.
// Searched for streamMessage/nsIStreamListener fixtures -> found server-lifecycle
// callback mocks, but no shared IMAP raw retrieval fixture. Reuse that callback
// pattern here alongside the existing raw-stream and controlled-clock fixture.
function loadPreflight(options = {}) {
  const attachments = makeDirectSendAttachments();
  const calls = {
    rawOpenAttempts: 0, rawOpens: 0, rawCloses: 0, rawReads: [], bytesRead: 0, imports: 0,
    rawServices: [], rawFetches: [], fetchReads: [], fetchBytesRead: 0, fetchCancellations: [],
    mime: [], timers: [], sends: 0, reviews: 0, settlements: 0,
    created: { params: 0, fields: 0, compose: 0, sender: 0 },
  };
  const clock = { now: 1000 };
  const raw = options.raw ?? RAW_HEADER + "Original text";
  const mime = options.mime ?? { contentType: "text/plain", body: "Original text" };
  const identity = { key: "sender", email: "sender@example.test", composeHtml: false };
  const folder = {
    server: { type: options.serverType ?? (options.uncached ? "imap" : undefined) },
    getUriForMsg: () => "message-uri",
    getMsgInputStream() {
      calls.rawOpenAttempts++;
      clock.now += options.openDuration || 0;
      if (options.uncached) throw new Error("Original is not available offline");
      calls.rawOpens++;
      let cursor = 0;
      return {
        available() {
          if (options.availableError) throw new Error("Raw stream unavailable");
          return raw.length - cursor;
        },
        read(count) {
          calls.rawReads.push(count);
          clock.now += options.readDuration || 0;
          if (options.readError) throw new Error("Raw stream failed");
          if (options.noProgress) return "";
          const value = raw.slice(cursor, cursor + count);
          cursor += value.length;
          calls.bytesRead += value.length;
          return value;
        },
        close() { calls.rawCloses++; },
      };
    },
  };
  const hdr = {
    folder, messageId: "original@example.test", subject: "Original",
    author: "author@example.test", recipients: "recipient@example.test", date: 0,
    flags: options.offline ?? !options.uncached ? 0x80 : 0,
  };
  const sandbox = {
    TextDecoder, Uint8Array, atob, btoa,
    DOMParser: options.DOMParser,
    Date: class extends Date { static now() { return clock.now; } },
    Cr: { NS_BASE_STREAM_CLOSED: "NS_BASE_STREAM_CLOSED", NS_BINDING_ABORTED: "NS_BINDING_ABORTED" },
    // S4 completeness limits are independent of the user's size warning gate.
    Services: { prefs: makeDirectSendPrefs({ [SIZE_WARNING_PREF]: 0, ...options.prefs }) },
    Ci: {
      nsIPrefBranch: prefTypes, nsITimer: { TYPE_ONE_SHOT: 0 },
      nsMsgMessageFlags: { Offline: 0x80 },
      nsIMsgCompType: { New: 0, Reply: 1, ReplyAll: 2, ForwardInline: 3 },
      nsIMsgCompFormat: { Default: 0, HTML: 1, PlainText: 2 },
      nsIMsgCompDeliverMode: { Now: 0 },
    },
    NetUtil: { readInputStreamToString: (stream, count) => stream.read(count) },
    Components: { isSuccessCode: status => status === 0 },
    ChromeUtils: {
      generateQI: () => () => {},
      importESModule() {
        calls.imports++;
        clock.now += options.importDuration || 0;
        if (options.importError) throw new Error("MIME module unavailable");
        return {
          MsgHdrToMimeMessage(msg, _listener, callback, allowDownload, mimeOptions) {
            calls.mime.push({ allowDownload, ...mimeOptions });
            if (options.mimeError) throw new Error("MIME parser failed");
            calls.deliverMime = value => callback(msg, value);
            if (!options.deferMime) callback(msg, mime);
          },
        };
      },
    },
    MailServices: {
      accounts: { accounts: [{ key: "account", identities: [identity] }] },
      messageServiceFromURI(uri) {
        calls.rawServices.push(uri);
        clock.now += options.serviceDuration || 0;
        if (options.serviceError) throw new Error("Message service unavailable");
        return {
          streamMessage(...args) {
            calls.rawFetches.push(args);
            if (options.fetchError) throw new Error("Raw retrieval unavailable");
            const listener = args[1];
            let offset = 0;
            const request = {
              QueryInterface() { return this; },
              cancel(status) {
                calls.fetchCancellations.push(status);
                if (options.cancelError) throw new Error("Request cancellation failed");
              },
            };
            calls.fetchRequest = request;
            calls.fetchListener = listener;
            calls.startFetch = () => listener.onStartRequest(request);
            calls.deliverRaw = (chunk, count = chunk.length) => {
              let cursor = 0;
              const stream = {
                available() { return chunk.length - cursor; },
                read(readCount) {
                  calls.fetchReads.push(readCount);
                  clock.now += options.fetchReadDuration || 0;
                  if (options.fetchReadError) throw new Error("Raw retrieval read failed");
                  if (options.fetchNoProgress) return "";
                  const size = Math.min(readCount, options.fetchReadLimit ?? readCount);
                  const value = chunk.slice(cursor, cursor + size);
                  cursor += value.length;
                  calls.fetchBytesRead += value.length;
                  return value;
                },
              };
              listener.onDataAvailable(request, stream, offset, count);
              offset += count;
            };
            calls.stopFetch = (status = 0) => listener.onStopRequest(request, status);
            const finish = () => {
              calls.startFetch();
              for (const chunk of options.fetchChunks || [raw]) calls.deliverRaw(chunk);
              calls.stopFetch(options.fetchStatus ?? 0);
            };
            if (!options.deferFetch) {
              if (options.synchronousFetch) finish();
              else queueMicrotask(finish);
            }
            return { spec: "imap://fixture/original" };
          },
        };
      },
    },
    Cc: {
      "@mozilla.org/timer;1": { createInstance() {
        if (options.timerCreateError) throw new Error("Timer unavailable");
        const timer = {
          cancelled: false, cancellations: 0,
          initWithCallback(callback, delay) {
            this.callback = callback;
            this.delay = delay;
            clock.now += options.timerInitDuration || 0;
            if (options.timerInitError) throw new Error("Timer initialization failed");
          },
          cancel() { this.cancelled = true; this.cancellations++; },
        };
        calls.timers.push(timer);
        return timer;
      } },
      "@mozilla.org/messengercompose/composeparams;1": { createInstance() {
        calls.created.params++;
        return {};
      } },
      "@mozilla.org/messengercompose/composefields;1": { createInstance() {
        calls.created.fields++;
        return { addAttachment() {}, setHeader() {} };
      } },
      "@mozilla.org/messengercompose/compose;1": { createInstance() {
        calls.created.compose++;
        return {
          initialize(params) { this.params = params; },
          addMsgSendListener(listener) { this.listener = listener; },
          removeMsgSendListener() {},
          RegisterStateListener() {},
          UnregisterStateListener() {},
          sendMsg() {
            calls.sends++;
            this.listener.onStopSending("sent@example.test", 0);
            return Promise.resolve();
          },
        };
      } },
      "@mozilla.org/messengercompose/send;1": { createInstance() {
        calls.created.sender++;
        throw new Error("Unexpected direct sender");
      } },
    },
    findMessage: () => ({ msgHdr: hdr, folder }),
    setComposeIdentity(params) { params.identity = identity; },
    filePathsToAttachDescs: attachments.convert,
    materializeAttachmentDescs: attachments.materialize,
    descsToMsgAttachments: attachments.toNative,
    isSkipReviewBlocked: () => false,
    isToolEnabled: () => true,
    getReplyAllCcRecipients: () => "",
    PREF_ALLOW_ENCRYPTED_MESSAGES: "encrypted",
    isPrivacyOptInEnabled: () => true,
    async openComposeWindowWithCustomizations() {
      calls.reviews++;
      return { success: true };
    },
  };
  vm.createContext(sandbox);
  vm.runInContext([
    snippet("MCP TEXT SANITIZATION"), snippet("MESSAGE TEXT CONVERSION"),
    snippet("INLINE ATTACHMENT BASE64 HELPERS"), snippet("INLINE IMAGE CONTENT HELPERS"),
    snippet("RAW MIME PARSING HELPERS"), snippet("RAW MIME ATTACHMENT HELPERS"),
    snippet("ENCRYPTED MESSAGE GUARD"), snippet("DIRECT SEND ENCRYPTION GUARD"),
    snippet("DIRECT SEND ORIGINAL GUARD"), snippet("COMPOSE SIGNATURE HELPERS"), snippet("COMPOSE FORMAT HELPER"),
    snippet("COMPOSE SEND HELPER"), snippet("REPLY TOOL"), snippet("FORWARD TOOL"),
  ].join("\n"), sandbox);
  return {
    api: sandbox, calls, clock, attachments,
    limits: vm.runInContext("({ raw: RAW_MESSAGE_MAX_BYTES, html: MAX_HTML_PARSE_BYTES, time: DIRECT_SEND_PREFLIGHT_TIMEOUT_MS })", sandbox),
    run(tool) {
      const result = tool === "replyToMessage"
        ? sandbox.replyToMessage(hdr.messageId, "folder", "Reply", false, false,
          undefined, undefined, undefined, undefined, attachments.entries, true)
        : sandbox.forwardMessage(hdr.messageId, "folder", "to@example.test", "Forward", false,
          undefined, undefined, undefined, attachments.entries, true);
      return result.then(value => { calls.settlements++; return value; });
    },
  };
}

// Searched for refusal checks with real attachments and timer cleanup -> found
// assertNoAttachmentEffects, which is reused; existing suite-local checks cover
// identity/MIME refusals but have no timer settlement or stream ownership checks.
function assertPreflightRefused(harness, result) {
  assert.equal(result.success, undefined);
  assert.match(result.error, /skipReview: false/);
  assert.match(result.error, /review window/);
  assert.match(result.error, /Thunderbird.*encryption/);
  assertNoAttachmentEffects(harness.attachments.calls);
  assert.deepEqual(harness.calls.created, { params: 0, fields: 0, compose: 0, sender: 0 });
  assert.equal(harness.calls.sends, 0);
  assert.equal(harness.calls.reviews, 0);
  assert.equal(harness.calls.settlements, 1);
  assert.equal(harness.calls.rawCloses, harness.calls.rawOpens);
  for (const timer of harness.calls.timers) {
    assert.equal(timer.cancelled, true, "Every created timer is cancelled on refusal");
    assert.equal(timer.cancellations, 1, "Settlement cancels a timer once");
  }
}

if (require.main === module) describe("direct-send preflight limits and ordinary completion", () => {
  for (const tool of tools) {
    it(`${tool}: completes clear MIME and cancels its bounded preflight timer`, async () => {
      const h = loadPreflight();
      assert.deepEqual({ ...h.limits }, { raw: RAW_LIMIT, html: HTML_LIMIT, time: TIME_LIMIT });
      const result = await h.run(tool);
      assert.equal(result.success, true, result.error);
      assert.equal(h.calls.settlements, 1);
      assert.equal(h.calls.sends, 1);
      assert.equal(h.attachments.calls.tempCreates.length, 1, "Allowed messages still convert inline attachments");
      assert.deepEqual(h.calls.created, { params: 1, fields: 1, compose: 1, sender: 0 });
      assert.equal(h.calls.rawOpens, 1);
      assert.equal(h.calls.rawCloses, 1);
      assert.deepEqual(h.calls.mime, [{ allowDownload: true, examineEncryptedParts: false, saneBodySize: false }]);
      assert.deepEqual(h.calls.timers.map(timer => timer.delay), [TIME_LIMIT, 120000]);
      assert.ok(h.calls.timers.every(timer => timer.cancelled));
      assert.ok(h.calls.timers.reduce((total, timer) => total + timer.delay, 0) < 150000);
    });
    for (const [label, options] of [
      ["raw encrypted content", { raw: "Content-Type: application/pkcs7-mime; smime-type=enveloped-data\r\n\r\nMIIB" }],
      ["incomplete raw MIME", { raw: "Content-Type: multipart/mixed; boundary=b\r\n\r\n--b\r\nContent-Type: text/plain\r\n\r\npartial" }],
      ["raw available failure", { availableError: true }],
      ["raw read failure", { readError: true }],
      ["raw read makes no progress", { noProgress: true }],
      ["timer creation failure", { timerCreateError: true }],
      ["timer initialization failure", { timerInitError: true }],
    ]) {
      it(`${tool}: ${label} settles before Gloda or attachment conversion`, async () => {
        const h = loadPreflight(options);
        assertPreflightRefused(h, await h.run(tool));
        assert.equal(h.calls.imports, 0);
        assert.equal(h.calls.mime.length, 0);
      });
    }
    for (const [label, options] of [
      ["MIME module failure", { importError: true }],
      ["MIME parser failure", { mimeError: true }],
      ["missing MIME result", { mime: false }],
    ]) {
      it(`${tool}: ${label} settles once and releases the timer`, async () => {
        const h = loadPreflight(options);
        assertPreflightRefused(h, await h.run(tool));
      });
    }
  }
});

if (require.main === module) describe("direct-send preflight retrieves uncached IMAP originals before Gloda", () => {
  for (const tool of tools) {
    for (const synchronousFetch of [false, true]) {
      it(`${tool}: clear uncached IMAP mail sends after ${synchronousFetch ? "synchronous" : "asynchronous"} raw retrieval`, async () => {
        const h = loadPreflight({ uncached: true, synchronousFetch });
        const result = await h.run(tool);
        assert.equal(result.success, true, result.error);
        assert.equal(h.calls.rawOpenAttempts, 0);
        assert.equal(h.calls.rawOpens, 0);
        assert.deepEqual(h.calls.rawServices, ["message-uri"]);
        assert.equal(h.calls.rawFetches.length, 1);
        const [uri, listener, ...args] = h.calls.rawFetches[0];
        assert.equal(uri, "message-uri");
        assert.equal(typeof listener.onDataAvailable, "function");
        assert.deepEqual(args, [null, null, false, "", false], "Retrieval streams the original raw MIME");
        assert.deepEqual(h.calls.mime, [{ allowDownload: true, examineEncryptedParts: false, saneBodySize: false }]);
        assert.equal(h.calls.fetchBytesRead, (RAW_HEADER + "Original text").length);
        assert.ok(h.calls.fetchReads.every(count => count <= 64 * 1024));
        assert.equal(h.calls.sends, 1);
        assert.equal(h.calls.settlements, 1);
        assert.equal(h.attachments.calls.tempCreates.length, 1);
        assert.deepEqual(h.calls.created, { params: 1, fields: 1, compose: 1, sender: 0 });
        assert.ok(h.calls.timers.every(timer => timer.cancelled));
        assert.equal(h.calls.fetchCancellations.length, 0);
      });
    }
    it(`${tool}: waits for successful completion of every raw chunk before importing Gloda`, async () => {
      const h = loadPreflight({ uncached: true, deferFetch: true });
      const pending = h.run(tool);
      await Promise.resolve();
      h.calls.startFetch();
      h.calls.deliverRaw(RAW_HEADER + "Original ");
      assert.equal(h.calls.imports, 0);
      assert.equal(h.calls.mime.length, 0);
      assertNoAttachmentEffects(h.attachments.calls);
      h.calls.deliverRaw("text");
      assert.equal(h.calls.imports, 0);
      assert.equal(h.calls.mime.length, 0);
      assertNoAttachmentEffects(h.attachments.calls);
      h.calls.stopFetch();
      const result = await pending;
      assert.equal(result.success, true, result.error);
      assert.equal(h.calls.sends, 1);
      assert.equal(h.calls.imports, 1);
      assert.equal(h.calls.mime.length, 1);
    });
    it(`${tool}: retries an unreadable cached IMAP stream after closing it`, async () => {
      const h = loadPreflight({ serverType: "imap", readError: true });
      const result = await h.run(tool);
      assert.equal(result.success, true, result.error);
      assert.equal(h.calls.rawOpens, 1);
      assert.equal(h.calls.rawCloses, 1);
      assert.equal(h.calls.rawFetches.length, 1);
      assert.equal(h.calls.sends, 1);
    });
    for (const serverType of ["pop3", "none", undefined]) {
      it(`${tool}: does not download an unavailable ${serverType || "unknown"} original`, async () => {
        const h = loadPreflight({ uncached: true, serverType: serverType || "" });
        assertPreflightRefused(h, await h.run(tool));
        assert.equal(h.calls.rawFetches.length, 0);
        assert.equal(h.calls.imports, 0);
      });
    }
    for (const [label, raw] of [
      ["encrypted S/MIME", "Content-Type: application/pkcs7-mime; smime-type=enveloped-data\r\n\r\nMIIB"],
      ["opaque-signed S/MIME", "Content-Type: application/pkcs7-mime; smime-type=signed-data\r\n\r\nMIIB"],
      ["encrypted OpenPGP", "Content-Type: multipart/encrypted; protocol=\"application/pgp-encrypted\"; boundary=p\r\n\r\n--p\r\nContent-Type: application/pgp-encrypted\r\n\r\nVersion: 1\r\n--p--\r\n"],
      ["inline OpenPGP", RAW_HEADER + "-----BEGIN PGP MESSAGE-----\r\nVersion: GnuPG\r\n\r\nYWJj\r\n-----END PGP MESSAGE-----"],
      ["incomplete MIME", "Content-Type: multipart/mixed; boundary=p\r\n\r\n--p\r\nContent-Type: text/plain\r\n\r\npartial"],
      ["empty raw content", ""],
    ]) {
      it(`${tool}: refuses downloaded ${label} before invoking any structured parser`, async () => {
        const h = loadPreflight({ uncached: true, raw });
        assertPreflightRefused(h, await h.run(tool));
        assert.equal(h.calls.rawFetches.length, 1);
        assert.equal(h.calls.imports, 0);
        assert.equal(h.calls.mime.length, 0);
      });
    }
    for (const [label, options] of [
      ["message service failure", { serviceError: true }],
      ["streamMessage failure", { fetchError: true }],
      ["failed request completion", { fetchStatus: 1 }],
      ["read failure", { fetchReadError: true }],
      ["read without progress", { fetchNoProgress: true }],
      ["short callback read", { fetchReadLimit: 7 }],
    ]) {
      it(`${tool}: raw retrieval ${label} refuses without attachment or native compose effects`, async () => {
        const h = loadPreflight({ uncached: true, ...options });
        assertPreflightRefused(h, await h.run(tool));
        assert.equal(h.calls.imports, 0);
        assert.equal(h.calls.mime.length, 0);
        if (options.fetchReadError || options.fetchNoProgress || options.fetchReadLimit) {
          assert.equal(h.calls.fetchCancellations.length, 1);
        }
      });
    }
    it(`${tool}: refuses raw callback counts larger than the bytes actually read`, async () => {
      const h = loadPreflight({ uncached: true, deferFetch: true });
      const pending = h.run(tool);
      await Promise.resolve();
      h.calls.startFetch();
      const raw = RAW_HEADER + "Original text";
      h.calls.deliverRaw(raw, raw.length + 1);
      h.calls.stopFetch();
      assertPreflightRefused(h, await pending);
      assert.equal(h.calls.imports, 0);
      assert.equal(h.calls.fetchCancellations.length, 1);
    });
    it(`${tool}: refuses a failed completion even after receiving complete clear MIME`, async () => {
      const h = loadPreflight({ uncached: true, deferFetch: true });
      const pending = h.run(tool);
      await Promise.resolve();
      h.calls.startFetch();
      h.calls.deliverRaw(RAW_HEADER + "Original text");
      assert.equal(h.calls.imports, 0);
      h.calls.stopFetch(1);
      assertPreflightRefused(h, await pending);
      assert.equal(h.calls.imports, 0);
    });
    it(`${tool}: downloaded raw overflow is bounded and cancels the request`, async () => {
      const raw = RAW_HEADER + "x".repeat(RAW_LIMIT + 1 - RAW_HEADER.length);
      const h = loadPreflight({ uncached: true, raw });
      assertPreflightRefused(h, await h.run(tool));
      assert.ok(h.calls.fetchBytesRead <= RAW_LIMIT + 1);
      assert.equal(h.calls.fetchCancellations.length, 1);
      assert.equal(h.calls.imports, 0);
    });
    it(`${tool}: downloaded raw content exactly at the limit remains eligible to send`, async () => {
      const raw = RAW_HEADER + "x".repeat(RAW_LIMIT - RAW_HEADER.length);
      const h = loadPreflight({ uncached: true, raw });
      const result = await h.run(tool);
      assert.equal(result.success, true, result.error);
      assert.equal(h.calls.fetchBytesRead, RAW_LIMIT);
      assert.ok(h.calls.fetchReads.every(count => count <= 64 * 1024));
      assert.equal(h.calls.fetchCancellations.length, 0);
      assert.equal(h.calls.sends, 1);
    });
    it(`${tool}: aggregate raw retrieval overflow cannot be hidden across callbacks`, async () => {
      const h = loadPreflight({
        uncached: true,
        fetchChunks: [RAW_HEADER + "x".repeat(RAW_LIMIT - RAW_HEADER.length), "x"],
      });
      assertPreflightRefused(h, await h.run(tool));
      assert.ok(h.calls.fetchBytesRead <= RAW_LIMIT + 1);
      assert.equal(h.calls.fetchCancellations.length, 1);
      assert.equal(h.calls.imports, 0);
    });
    it(`${tool}: cached raw overflow refuses without attempting an IMAP download`, async () => {
      const raw = RAW_HEADER + "x".repeat(RAW_LIMIT + 1 - RAW_HEADER.length);
      const h = loadPreflight({ serverType: "imap", raw });
      assertPreflightRefused(h, await h.run(tool));
      assert.equal(h.calls.rawFetches.length, 0);
      assert.equal(h.calls.imports, 0);
    });
    it(`${tool}: an expired cache read does not start an IMAP download`, async () => {
      const h = loadPreflight({ uncached: true, offline: true, openDuration: TIME_LIMIT });
      assertPreflightRefused(h, await h.run(tool));
      assert.equal(h.calls.rawFetches.length, 0);
      assert.equal(h.calls.imports, 0);
    });
    for (const elapsed of [TIME_LIMIT - 1, TIME_LIMIT, TIME_LIMIT + 1]) {
      it(`${tool}: raw request completion at ${elapsed} ms obeys the shared deadline`, async () => {
        const h = loadPreflight({ uncached: true, deferFetch: true });
        const pending = h.run(tool);
        await Promise.resolve();
        h.calls.startFetch();
        h.calls.deliverRaw(RAW_HEADER + "Original text");
        h.clock.now += elapsed;
        h.calls.stopFetch();
        const result = await pending;
        if (elapsed < TIME_LIMIT) {
          assert.equal(result.success, true, result.error);
          assert.equal(h.calls.sends, 1);
        } else {
          assertPreflightRefused(h, result);
          assert.equal(h.calls.imports, 0);
        }
      });
    }
    it(`${tool}: expiry during a raw callback read cancels retrieval before Gloda`, async () => {
      const h = loadPreflight({ uncached: true, fetchReadDuration: TIME_LIMIT });
      assertPreflightRefused(h, await h.run(tool));
      assert.equal(h.calls.fetchCancellations.length, 1);
      assert.equal(h.calls.imports, 0);
    });
    for (const started of [false, true]) {
      it(`${tool}: timeout ${started ? "after" : "before"} onStartRequest ignores all late raw callbacks`, async () => {
        const h = loadPreflight({ uncached: true, deferFetch: true });
        const pending = h.run(tool);
        await Promise.resolve();
        if (started) h.calls.startFetch();
        const timer = h.calls.timers[0];
        h.clock.now += TIME_LIMIT;
        timer.callback.notify();
        assertPreflightRefused(h, await pending);
        assert.equal(h.calls.fetchCancellations.length, started ? 1 : 0,
          "A request becomes cancellable when onStartRequest supplies it");
        let lateReads = 0;
        h.calls.fetchListener.onDataAvailable(h.calls.fetchRequest, {
          read() { lateReads++; throw new Error("Late raw data must not be read"); },
          available() { lateReads++; throw new Error("Late raw data must not be inspected"); },
        }, 0, 10);
        if (!started) h.calls.startFetch();
        h.calls.stopFetch();
        timer.callback.notify();
        await Promise.resolve();
        assertPreflightRefused(h, await pending);
        assert.equal(lateReads, 0);
        assert.equal(h.calls.imports, 0);
        assert.equal(h.calls.mime.length, 0);
        assert.equal(h.calls.fetchCancellations.length, 1);
      });
    }
    it(`${tool}: a cancellation error cannot reopen a timed out raw preflight`, async () => {
      const h = loadPreflight({ uncached: true, deferFetch: true, cancelError: true });
      const pending = h.run(tool);
      await Promise.resolve();
      h.calls.startFetch();
      h.clock.now += TIME_LIMIT;
      h.calls.timers[0].callback.notify();
      assertPreflightRefused(h, await pending);
      h.calls.deliverRaw(RAW_HEADER + "late clear text");
      h.calls.stopFetch();
      await Promise.resolve();
      assertPreflightRefused(h, await pending);
      assert.equal(h.calls.imports, 0);
    });
  }
});

if (require.main === module) describe("direct-send preflight byte limits", () => {
  for (const tool of tools) {
    it(`${tool}: raw overflow reads at most one byte beyond the limit and never invokes Gloda`, async () => {
      const raw = RAW_HEADER + "x".repeat(RAW_LIMIT + 1 - RAW_HEADER.length);
      const h = loadPreflight({ raw });
      assertPreflightRefused(h, await h.run(tool));
      assert.equal(h.calls.bytesRead, RAW_LIMIT + 1);
      assert.ok(h.calls.rawReads.every(count => count <= 64 * 1024));
      assert.equal(h.calls.imports, 0);
    });
    for (const [label, mime] of [
      ["one oversized body", { contentType: "text/plain", body: "x".repeat(RAW_LIMIT + 1) }],
      ["aggregate body bytes", { contentType: "multipart/mixed", parts: [
        { contentType: "text/plain", body: "x".repeat(RAW_LIMIT / 2) },
        { contentType: "text/plain", body: "y".repeat(RAW_LIMIT / 2 + 1) },
      ] }],
      ["multibyte body bytes", { contentType: "text/plain", body: "😀".repeat(RAW_LIMIT / 4 + 1) }],
      ["HTML above the presentation cap", { contentType: "text/html", body: "<p>" + " ".repeat(HTML_LIMIT) + "</p>" }],
      ["HTML armor beyond the presentation cap", { contentType: "text/html", body: "<p>" + " ".repeat(HTML_LIMIT) + "-----BEGIN PGP MESSAGE-----</p>" }],
      ["joined HTML above the presentation cap", { contentType: "multipart/mixed", parts: [
        { contentType: "text/html", body: "x".repeat(HTML_LIMIT / 2) },
        { contentType: "text/html", body: "y".repeat(HTML_LIMIT / 2 + 1) },
      ] }],
      ["multibyte HTML above the presentation cap", { contentType: "text/html", body: "é".repeat(HTML_LIMIT / 2 + 1) }],
    ]) {
      it(`${tool}: refuses ${label} without classifying a truncated prefix`, async () => {
        const h = loadPreflight({ mime });
        let quoteCalls = 0;
        const quote = h.api.getDirectQuoteContent;
        h.api.getDirectQuoteContent = (...args) => { quoteCalls++; return quote(...args); };
        assertPreflightRefused(h, await h.run(tool));
        assert.equal(h.calls.mime.length, 1);
        assert.equal(quoteCalls, 0, "Completeness is checked before quote extraction can truncate data");
      });
    }
    it(`${tool}: refuses an oversized plaintext-coercion result`, async () => {
      const h = loadPreflight({ mime: { parts: [], coerceBodyToPlaintext: () => "x".repeat(RAW_LIMIT + 1) } });
      assertPreflightRefused(h, await h.run(tool));
    });
  }

  it("accepts the exact raw byte limit without truncation", () => {
    const raw = RAW_HEADER + "x".repeat(RAW_LIMIT - RAW_HEADER.length);
    const h = loadPreflight({ raw });
    const stream = {
      cursor: 0,
      available() { return raw.length - this.cursor; },
      read(count) {
        const chunk = raw.slice(this.cursor, this.cursor + count);
        this.cursor += chunk.length;
        return chunk;
      },
    };
    assert.equal(h.api.readMessageStreamFully(stream, RAW_LIMIT, h.clock.now + TIME_LIMIT).length, RAW_LIMIT);
    assert.equal(stream.cursor, RAW_LIMIT);
  });
  for (const [label, mime] of [
    ["ASCII body at the exact byte limit", { contentType: "text/plain", body: "x".repeat(RAW_LIMIT) }],
    ["multibyte body at the exact byte limit", { contentType: "text/plain", body: "😀".repeat(RAW_LIMIT / 4) }],
    ["HTML at the exact presentation cap", { contentType: "text/html", body: "x".repeat(HTML_LIMIT) }],
    ["joined HTML at the exact presentation cap", { contentType: "multipart/mixed", parts: [
      { contentType: "text/html", body: "x".repeat(HTML_LIMIT / 2) },
      { contentType: "text/html", body: "y".repeat(HTML_LIMIT / 2) },
    ] }],
  ]) {
    it(`accepts ${label} as complete`, () => {
      const h = loadPreflight();
      assert.equal(h.api.isDirectSendMimeComplete(mime, h.clock.now + TIME_LIMIT), true);
    });
  }
});

if (require.main === module) describe("direct-send raw HTML completeness before Gloda", () => {
  const signedType = 'multipart/signed; protocol="application/pgp-signature"; boundary="sig"';
  const signedMime = {
    contentType: "message/rfc822", headers: { "content-type": [signedType] },
    parts: [{ contentType: "multipart/signed", headers: { "content-type": [signedType] }, isEncrypted: true, parts: [] }],
  };
  const signature = "Content-Type: application/pgp-signature\r\n\r\n-----BEGIN PGP SIGNATURE-----\r\nwsB5\r\n-----END PGP SIGNATURE-----";
  for (const tool of tools) {
    for (const [content, html] of [
      ["clear HTML one byte over the cap", "<p>" + "x".repeat(HTML_LIMIT - 6) + "</p>"],
      ["HTML with valid armor beyond the cap", "<p>" + "x".repeat(HTML_LIMIT) + "</p><p>-----BEGIN PGP MESSAGE-----</p>"],
    ]) {
      for (const encoding of ["8bit", "base64"]) {
        const encoded = encoding === "base64" ? Buffer.from(html).toString("base64") : html;
        const leaf = `Content-Type: text/html; charset=utf-8\r\nContent-Transfer-Encoding: ${encoding}\r\n\r\n${encoded}`;
        for (const [wrapper, raw, mime] of [
          ["ordinary raw message", leaf, undefined],
          ["flagged clear-signed wrapper", `Content-Type: ${signedType}\r\n\r\n--sig\r\n${leaf}\r\n--sig\r\n${signature}\r\n--sig--\r\n`, signedMime],
          ["nested multipart", `Content-Type: multipart/mixed; boundary=mix\r\n\r\n--mix\r\n${leaf}\r\n--mix--\r\n`, undefined],
          ["embedded message", `Content-Type: message/rfc822\r\n\r\n${leaf}`, undefined],
        ]) {
          it(`${tool}: refuses ${content}, ${encoding}, ${wrapper}, before parsing HTML or importing Gloda`, async () => {
            let parses = 0;
            // Searched for HTML parser fixtures -> privacy-messages has a local
            // DOM fixture. Retain its text-node shape as a spy; refusing these
            // bytes must precede any presentation parser call.
            const h = loadPreflight({ raw, mime, DOMParser: class {
              parseFromString() {
                parses++;
                return {
                  querySelectorAll: () => [],
                  body: { nodeType: 1, tagName: "BODY", childNodes: [{ nodeType: 3, textContent: "clear prefix" }] },
                };
              }
            } });
            assertPreflightRefused(h, await h.run(tool));
            assert.equal(parses, 0, "No prefix can be silently truncated and classified as complete");
            assert.equal(h.calls.imports, 0);
            assert.equal(h.calls.mime.length, 0);
            assert.equal(h.calls.rawOpens, 1);
          });
        }
      }
    }
    for (const signed of [false, true]) {
      it(`${tool}: accepts raw HTML at the exact cap, clear-signed=${signed}, using one raw snapshot`, async () => {
        const text = "x".repeat(HTML_LIMIT - 7);
        const html = `<p>${text}</p>`;
        const leaf = `Content-Type: text/html; charset=utf-8\r\n\r\n${html}`;
        const raw = signed
          ? `Content-Type: ${signedType}\r\n\r\n--sig\r\n${leaf}\r\n--sig\r\n${signature}\r\n--sig--\r\n`
          : leaf;
        let parses = 0;
        const h = loadPreflight({ raw, mime: signed ? signedMime : undefined, DOMParser: class {
          parseFromString(input, type) {
            parses++;
            assert.equal(input, html);
            assert.equal(type, "text/html");
            return {
              querySelectorAll: () => [],
              body: { nodeType: 1, tagName: "BODY", childNodes: [{ nodeType: 3, textContent: text }] },
            };
          }
        } });
        const result = await h.run(tool);
        assert.equal(result.success, true, result.error);
        assert.equal(h.calls.sends, 1);
        assert.equal(h.calls.rawOpens, 1, "Quoting must reuse the already bounded raw snapshot");
        assert.equal(h.calls.rawCloses, 1);
        assert.equal(h.calls.mime.length, 1);
        assert.ok(parses > 0);
      });
    }
    it(`${tool}: a clear-signed snapshot classification that exhausts the deadline refuses without a second raw read`, async () => {
      const leaf = RAW_HEADER + "signed original";
      const raw = `Content-Type: ${signedType}\r\n\r\n--sig\r\n${leaf}\r\n--sig\r\n${signature}\r\n--sig--\r\n`;
      const h = loadPreflight({ raw, mime: signedMime });
      let classifications = 0;
      const classify = h.api.classifyRawMessageEncryption;
      h.api.classifyRawMessageEncryption = (value, depth = 0, options) => {
        const result = classify(value, depth, options);
        if (depth === 0 && ++classifications === 2) h.clock.now += TIME_LIMIT;
        return result;
      };
      assertPreflightRefused(h, await h.run(tool));
      assert.equal(h.calls.rawOpens, 1);
      assert.equal(classifications, 2);
    });
  }
});

if (require.main === module) describe("direct-send preflight refuses incomplete structured MIME", () => {
  for (const tool of tools) {
    for (const flag of ["truncated", "isTruncated", "bodyTruncated"]) {
      for (const nested of [false, true]) {
        it(`${tool}: refuses ${flag} on ${nested ? "a nested part" : "the root"}`, async () => {
          const part = { contentType: "text/plain", body: "apparently clear prefix", [flag]: true };
          const mime = nested ? { contentType: "multipart/mixed", parts: [part] } : part;
          const h = loadPreflight({ mime });
          assertPreflightRefused(h, await h.run(tool));
        });
      }
    }
    for (const [label, mime] of [
      ["non-array parts", { contentType: "multipart/mixed", parts: {} }],
      ["non-string body", { contentType: "text/plain", body: 42 }],
      ["throwing body getter", { contentType: "text/plain", get body() { throw new Error("Unavailable body"); } }],
      ["throwing completeness flag", { contentType: "text/plain", body: "clear", get truncated() { throw new Error("Unavailable state"); } }],
    ]) {
      it(`${tool}: refuses ${label}`, async () => {
        const h = loadPreflight({ mime });
        assertPreflightRefused(h, await h.run(tool));
      });
    }
    it(`${tool}: refuses a cyclic MIME tree`, async () => {
      const mime = { contentType: "multipart/mixed", parts: [] };
      mime.parts.push(mime);
      const h = loadPreflight({ mime });
      assertPreflightRefused(h, await h.run(tool));
    });
  }
});

if (require.main === module) describe("direct-send preflight deadline and late callbacks", () => {
  for (const tool of tools) {
    for (const [label, options] of [
      ["before opening the raw stream", { timerInitDuration: TIME_LIMIT }],
      ["while opening the raw stream", { openDuration: TIME_LIMIT }],
      ["during raw streaming", { raw: RAW_HEADER + "x".repeat(128 * 1024), readDuration: TIME_LIMIT }],
    ]) {
      it(`${tool}: elapsed deadline ${label} refuses synchronously`, async () => {
        const h = loadPreflight(options);
        assertPreflightRefused(h, await h.run(tool));
        assert.equal(h.calls.imports, 0);
        if (options.timerInitDuration) assert.equal(h.calls.rawOpens, 0);
        if (options.readDuration) assert.equal(h.calls.rawReads.length, 1);
      });
    }
    for (const stage of ["classifyRawMessageEncryption", "getDirectQuoteContent"]) {
      it(`${tool}: elapsed deadline after ${stage} refuses before side effects`, async () => {
        const h = loadPreflight();
        const original = h.api[stage];
        h.api[stage] = (...args) => {
          const result = original(...args);
          h.clock.now += TIME_LIMIT;
          return result;
        };
        assertPreflightRefused(h, await h.run(tool));
        if (stage === "classifyRawMessageEncryption") assert.equal(h.calls.imports, 0);
      });
    }
    it(`${tool}: elapsed deadline during module import refuses before calling Gloda`, async () => {
      const h = loadPreflight({ importDuration: TIME_LIMIT });
      assertPreflightRefused(h, await h.run(tool));
      assert.equal(h.calls.imports, 1);
      assert.equal(h.calls.mime.length, 0);
    });
    it(`${tool}: elapsed deadline during structured body inspection refuses`, async () => {
      const mime = { contentType: "text/plain" };
      const h = loadPreflight({ mime });
      Object.defineProperty(mime, "body", { get() { h.clock.now += TIME_LIMIT; return "visible"; } });
      assertPreflightRefused(h, await h.run(tool));
    });
    for (const elapsed of [TIME_LIMIT - 1, TIME_LIMIT, TIME_LIMIT + 1]) {
      it(`${tool}: callback at ${elapsed} ms obeys the deadline without relying on timer scheduling`, async () => {
        const h = loadPreflight({ deferMime: true });
        const pending = h.run(tool);
        await Promise.resolve();
        assert.equal(h.calls.mime.length, 1);
        h.clock.now += elapsed;
        h.calls.deliverMime({ contentType: "text/plain", body: "Original text" });
        const result = await pending;
        if (elapsed < TIME_LIMIT) {
          assert.equal(result.success, true, result.error);
          assert.equal(h.calls.sends, 1);
          assert.ok(h.calls.timers.every(timer => timer.cancelled));
        } else {
          assertPreflightRefused(h, result);
        }
      });
    }
    it(`${tool}: timeout settles once and ignores late callbacks without reading their MIME`, async () => {
      const h = loadPreflight({ deferMime: true });
      const pending = h.run(tool);
      await Promise.resolve();
      const timer = h.calls.timers[0];
      assert.equal(timer.delay, TIME_LIMIT);
      h.clock.now += TIME_LIMIT;
      timer.callback.notify();
      assertPreflightRefused(h, await pending);
      let getterReads = 0;
      const lateMime = {
        get body() { getterReads++; throw new Error("Late MIME must not be inspected"); },
        get contentType() { getterReads++; throw new Error("Late MIME must not be inspected"); },
      };
      h.calls.deliverMime(lateMime);
      h.calls.deliverMime({ contentType: "text/plain", body: "late clear callback" });
      timer.callback.notify();
      await Promise.resolve();
      assertPreflightRefused(h, await pending);
      assert.equal(getterReads, 0);
    });
    it(`${tool}: elapsed deadline before the awaiting tool resumes prevents attachment conversion`, async () => {
      const h = loadPreflight();
      const pending = h.run(tool);
      await Promise.resolve();
      assert.equal(h.calls.mime.length, 1);
      assert.equal(h.calls.timers[0].cancelled, true, "The callback completed before the tool resumed");
      h.clock.now += TIME_LIMIT;
      assertPreflightRefused(h, await pending);
    });
    it(`${tool}: repeated callbacks after success cannot compose or send again`, async () => {
      const h = loadPreflight({ deferMime: true });
      const pending = h.run(tool);
      await Promise.resolve();
      h.calls.deliverMime({ contentType: "text/plain", body: "Original text" });
      assert.equal((await pending).success, true);
      h.calls.deliverMime({ contentType: "text/plain", body: "duplicate callback" });
      h.calls.timers[0].callback.notify();
      await Promise.resolve();
      assert.equal(h.calls.sends, 1);
      assert.equal(h.calls.settlements, 1);
      assert.equal(h.attachments.calls.tempCreates.length, 1);
      assert.deepEqual(h.calls.created, { params: 1, fields: 1, compose: 1, sender: 0 });
    });
  }
});

module.exports = { loadPreflight, assertPreflightRefused };

if (require.main === module) describe("reply/forward size assessment counts inherited content correctly", () => {
  for (const tool of tools) {
    it(`${tool}: a large original raw body does not dominate a short MIME quote`, async () => {
      const raw = RAW_HEADER + "x".repeat(200000);
      const h = loadPreflight({ raw, prefs: { [SIZE_WARNING_PREF]: SMALL_SIZE_WARNING_BYTES } });
      const plans = [];
      const assess = h.api.assessComposeSize;
      h.api.assessComposeSize = plan => { plans.push(plan); return assess(plan); };
      const result = await h.run(tool);
      assert.equal(result.success, true, result.error);
      assert.equal(plans.length, 2);
      assert.equal(plans[0].inheritedRawBytes, undefined);
      assert.equal(plans[1].inheritedRawBytes, undefined);
      assert.ok(plans[0].body.length < 100);
      assert.equal(h.calls.rawOpens + h.calls.rawFetches.length, 1);
      assert.equal(h.calls.mime.length, 1);
      assert.equal(h.calls.sends, 1);
    });
    it(`${tool}: both checks use the same estimator without an original-byte floor`, async () => {
      const h = loadPreflight({ prefs: { [SIZE_WARNING_PREF]: DEFAULT_SIZE_WARNING_BYTES } });
      const plans = [];
      const assess = h.api.assessComposeSize;
      h.api.assessComposeSize = plan => { plans.push(plan); return assess(plan); };
      const result = await h.run(tool);
      assert.equal(result.success, true, result.error);
      assert.equal(plans.length, 2);
      assert.equal(plans[0].inheritedRawBytes, undefined);
      assert.equal(plans[1].inheritedRawBytes, undefined);
      assert.ok(plans[1].body.length > plans[0].body.length);
      assert.equal(h.calls.sends, 1);
    });
    it(`${tool}: unreadable thresholds refuse before native compose and attachment writes`, async () => {
      const h = loadPreflight({ prefs: { [SIZE_WARNING_PREF]: undefined } });
      const result = await h.run(tool);
      assert.match(result.error, /threshold could not be read/);
      assert.equal(h.calls.sends, 0);
      assert.deepEqual(h.calls.created, { params: 0, fields: 0, compose: 0, sender: 0 });
      assert.equal(h.attachments.calls.tempCreates.length + h.attachments.calls.writes.length + h.attachments.calls.nativeCreates, 0);
    });
  }
});


if (require.main === module) describe("final reply/forward size boundaries before attachment writes", () => {
  for (const tool of tools) {
    it(`${tool}: equality sends; a final estimate one byte above the threshold refuses before writes`, async () => {
      const warm = loadPreflight();
      let finalEstimate;
      const assess = warm.api.assessComposeSize;
      warm.api.assessComposeSize = plan => {
        const result = assess(plan);
        finalEstimate = result.estimatedEncodedBytes;
        return result;
      };
      assert.equal((await warm.run(tool)).success, true);
      for (const delta of [-1, 0]) {
        const h = loadPreflight({ prefs: { [SIZE_WARNING_PREF]: finalEstimate + delta } });
        const result = await h.run(tool);
        if (delta === 0) {
          assert.equal(result.success, true, result.error);
          assert.equal(h.calls.sends, 1);
        } else {
          assert.match(result.error, /encoded size estimate.*exceeds/);
          assert.equal(h.calls.created.compose + h.calls.sends, 0);
          assert.equal(h.attachments.calls.tempCreates.length + h.attachments.calls.writes.length + h.attachments.calls.nativeCreates, 0);
        }
      }
    });
  }
  const rawWithAttachment = [
    'Content-Type: multipart/mixed; boundary="parts"', '',
    '--parts', 'Content-Type: text/plain', '', 'intro',
    '--parts', 'Content-Type: application/octet-stream', 'Content-Transfer-Encoding: base64',
    'Content-Disposition: attachment; filename="original.bin"', '', 'A'.repeat(200000), '--parts--', '',
  ].join('\r\n');
  const originalAttachment = { name: "original.bin", contentType: "application/octet-stream",
    url: "imap://fixture/original?part=1.2", size: 150000 };
  const originalMime = { contentType: "multipart/mixed", parts: [
    { contentType: "text/plain", body: "intro" }, { contentType: originalAttachment.contentType },
  ], allUserAttachments: [originalAttachment] };

  it("counts original forward attachments by decoded size, not the raw snapshot", async () => {
    const h = loadPreflight({ raw: rawWithAttachment, mime: originalMime,
      prefs: { [SIZE_WARNING_PREF]: 100000 } });
    const plans = [];
    const assess = h.api.assessComposeSize;
    h.api.assessComposeSize = plan => { plans.push(plan); return assess(plan); };
    const result = await h.run('forwardMessage');
    assert.match(result.error, /encoded size estimate.*exceeds/);
    assert.equal(plans[0].attachments.some(att => att.name === originalAttachment.name
      && att.size === originalAttachment.size), true);
    assert.equal(plans[0].inheritedRawBytes, undefined);
    assert.equal(h.calls.bytesRead, rawWithAttachment.length);
    assert.equal(h.calls.sends + h.calls.created.compose, 0);
    assert.equal(h.attachments.calls.writes.length + h.attachments.calls.nativeCreates, 0);
  });

  it("a reply does not count the original attachment bytes", async () => {
    const h = loadPreflight({ raw: rawWithAttachment, mime: originalMime,
      prefs: { [SIZE_WARNING_PREF]: 150000 } });
    const plans = [];
    const assess = h.api.assessComposeSize;
    h.api.assessComposeSize = plan => { plans.push(plan); return assess(plan); };
    const result = await h.run('replyToMessage');
    assert.equal(result.success, true, result.error);
    assert.equal(plans[0].inheritedRawBytes, undefined);
    assert.equal(h.calls.sends, 1);
  });

  it("a forward with an ordinary 150 KB attachment fits the default threshold", async () => {
    const h = loadPreflight({ raw: rawWithAttachment, mime: originalMime,
      prefs: { [SIZE_WARNING_PREF]: DEFAULT_SIZE_WARNING_BYTES } });
    const estimates = [];
    const assess = h.api.assessComposeSize;
    h.api.assessComposeSize = plan => {
      const result = assess(plan);
      estimates.push(result.estimatedEncodedBytes);
      return result;
    };
    assert.equal((await h.run('forwardMessage')).success, true);
    assert.equal(h.calls.sends, 1);
    assert.equal(estimates.length, 2);
    assert.ok(estimates.every(bytes => bytes > originalAttachment.size && bytes < 2 * 1024 * 1024));
  });

  const videoAttachment = { name: "video.mp4", contentType: "video/mp4",
    url: "imap://fixture/original?part=1.3", size: 10 * 1024 * 1024 };
  const videoMime = { contentType: "multipart/mixed", parts: [
    { contentType: "text/plain", body: "intro" }, { contentType: videoAttachment.contentType },
  ], allUserAttachments: [videoAttachment] };

  it("a 10 MiB video forward fits the 20 MiB threshold", async () => {
    const h = loadPreflight({ raw: RAW_HEADER + "intro", mime: videoMime,
      prefs: { [SIZE_WARNING_PREF]: DEFAULT_SIZE_WARNING_BYTES } });
    const result = await h.run('forwardMessage');
    assert.equal(result.success, true, result.error);
    assert.equal(h.calls.sends, 1);
  });

  it("a 10 MiB video is still blocked by a lower threshold", async () => {
    const h = loadPreflight({ raw: RAW_HEADER + "intro", mime: videoMime,
      prefs: { [SIZE_WARNING_PREF]: 15 * 1024 * 1024 } });
    const result = await h.run('forwardMessage');
    assert.match(result.error, /encoded size estimate.*exceeds/);
    assert.equal(h.calls.sends, 0);
    assert.equal(h.calls.created.compose, 0);
  });

  it("a reply to a 10 MiB video original is not blocked", async () => {
    const h = loadPreflight({ raw: RAW_HEADER + "intro", mime: videoMime,
      prefs: { [SIZE_WARNING_PREF]: 1024 * 1024 } });
    const result = await h.run('replyToMessage');
    assert.equal(result.success, true, result.error);
    assert.equal(h.calls.sends, 1);
  });

  it("bounds new headers from inherited attachments before materializing caller attachments", async () => {
    for (const metadata of [
      { name: "名".repeat(100000) },
      { url: "owl://fixture/" + "x".repeat(100000) },
    ]) {
      const h = loadPreflight({ raw: rawWithAttachment,
        mime: { ...originalMime, allUserAttachments: [{ ...originalAttachment, ...metadata }] },
        prefs: { [SIZE_WARNING_PREF]: 1000000 } });
      const result = await h.run("forwardMessage");
      assert.match(result.error, /encoded size estimate.*exceeds/);
      assert.ok(h.calls.created.fields <= 1, "metadata is counted before or at the assembled-quote check");
      assert.equal(h.calls.created.compose + h.calls.sends, 0);
      assert.equal(h.attachments.calls.writes.length + h.attachments.calls.nativeCreates, 0);
    }
  });
});
