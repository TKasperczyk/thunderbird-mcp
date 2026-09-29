"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const NS_BASE_STREAM_CLOSED = "NS_BASE_STREAM_CLOSED";
const FIXED_CHUNK_SIZE = 64 * 1024;

function loadReadMessageStreamFully() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN RAW MIME ATTACHMENT HELPERS";
  const endMarker = "// END RAW MIME ATTACHMENT HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "raw MIME attachment helper start marker missing");
  assert.ok(end > start, "raw MIME attachment helper end marker missing");

  const snippet = source.slice(start, end);
  const sandbox = {
    Cr: { NS_BASE_STREAM_CLOSED },
    NetUtil: {
      readInputStreamToString(stream, count) {
        return stream.read(count);
      },
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.readMessageStreamFully = readMessageStreamFully;`,
    sandbox
  );
  return sandbox.readMessageStreamFully;
}

const readMessageStreamFully = loadReadMessageStreamFully();

function loadInlineImageFetcher({
  stream,
  contentLength,
  log,
  newChannelError,
  asyncFetchError,
}) {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  function markedSnippet(startMarker, endMarker) {
    const start = source.indexOf(startMarker);
    const end = source.indexOf(endMarker);
    assert.ok(start >= 0, `${startMarker} missing`);
    assert.ok(end > start, `${endMarker} missing`);
    return source.slice(start, end);
  }

  const fetchStart = source.indexOf("function fetchInlineImageBase64(source) {");
  const fetchEnd = source.indexOf(
    "async function appendInlineImageContent() {",
    fetchStart
  );
  assert.ok(fetchStart >= 0, "fetchInlineImageBase64 missing");
  assert.ok(fetchEnd > fetchStart, "fetchInlineImageBase64 end missing");

  const snippet = [
    markedSnippet(
      "// BEGIN INLINE IMAGE CONTENT HELPERS",
      "// END INLINE IMAGE CONTENT HELPERS"
    ),
    markedSnippet(
      "// BEGIN RAW MIME ATTACHMENT HELPERS",
      "// END RAW MIME ATTACHMENT HELPERS"
    ),
    source.slice(fetchStart, fetchEnd),
  ].join("\n");
  const sandbox = {
    btoa: globalThis.btoa,
    Cr: { NS_BASE_STREAM_CLOSED },
    console: { error: (...args) => log.push(args) },
    NetUtil: {
      newChannel() {
        if (newChannelError) throw newChannelError;
        return {};
      },
      asyncFetch(channel, callback) {
        if (asyncFetchError) throw asyncFetchError;
        callback(stream, 0, { contentLength });
      },
      readInputStreamToString(inputStream, count) {
        return inputStream.read(count);
      },
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.fetchInlineImageBase64 = fetchInlineImageBase64;`,
    sandbox
  );
  return sandbox.fetchInlineImageBase64;
}

function loadGetMessage({ stream, log, findMessageError }) {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const helperStart = source.indexOf("// BEGIN RAW MIME ATTACHMENT HELPERS");
  const helperEnd = source.indexOf("// END RAW MIME ATTACHMENT HELPERS");
  const getMessageStart = source.indexOf("function getMessage(messageId,");
  const getMessageEnd = source.indexOf(
    "async function getMessages(",
    getMessageStart
  );
  assert.ok(helperStart >= 0, "raw MIME attachment helper start marker missing");
  assert.ok(helperEnd > helperStart, "raw MIME attachment helper end marker missing");
  assert.ok(getMessageStart >= 0, "getMessage missing");
  assert.ok(getMessageEnd > getMessageStart, "getMessage end missing");

  const sandbox = {
    Cr: { NS_BASE_STREAM_CLOSED },
    console: { error: (...args) => log.push(args) },
    findMessage() {
      if (findMessageError) throw findMessageError;
      return {
        msgHdr: {
          folder: {
            getMsgInputStream() { return stream; },
          },
          messageId: "message-1",
          mime2DecodedSubject: "subject",
        },
      };
    },
    NetUtil: {
      readInputStreamToString(inputStream, count) {
        return inputStream.read(count);
      },
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(helperStart, helperEnd)}
${source.slice(getMessageStart, getMessageEnd)}
this.getMessage = getMessage;`,
    sandbox
  );
  return sandbox.getMessage;
}

function xpcomError(result, message = String(result)) {
  const error = new Error(message);
  error.result = result;
  return error;
}

function makePipeStream(bytes, chunkSize = bytes.length || 1) {
  let cursor = 0;
  const readRequests = [];
  return {
    available() {
      if (cursor >= bytes.length) {
        throw xpcomError(NS_BASE_STREAM_CLOSED);
      }
      return bytes.length - cursor;
    },
    read(count) {
      readRequests.push(count);
      const length = Math.min(count, chunkSize, bytes.length - cursor);
      const chunk = bytes.slice(cursor, cursor + length);
      cursor += length;
      return chunk;
    },
    get bytesRead() { return cursor; },
    get readRequests() { return readRequests; },
  };
}

describe("readMessageStreamFully", () => {
  it("treats NS_BASE_STREAM_CLOSED after buffered pipe data as EOF", () => {
    const stream = makePipeStream("hello world");

    assert.equal(readMessageStreamFully(stream), "hello world");
  });

  it("treats NS_BASE_STREAM_CLOSED on an empty pipe as EOF", () => {
    const stream = makePipeStream("");

    assert.equal(readMessageStreamFully(stream), "");
  });

  it("also accepts NS_BASE_STREAM_CLOSED thrown as a direct result code", () => {
    const stream = {
      available() { throw NS_BASE_STREAM_CLOSED; },
    };

    assert.equal(readMessageStreamFully(stream), "");
  });

  it("propagates other available() failures", () => {
    const failure = xpcomError("NS_ERROR_UNEXPECTED", "unexpected stream failure");
    const stream = {
      available() { throw failure; },
    };

    assert.throws(() => readMessageStreamFully(stream), error => error === failure);
  });

  it("reads one byte past the actual limit before throwing a tagged error", () => {
    const stream = makePipeStream("X".repeat(2000));
    let error;

    try {
      readMessageStreamFully(stream, 1000);
    } catch (caught) {
      error = caught;
    }

    assert.ok(error, "expected a size-limit error");
    assert.equal(error.isStreamSizeLimit, true);
    assert.equal(stream.bytesRead, 1001);
    assert.deepEqual(stream.readRequests, [1001]);
  });

  it("accepts the exact byte limit when the pipe then closes", () => {
    const stream = makePipeStream("X".repeat(1000));

    assert.equal(readMessageStreamFully(stream, 1000).length, 1000);
  });

  it("bounds individual reads to a fixed chunk size", () => {
    const stream = makePipeStream("X".repeat(FIXED_CHUNK_SIZE * 2 + 17));

    assert.equal(
      readMessageStreamFully(stream).length,
      FIXED_CHUNK_SIZE * 2 + 17
    );
    assert.ok(stream.readRequests.every(count => count <= FIXED_CHUNK_SIZE));
  });

  it("throws instead of returning truncated data when a read makes no progress", () => {
    const stream = {
      available() { return 10; },
      read() { return ""; },
    };

    assert.throws(
      () => readMessageStreamFully(stream),
      /made no progress/
    );
  });

  it("does not silently truncate streams that need more than 1024 reads", () => {
    const stream = makePipeStream("a".repeat(1300), 1);

    assert.equal(readMessageStreamFully(stream).length, 1300);
  });
});

describe("inline image stream errors", () => {
  it("treats message-part channel contentLength as advisory", async () => {
    const stream = makePipeStream("small image bytes");
    const log = [];
    const fetchInlineImageBase64 = loadInlineImageFetcher({
      stream,
      contentLength: 2 * 1024 * 1024,
      log,
    });

    const result = await fetchInlineImageBase64({ url: "mailbox-message-part" });

    assert.equal(result.error, undefined);
    assert.equal(result.data, btoa("small image bytes"));
    assert.equal(log.length, 0);
  });

  it("logs non-limit stream errors and returns a stable non-sensitive message", async () => {
    const failure = xpcomError(
      "NS_ERROR_UNEXPECTED",
      "failed to read mailbox:///private/internal-message"
    );
    const stream = {
      available() { throw failure; },
      close() {},
    };
    const log = [];
    const fetchInlineImageBase64 = loadInlineImageFetcher({
      stream,
      contentLength: -1,
      log,
    });

    const result = await fetchInlineImageBase64({ url: "mailbox-message-part" });

    assert.equal(result.error, "Inline image read failed");
    assert.equal(log.length, 1);
    assert.equal(log[0].at(-1), failure);
    assert.doesNotMatch(result.error, /mailbox:/);
  });

  it("sanitizes synchronous newChannel failures", async () => {
    const internalPath = "/home/alice/.thunderbird/private/Inbox";
    const failure = new Error(`invalid mailbox URI mailbox://${internalPath}`);
    const log = [];
    const fetchInlineImageBase64 = loadInlineImageFetcher({
      log,
      newChannelError: failure,
    });

    const result = await fetchInlineImageBase64({ url: "mailbox-message-part" });

    assert.equal(result.error, "Inline image fetch failed");
    assert.equal(JSON.stringify(result).includes(internalPath), false);
    assert.equal(log.length, 1);
    assert.equal(log[0].at(-1), failure);
  });

  it("sanitizes synchronous asyncFetch failures", async () => {
    const internalUri = "mailbox:///home/alice/.thunderbird/private/Inbox";
    const failure = new Error(`Failed to open input source '${internalUri}'`);
    const log = [];
    const fetchInlineImageBase64 = loadInlineImageFetcher({
      log,
      asyncFetchError: failure,
    });

    const result = await fetchInlineImageBase64({ url: "mailbox-message-part" });

    assert.equal(result.error, "Inline image fetch failed");
    assert.equal(JSON.stringify(result).includes(internalUri), false);
    assert.equal(log.length, 1);
    assert.equal(log[0].at(-1), failure);
  });

  it("preserves the genuine size-limit error", async () => {
    const maxRawBytes = Math.floor((1 * 1024 * 1024) / 4) * 3;
    const stream = makePipeStream("X".repeat(maxRawBytes + 1));
    const log = [];
    const fetchInlineImageBase64 = loadInlineImageFetcher({
      stream,
      contentLength: -1,
      log,
    });

    const result = await fetchInlineImageBase64({ url: "mailbox-message-part" });

    assert.equal(
      result.error,
      "Image exceeds per-image base64 limit (1048576 bytes)"
    );
    assert.equal(log.length, 0);
  });
});

describe("getMessage error sanitization", () => {
  it("logs raw-source stream failures without returning internal paths", async () => {
    const internalPath = "/home/alice/.thunderbird/private/Inbox";
    const failure = xpcomError(
      "NS_ERROR_UNEXPECTED",
      `failed to read ${internalPath}`
    );
    const stream = {
      available() { throw failure; },
      close() {},
    };
    const log = [];
    const getMessage = loadGetMessage({ stream, log });

    const result = await getMessage(
      "message-1",
      "folder",
      false,
      "markdown",
      true,
      false
    );

    assert.equal(result.error, "Failed to read raw source");
    assert.equal(JSON.stringify(result).includes(internalPath), false);
    assert.equal(log.length, 1);
    assert.equal(log[0].at(-1), failure);
  });

  it("sanitizes the final getMessage exception boundary", async () => {
    const internalUri = "mailbox:///home/alice/.thunderbird/private/Inbox";
    const failure = new Error(`failed to resolve ${internalUri}`);
    const log = [];
    const getMessage = loadGetMessage({ log, findMessageError: failure });

    const result = await getMessage(
      "message-1",
      "folder",
      false,
      "markdown",
      true,
      false
    );

    assert.equal(result.error, "Failed to get message");
    assert.equal(JSON.stringify(result).includes(internalUri), false);
    assert.equal(log.length, 1);
    assert.equal(log[0].at(-1), failure);
  });
});
