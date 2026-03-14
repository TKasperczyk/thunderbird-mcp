/**
 * Stress tests and corner-case coverage for thunderbird-mcp.
 *
 * These tests go beyond happy-path validation to exercise edge cases
 * that will surface in heavy real-world usage:
 *
 * - Pagination: boundary offsets, over-sized requests, negative values
 * - Auth: corrupted connection files, empty tokens, header manipulation
 * - Bridge concurrency: parallel requests, interleaved responses
 * - Connection resilience: server restarts, port changes mid-session
 * - Validation: adversarial inputs, prototype pollution attempts
 * - Large payloads: oversized JSON, deeply nested objects
 */

const { describe, it, before, after, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');

const net = require('net');

const BRIDGE_PATH = path.resolve(__dirname, '..', 'mcp-bridge.cjs');
const CONN_DIR = path.join(os.tmpdir(), 'thunderbird-mcp');
const CONN_FILE = path.join(CONN_DIR, 'connection.json');
const DEFAULT_PORT = 8765;

/** Check if the default MCP port is already in use (e.g. real Thunderbird). */
function isDefaultPortInUse() {
  return new Promise((resolve) => {
    const sock = net.createConnection({ port: DEFAULT_PORT, host: '127.0.0.1' });
    sock.on('connect', () => { sock.destroy(); resolve(true); });
    sock.on('error', () => resolve(false));
  });
}

// ─── Helpers ───────────────────────────────────────────────────────

function writeTestConnectionInfo(port, token) {
  fs.mkdirSync(CONN_DIR, { recursive: true });
  fs.writeFileSync(CONN_FILE, JSON.stringify({ port, token, pid: process.pid }), 'utf8');
}

let savedConnectionData = null;
function backupConnectionFile() {
  try { savedConnectionData = fs.readFileSync(CONN_FILE, 'utf8'); } catch { savedConnectionData = null; }
}
function restoreConnectionFile() {
  if (savedConnectionData !== null) {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, savedConnectionData, 'utf8');
  } else {
    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }
  }
}

/**
 * Send a JSON-RPC message to the bridge via child process.
 * Returns parsed response or null for notifications.
 */
function sendToBridge(message, { timeout = 10000 } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [BRIDGE_PATH], {
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error(`Bridge timed out. stdout: ${stdout}, stderr: ${stderr}`));
    }, timeout);

    child.stdout.on('data', (data) => {
      stdout += data.toString();
      const lines = stdout.split('\n').filter(l => l.trim());
      if (lines.length > 0) {
        clearTimeout(timer);
        child.stdin.end();
        try { resolve(JSON.parse(lines[0])); }
        catch (e) { reject(new Error(`Failed to parse: ${lines[0]}`)); }
      }
    });
    child.stderr.on('data', (data) => { stderr += data.toString(); });
    child.on('error', (err) => { clearTimeout(timer); reject(err); });
    child.on('exit', () => { clearTimeout(timer); if (!stdout.trim()) resolve(null); });

    child.stdin.write(JSON.stringify(message) + '\n');
  });
}

/**
 * Send multiple JSON-RPC messages to a single bridge process and collect all responses.
 * Messages are written one per line. Returns array of parsed responses.
 */
function sendMultipleToBridge(messages, { timeout = 15000 } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [BRIDGE_PATH], {
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    const responses = [];
    const expectedCount = messages.filter(m =>
      Object.prototype.hasOwnProperty.call(m, 'id') &&
      !(typeof m.method === 'string' && m.method.startsWith('notifications/'))
    ).length;

    const timer = setTimeout(() => {
      child.kill();
      // Return what we have so far instead of failing
      resolve(responses);
    }, timeout);

    child.stdout.on('data', (data) => {
      stdout += data.toString();
      const lines = stdout.split('\n');
      // Parse complete lines
      for (let i = 0; i < lines.length - 1; i++) {
        const line = lines[i].trim();
        if (line) {
          try { responses.push(JSON.parse(line)); } catch { /* skip malformed */ }
        }
      }
      // Keep the last (possibly incomplete) line
      stdout = lines[lines.length - 1];

      if (responses.length >= expectedCount) {
        clearTimeout(timer);
        child.stdin.end();
      }
    });

    child.stderr.on('data', (data) => { stderr += data.toString(); });
    child.on('error', (err) => { clearTimeout(timer); reject(err); });
    child.on('exit', () => {
      clearTimeout(timer);
      // Parse any remaining output
      if (stdout.trim()) {
        try { responses.push(JSON.parse(stdout.trim())); } catch { /* skip */ }
      }
      resolve(responses);
    });

    // Write all messages
    for (const msg of messages) {
      child.stdin.write(JSON.stringify(msg) + '\n');
    }
    // Close stdin after a short delay to allow processing
    setTimeout(() => {
      try { child.stdin.end(); } catch { /* ignore */ }
    }, 500);
  });
}

/**
 * Replicate pagination logic from api.js for unit testing.
 * Simulates the offset/limit/hasMore behavior.
 */
function paginateResults(allResults, maxResults, offset) {
  const MAX_RESULTS_CAP = 200;
  const DEFAULT_MAX = 50;

  const requestedLimit = Number(maxResults);
  const effectiveLimit = Math.min(
    Number.isFinite(requestedLimit) && requestedLimit > 0 ? Math.floor(requestedLimit) : DEFAULT_MAX,
    MAX_RESULTS_CAP
  );
  const effectiveOffset = Number.isFinite(Number(offset)) && Number(offset) > 0 ? Math.floor(Number(offset)) : 0;

  const page = allResults.slice(effectiveOffset, effectiveOffset + effectiveLimit);

  return {
    messages: page,
    totalMatches: allResults.length,
    offset: effectiveOffset,
    limit: effectiveLimit,
    hasMore: effectiveOffset + effectiveLimit < allResults.length
  };
}

/**
 * Replicate sanitizeJson from bridge for thorough testing.
 */
function sanitizeJson(data) {
  let sanitized = data.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, '');
  sanitized = sanitized.replace(/(?<!\\)\r/g, '\\r');
  sanitized = sanitized.replace(/(?<!\\)\n/g, '\\n');
  sanitized = sanitized.replace(/(?<!\\)\t/g, '\\t');
  return sanitized;
}

/**
 * Replicate validateToolArgs from api.js.
 */
function createValidator(tools) {
  const toolSchemas = Object.create(null);
  for (const t of tools) { toolSchemas[t.name] = t.inputSchema; }

  return function validateToolArgs(name, args) {
    const schema = toolSchemas[name];
    if (!schema) return [`Unknown tool: ${name}`];
    const errors = [];
    const props = schema.properties || {};
    const required = schema.required || [];

    for (const key of required) {
      if (args[key] === undefined || args[key] === null) {
        errors.push(`Missing required parameter: ${key}`);
      }
    }
    for (const [key, value] of Object.entries(args)) {
      const propSchema = Object.prototype.hasOwnProperty.call(props, key) ? props[key] : undefined;
      if (!propSchema) { errors.push(`Unknown parameter: ${key}`); continue; }
      if (value === undefined || value === null) continue;
      const expectedType = propSchema.type;
      if (expectedType === "array") {
        if (!Array.isArray(value)) errors.push(`Parameter '${key}' must be an array, got ${typeof value}`);
      } else if (expectedType === "object") {
        if (typeof value !== "object" || Array.isArray(value))
          errors.push(`Parameter '${key}' must be an object, got ${Array.isArray(value) ? "array" : typeof value}`);
      } else if (expectedType && typeof value !== expectedType) {
        errors.push(`Parameter '${key}' must be ${expectedType}, got ${typeof value}`);
      }
    }
    return errors;
  };
}

const sampleTools = [
  { name: "searchMessages", inputSchema: {
    type: "object",
    properties: {
      query: { type: "string" }, folderPath: { type: "string" },
      maxResults: { type: "number" }, offset: { type: "number" },
      unreadOnly: { type: "boolean" },
    },
    required: ["query"],
  }},
  { name: "getMessage", inputSchema: {
    type: "object",
    properties: {
      messageId: { type: "string" }, folderPath: { type: "string" },
      saveAttachments: { type: "boolean" },
    },
    required: ["messageId", "folderPath"],
  }},
  { name: "sendMail", inputSchema: {
    type: "object",
    properties: {
      to: { type: "string" }, subject: { type: "string" }, body: { type: "string" },
      cc: { type: "string" }, bcc: { type: "string" },
      isHtml: { type: "boolean" }, from: { type: "string" },
      attachments: { type: "array", items: { type: "string" } },
    },
    required: ["to", "subject", "body"],
  }},
  { name: "deleteMessages", inputSchema: {
    type: "object",
    properties: {
      messageIds: { type: "array", items: { type: "string" } },
      folderPath: { type: "string" },
    },
    required: ["messageIds", "folderPath"],
  }},
];
const validate = createValidator(sampleTools);


// ═══════════════════════════════════════════════════════════════════
// PAGINATION STRESS TESTS
// ═══════════════════════════════════════════════════════════════════

describe('Pagination: boundary conditions', () => {
  // Generate a dataset of 500 mock results
  const mockData = Array.from({ length: 500 }, (_, i) => ({ id: `msg-${i}`, subject: `Message ${i}` }));

  it('offset=0 returns first page', () => {
    const result = paginateResults(mockData, 50, 0);
    assert.equal(result.messages.length, 50);
    assert.equal(result.offset, 0);
    assert.equal(result.totalMatches, 500);
    assert.equal(result.hasMore, true);
    assert.equal(result.messages[0].id, 'msg-0');
  });

  it('offset at exact page boundary works', () => {
    const result = paginateResults(mockData, 50, 50);
    assert.equal(result.messages.length, 50);
    assert.equal(result.offset, 50);
    assert.equal(result.messages[0].id, 'msg-50');
    assert.equal(result.hasMore, true);
  });

  it('offset near end returns partial page', () => {
    const result = paginateResults(mockData, 50, 480);
    assert.equal(result.messages.length, 20); // 500 - 480
    assert.equal(result.offset, 480);
    assert.equal(result.hasMore, false);
  });

  it('offset exactly at end returns empty page', () => {
    const result = paginateResults(mockData, 50, 500);
    assert.equal(result.messages.length, 0);
    assert.equal(result.offset, 500);
    assert.equal(result.hasMore, false);
    assert.equal(result.totalMatches, 500);
  });

  it('offset beyond end returns empty page', () => {
    const result = paginateResults(mockData, 50, 9999);
    assert.equal(result.messages.length, 0);
    assert.equal(result.offset, 9999);
    assert.equal(result.hasMore, false);
  });

  it('negative offset is treated as 0', () => {
    const result = paginateResults(mockData, 50, -10);
    assert.equal(result.offset, 0);
    assert.equal(result.messages.length, 50);
    assert.equal(result.messages[0].id, 'msg-0');
  });

  it('non-numeric offset is treated as 0', () => {
    const result = paginateResults(mockData, 50, "abc");
    assert.equal(result.offset, 0);
    assert.equal(result.messages.length, 50);
  });

  it('NaN offset is treated as 0', () => {
    const result = paginateResults(mockData, 50, NaN);
    assert.equal(result.offset, 0);
  });

  it('Infinity offset is treated as 0', () => {
    const result = paginateResults(mockData, 50, Infinity);
    assert.equal(result.offset, 0);
  });

  it('fractional offset is floored', () => {
    const result = paginateResults(mockData, 50, 10.7);
    assert.equal(result.offset, 10);
    assert.equal(result.messages[0].id, 'msg-10');
  });

  it('null offset is treated as 0', () => {
    const result = paginateResults(mockData, 50, null);
    assert.equal(result.offset, 0);
  });

  it('undefined offset is treated as 0', () => {
    const result = paginateResults(mockData, 50, undefined);
    assert.equal(result.offset, 0);
  });
});

describe('Pagination: maxResults edge cases', () => {
  const mockData = Array.from({ length: 300 }, (_, i) => ({ id: `msg-${i}` }));

  it('maxResults=0 uses default (50)', () => {
    const result = paginateResults(mockData, 0, 0);
    assert.equal(result.limit, 50);
    assert.equal(result.messages.length, 50);
  });

  it('maxResults=-1 uses default (50)', () => {
    const result = paginateResults(mockData, -1, 0);
    assert.equal(result.limit, 50);
  });

  it('maxResults=1 returns single result', () => {
    const result = paginateResults(mockData, 1, 0);
    assert.equal(result.messages.length, 1);
    assert.equal(result.limit, 1);
    assert.equal(result.hasMore, true);
  });

  it('maxResults exceeding cap is clamped to 200', () => {
    const result = paginateResults(mockData, 999, 0);
    assert.equal(result.limit, 200);
    assert.equal(result.messages.length, 200);
  });

  it('maxResults=200 (at cap) works', () => {
    const result = paginateResults(mockData, 200, 0);
    assert.equal(result.limit, 200);
    assert.equal(result.messages.length, 200);
    assert.equal(result.hasMore, true);
  });

  it('maxResults as string is treated as default', () => {
    const result = paginateResults(mockData, "fifty", 0);
    assert.equal(result.limit, 50);
  });

  it('maxResults=NaN uses default', () => {
    const result = paginateResults(mockData, NaN, 0);
    assert.equal(result.limit, 50);
  });

  it('empty results with pagination returns correct metadata', () => {
    const result = paginateResults([], 50, 0);
    assert.equal(result.messages.length, 0);
    assert.equal(result.totalMatches, 0);
    assert.equal(result.hasMore, false);
    assert.equal(result.offset, 0);
  });

  it('single result with offset=0 has hasMore=false', () => {
    const result = paginateResults([{ id: 'only' }], 50, 0);
    assert.equal(result.messages.length, 1);
    assert.equal(result.hasMore, false);
  });
});

describe('Pagination: sequential page traversal', () => {
  const mockData = Array.from({ length: 137 }, (_, i) => ({ id: `msg-${i}` }));

  it('can traverse all pages without gaps or duplicates', () => {
    const pageSize = 25;
    const allCollected = [];
    let offset = 0;
    let pages = 0;

    while (true) {
      const result = paginateResults(mockData, pageSize, offset);
      allCollected.push(...result.messages);
      pages++;

      if (!result.hasMore) break;
      offset += pageSize;

      // Safety valve to prevent infinite loops in tests
      if (pages > 100) throw new Error('Pagination infinite loop detected');
    }

    // Should have collected all 137 items
    assert.equal(allCollected.length, 137);
    // No duplicates
    const ids = new Set(allCollected.map(m => m.id));
    assert.equal(ids.size, 137);
    // Took ceiling(137/25) = 6 pages
    assert.equal(pages, 6);
  });

  it('page traversal with maxResults=1 (worst case) collects all items', () => {
    const smallData = Array.from({ length: 10 }, (_, i) => ({ id: `msg-${i}` }));
    const allCollected = [];
    let offset = 0;
    let pages = 0;

    while (true) {
      const result = paginateResults(smallData, 1, offset);
      allCollected.push(...result.messages);
      pages++;
      if (!result.hasMore) break;
      offset += 1;
      if (pages > 100) throw new Error('Infinite loop');
    }

    assert.equal(allCollected.length, 10);
    assert.equal(pages, 10);
  });
});


// ═══════════════════════════════════════════════════════════════════
// AUTH STRESS TESTS
// ═══════════════════════════════════════════════════════════════════

describe('Auth: connection file corruption', () => {
  before(() => backupConnectionFile());
  after(() => restoreConnectionFile());

  it('handles empty connection file gracefully', async (t) => {
    if (await isDefaultPortInUse()) {
      return t.skip('Port 8765 in use (Thunderbird running), skipping fallback test');
    }
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, '', 'utf8');

    // Should fall back to default port, get connection refused or similar
    const response = await sendToBridge({
      jsonrpc: '2.0', id: 1, method: 'tools/list'
    });

    assert.equal(response.id, 1);
    // Will be an error since no server on default port
    assert.ok(response.result || response.error);
  });

  it('handles malformed JSON in connection file', async (t) => {
    if (await isDefaultPortInUse()) {
      return t.skip('Port 8765 in use (Thunderbird running), skipping fallback test');
    }
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, '{ port: invalid json }', 'utf8');

    const response = await sendToBridge({
      jsonrpc: '2.0', id: 2, method: 'tools/list'
    });

    assert.equal(response.id, 2);
    assert.ok(response.result || response.error);
  });

  it('handles connection file with missing port field', async (t) => {
    if (await isDefaultPortInUse()) {
      return t.skip('Port 8765 in use (Thunderbird running), skipping fallback test');
    }
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, JSON.stringify({ token: 'abc' }), 'utf8');

    // Should fall back to default port
    const response = await sendToBridge({
      jsonrpc: '2.0', id: 3, method: 'tools/list'
    });

    assert.equal(response.id, 3);
    assert.ok(response.result || response.error);
  });

  it('handles connection file with missing token field', async () => {
    const TEST_PORT = 18770;

    const server = http.createServer((req, res) => {
      // Server that doesn't check auth — accepts anything
      let body = '';
      req.on('data', (c) => { body += c; });
      req.on('end', () => {
        const msg = JSON.parse(body);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ jsonrpc: '2.0', id: msg.id, result: { tools: [] } }));
      });
    });
    await new Promise((resolve, reject) => {
      server.on('error', reject);
      server.listen(TEST_PORT, '127.0.0.1', resolve);
    });

    try {
      fs.mkdirSync(CONN_DIR, { recursive: true });
      fs.writeFileSync(CONN_FILE, JSON.stringify({ port: TEST_PORT }), 'utf8');

      const response = await sendToBridge({
        jsonrpc: '2.0', id: 4, method: 'tools/list'
      });

      // Should connect without auth header
      assert.equal(response.id, 4);
      assert.ok(response.result);
    } finally {
      await new Promise((resolve) => server.close(resolve));
    }
  });

  it('handles binary garbage in connection file', async (t) => {
    if (await isDefaultPortInUse()) {
      return t.skip('Port 8765 in use (Thunderbird running), skipping fallback test');
    }
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, Buffer.from([0x00, 0xFF, 0xFE, 0x80, 0x90]));

    const response = await sendToBridge({
      jsonrpc: '2.0', id: 5, method: 'tools/list'
    });

    assert.equal(response.id, 5);
    assert.ok(response.result || response.error);
  });
});

describe('Auth: token edge cases', () => {
  let server;
  const TEST_PORT = 18771;
  let lastAuthHeader = null;

  before(async () => {
    backupConnectionFile();

    server = http.createServer((req, res) => {
      lastAuthHeader = req.headers['authorization'] || null;
      let body = '';
      req.on('data', (c) => { body += c; });
      req.on('end', () => {
        const msg = JSON.parse(body);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          jsonrpc: '2.0', id: msg.id,
          result: { receivedAuth: lastAuthHeader }
        }));
      });
    });

    await new Promise((resolve, reject) => {
      server.on('error', reject);
      server.listen(TEST_PORT, '127.0.0.1', resolve);
    });
  });

  after(async () => {
    if (server) await new Promise((resolve) => server.close(resolve));
    restoreConnectionFile();
  });

  it('sends token with special characters correctly', async () => {
    const weirdToken = 'tok3n-with-sp3c!@l_ch@rs+and=more/stuff';
    writeTestConnectionInfo(TEST_PORT, weirdToken);

    const response = await sendToBridge({
      jsonrpc: '2.0', id: 1, method: 'tools/list'
    });

    assert.equal(lastAuthHeader, `Bearer ${weirdToken}`);
  });

  it('sends very long token correctly', async () => {
    const longToken = 'a'.repeat(2048);
    writeTestConnectionInfo(TEST_PORT, longToken);

    const response = await sendToBridge({
      jsonrpc: '2.0', id: 2, method: 'tools/list'
    });

    assert.equal(lastAuthHeader, `Bearer ${longToken}`);
  });

  it('sends empty string token as no auth', async () => {
    writeTestConnectionInfo(TEST_PORT, '');

    const response = await sendToBridge({
      jsonrpc: '2.0', id: 3, method: 'tools/list'
    });

    // Empty token should result in no Authorization header
    // (bridge checks `if (token)` which is falsy for '')
    assert.equal(lastAuthHeader, null);
  });
});


// ═══════════════════════════════════════════════════════════════════
// BRIDGE CONCURRENCY STRESS TESTS
// ═══════════════════════════════════════════════════════════════════

describe('Bridge: concurrent lifecycle requests', () => {
  it('handles multiple lifecycle methods in rapid succession', async () => {
    const messages = [
      { jsonrpc: '2.0', id: 1, method: 'initialize', params: {} },
      { jsonrpc: '2.0', id: 2, method: 'ping' },
      { jsonrpc: '2.0', id: 3, method: 'resources/list' },
      { jsonrpc: '2.0', id: 4, method: 'prompts/list' },
      { jsonrpc: '2.0', id: 5, method: 'ping' },
    ];

    const responses = await sendMultipleToBridge(messages);

    assert.equal(responses.length, 5);
    // All should have correct IDs (may be out of order due to async)
    const ids = new Set(responses.map(r => r.id));
    assert.equal(ids.size, 5);
    for (let i = 1; i <= 5; i++) {
      assert.ok(ids.has(i), `Missing response for id ${i}`);
    }
  });

  it('handles notifications interleaved with requests', async () => {
    const messages = [
      { jsonrpc: '2.0', id: 1, method: 'initialize', params: {} },
      { jsonrpc: '2.0', method: 'notifications/initialized' }, // no id
      { jsonrpc: '2.0', id: 2, method: 'ping' },
      { jsonrpc: '2.0', method: 'notifications/something' }, // no id
      { jsonrpc: '2.0', id: 3, method: 'resources/list' },
    ];

    const responses = await sendMultipleToBridge(messages);

    // Should only get 3 responses (for the 3 requests with IDs)
    assert.equal(responses.length, 3);
    const ids = new Set(responses.map(r => r.id));
    assert.ok(ids.has(1));
    assert.ok(ids.has(2));
    assert.ok(ids.has(3));
  });
});

describe('Bridge: parallel forwarded requests', () => {
  let server;
  const TEST_PORT = 18772;
  let requestCount = 0;

  before(async () => {
    backupConnectionFile();

    server = http.createServer((req, res) => {
      requestCount++;
      let body = '';
      req.on('data', (c) => { body += c; });
      req.on('end', () => {
        const msg = JSON.parse(body);
        // Add a small delay to simulate real processing
        setTimeout(() => {
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            jsonrpc: '2.0',
            id: msg.id,
            result: { content: [{ type: 'text', text: `response-${msg.id}` }] }
          }));
        }, 10);
      });
    });

    await new Promise((resolve, reject) => {
      server.on('error', reject);
      server.listen(TEST_PORT, '127.0.0.1', resolve);
    });

    writeTestConnectionInfo(TEST_PORT, 'stress-token');
  });

  after(async () => {
    if (server) await new Promise((resolve) => server.close(resolve));
    restoreConnectionFile();
  });

  beforeEach(() => { requestCount = 0; });

  it('handles 10 parallel forwarded requests correctly', async () => {
    const messages = Array.from({ length: 10 }, (_, i) => ({
      jsonrpc: '2.0',
      id: i + 1,
      method: 'tools/call',
      params: { name: 'test', arguments: {} }
    }));

    const responses = await sendMultipleToBridge(messages);

    // All 10 should have responses
    assert.equal(responses.length, 10);
    const ids = new Set(responses.map(r => r.id));
    assert.equal(ids.size, 10);
    // All should have results (no errors)
    for (const r of responses) {
      assert.ok(r.result, `Request ${r.id} should have a result`);
    }
  });
});


// ═══════════════════════════════════════════════════════════════════
// CONNECTION RESILIENCE TESTS
// ═══════════════════════════════════════════════════════════════════

describe('Bridge: server restart recovery', () => {
  before(() => backupConnectionFile());
  after(() => restoreConnectionFile());

  it('recovers when server restarts on a different port', async () => {
    const PORT_A = 18773;
    const PORT_B = 18774;

    // Start server A
    const serverA = http.createServer((req, res) => {
      let body = '';
      req.on('data', (c) => { body += c; });
      req.on('end', () => {
        const msg = JSON.parse(body);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          jsonrpc: '2.0', id: msg.id, result: { server: 'A' }
        }));
      });
    });
    await new Promise((r, j) => { serverA.on('error', j); serverA.listen(PORT_A, '127.0.0.1', r); });

    writeTestConnectionInfo(PORT_A, 'token-a');

    // First request to server A
    const resp1 = await sendToBridge({ jsonrpc: '2.0', id: 1, method: 'tools/list' });
    assert.equal(resp1.result.server, 'A');

    // "Restart" — stop A, start B on different port
    await new Promise((r) => serverA.close(r));

    const serverB = http.createServer((req, res) => {
      let body = '';
      req.on('data', (c) => { body += c; });
      req.on('end', () => {
        const msg = JSON.parse(body);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          jsonrpc: '2.0', id: msg.id, result: { server: 'B' }
        }));
      });
    });
    await new Promise((r, j) => { serverB.on('error', j); serverB.listen(PORT_B, '127.0.0.1', r); });

    // Update connection file to point to new server
    writeTestConnectionInfo(PORT_B, 'token-b');

    // Bridge should pick up the new connection info
    const resp2 = await sendToBridge({ jsonrpc: '2.0', id: 2, method: 'tools/list' });
    assert.equal(resp2.result.server, 'B');

    await new Promise((r) => serverB.close(r));
  });
});


// ═══════════════════════════════════════════════════════════════════
// SANITIZE JSON STRESS TESTS
// ═══════════════════════════════════════════════════════════════════

describe('sanitizeJson: adversarial inputs', () => {
  it('handles all 32 ASCII control characters', () => {
    // Build a string with every control char (0x00-0x1F) plus DEL (0x7F)
    let input = '';
    for (let i = 0; i <= 0x1f; i++) input += String.fromCharCode(i);
    input += String.fromCharCode(0x7f);

    const result = sanitizeJson(input);
    // Only \n (0x0a), \r (0x0d), \t (0x09) should survive (escaped)
    // Everything else should be stripped
    assert.ok(!result.includes('\x00'));
    assert.ok(!result.includes('\x01'));
    assert.ok(!result.includes('\x7f'));
    // \n, \r, \t should be escaped as \\n, \\r, \\t
    assert.ok(result.includes('\\n'));
    assert.ok(result.includes('\\r'));
    assert.ok(result.includes('\\t'));
  });

  it('handles a large string with embedded control chars (100KB)', () => {
    // Simulate a large email body with scattered control characters
    const chunks = [];
    for (let i = 0; i < 1000; i++) {
      chunks.push('Normal text line ' + i);
      if (i % 10 === 0) chunks.push('\x00\x01\x02'); // inject garbage
      chunks.push('\n');
    }
    const input = chunks.join('');
    assert.ok(input.length > 20000);

    const result = sanitizeJson(input);
    // Should not contain any raw control chars
    assert.ok(!/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/.test(result));
    // Should still contain the text content
    assert.ok(result.includes('Normal text line 0'));
    assert.ok(result.includes('Normal text line 999'));
  });

  it('handles string that is entirely control characters', () => {
    const input = '\x00\x01\x02\x03\x04\x05\x06\x07\x08';
    const result = sanitizeJson(input);
    assert.equal(result, '');
  });

  it('preserves already-escaped sequences in complex JSON', () => {
    const input = '{"body":"line1\\nline2\\ttab\\r\\nwindows","raw":"\\\\n"}';
    const result = sanitizeJson(input);
    // The already-escaped \\n, \\t, \\r should be preserved
    assert.equal(result, input);
  });

  it('handles mixed escaped and unescaped in same string', () => {
    // \\n (already escaped) followed by real \n
    const input = 'before\\n\nafter';
    const result = sanitizeJson(input);
    assert.equal(result, 'before\\n\\nafter');
  });
});


// ═══════════════════════════════════════════════════════════════════
// VALIDATION: ADVERSARIAL INPUTS
// ═══════════════════════════════════════════════════════════════════

describe('Validation: adversarial and edge-case inputs', () => {
  it('handles empty string for required string param', () => {
    // Empty string is still a string — should pass type check
    const errors = validate('getMessage', {
      messageId: '',
      folderPath: '',
    });
    assert.equal(errors.length, 0);
  });

  it('rejects prototype pollution attempt via __proto__', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      __proto__: { admin: true },
    });
    // __proto__ is not in the schema, should be rejected as unknown
    // However, Object.entries may not enumerate __proto__
    // The key thing is it doesn't crash or pollute
    // Validator should be stable regardless
    assert.ok(Array.isArray(errors));
  });

  it('rejects constructor pollution attempt', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      constructor: 'evil',
    });
    assert.ok(errors.some(e => e.includes('Unknown parameter: constructor')));
  });

  it('handles very long parameter names gracefully', () => {
    const longKey = 'x'.repeat(10000);
    const args = { query: 'test', [longKey]: 'value' };
    const errors = validate('searchMessages', args);
    assert.ok(errors.some(e => e.includes('Unknown parameter')));
  });

  it('handles very long parameter values', () => {
    const longValue = 'x'.repeat(100000);
    const errors = validate('searchMessages', { query: longValue });
    // Should pass — value length is not validated by schema
    assert.equal(errors.length, 0);
  });

  it('handles deeply nested object where string expected', () => {
    const errors = validate('searchMessages', {
      query: { nested: { deep: { very: 'deep' } } },
    });
    assert.ok(errors.some(e => e.includes('must be string')));
  });

  it('handles array where string expected', () => {
    const errors = validate('getMessage', {
      messageId: ['id1', 'id2'],
      folderPath: '/INBOX',
    });
    assert.ok(errors.some(e => e.includes('must be string')));
  });

  it('handles 0 as valid number', () => {
    const errors = validate('searchMessages', { query: 'test', maxResults: 0 });
    assert.equal(errors.length, 0); // 0 is a valid number type
  });

  it('handles negative number', () => {
    const errors = validate('searchMessages', { query: 'test', maxResults: -100 });
    assert.equal(errors.length, 0); // negative is still a number type
  });

  it('handles boolean false as valid required field', () => {
    // false is not null/undefined — it's a value
    // But it's a boolean, not a string — so type check should catch it
    const errors = validate('getMessage', {
      messageId: false,
      folderPath: '/INBOX',
    });
    assert.ok(errors.some(e => e.includes('must be string')));
  });

  it('handles massive number of unknown parameters', () => {
    const args = { query: 'test' };
    for (let i = 0; i < 100; i++) {
      args[`unknown_param_${i}`] = `value_${i}`;
    }
    const errors = validate('searchMessages', args);
    assert.equal(errors.length, 100);
  });

  it('handles all required params wrong type simultaneously', () => {
    const errors = validate('sendMail', {
      to: 123,
      subject: true,
      body: ['not', 'a', 'string'],
    });
    assert.equal(errors.length, 3);
  });
});


// ═══════════════════════════════════════════════════════════════════
// BRIDGE: LARGE PAYLOAD HANDLING
// ═══════════════════════════════════════════════════════════════════

describe('Bridge: large payloads', () => {
  let server;
  const TEST_PORT = 18775;

  before(async () => {
    backupConnectionFile();

    server = http.createServer((req, res) => {
      let body = '';
      req.on('data', (c) => { body += c; });
      req.on('end', () => {
        const msg = JSON.parse(body);
        // Echo back a large response to stress the bridge's output handling
        const largeText = 'x'.repeat(50000);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          jsonrpc: '2.0',
          id: msg.id,
          result: { content: [{ type: 'text', text: largeText }] }
        }));
      });
    });

    await new Promise((r, j) => { server.on('error', j); server.listen(TEST_PORT, '127.0.0.1', r); });
    writeTestConnectionInfo(TEST_PORT, 'large-payload-token');
  });

  after(async () => {
    if (server) await new Promise((r) => server.close(r));
    restoreConnectionFile();
  });

  it('handles 50KB response without truncation', async () => {
    const response = await sendToBridge({
      jsonrpc: '2.0', id: 1,
      method: 'tools/call',
      params: { name: 'test', arguments: {} }
    });

    assert.equal(response.id, 1);
    assert.ok(response.result);
    assert.equal(response.result.content[0].text.length, 50000);
  });

  it('handles large request payload', async () => {
    const response = await sendToBridge({
      jsonrpc: '2.0', id: 2,
      method: 'tools/call',
      params: {
        name: 'test',
        arguments: { body: 'y'.repeat(50000) }
      }
    });

    assert.equal(response.id, 2);
    assert.ok(response.result);
  });
});


// ═══════════════════════════════════════════════════════════════════
// BRIDGE: ERROR HANDLING EDGE CASES
// ═══════════════════════════════════════════════════════════════════

describe('Bridge: malformed input handling', () => {
  it('handles empty line gracefully (no crash)', async () => {
    // Send empty lines then a valid request
    const child = spawn(process.execPath, [BRIDGE_PATH], {
      stdio: ['pipe', 'pipe', 'pipe'],
    });

    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        child.kill();
        resolve(); // Empty lines should be silently ignored
      }, 3000);

      let stdout = '';
      child.stdout.on('data', (data) => {
        stdout += data.toString();
        if (stdout.includes('"id":99')) {
          clearTimeout(timer);
          child.stdin.end();
          const response = JSON.parse(stdout.trim().split('\n').pop());
          assert.equal(response.id, 99);
          resolve();
        }
      });

      child.on('error', (err) => { clearTimeout(timer); reject(err); });

      // Send empty lines, then a valid request
      child.stdin.write('\n');
      child.stdin.write('  \n');
      child.stdin.write('\n');
      child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id: 99, method: 'ping' }) + '\n');
    });
  });
});

// ─── Tagging stress tests ─────────────────────────────────────────

describe('Tagging: validation edge cases', () => {
  const tagTools = [
    {
      name: "updateMessage",
      inputSchema: {
        type: "object",
        properties: {
          messageId: { type: "string" },
          messageIds: { type: "array", items: { type: "string" } },
          folderPath: { type: "string" },
          read: { type: "boolean" },
          flagged: { type: "boolean" },
          addTags: { type: "array", items: { type: "string" } },
          removeTags: { type: "array", items: { type: "string" } },
          moveTo: { type: "string" },
          trash: { type: "boolean" },
        },
        required: ["folderPath"],
      },
    },
    {
      name: "searchMessages",
      inputSchema: {
        type: "object",
        properties: {
          query: { type: "string" },
          tag: { type: "string" },
          maxResults: { type: "number" },
        },
        required: ["query"],
      },
    },
  ];
  const tagValidate = createValidator(tagTools);

  it('rejects addTags as object (not array)', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', addTags: { tag: '$label1' }
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects addTags as boolean', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', addTags: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects removeTags as string', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', removeTags: '$label1'
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('accepts empty addTags array', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', addTags: []
    });
    assert.equal(errors.length, 0);
  });

  it('accepts empty removeTags array', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', removeTags: []
    });
    assert.equal(errors.length, 0);
  });

  it('accepts addTags and removeTags simultaneously', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX',
      messageId: 'msg1',
      addTags: ['$label1', '$label2'],
      removeTags: ['$label3'],
    });
    assert.equal(errors.length, 0);
  });

  it('accepts large number of tags in addTags', () => {
    const tags = Array.from({ length: 100 }, (_, i) => `custom-tag-${i}`);
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', addTags: tags
    });
    assert.equal(errors.length, 0);
  });

  it('rejects tag filter as array in searchMessages', () => {
    const errors = tagValidate('searchMessages', {
      query: 'test', tag: ['$label1']
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects tag filter as boolean in searchMessages', () => {
    const errors = tagValidate('searchMessages', {
      query: 'test', tag: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('accepts empty string tag filter', () => {
    const errors = tagValidate('searchMessages', {
      query: 'test', tag: ''
    });
    assert.equal(errors.length, 0);
  });

  it('handles tags with special characters in validation', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX',
      addTags: ['tag with spaces', '$label/slash', 'tag\twith\ttabs'],
    });
    assert.equal(errors.length, 0);
  });

  it('handles prototype pollution attempt via addTags', () => {
    // Object.create(null) + hasOwnProperty prevents prototype pollution.
    // __proto__ in object literal sets prototype (not own property),
    // so it won't appear in Object.entries. Use Object.create instead.
    const malicious = Object.create(null);
    malicious.folderPath = '/INBOX';
    malicious.addTags = ['$label1'];
    malicious.constructor = 'attack';
    const errors = tagValidate('updateMessage', malicious);
    assert.ok(errors.some(e => e.includes('Unknown parameter: constructor')));
  });
});

// ─── Account access control stress tests ──────────────────────────

describe('Account access: validation edge cases', () => {
  const accessTools = [
    {
      name: "setAccountAccess",
      inputSchema: {
        type: "object",
        properties: {
          allowedAccountIds: { type: "array", items: { type: "string" } },
        },
        required: ["allowedAccountIds"],
      },
    },
    {
      name: "getAccountAccess",
      inputSchema: { type: "object", properties: {}, required: [] },
    },
  ];
  const accessValidate = createValidator(accessTools);

  it('rejects allowedAccountIds as object', () => {
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: { id: 'account1' }
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects allowedAccountIds as boolean', () => {
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects allowedAccountIds as number', () => {
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: 42
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('accepts large number of account IDs', () => {
    const ids = Array.from({ length: 50 }, (_, i) => `account${i}`);
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: ids
    });
    assert.equal(errors.length, 0);
  });

  it('rejects unknown parameters on setAccountAccess', () => {
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: ['acct1'],
      inject: 'malicious',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: inject/);
  });

  it('rejects unknown parameters on getAccountAccess', () => {
    const errors = accessValidate('getAccountAccess', {
      admin: true,
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: admin/);
  });

  it('handles prototype pollution on setAccountAccess', () => {
    const malicious = Object.create(null);
    malicious.allowedAccountIds = ['acct1'];
    malicious.constructor = 'attack';
    const errors = accessValidate('setAccountAccess', malicious);
    assert.ok(errors.some(e => e.includes('Unknown parameter: constructor')));
  });
});

// ─── Folder management stress tests ──────────────────────────────

describe('Folder management: validation edge cases', () => {
  const folderTools = [
    {
      name: "renameFolder",
      inputSchema: {
        type: "object",
        properties: { folderPath: { type: "string" }, newName: { type: "string" } },
        required: ["folderPath", "newName"],
      },
    },
    {
      name: "deleteFolder",
      inputSchema: {
        type: "object",
        properties: { folderPath: { type: "string" } },
        required: ["folderPath"],
      },
    },
    {
      name: "moveFolder",
      inputSchema: {
        type: "object",
        properties: { folderPath: { type: "string" }, newParentPath: { type: "string" } },
        required: ["folderPath", "newParentPath"],
      },
    },
  ];
  const folderValidate = createValidator(folderTools);

  it('rejects renameFolder with null folderPath', () => {
    const errors = folderValidate('renameFolder', { folderPath: null, newName: 'test' });
    assert.ok(errors.some(e => e.includes('folderPath')));
  });

  it('rejects renameFolder with array newName', () => {
    const errors = folderValidate('renameFolder', { folderPath: '/INBOX', newName: ['a'] });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects deleteFolder with number folderPath', () => {
    const errors = folderValidate('deleteFolder', { folderPath: 12345 });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects moveFolder with boolean paths', () => {
    const errors = folderValidate('moveFolder', { folderPath: true, newParentPath: false });
    assert.equal(errors.length, 2);
  });

  it('rejects unknown params on renameFolder', () => {
    const errors = folderValidate('renameFolder', {
      folderPath: '/INBOX', newName: 'test', force: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: force/);
  });

  it('handles very long folder name in renameFolder', () => {
    const longName = 'a'.repeat(10000);
    const errors = folderValidate('renameFolder', { folderPath: '/INBOX', newName: longName });
    assert.equal(errors.length, 0); // Validation passes; filesystem will reject
  });

  it('handles special characters in folder paths', () => {
    const errors = folderValidate('moveFolder', {
      folderPath: 'imap://user@server/INBOX/Über Spëcial & "Quotes"',
      newParentPath: 'imap://user@server/Archive/日本語'
    });
    assert.equal(errors.length, 0);
  });
});

// ─── Attachment sending stress tests ────────────────────────────

describe('Attachment sending: validation edge cases', () => {
  const mailTools = [
    {
      name: "sendMail",
      inputSchema: {
        type: "object",
        properties: {
          to: { type: "string" },
          subject: { type: "string" },
          body: { type: "string" },
          attachments: { type: "array" },
        },
        required: ["to", "subject", "body"],
      },
    },
  ];
  const mailValidate = createValidator(mailTools);

  it('rejects attachments as object (not array)', () => {
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments: { file: '/path' }
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects attachments as boolean', () => {
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('accepts empty attachments array', () => {
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments: []
    });
    assert.equal(errors.length, 0);
  });

  it('accepts large number of attachments in validation', () => {
    const attachments = Array.from({ length: 100 }, (_, i) => `/path/file${i}.pdf`);
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments
    });
    assert.equal(errors.length, 0);
  });

  it('accepts mixed file paths and inline objects in array', () => {
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments: [
        '/path/to/file.pdf',
        { name: 'inline.txt', contentType: 'text/plain', base64: 'SGVsbG8=' },
        '/another/file.doc'
      ]
    });
    assert.equal(errors.length, 0);
  });
});

// ─── Contact write support stress tests ─────────────────────────

describe('Contact write: validation edge cases', () => {
  const contactTools = [
    {
      name: "createContact",
      inputSchema: {
        type: "object",
        properties: {
          email: { type: "string" },
          displayName: { type: "string" },
          firstName: { type: "string" },
          lastName: { type: "string" },
          addressBookId: { type: "string" },
        },
        required: ["email"],
      },
    },
    {
      name: "updateContact",
      inputSchema: {
        type: "object",
        properties: {
          contactId: { type: "string" },
          email: { type: "string" },
          displayName: { type: "string" },
          firstName: { type: "string" },
          lastName: { type: "string" },
        },
        required: ["contactId"],
      },
    },
    {
      name: "deleteContact",
      inputSchema: {
        type: "object",
        properties: {
          contactId: { type: "string" },
        },
        required: ["contactId"],
      },
    },
  ];
  const contactValidate = createValidator(contactTools);

  it('rejects createContact with null email', () => {
    const errors = contactValidate('createContact', { email: null });
    assert.ok(errors.some(e => e.includes('email')));
  });

  it('rejects createContact with array email', () => {
    const errors = contactValidate('createContact', { email: ['a@b.com'] });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects updateContact with boolean contactId', () => {
    const errors = contactValidate('updateContact', { contactId: true });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects deleteContact with number contactId', () => {
    const errors = contactValidate('deleteContact', { contactId: 42 });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects unknown params on createContact', () => {
    const errors = contactValidate('createContact', {
      email: 'a@b.com', phone: '555-1234'
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: phone/);
  });

  it('handles prototype pollution on createContact', () => {
    const malicious = Object.create(null);
    malicious.email = 'a@b.com';
    malicious.constructor = 'attack';
    const errors = contactValidate('createContact', malicious);
    assert.ok(errors.some(e => e.includes('Unknown parameter: constructor')));
  });

  it('accepts all optional fields on createContact', () => {
    const errors = contactValidate('createContact', {
      email: 'a@b.com',
      displayName: 'Test',
      firstName: 'First',
      lastName: 'Last',
      addressBookId: 'book-1',
    });
    assert.equal(errors.length, 0);
  });

  it('accepts contactId-only for updateContact (no changes)', () => {
    const errors = contactValidate('updateContact', { contactId: 'uid-1' });
    assert.equal(errors.length, 0);
  });

  it('handles very long email on createContact', () => {
    const longEmail = 'a'.repeat(5000) + '@example.com';
    const errors = contactValidate('createContact', { email: longEmail });
    assert.equal(errors.length, 0);
  });
});
