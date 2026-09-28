/**
 * Tests for auth token and connection info discovery.
 *
 * Covers:
 * - Connection info file reading and caching
 * - Auth token inclusion in HTTP requests
 * - Fail-closed behavior when connection file is missing
 * - Cache invalidation on connection errors
 * - Bridge behavior with and without auth
 * - Timing-safe comparison correctness
 * - Connection file corruption/edge cases
 * - Bridge retry behavior
 */

const { describe, it, before, after, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');
const { isSensitiveFilePath } = require('./helpers/bridge.cjs');

const BRIDGE_PATH = path.resolve(__dirname, '..', 'mcp-bridge.cjs');
let CONN_DIR;
let CONN_FILE;
let BRIDGE_ENV;

beforeEach(() => {
  // Windows temp normally lives in denied AppData; outgoing attachment
  // fixtures need a benign path. Every directory is unique and removed below.
  const base = isSensitiveFilePath(os.tmpdir()) ? path.resolve(__dirname, '..') : os.tmpdir();
  CONN_DIR = fs.mkdtempSync(path.join(base, 'tb-mcp-auth-'));
  CONN_FILE = path.join(CONN_DIR, 'connection.json');
  // A missing/invalid pin must never fall through to a running instance.
  BRIDGE_ENV = { ...process.env, THUNDERBIRD_MCP_CONNECTION_FILE: CONN_FILE };
});

afterEach(() => {
  fs.rmSync(CONN_DIR, { recursive: true, force: true });
});

/**
 * Helper: send a JSON-RPC message to the bridge and get the response.
 */
function sendToBridge(message, { timeout = 10000 } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [BRIDGE_PATH], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: BRIDGE_ENV,
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
        try {
          resolve(JSON.parse(lines[0]));
        } catch (e) {
          reject(new Error(`Failed to parse: ${lines[0]}`));
        }
      }
    });

    child.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    child.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });

    child.on('exit', () => {
      clearTimeout(timer);
      if (!stdout.trim()) resolve(null);
    });

    child.stdin.write(JSON.stringify(message) + '\n');
  });
}

/**
 * Write a test connection.json file.
 */
function writeTestConnectionInfo(port, token) {
  fs.mkdirSync(CONN_DIR, { recursive: true });
  fs.writeFileSync(CONN_FILE, JSON.stringify({ port, token, pid: process.pid }), { encoding: 'utf8', mode: 0o600 });
}

describe('Auth: connection info file', () => {
  it('bridge reads port and token from connection.json', async () => {
    let TEST_PORT;
    const TEST_TOKEN = 'a'.repeat(64);
    let receivedHeaders = null;
    let receivedPort = null;

    // Start a mock server on the test port
    const server = http.createServer((req, res) => {
      receivedHeaders = req.headers;
      receivedPort = TEST_PORT;
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        jsonrpc: '2.0',
        id: 1,
        result: { tools: [] }
      }));
    });

    await new Promise((resolve, reject) => {
      server.on('error', reject);
      server.listen(0, '127.0.0.1', resolve);
    });
    TEST_PORT = server.address().port;

    try {
      // Write connection info pointing to our mock server
      writeTestConnectionInfo(TEST_PORT, TEST_TOKEN);

      const response = await sendToBridge({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/list'
      });

      // Verify the bridge connected to our mock server (correct port)
      assert.equal(receivedPort, TEST_PORT);

      // Verify the auth token was sent
      assert.equal(receivedHeaders['authorization'], `Bearer ${TEST_TOKEN}`);

      // Verify we got a valid response
      assert.equal(response.id, 1);
    } finally {
      await new Promise((resolve) => server.close(resolve));
    }
  });

  it('bridge fails closed when connection file is missing', async () => {
    // Remove connection file
    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }

    // With fail-closed auth, the bridge must refuse to forward requests
    // when it can't find the connection file (no fallback to default port).
    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/list'
    });

    assert.equal(response.id, 2);
    assert.ok(response.error, 'should return an error when connection file is missing');
    assert.match(response.error.message, /Connection file not found|Bridge error/);
  });
});

describe('Auth: token verification', () => {
  let server;
  let TEST_PORT;
  const CORRECT_TOKEN = 'b'.repeat(64);

  before(async () => {
    // Mock server that checks auth like the extension does
    server = http.createServer((req, res) => {
      let authHeader = req.headers['authorization'] || '';
      if (authHeader !== `Bearer ${CORRECT_TOKEN}`) {
        res.writeHead(403, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          jsonrpc: '2.0',
          id: null,
          error: { code: -32600, message: 'Invalid or missing auth token' }
        }));
        return;
      }

      let body = '';
      req.on('data', (chunk) => { body += chunk; });
      req.on('end', () => {
        const msg = JSON.parse(body);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          jsonrpc: '2.0',
          id: msg.id,
          result: { tools: [{ name: 'authenticated' }] }
        }));
      });
    });

    await new Promise((resolve, reject) => {
      server.on('error', reject);
      server.listen(0, '127.0.0.1', resolve);
    });
    TEST_PORT = server.address().port;
  });

  after(async () => {
    if (server) await new Promise((resolve) => server.close(resolve));
  });

  it('succeeds with correct token', async () => {
    writeTestConnectionInfo(TEST_PORT, CORRECT_TOKEN);

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 10,
      method: 'tools/list'
    });

    assert.equal(response.id, 10);
    assert.ok(response.result);
    assert.equal(response.result.tools[0].name, 'authenticated');
  });

  it('fails with wrong token', async () => {
    writeTestConnectionInfo(TEST_PORT, 'c'.repeat(64));

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 11,
      method: 'tools/list'
    });

    // The bridge clears its cache and rejects on 403 responses
    assert.equal(response.id, 11);
    assert.ok(response.error);
    assert.match(response.error.message, /authentication failed/i);
  });
});

// ═══════════════════════════════════════════════════════════════════
// TIMING-SAFE COMPARISON TESTS
// ═══════════════════════════════════════════════════════════════════

/**
 * Exact copy of timingSafeEqual from api.js.
 * Tests verify correctness, not timing (timing tests are fragile in CI).
 */
function timingSafeEqual(a, b) {
  const aStr = String(a);
  const bStr = String(b);
  const len = Math.max(aStr.length, bStr.length);
  let result = aStr.length ^ bStr.length;
  for (let i = 0; i < len; i++) {
    result |= (aStr.charCodeAt(i) || 0) ^ (bStr.charCodeAt(i) || 0);
  }
  return result === 0;
}

describe('Timing-safe comparison: correctness', () => {
  it('equal strings return true', () => {
    assert.equal(timingSafeEqual('abc', 'abc'), true);
  });

  it('different strings return false', () => {
    assert.equal(timingSafeEqual('abc', 'xyz'), false);
  });

  it('different lengths return false', () => {
    assert.equal(timingSafeEqual('short', 'longer-string'), false);
  });

  it('empty strings are equal', () => {
    assert.equal(timingSafeEqual('', ''), true);
  });

  it('empty vs non-empty returns false', () => {
    assert.equal(timingSafeEqual('', 'x'), false);
  });

  it('single char difference returns false', () => {
    assert.equal(timingSafeEqual('Bearer token-abc', 'Bearer token-abd'), false);
  });

  it('handles realistic auth token comparison', () => {
    const token = 'a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2';
    assert.equal(timingSafeEqual(`Bearer ${token}`, `Bearer ${token}`), true);
    assert.equal(timingSafeEqual(`Bearer ${token}`, `Bearer wrong-token`), false);
  });

  it('handles unicode strings', () => {
    assert.equal(timingSafeEqual('héllo', 'héllo'), true);
    assert.equal(timingSafeEqual('héllo', 'hello'), false);
  });

  it('handles null/undefined coercion via String()', () => {
    assert.equal(timingSafeEqual(null, 'null'), true);
    assert.equal(timingSafeEqual(undefined, 'undefined'), true);
    assert.equal(timingSafeEqual(null, undefined), false);
  });

  it('handles numeric coercion via String()', () => {
    assert.equal(timingSafeEqual(123, '123'), true);
    assert.equal(timingSafeEqual(0, '0'), true);
  });

  it('prefix of another string returns false', () => {
    assert.equal(timingSafeEqual('Bearer abc', 'Bearer abcdef'), false);
  });

  it('handles very long equal strings', () => {
    const long = 'x'.repeat(10000);
    assert.equal(timingSafeEqual(long, long), true);
  });

  it('handles very long strings differing only at end', () => {
    const base = 'x'.repeat(9999);
    assert.equal(timingSafeEqual(base + 'a', base + 'b'), false);
  });

  it('handles special characters', () => {
    assert.equal(timingSafeEqual('a\x00b', 'a\x00b'), true);
    assert.equal(timingSafeEqual('a\x00b', 'a\x00c'), false);
  });
});

// ═══════════════════════════════════════════════════════════════════
// CONNECTION FILE EDGE CASES
// ═══════════════════════════════════════════════════════════════════

describe('Auth: connection file corruption', () => {

  it('rejects empty connection file', async () => {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, '', { encoding: 'utf8', mode: 0o600 });

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 20,
      method: 'tools/list'
    });

    assert.equal(response.id, 20);
    assert.ok(response.error);
  });

  it('rejects connection file with invalid JSON', async () => {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, '{not valid json!!!', { encoding: 'utf8', mode: 0o600 });

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 21,
      method: 'tools/list'
    });

    assert.equal(response.id, 21);
    assert.ok(response.error);
  });

  it('rejects connection file with missing port', async () => {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, JSON.stringify({ token: 'abc' }), { encoding: 'utf8', mode: 0o600 });

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 22,
      method: 'tools/list'
    });

    assert.equal(response.id, 22);
    assert.ok(response.error);
    assert.match(response.error.message, /missing port or token|Bridge error/);
  });

  it('rejects connection file with missing token', async () => {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, JSON.stringify({ port: 19999 }), { encoding: 'utf8', mode: 0o600 });

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 23,
      method: 'tools/list'
    });

    assert.equal(response.id, 23);
    assert.ok(response.error);
    assert.match(response.error.message, /missing port or token|Bridge error/);
  });

  it('rejects connection file with null port', async () => {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, JSON.stringify({ port: null, token: 'abc' }), { encoding: 'utf8', mode: 0o600 });

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 24,
      method: 'tools/list'
    });

    assert.equal(response.id, 24);
    assert.ok(response.error);
  });

  it('rejects connection file with empty string token', async () => {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, JSON.stringify({ port: 19999, token: '' }), { encoding: 'utf8', mode: 0o600 });

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 25,
      method: 'tools/list'
    });

    assert.equal(response.id, 25);
    assert.ok(response.error);
  });

  it('rejects connection file with port=0', async () => {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, JSON.stringify({ port: 0, token: 'abc' }), { encoding: 'utf8', mode: 0o600 });

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 26,
      method: 'tools/list'
    });

    assert.equal(response.id, 26);
    assert.ok(response.error);
  });

  it('handles binary garbage in connection file', async () => {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, Buffer.from([0x00, 0xff, 0xfe, 0x80, 0x90]), { mode: 0o600 });

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 27,
      method: 'tools/list'
    });

    assert.equal(response.id, 27);
    assert.ok(response.error);
  });

  it('accepts connection file with extra fields (forward compat)', async () => {
    let TEST_PORT;
    const TEST_TOKEN = 'd'.repeat(64);

    const server = http.createServer((req, res) => {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', id: 28, result: { tools: [] } }));
    });

    await new Promise((resolve, reject) => {
      server.on('error', reject);
      server.listen(0, '127.0.0.1', resolve);
    });
    TEST_PORT = server.address().port;

    try {
      fs.mkdirSync(CONN_DIR, { recursive: true });
      fs.writeFileSync(CONN_FILE, JSON.stringify({
        port: TEST_PORT, token: TEST_TOKEN, pid: 12345,
        version: '2.0', extraField: 'should be ignored'
      }), { encoding: 'utf8', mode: 0o600 });

      const response = await sendToBridge({
        jsonrpc: '2.0',
        id: 28,
        method: 'tools/list'
      });

      assert.equal(response.id, 28);
      assert.ok(response.result);
    } finally {
      await new Promise((resolve) => server.close(resolve));
    }
  });

  it('rejects malformed tokens before contacting Thunderbird', async () => {
    let TEST_PORT;
    let requestCount = 0;
    const cases = [
      { name: 'whitespace-only', token: ' '.repeat(64) },
      { name: 'uppercase hex', token: 'A'.repeat(64) },
      { name: 'too short', token: 'a'.repeat(63) },
      { name: 'too long', token: 'a'.repeat(65) },
      { name: 'trailing newline', token: `${'a'.repeat(64)}\n` },
      { name: 'non-hex characters', token: 'g'.repeat(64) },
      { name: 'very large', token: 'x'.repeat(2048) },
      { name: 'special characters', token: 'tok3n+with/special=chars&more!@#$%' },
    ];

    const server = http.createServer((req, res) => {
      requestCount += 1;
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', id: null, result: { shouldNotReach: true } }));
    });

    await new Promise((resolve, reject) => {
      server.on('error', reject);
      server.listen(0, '127.0.0.1', resolve);
    });
    TEST_PORT = server.address().port;

    try {
      for (let i = 0; i < cases.length; i++) {
        const testCase = cases[i];
        const id = 29 + i;
        requestCount = 0;
        writeTestConnectionInfo(TEST_PORT, testCase.token);

        const response = await sendToBridge({
          jsonrpc: '2.0',
          id,
          method: 'tools/list'
        });

        assert.equal(response.id, id, testCase.name);
        assert.ok(response.error, testCase.name);
        assert.match(response.error.message, /token must be 64 lowercase hex characters/, testCase.name);
        assert.equal(requestCount, 0, `${testCase.name} should be rejected before HTTP`);
      }
    } finally {
      await new Promise((resolve) => server.close(resolve));
    }
  });
});

// ═══════════════════════════════════════════════════════════════════
// BRIDGE LIFECYCLE & RETRY BEHAVIOR
// ═══════════════════════════════════════════════════════════════════

describe('Auth: bridge handles MCP lifecycle locally', () => {
  // These methods are handled by the bridge directly without
  // contacting Thunderbird, so they should work even without
  // a connection file.

  it('initialize succeeds without connection file', async () => {
    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 40,
      method: 'initialize',
      params: { protocolVersion: '2024-11-05', capabilities: {} }
    });

    assert.equal(response.id, 40);
    assert.ok(response.result);
    assert.equal(response.result.serverInfo.name, 'thunderbird-mcp');
  });

  it('ping succeeds without connection file', async () => {
    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 41,
      method: 'ping'
    });

    assert.equal(response.id, 41);
    assert.ok(response.result);
  });

  it('resources/list succeeds without connection file', async () => {
    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 42,
      method: 'resources/list'
    });

    assert.equal(response.id, 42);
    assert.deepStrictEqual(response.result, { resources: [] });
  });

  it('prompts/list succeeds without connection file', async () => {
    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 43,
      method: 'prompts/list'
    });

    assert.equal(response.id, 43);
    assert.deepStrictEqual(response.result, { prompts: [] });
  });

  it('notifications are silently dropped (no response)', async () => {
    // Notifications have no id — bridge should not respond
    const child = spawn(process.execPath, [BRIDGE_PATH], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: BRIDGE_ENV,
    });

    let stdout = '';
    const timer = setTimeout(() => {
      child.kill();
    }, 2000);

    child.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    // Send a notification (no id)
    child.stdin.write(JSON.stringify({
      jsonrpc: '2.0',
      method: 'notifications/initialized'
    }) + '\n');

    // Then send a ping to verify bridge is still alive
    child.stdin.write(JSON.stringify({
      jsonrpc: '2.0',
      id: 44,
      method: 'ping'
    }) + '\n');

    await new Promise((resolve) => {
      child.stdout.on('data', () => {
        const lines = stdout.split('\n').filter(l => l.trim());
        if (lines.length >= 1) {
          clearTimeout(timer);
          child.stdin.end();
          resolve();
        }
      });
    });

    // Should only have the ping response, not the notification
    const lines = stdout.split('\n').filter(l => l.trim());
    assert.equal(lines.length, 1);
    const response = JSON.parse(lines[0]);
    assert.equal(response.id, 44);
    assert.equal(response.method, undefined); // it's a response, not an echo
  });
});

describe('Auth: bridge retry then fail', () => {

  it('retries and succeeds when connection file appears mid-retry', async () => {
    // Remove connection file first
    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }

    let TEST_PORT;
    const TEST_TOKEN = 'e'.repeat(64);

    const server = http.createServer((req, res) => {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ jsonrpc: '2.0', id: 50, result: { delayed: true } }));
    });

    await new Promise((resolve, reject) => {
      server.on('error', reject);
      server.listen(0, '127.0.0.1', resolve);
    });
    TEST_PORT = server.address().port;

    try {
      // Write the connection file after 2 seconds (bridge retries every 1s)
      setTimeout(() => {
        writeTestConnectionInfo(TEST_PORT, TEST_TOKEN);
      }, 2000);

      const response = await sendToBridge({
        jsonrpc: '2.0',
        id: 50,
        method: 'tools/list'
      });

      assert.equal(response.id, 50);
      assert.ok(response.result);
      assert.equal(response.result.delayed, true);
    } finally {
      await new Promise((resolve) => server.close(resolve));
    }
  });
});


describe('Bridge attachment dispatch', () => {
  it('refuses mixed attachments for every mail tool before forwarding to HTTP', async () => {
    let requests = 0;
    const server = http.createServer((_req, res) => {
      requests++;
      res.end(JSON.stringify({ jsonrpc: '2.0', id: 1, result: {} }));
    });
    await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    try {
      writeTestConnectionInfo(server.address().port, 'a'.repeat(64));
      const valid = path.join(CONN_DIR, 'report.pdf');
      const refused = path.join(CONN_DIR, '.env');
      const missing = path.join(CONN_DIR, 'missing.pdf');
      fs.writeFileSync(valid, 'safe attachment');
      for (const name of ['sendMail', 'saveDraft', 'replyToMessage', 'forwardMessage']) {
        for (const skipReview of [false, true]) {
          const response = await sendToBridge({
            jsonrpc: '2.0', id: 1, method: 'tools/call',
            params: { name, arguments: { attachments: [valid, refused, missing], skipReview } },
          });
          assert.equal(response.error?.code, -32602, name);
          assert.ok(response.error.message.includes(refused), name);
          assert.ok(response.error.message.includes(missing), name);
        }
      }
      assert.equal(requests, 0);
    } finally {
      await new Promise(resolve => server.close(resolve));
    }
  });

  it('rejects JSON-encoded attachment strings for every mail tool before HTTP', async () => {
    let requests = 0;
    const server = http.createServer((_req, res) => {
      requests++;
      res.end(JSON.stringify({ jsonrpc: '2.0', id: 1, result: {} }));
    });
    await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    try {
      writeTestConnectionInfo(server.address().port, 'a'.repeat(64));
      const file = path.join(CONN_DIR, 'report.txt');
      fs.writeFileSync(file, 'attachment fixture');
      for (const name of ['sendMail', 'saveDraft', 'replyToMessage', 'forwardMessage']) {
        for (const skipReview of [false, true]) {
          const response = await sendToBridge({
            jsonrpc: '2.0', id: 1, method: 'tools/call',
            params: { name, arguments: { attachments: JSON.stringify([file]), skipReview } },
          });
          assert.equal(response.error?.code, -32602, name);
          assert.match(response.error.message, /attachments must be an array/, name);
        }
      }
      assert.equal(requests, 0);
    } finally {
      await new Promise(resolve => server.close(resolve));
    }
  });

  it('inlines allowed saveDraft paths and applies the existing 18 MiB bridge limit', async () => {
    const received = [];
    const server = http.createServer((req, res) => {
      let body = '';
      req.on('data', chunk => { body += chunk; });
      req.on('end', () => {
        received.push(JSON.parse(body));
        res.end(JSON.stringify({ jsonrpc: '2.0', id: 1, result: {} }));
      });
    });
    await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    try {
      writeTestConnectionInfo(server.address().port, 'a'.repeat(64));
      const file = path.join(CONN_DIR, 'report.txt');
      fs.writeFileSync(file, 'draft attachment');
      const message = { jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'saveDraft', arguments: { attachments: [file] } } };
      const response = await sendToBridge(message);
      assert.ok(response.result);
      assert.equal(received.length, 1);
      assert.deepEqual(received[0].params.arguments.attachments, [{
        name: 'report.txt', contentType: 'text/plain', base64: Buffer.from('draft attachment').toString('base64'),
      }]);
      fs.truncateSync(file, 18 * 1024 * 1024 + 1);
      assert.match((await sendToBridge(message)).error.message, /Attachment too large/);
      assert.equal(received.length, 1);
    } finally {
      await new Promise(resolve => server.close(resolve));
    }
  });
});
