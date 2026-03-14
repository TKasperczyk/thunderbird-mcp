/**
 * Tests for auth token and connection info discovery.
 *
 * Covers:
 * - Connection info file reading and caching
 * - Auth token inclusion in HTTP requests
 * - Fallback behavior when connection file is missing
 * - Cache invalidation on connection errors
 * - Bridge behavior with and without auth
 */

const { describe, it, before, after, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');

const BRIDGE_PATH = path.resolve(__dirname, '..', 'mcp-bridge.cjs');
const CONN_DIR = path.join(os.tmpdir(), 'thunderbird-mcp');
const CONN_FILE = path.join(CONN_DIR, 'connection.json');

/**
 * Helper: send a JSON-RPC message to the bridge and get the response.
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
  fs.writeFileSync(CONN_FILE, JSON.stringify({ port, token, pid: process.pid }), 'utf8');
}

/**
 * Back up and restore any existing connection file to avoid
 * interfering with a running Thunderbird instance.
 */
let savedConnectionData = null;

function backupConnectionFile() {
  try {
    savedConnectionData = fs.readFileSync(CONN_FILE, 'utf8');
  } catch {
    savedConnectionData = null;
  }
}

function restoreConnectionFile() {
  if (savedConnectionData !== null) {
    fs.mkdirSync(CONN_DIR, { recursive: true });
    fs.writeFileSync(CONN_FILE, savedConnectionData, 'utf8');
  } else {
    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }
  }
}

describe('Auth: connection info file', () => {
  before(() => backupConnectionFile());
  after(() => restoreConnectionFile());

  it('bridge reads port and token from connection.json', async () => {
    const TEST_PORT = 18765;
    const TEST_TOKEN = 'test-secret-token-abc123';
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
      server.listen(TEST_PORT, '127.0.0.1', resolve);
    });

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

  it('bridge falls back to default port when connection file is missing', async () => {
    // Remove connection file
    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }

    // The bridge should try port 8765 — we won't mock it, so we expect
    // either a connection error or a response from real Thunderbird.
    // We just verify it doesn't crash.
    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/list'
    });

    assert.equal(response.id, 2);
    // Either a valid result (real Thunderbird) or an error (connection refused)
    assert.ok(response.result || response.error);
  });
});

describe('Auth: token verification', () => {
  let server;
  const TEST_PORT = 18766;
  const CORRECT_TOKEN = 'correct-token-xyz';

  before(async () => {
    backupConnectionFile();

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
      server.listen(TEST_PORT, '127.0.0.1', resolve);
    });
  });

  after(async () => {
    if (server) await new Promise((resolve) => server.close(resolve));
    restoreConnectionFile();
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
    writeTestConnectionInfo(TEST_PORT, 'wrong-token');

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 11,
      method: 'tools/list'
    });

    // The bridge receives the 403 response body which has an error
    // The mock returns a JSON-RPC error, so the bridge should parse it
    assert.equal(response.id, null); // 403 response uses id: null
    assert.ok(response.error);
    assert.match(response.error.message, /auth token/i);
  });
});
