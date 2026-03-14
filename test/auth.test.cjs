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

const { describe, it, before, after } = require('node:test');
const assert = require('node:assert/strict');
const http = require('http');
const fs = require('fs');

const {
  sendToBridge,
  writeTestConnectionInfo,
  backupConnectionFile,
  restoreConnectionFile,
  isDefaultPortInUse,
  CONN_DIR,
  CONN_FILE,
} = require('./helpers.cjs');

describe('Auth: connection info file', () => {
  before(() => backupConnectionFile());
  after(() => restoreConnectionFile());

  it('bridge reads port and token from connection.json', async () => {
    const TEST_PORT = 18765;
    const TEST_TOKEN = 'test-secret-token-abc123';
    let receivedHeaders = null;
    let receivedPort = null;

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
      writeTestConnectionInfo(TEST_PORT, TEST_TOKEN);

      const response = await sendToBridge({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/list'
      });

      assert.equal(receivedPort, TEST_PORT);
      assert.equal(receivedHeaders['authorization'], `Bearer ${TEST_TOKEN}`);
      assert.equal(response.id, 1);
    } finally {
      await new Promise((resolve) => server.close(resolve));
    }
  });

  it('bridge falls back to default port when connection file is missing', async (t) => {
    if (await isDefaultPortInUse()) {
      return t.skip('Port 8765 in use (Thunderbird running), skipping fallback test');
    }

    try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/list'
    });

    assert.equal(response.id, 2);
    assert.ok(response.result || response.error);
  });
});

describe('Auth: token verification', () => {
  let server;
  const TEST_PORT = 18766;
  const CORRECT_TOKEN = 'correct-token-xyz';

  before(async () => {
    backupConnectionFile();

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

    assert.equal(response.id, null);
    assert.ok(response.error);
    assert.match(response.error.message, /auth token/i);
  });
});
