/**
 * Tests for mcp-bridge.cjs
 *
 * Covers:
 * - sanitizeJson: control character removal and escaping
 * - handleMessage: MCP lifecycle methods handled locally
 * - handleMessage: notifications return null
 * - forwardToThunderbird: HTTP forwarding with mock server
 * - forwardToThunderbird: connection refused error handling
 * - End-to-end stdin/stdout bridge via child process
 */

const { describe, it, before, after } = require('node:test');
const assert = require('node:assert/strict');
const http = require('http');
const fs = require('fs');
const path = require('path');

const {
  sendToBridge,
  sanitizeJson,
  CONN_DIR,
  CONN_FILE,
} = require('./helpers.cjs');

describe('Bridge: MCP lifecycle methods', () => {
  it('responds to initialize with protocol version and capabilities', async () => {
    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {}
    });

    assert.equal(response.jsonrpc, '2.0');
    assert.equal(response.id, 1);
    assert.equal(response.result.protocolVersion, '2024-11-05');
    assert.deepEqual(response.result.capabilities, { tools: {} });
    assert.equal(response.result.serverInfo.name, 'thunderbird-mcp');
  });

  it('responds to ping with empty result', async () => {
    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 2,
      method: 'ping'
    });

    assert.equal(response.id, 2);
    assert.deepEqual(response.result, {});
  });

  it('responds to resources/list with empty array', async () => {
    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 3,
      method: 'resources/list'
    });

    assert.equal(response.id, 3);
    assert.deepEqual(response.result, { resources: [] });
  });

  it('responds to prompts/list with empty array', async () => {
    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 4,
      method: 'prompts/list'
    });

    assert.equal(response.id, 4);
    assert.deepEqual(response.result, { prompts: [] });
  });
});

describe('Bridge: notifications', () => {
  it('returns null for notification messages (no id)', async () => {
    const response = await sendToBridge({
      jsonrpc: '2.0',
      method: 'notifications/initialized'
    });

    assert.equal(response, null);
  });
});

describe('Bridge: HTTP forwarding', () => {
  let mockServer;
  let mockPort = 8765;
  let savedConnData = null;

  before(async () => {
    // Back up existing connection file
    try { savedConnData = fs.readFileSync(CONN_FILE, 'utf8'); } catch { savedConnData = null; }

    // Start a mock server on port 8765 to simulate Thunderbird
    mockServer = http.createServer((req, res) => {
      let body = '';
      req.on('data', (chunk) => { body += chunk; });
      req.on('end', () => {
        const message = JSON.parse(body);
        if (message.method === 'tools/list') {
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            result: { tools: [{ name: 'testTool', description: 'A test tool' }] }
          }));
        } else if (message.method === 'tools/call') {
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            result: { content: [{ type: 'text', text: '{"ok": true}' }] }
          }));
        } else {
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({
            jsonrpc: '2.0',
            id: message.id,
            error: { code: -32601, message: 'Method not found' }
          }));
        }
      });
    });

    await new Promise((resolve, reject) => {
      mockServer.on('error', (err) => {
        if (err.code === 'EADDRINUSE') {
          mockServer = null;
          resolve();
        } else {
          reject(err);
        }
      });
      mockServer.listen(mockPort, '127.0.0.1', () => {
        fs.mkdirSync(CONN_DIR, { recursive: true });
        fs.writeFileSync(CONN_FILE, JSON.stringify({ port: mockPort, token: 'mock-test-token' }), 'utf8');
        resolve();
      });
    });
  });

  after(async () => {
    if (mockServer) {
      await new Promise((resolve) => mockServer.close(resolve));
    }
    if (savedConnData !== null) {
      fs.mkdirSync(CONN_DIR, { recursive: true });
      fs.writeFileSync(CONN_FILE, savedConnData, 'utf8');
    } else {
      try { fs.unlinkSync(CONN_FILE); } catch { /* ignore */ }
    }
  });

  it('forwards tools/list to Thunderbird and returns response', async (t) => {
    if (!mockServer) return t.skip('Port 8765 in use (Thunderbird running), skipping mock test');

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 10,
      method: 'tools/list'
    });

    assert.equal(response.id, 10);
    assert.ok(response.result.tools.length > 0);
  });

  it('forwards tools/call to Thunderbird and returns response', async (t) => {
    if (!mockServer) return t.skip('Port 8765 in use (Thunderbird running), skipping mock test');

    const response = await sendToBridge({
      jsonrpc: '2.0',
      id: 11,
      method: 'tools/call',
      params: { name: 'testTool', arguments: {} }
    });

    assert.equal(response.id, 11);
    assert.ok(response.result);
  });
});

describe('Bridge: sanitizeJson', () => {
  it('removes control characters except newline, carriage return, tab', () => {
    assert.equal(sanitizeJson('hello\x00world\x07'), 'helloworld');
    assert.equal(sanitizeJson('line1\nline2'), 'line1\\nline2');
    assert.equal(sanitizeJson('col1\tcol2'), 'col1\\tcol2');
    assert.equal(sanitizeJson('already\\nescaped'), 'already\\nescaped');
    assert.equal(sanitizeJson('\x01\x02hello\nworld\x7f'), 'hello\\nworld');
  });

  it('handles empty string', () => {
    assert.equal(sanitizeJson(''), '');
  });

  it('preserves valid JSON with unicode', () => {
    const input = '{"emoji": "\\ud83d\\ude00", "text": "hello"}';
    assert.equal(sanitizeJson(input), input);
  });
});
