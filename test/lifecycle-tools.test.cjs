/**
 * Tests for bridge-local Thunderbird lifecycle management tools.
 *
 * These tools (launchThunderbird, stopThunderbird, thunderbirdStatus) are
 * handled entirely within the bridge and never forwarded to the extension.
 */

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { spawn } = require('child_process');
const path = require('path');

const {
  LIFECYCLE_TOOL_NAMES,
  LIFECYCLE_TOOL_SCHEMAS,
} = require('../mcp-bridge.cjs');

const BRIDGE_PATH = path.resolve(__dirname, '..', 'mcp-bridge.cjs');

/**
 * Spawn the bridge, send an initialize handshake followed by one request,
 * and return the parsed JSON-RPC response for the request (skipping the
 * initialize response).
 */
function callBridge(request, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [BRIDGE_PATH], {
      stdio: ['pipe', 'pipe', 'pipe'],
    });

    let buf = '';
    const parsed = [];
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error(`Bridge timed out. Collected ${parsed.length} responses. buf="${buf}"`));
    }, timeoutMs);

    child.stdout.setEncoding('utf8');
    child.stdout.on('data', (chunk) => {
      buf += chunk;
      let idx;
      while ((idx = buf.indexOf('\n')) !== -1) {
        const line = buf.slice(0, idx).trim();
        buf = buf.slice(idx + 1);
        if (!line) continue;
        try {
          parsed.push(JSON.parse(line));
        } catch {
          // partial data, ignore
        }
      }
      // We expect 2 responses: initialize (id=0) + the actual request (id=1)
      if (parsed.length >= 2) {
        clearTimeout(timer);
        child.stdin.end();
        child.kill();
        resolve(parsed[1]); // return the second response
      }
    });

    child.stderr.on('data', () => {});
    child.on('error', (err) => { clearTimeout(timer); reject(err); });
    child.on('exit', (code) => {
      clearTimeout(timer);
      if (parsed.length >= 2) return; // already resolved
      if (parsed.length === 1) {
        resolve(parsed[0]); // only got one response
      } else {
        reject(new Error(`Bridge exited (code ${code}) with ${parsed.length} responses`));
      }
    });

    const init = JSON.stringify({
      jsonrpc: '2.0', id: 0, method: 'initialize',
      params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '0.0.0' } },
    });
    const msg = JSON.stringify(request);
    child.stdin.write(init + '\n' + msg + '\n');
  });
}

// ── Schema validation ────────────────────────────────────────────────────────

describe('lifecycle tool schemas', () => {
  it('defines exactly 3 tools', () => {
    assert.equal(LIFECYCLE_TOOL_SCHEMAS.length, 3);
  });

  it('tool names match the Set', () => {
    const names = LIFECYCLE_TOOL_SCHEMAS.map((t) => t.name);
    assert.deepEqual(new Set(names), LIFECYCLE_TOOL_NAMES);
  });

  for (const schema of LIFECYCLE_TOOL_SCHEMAS) {
    it(`${schema.name} has valid MCP tool schema`, () => {
      assert.equal(typeof schema.name, 'string');
      assert.equal(typeof schema.description, 'string');
      assert.ok(schema.description.length > 0, 'description should not be empty');
      assert.deepEqual(typeof schema.inputSchema, 'object');
      assert.equal(schema.inputSchema.type, 'object');
    });
  }
});

// ── tools/list includes lifecycle tools ──────────────────────────────────────

describe('tools/list with lifecycle tools', () => {
  it('returns lifecycle tools even when Thunderbird is not running', async () => {
    const resp = await callBridge({
      jsonrpc: '2.0', id: 1, method: 'tools/list', params: {},
    });

    assert.ok(resp.result, 'tools/list should return a result');
    assert.ok(Array.isArray(resp.result.tools), 'result.tools should be an array');

    const toolNames = resp.result.tools.map((t) => t.name);
    assert.ok(toolNames.includes('launchThunderbird'), 'should include launchThunderbird');
    assert.ok(toolNames.includes('stopThunderbird'), 'should include stopThunderbird');
    assert.ok(toolNames.includes('thunderbirdStatus'), 'should include thunderbirdStatus');
  });
});

// ── thunderbirdStatus tool ───────────────────────────────────────────────────

describe('thunderbirdStatus tool', () => {
  it('returns valid status shape', async () => {
    const resp = await callBridge({
      jsonrpc: '2.0', id: 1, method: 'tools/call',
      params: { name: 'thunderbirdStatus', arguments: {} },
    });

    assert.ok(resp.result, 'should return a result');
    assert.ok(Array.isArray(resp.result.content), 'result.content should be an array');
    assert.equal(resp.result.content[0].type, 'text');

    const status = JSON.parse(resp.result.content[0].text);
    assert.equal(typeof status.running, 'boolean', 'running should be boolean');
    assert.equal(typeof status.extensionReachable, 'boolean', 'extensionReachable should be boolean');
  });
});

// ── stopThunderbird response shape ───────────────────────────────────────────

describe('stopThunderbird response shape', () => {
  it('returns valid response', async () => {
    const resp = await callBridge({
      jsonrpc: '2.0', id: 1, method: 'tools/call',
      params: { name: 'stopThunderbird', arguments: {} },
    });

    assert.ok(resp.result, 'should return a result');
    assert.ok(Array.isArray(resp.result.content), 'result.content should be an array');

    const result = JSON.parse(resp.result.content[0].text);
    assert.equal(typeof result.stopped, 'boolean');
    assert.equal(typeof result.message, 'string');
  });
});

// ── Routing ──────────────────────────────────────────────────────────────────

describe('lifecycle tool routing', () => {
  it('does not intercept non-lifecycle tools', () => {
    assert.ok(!LIFECYCLE_TOOL_NAMES.has('listAccounts'));
    assert.ok(!LIFECYCLE_TOOL_NAMES.has('searchMessages'));
    assert.ok(!LIFECYCLE_TOOL_NAMES.has('getMessage'));
  });

  it('intercepts all lifecycle tools', () => {
    assert.ok(LIFECYCLE_TOOL_NAMES.has('launchThunderbird'));
    assert.ok(LIFECYCLE_TOOL_NAMES.has('stopThunderbird'));
    assert.ok(LIFECYCLE_TOOL_NAMES.has('thunderbirdStatus'));
  });
});
