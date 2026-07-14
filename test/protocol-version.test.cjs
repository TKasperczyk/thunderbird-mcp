/**
 * Tests for spec-compliant MCP protocolVersion negotiation.
 *
 * Per MCP lifecycle spec, the server MUST respond with the requested
 * protocolVersion if it supports it, otherwise with the latest version
 * it supports. The bridge is a transparent JSON-RPC relay — behavior
 * never changes by version — so it accepts every published version,
 * but does NOT echo unknown future versions back as if it knew them.
 */

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const BRIDGE_PATH = path.resolve(__dirname, '..', 'mcp-bridge.cjs');
const API_PATH = path.resolve(__dirname, '..', 'extension', 'mcp_server', 'api.js');

function runProductionNotificationBranch(method) {
  const source = fs.readFileSync(API_PATH, 'utf8');
  const startMarker = '// BEGIN MCP NOTIFICATION HTTP RESPONSE';
  const endMarker = '// END MCP NOTIFICATION HTTP RESPONSE';
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0, 'notification response start marker missing');
  assert.ok(end > start, 'notification response end marker missing');

  const calls = [];
  const sandbox = {
    method,
    res: {
      setStatusLine(...args) {
        calls.push({ name: 'setStatusLine', args });
      },
      finish() {
        calls.push({ name: 'finish', args: [] });
      },
    },
  };
  vm.runInNewContext(`(function () {\n${source.slice(start, end)}\n})()`, sandbox);
  return calls;
}

function sendInitialize(protocolVersion) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [BRIDGE_PATH], {
      stdio: ['pipe', 'pipe', 'pipe'],
    });

    let stdout = '';
    let stderr = '';
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error(`Bridge timed out. stderr: ${stderr}`));
    }, 5000);

    child.stdout.on('data', (data) => {
      stdout += data.toString();
      const lines = stdout.split('\n').filter(l => l.trim());
      if (lines.length > 0) {
        clearTimeout(timer);
        child.stdin.end();
        child.kill();
        try {
          resolve(JSON.parse(lines[0]));
        } catch (e) {
          reject(new Error(`Failed to parse response: ${lines[0]}`));
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

    child.on('exit', (code) => {
      clearTimeout(timer);
      if (code !== 0 && !stdout.trim()) {
        reject(new Error(`Bridge exited with code ${code}: ${stderr}`));
      }
    });

    const message = JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: protocolVersion,
        capabilities: {},
        clientInfo: { name: 'test-client', version: '1.0' }
      }
    });
    child.stdin.write(message + '\n');
  });
}

describe('protocolVersion negotiation', () => {
  it('responds with the requested known version', async () => {
    const response = await sendInitialize('2024-11-05');
    assert.equal(response.result.protocolVersion, '2024-11-05');
  });

  it('responds with the requested known version 2025-03-26', async () => {
    const response = await sendInitialize('2025-03-26');
    assert.equal(response.result.protocolVersion, '2025-03-26');
  });

  it('responds with the requested known version 2025-06-18', async () => {
    const response = await sendInitialize('2025-06-18');
    assert.equal(response.result.protocolVersion, '2025-06-18');
  });

  it('responds with the requested known version 2025-11-25', async () => {
    const response = await sendInitialize('2025-11-25');
    assert.equal(response.result.protocolVersion, '2025-11-25');
  });

  it('responds with latest version when client requests unknown future version', async () => {
    const response = await sendInitialize('2026-01-01');
    assert.equal(response.result.protocolVersion, '2025-11-25');
  });

  it('responds with latest version when client requests unknown old version', async () => {
    const response = await sendInitialize('2020-01-01');
    assert.equal(response.result.protocolVersion, '2025-11-25');
  });

  it('responds with latest version when no protocolVersion is provided', async () => {
    const message = JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '',
        capabilities: {},
        clientInfo: { name: 'test-client', version: '1.0' }
      }
    });

    return new Promise((resolve, reject) => {
      const child = spawn(process.execPath, [BRIDGE_PATH], {
        stdio: ['pipe', 'pipe', 'pipe'],
      });

      let stdout = '';
      let stderr = '';
      const timer = setTimeout(() => {
        child.kill();
        reject(new Error(`Bridge timed out. stderr: ${stderr}`));
      }, 5000);

      child.stdout.on('data', (data) => {
        stdout += data.toString();
        const lines = stdout.split('\n').filter(l => l.trim());
        if (lines.length > 0) {
          clearTimeout(timer);
          child.stdin.end();
          child.kill();
          try {
            const response = JSON.parse(lines[0]);
            assert.equal(response.result.protocolVersion, '2025-11-25');
            resolve();
          } catch (e) {
            reject(new Error(`Failed to parse response: ${lines[0]}`));
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

      child.on('exit', (code) => {
        clearTimeout(timer);
        if (code !== 0 && !stdout.trim()) {
          reject(new Error(`Bridge exited with code ${code}: ${stderr}`));
        }
      });

      child.stdin.write(message + '\n');
    });
  });

  it('includes serverInfo with name and version', async () => {
    const response = await sendInitialize('2024-11-05');
    assert.equal(response.result.serverInfo.name, 'thunderbird-mcp');
    assert.ok(response.result.serverInfo.version, 'serverInfo.version should be non-empty');
  });

  it('includes tools capability', async () => {
    const response = await sendInitialize('2024-11-05');
    assert.ok(response.result.capabilities.tools, 'should include tools capability');
  });
});

describe('Streamable HTTP notifications', () => {
  it('returns 202 Accepted without a response body', () => {
    const calls = runProductionNotificationBranch('notifications/initialized');
    assert.deepEqual(calls, [
      { name: 'setStatusLine', args: ['1.1', 202, 'Accepted'] },
      { name: 'finish', args: [] },
    ]);
  });

  it('does not handle ordinary JSON-RPC methods as notifications', () => {
    assert.deepEqual(runProductionNotificationBranch('tools/list'), []);
  });
});
