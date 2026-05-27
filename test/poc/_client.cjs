/**
 * Shared MCP HTTP client used by the PoC scripts in this directory.
 *
 * Discovers Thunderbird's connection.json the same way mcp-bridge.cjs does
 * (env override, then native tmp, then macOS/Snap/Flatpak fallbacks via the
 * bridge module) and exposes a single `callTool(name, args)` helper that
 * speaks the MCP JSON-RPC `tools/call` shape.
 *
 * The PoCs in this folder are illustrative scripts for the security review --
 * they exercise the attack chains described in PR #102. They never target
 * external addresses; the dummy recipient is `attacker@dummy.invalid`, an
 * RFC 6761 reserved TLD that cannot resolve or receive mail.
 */
'use strict';

const http = require('http');
const path = require('path');

// Reuse the bridge's discovery logic so the PoC works on Windows / macOS /
// Snap / Flatpak without reimplementation. The bridge is a sibling file
// in the repo root.
const bridge = require(path.resolve(__dirname, '..', '..', 'mcp-bridge.cjs'));

let cached = null;

function loadConnection() {
  if (cached) return cached;
  const info = bridge.readConnectionInfo();
  if (!info) {
    const attempts = bridge.formatDiscoveryAttempts();
    throw new Error(
      `Could not find Thunderbird MCP connection file. Is Thunderbird running with the extension installed?\nTried: ${attempts}`
    );
  }
  if (!bridge.isValidAuthToken(info.token)) {
    throw new Error('Connection file token is not a valid 64-hex auth token.');
  }
  cached = info;
  return info;
}

function postJson(port, token, body) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify(body);
    const req = http.request({
      hostname: '127.0.0.1',
      port,
      path: '/',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
        Authorization: `Bearer ${token}`,
      },
    }, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        try {
          resolve({ status: res.statusCode, body: JSON.parse(raw) });
        } catch (e) {
          reject(new Error(`Non-JSON response (status ${res.statusCode}): ${raw.slice(0, 400)}`));
        }
      });
    });
    req.on('error', reject);
    req.setTimeout(30000, () => {
      req.destroy();
      reject(new Error('Request timed out after 30s'));
    });
    req.write(payload);
    req.end();
  });
}

let nextId = 1;

async function rpc(method, params) {
  const info = loadConnection();
  const { status, body } = await postJson(info.port, info.token, {
    jsonrpc: '2.0',
    id: nextId++,
    method,
    params,
  });
  if (status !== 200) {
    throw new Error(`HTTP ${status}: ${JSON.stringify(body)}`);
  }
  return body;
}

/**
 * Invoke a tool via MCP and unwrap the structured result. Returns the parsed
 * tool response (i.e. what the tool's handler returned), or throws if the
 * server reported a JSON-RPC error.
 */
async function callTool(name, args) {
  const resp = await rpc('tools/call', { name, arguments: args });
  if (resp.error) {
    const e = new Error(resp.error.message || JSON.stringify(resp.error));
    e.jsonrpcError = resp.error;
    throw e;
  }
  // Tool results are returned as { content: [{type:'text', text: '<json>'}] }
  const text = resp?.result?.content?.[0]?.text;
  if (typeof text !== 'string') return resp.result;
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function banner(title) {
  const bar = '─'.repeat(Math.max(8, title.length + 4));
  process.stdout.write(`\n${bar}\n  ${title}\n${bar}\n`);
}

function log(msg) {
  process.stdout.write(`${msg}\n`);
}

module.exports = { rpc, callTool, banner, log, loadConnection };
