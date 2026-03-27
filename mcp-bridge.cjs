#!/usr/bin/env node
/**
 * MCP Bridge for Thunderbird
 *
 * Converts stdio MCP protocol to HTTP requests for the Thunderbird MCP extension.
 * The extension exposes an HTTP endpoint on localhost:8765.
 */

const http = require('http');
const readline = require('readline');
const fs = require('fs');
const path = require('path');
const os = require('os');

const THUNDERBIRD_HOSTS = ['127.0.0.1'];
const REQUEST_TIMEOUT = 30000;
const DEFAULT_CONNECTION_FILE = path.join(os.tmpdir(), 'thunderbird-mcp', 'connection.json');
const CONNECTION_RETRY_DELAY_MS = 1000;
const CONNECTION_MAX_RETRIES = 5;

function resolveConnectionFile() {
  const envFile = process.env.THUNDERBIRD_MCP_CONNECTION_FILE;
  if (envFile) {
    return envFile;
  }

  const runtimeDir = process.env.XDG_RUNTIME_DIR || `/run/user/${process.getuid()}`;
  const flatpakPath = path.join(
    runtimeDir,
    'app',
    'org.mozilla.Thunderbird',
    'thunderbird-mcp',
    'connection.json'
  );
  if (fs.existsSync(flatpakPath)) {
    return flatpakPath;
  }

  if (fs.existsSync(DEFAULT_CONNECTION_FILE)) {
    return DEFAULT_CONNECTION_FILE;
  }

  return DEFAULT_CONNECTION_FILE;
}

const CONNECTION_FILE = resolveConnectionFile();

/**
 * Read connection info (port + auth token) written by the Thunderbird extension.
 * Returns { port, token } or null if the file doesn't exist.
 * Caches the result for a short TTL to avoid hitting the filesystem on every request.
 * Cache is cleared on connection errors (see clearConnectionCache).
 */
let cachedConnectionInfo = null;
let connectionCacheExpiry = 0;
const CONNECTION_CACHE_TTL_MS = 5000; // 5 seconds

function readConnectionInfo() {
  if (cachedConnectionInfo && Date.now() < connectionCacheExpiry) {
    return cachedConnectionInfo;
  }
  try {
    const data = JSON.parse(fs.readFileSync(CONNECTION_FILE, 'utf8'));
    cachedConnectionInfo = data;
    connectionCacheExpiry = Date.now() + CONNECTION_CACHE_TTL_MS;
    return data;
  } catch {
    return null;
  }
}

function clearConnectionCache() {
  cachedConnectionInfo = null;
  connectionCacheExpiry = 0;
}

// Ensure stdout doesn't buffer - critical for MCP protocol
if (process.stdout._handle?.setBlocking) {
  process.stdout._handle.setBlocking(true);
}

let pendingRequests = 0;
let stdinClosed = false;

function checkExit() {
  if (stdinClosed && pendingRequests === 0) {
    process.exit(0);
  }
}

// Write with backpressure handling
function writeOutput(data) {
  return new Promise((resolve) => {
    if (process.stdout.write(data)) {
      resolve();
    } else {
      process.stdout.once('drain', resolve);
    }
  });
}

/**
 * Sanitize JSON response that may contain invalid control characters.
 * Email bodies often contain raw control chars that break JSON parsing.
 * api.js now pre-encodes non-ASCII for Thunderbird's raw-byte HTTP writer;
 * this remains a fallback for malformed responses.
 */
function sanitizeJson(data) {
  // Remove control chars except \n, \r, \t
  let sanitized = data.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, '');
  // Escape raw newlines/carriage returns/tabs that aren't already escaped
  sanitized = sanitized.replace(/(?<!\\)\r/g, '\\r');
  sanitized = sanitized.replace(/(?<!\\)\n/g, '\\n');
  sanitized = sanitized.replace(/(?<!\\)\t/g, '\\t');
  return sanitized;
}

async function handleMessage(line) {
  const message = JSON.parse(line);
  const hasId = Object.prototype.hasOwnProperty.call(message, 'id');
  const isNotification =
    !hasId ||
    (typeof message.method === 'string' && message.method.startsWith('notifications/'));

  if (isNotification) {
    return null;
  }

  // Handle MCP lifecycle methods locally so the bridge can complete
  // handshake even when Thunderbird isn't running yet.
  switch (message.method) {
    case 'initialize':
      return {
        jsonrpc: '2.0',
        id: message.id,
        result: {
          protocolVersion: '2024-11-05',
          capabilities: { tools: {} },
          serverInfo: { name: 'thunderbird-mcp', version: '0.1.0' }
        }
      };
    case 'ping':
      return { jsonrpc: '2.0', id: message.id, result: {} };
    case 'resources/list':
      return { jsonrpc: '2.0', id: message.id, result: { resources: [] } };
    case 'prompts/list':
      return { jsonrpc: '2.0', id: message.id, result: { prompts: [] } };
  }

  return forwardToThunderbird(message);
}

function tryRequest(hostname, postData, port, token) {
  return new Promise((resolve, reject) => {
    const headers = {
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(postData)
    };
    if (token) {
      headers['Authorization'] = `Bearer ${token}`;
    }
    const req = http.request({
      hostname,
      port,
      path: '/',
      method: 'POST',
      headers
    }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        if (res.statusCode === 403) {
          clearConnectionCache();
          reject(new Error('Authentication failed (403). Token may be stale — retrying with fresh connection info.'));
          return;
        }
        const data = Buffer.concat(chunks).toString('utf8');
        try {
          resolve(JSON.parse(data));
        } catch {
          try {
            resolve(JSON.parse(sanitizeJson(data)));
          } catch (e) {
            reject(new Error(`Invalid JSON from Thunderbird: ${e.message}`));
          }
        }
      });
    });

    req.on('error', reject);

    req.setTimeout(REQUEST_TIMEOUT, () => {
      req.destroy();
      reject(new Error('Request to Thunderbird timed out'));
    });

    req.write(postData);
    req.end();
  });
}

async function forwardToThunderbird(message, _retried) {
  const postData = JSON.stringify(message);

  // Read connection info (port + auth token) from the file written by the extension.
  // Fail-closed: if the connection file is missing, retry a few times
  // (Thunderbird may still be starting), then fail with an error.
  // Never forward requests without authentication.
  let connInfo = readConnectionInfo();
  if (!connInfo) {
    for (let attempt = 0; attempt < CONNECTION_MAX_RETRIES; attempt++) {
      await new Promise(r => setTimeout(r, CONNECTION_RETRY_DELAY_MS));
      connInfo = readConnectionInfo();
      if (connInfo) break;
    }
    if (!connInfo) {
      throw new Error(
        'Connection file not found. Is Thunderbird running with the MCP extension? ' +
        'The extension must be started first to create the connection file.'
      );
    }
  }

  if (!connInfo.port || !connInfo.token) {
    throw new Error('Invalid connection file: missing port or token');
  }

  const { port, token } = connInfo;

  // Try each host in order - handles platforms where 'localhost' resolves to
  // IPv6 (::1) but the extension only listens on IPv4 (127.0.0.1).
  const tryNext = (hosts) => {
    const [hostname, ...rest] = hosts;
    return tryRequest(hostname, postData, port, token).catch((err) => {
      if (rest.length > 0 && (err.code === 'ECONNREFUSED' || err.code === 'EADDRNOTAVAIL')) {
        return tryNext(rest);
      }
      // On 403 or connection refused, clear cache and retry once with fresh
      // connection info (Thunderbird may have restarted on a new port/token).
      if (!_retried) {
        if (err.message && err.message.includes('403')) {
          clearConnectionCache();
          return forwardToThunderbird(message, true);
        }
        if (err.code === 'ECONNREFUSED' || err.code === 'EADDRNOTAVAIL' || err.code === 'EAFNOSUPPORT') {
          clearConnectionCache();
          return forwardToThunderbird(message, true);
        }
      }
      // Already retried or non-recoverable error
      if (err.code === 'ECONNREFUSED' || err.code === 'EADDRNOTAVAIL' || err.code === 'EAFNOSUPPORT') {
        throw new Error(`Connection failed: ${err.message}. Is Thunderbird running with the MCP extension?`);
      }
      throw err;
    });
  };

  return tryNext(THUNDERBIRD_HOSTS);
}

// Process stdin as JSON-RPC messages
const rl = readline.createInterface({ input: process.stdin, terminal: false });

rl.on('line', (line) => {
  if (!line.trim()) return;

  let messageId = null;
  try {
    messageId = JSON.parse(line).id ?? null;
  } catch {
    // Leave as null when request cannot be parsed
  }

  pendingRequests++;
  handleMessage(line)
    .then(async (response) => {
      if (response !== null) {
        await writeOutput(JSON.stringify(response) + '\n');
      }
    })
    .catch(async (err) => {
      await writeOutput(JSON.stringify({
        jsonrpc: '2.0',
        id: messageId,
        error: { code: -32700, message: `Bridge error: ${err.message}` }
      }) + '\n');
    })
    .finally(() => {
      pendingRequests--;
      checkExit();
    });
});

rl.on('close', () => {
  stdinClosed = true;
  checkExit();
});

process.on('SIGINT', () => process.exit(0));
process.on('SIGTERM', () => process.exit(0));
