/**
 * Shared test helpers for thunderbird-mcp test suite.
 *
 * Consolidates common utilities used across bridge, auth, and stress tests
 * to eliminate duplication and ensure consistent behavior.
 */

const fs = require('fs');
const path = require('path');
const os = require('os');
const net = require('net');
const { spawn } = require('child_process');

const BRIDGE_PATH = path.resolve(__dirname, '..', 'mcp-bridge.cjs');
const CONN_DIR = path.join(os.tmpdir(), 'thunderbird-mcp');
const CONN_FILE = path.join(CONN_DIR, 'connection.json');
const DEFAULT_PORT = 8765;

/**
 * Check if the default MCP port is already in use (e.g. real Thunderbird).
 */
function isDefaultPortInUse() {
  return new Promise((resolve) => {
    const sock = net.createConnection({ port: DEFAULT_PORT, host: '127.0.0.1' });
    sock.on('connect', () => { sock.destroy(); resolve(true); });
    sock.on('error', () => resolve(false));
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
 * Back up any existing connection file to avoid interfering with a running Thunderbird.
 * Returns the saved data (or null if no file existed).
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
      reject(new Error(`Bridge timed out after ${timeout}ms. stdout: ${stdout}, stderr: ${stderr}`));
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
          reject(new Error(`Failed to parse bridge output: ${lines[0]}`));
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
      if (!stdout.trim()) {
        // For notifications, the bridge returns nothing and exits cleanly
        resolve(null);
      }
    });

    child.stdin.write(JSON.stringify(message) + '\n');
    // For notifications, close stdin immediately so the process exits
    if (!Object.prototype.hasOwnProperty.call(message, 'id') ||
        (typeof message.method === 'string' && message.method.startsWith('notifications/'))) {
      child.stdin.end();
    }
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
      for (let i = 0; i < lines.length - 1; i++) {
        const line = lines[i].trim();
        if (line) {
          try { responses.push(JSON.parse(line)); } catch { /* skip malformed */ }
        }
      }
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
      if (stdout.trim()) {
        try { responses.push(JSON.parse(stdout.trim())); } catch { /* skip */ }
      }
      resolve(responses);
    });

    for (const msg of messages) {
      child.stdin.write(JSON.stringify(msg) + '\n');
    }
    setTimeout(() => {
      try { child.stdin.end(); } catch { /* ignore */ }
    }, 500);
  });
}

/**
 * Replicate sanitizeJson from bridge for thorough testing.
 * Must be kept in sync with mcp-bridge.cjs sanitizeJson.
 */
function sanitizeJson(data) {
  let sanitized = data.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, '');
  sanitized = sanitized.replace(/(?<!\\)\r/g, '\\r');
  sanitized = sanitized.replace(/(?<!\\)\n/g, '\\n');
  sanitized = sanitized.replace(/(?<!\\)\t/g, '\\t');
  return sanitized;
}

/**
 * Replicate pagination logic from api.js for unit testing.
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
 * Create a validator function from tool schemas.
 * Replicates validateToolArgs from api.js.
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

module.exports = {
  BRIDGE_PATH,
  CONN_DIR,
  CONN_FILE,
  DEFAULT_PORT,
  isDefaultPortInUse,
  writeTestConnectionInfo,
  backupConnectionFile,
  restoreConnectionFile,
  sendToBridge,
  sendMultipleToBridge,
  sanitizeJson,
  paginateResults,
  createValidator,
};
