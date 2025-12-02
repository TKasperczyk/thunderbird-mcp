#!/usr/bin/env node
/**
 * MCP Bridge: Converts stdio MCP protocol to HTTP requests for Thunderbird MCP extension
 *
 * This bridge allows Claude Code to communicate with the Thunderbird MCP extension
 * which exposes an HTTP endpoint on localhost:8765
 */

const http = require('http');
const readline = require('readline');

const THUNDERBIRD_PORT = 8765;

// Track pending requests to wait for them before exiting
let pendingRequests = 0;
let stdinClosed = false;

function checkExit() {
  if (stdinClosed && pendingRequests === 0) {
    process.exit(0);
  }
}

// Create readline interface for persistent line-by-line reading
const rl = readline.createInterface({
  input: process.stdin,
  terminal: false
});

// Process each line as a JSON-RPC message
rl.on('line', (line) => {
  if (!line.trim()) return;

  pendingRequests++;
  handleMessage(line).then(response => {
    process.stdout.write(JSON.stringify(response) + '\n');
  }).catch(err => {
    const errorResponse = {
      jsonrpc: '2.0',
      id: null,
      error: { code: -32700, message: `Bridge error: ${err.message}` }
    };
    process.stdout.write(JSON.stringify(errorResponse) + '\n');
  }).finally(() => {
    pendingRequests--;
    checkExit();
  });
});

rl.on('close', () => {
  stdinClosed = true;
  checkExit();
});

async function handleMessage(line) {
  const message = JSON.parse(line);

  // Handle MCP protocol methods locally
  if (message.method === 'initialize') {
    return {
      jsonrpc: '2.0',
      id: message.id,
      result: {
        protocolVersion: '2024-11-05',
        capabilities: { tools: {} },
        serverInfo: {
          name: 'thunderbird-mcp',
          version: '0.1.0'
        }
      }
    };
  }

  if (message.method === 'notifications/initialized') {
    // This is a notification, no response needed but we return empty for safety
    return { jsonrpc: '2.0', id: message.id, result: {} };
  }

  // Forward everything else to Thunderbird
  return forwardToThunderbird(message);
}

function forwardToThunderbird(message) {
  return new Promise((resolve, reject) => {
    const postData = JSON.stringify(message);

    const req = http.request({
      hostname: 'localhost',
      port: THUNDERBIRD_PORT,
      path: '/',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(postData)
      }
    }, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          reject(new Error(`Invalid JSON from Thunderbird: ${data}`));
        }
      });
    });

    req.on('error', (e) => {
      reject(new Error(`Connection failed: ${e.message}. Is Thunderbird running with the MCP extension?`));
    });

    req.setTimeout(30000, () => {
      req.destroy();
      reject(new Error('Request to Thunderbird timed out'));
    });

    req.write(postData);
    req.end();
  });
}

// Handle process signals gracefully
process.on('SIGINT', () => process.exit(0));
process.on('SIGTERM', () => process.exit(0));
