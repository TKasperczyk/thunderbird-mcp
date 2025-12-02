#!/usr/bin/env node
/**
 * MCP Bridge: Converts stdio MCP protocol to HTTP requests for Thunderbird MCP extension
 *
 * This bridge allows Claude Code to communicate with the Thunderbird MCP extension
 * which exposes an HTTP endpoint on localhost:8765
 */

const http = require('http');

const THUNDERBIRD_PORT = 8765;

// Read all of stdin
let input = '';
process.stdin.setEncoding('utf8');

process.stdin.on('readable', () => {
  let chunk;
  while ((chunk = process.stdin.read()) !== null) {
    input += chunk;
  }
});

process.stdin.on('end', async () => {
  const lines = input.trim().split('\n').filter(l => l.trim());

  for (const line of lines) {
    try {
      const message = JSON.parse(line);
      const response = await forwardToThunderbird(message);
      process.stdout.write(JSON.stringify(response) + '\n');
    } catch (e) {
      const errorResponse = {
        jsonrpc: '2.0',
        id: null,
        error: { code: -32700, message: `Bridge error: ${e.message}` }
      };
      process.stdout.write(JSON.stringify(errorResponse) + '\n');
    }
  }
});

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
