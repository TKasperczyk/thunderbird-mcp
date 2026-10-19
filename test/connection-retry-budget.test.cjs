/**
 * Tests for the connection retry budgets.
 *
 * Covers:
 * - Defaults for delay, warm budget and cold-start budget
 * - Environment overrides, including invalid values falling back
 * - The cold-start budget never being shorter than the warm one
 * - Fail-closed behaviour still terminating with a pinned, missing file
 */

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const { makeAllowedTempDir } = require('./helpers/fixture-dir.cjs');

const BRIDGE_PATH = path.resolve(__dirname, '..', 'mcp-bridge.cjs');

/**
 * Load the bridge with a specific environment and return its exported
 * budgets. The module reads the environment once at load time, so the require
 * cache has to be dropped between cases.
 */
function loadBudgets(env = {}) {
  const saved = {};
  const keys = [
    'THUNDERBIRD_MCP_CONNECTION_RETRY_DELAY_MS',
    'THUNDERBIRD_MCP_CONNECTION_RETRIES',
    'THUNDERBIRD_MCP_COLD_START_CONNECTION_RETRIES',
    'THUNDERBIRD_MCP_NO_AUTOSTART',
  ];
  for (const key of keys) {
    saved[key] = process.env[key];
  }
  try {
    for (const key of keys) {
      delete process.env[key];
    }
    // Requiring the module must not start the stdio loop.
    process.env.THUNDERBIRD_MCP_NO_AUTOSTART = '1';
    Object.assign(process.env, env);
    delete require.cache[require.resolve(BRIDGE_PATH)];
    return require(BRIDGE_PATH).connectionRetryBudgets;
  } finally {
    for (const key of keys) {
      if (saved[key] === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = saved[key];
      }
    }
    delete require.cache[require.resolve(BRIDGE_PATH)];
  }
}

describe('Connection retry budgets: defaults', () => {
  it('keeps the historical warm budget and delay', () => {
    const budgets = loadBudgets();
    assert.equal(budgets.CONNECTION_RETRY_DELAY_MS, 1000);
    assert.equal(budgets.CONNECTION_MAX_RETRIES, 5);
  });

  it('gives a cold start a longer budget than a warm one', () => {
    const budgets = loadBudgets();
    assert.equal(budgets.CONNECTION_COLD_START_MAX_RETRIES, 20);
    assert.ok(
      budgets.CONNECTION_COLD_START_MAX_RETRIES > budgets.CONNECTION_MAX_RETRIES,
      'cold start must outlast the warm budget'
    );
  });
});

describe('Connection retry budgets: environment overrides', () => {
  it('accepts non-negative integers', () => {
    const budgets = loadBudgets({
      THUNDERBIRD_MCP_CONNECTION_RETRY_DELAY_MS: '250',
      THUNDERBIRD_MCP_CONNECTION_RETRIES: '2',
      THUNDERBIRD_MCP_COLD_START_CONNECTION_RETRIES: '40',
    });
    assert.equal(budgets.CONNECTION_RETRY_DELAY_MS, 250);
    assert.equal(budgets.CONNECTION_MAX_RETRIES, 2);
    assert.equal(budgets.CONNECTION_COLD_START_MAX_RETRIES, 40);
  });

  it('accepts zero to disable retrying', () => {
    const budgets = loadBudgets({
      THUNDERBIRD_MCP_CONNECTION_RETRIES: '0',
      THUNDERBIRD_MCP_COLD_START_CONNECTION_RETRIES: '0',
    });
    assert.equal(budgets.CONNECTION_MAX_RETRIES, 0);
    assert.equal(budgets.CONNECTION_COLD_START_MAX_RETRIES, 0);
  });

  for (const bad of ['abc', '-1', '1.5', '', '  ', 'Infinity']) {
    it(`falls back to the default for ${JSON.stringify(bad)}`, () => {
      const budgets = loadBudgets({ THUNDERBIRD_MCP_CONNECTION_RETRIES: bad });
      assert.equal(budgets.CONNECTION_MAX_RETRIES, 5);
    });
  }

  it('never lets the cold-start budget fall below the warm budget', () => {
    const budgets = loadBudgets({
      THUNDERBIRD_MCP_CONNECTION_RETRIES: '30',
      THUNDERBIRD_MCP_COLD_START_CONNECTION_RETRIES: '3',
    });
    assert.equal(budgets.CONNECTION_MAX_RETRIES, 30);
    assert.equal(budgets.CONNECTION_COLD_START_MAX_RETRIES, 30);
  });
});

describe('Connection retry budgets: fail-closed is still bounded', () => {
  it('errors out instead of hanging when the pinned file never appears', async () => {
    const dir = makeAllowedTempDir('tb-mcp-retry-');
    const connFile = path.join(dir, 'connection.json');
    try {
      const response = await sendToBridge({
        jsonrpc: '2.0',
        id: 7,
        method: 'tools/list',
      }, {
        THUNDERBIRD_MCP_CONNECTION_FILE: connFile,
        THUNDERBIRD_MCP_CONNECTION_RETRIES: '1',
        THUNDERBIRD_MCP_COLD_START_CONNECTION_RETRIES: '2',
        THUNDERBIRD_MCP_CONNECTION_RETRY_DELAY_MS: '10',
      });
      assert.equal(response.id, 7);
      assert.ok(response.error, 'a missing pinned file must fail closed');
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});

/**
 * Helper: send one JSON-RPC message to a freshly spawned bridge.
 */
function sendToBridge(message, env, { timeout = 10000 } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [BRIDGE_PATH], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: { ...process.env, ...env },
    });

    let stdout = '';
    let stderr = '';
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error(`Bridge timed out. stdout: ${stdout}, stderr: ${stderr}`));
    }, timeout);

    child.stdout.on('data', (data) => {
      stdout += data.toString();
      const lines = stdout.split('\n').filter((l) => l.trim());
      if (lines.length > 0) {
        clearTimeout(timer);
        child.stdin.end();
        try {
          resolve(JSON.parse(lines[0]));
        } catch {
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

    child.stdin.write(JSON.stringify(message) + '\n');
  });
}
