const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const {
  buildConnectionDiscoveryErrorMessage,
  clearConnectionCache,
  readConnectionInfo,
} = require('../mcp-bridge.cjs');

function makeTempRoot() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'tb-mcp-bridge-'));
}

function cleanupTempRoot(root) {
  fs.rmSync(root, { recursive: true, force: true });
}

function writeConnectionFile(filePath, { port, token, pid = process.pid }) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, JSON.stringify({ port, token, pid }), 'utf8');
}

function makeTestOptions(root, overrides = {}) {
  const uid = typeof process.getuid === 'function' ? process.getuid() : 0;
  const homeDir = path.join(root, 'home');
  const tmpDir = path.join(root, 'tmp');
  const runtimeDir = path.join(root, 'runtime');

  fs.mkdirSync(homeDir, { recursive: true });
  fs.mkdirSync(tmpDir, { recursive: true });
  fs.mkdirSync(runtimeDir, { recursive: true });

  return {
    env: overrides.env || {},
    fsImpl: overrides.fsImpl || fs,
    homeDir,
    osImpl: overrides.osImpl || {
      tmpdir: () => tmpDir,
      homedir: () => homeDir,
    },
    pathImpl: path,
    platform: overrides.platform || 'linux',
    procRoot: overrides.procRoot || path.join(root, 'proc'),
    processImpl: overrides.processImpl || { env: overrides.env || {}, platform: overrides.platform || 'linux' },
    runtimeDir: Object.prototype.hasOwnProperty.call(overrides, 'runtimeDir')
      ? overrides.runtimeDir
      : runtimeDir,
    uid: Object.prototype.hasOwnProperty.call(overrides, 'uid') ? overrides.uid : uid,
    darwinFoldersRoot: overrides.darwinFoldersRoot || path.join(root, 'var', 'folders'),
  };
}

function makeFsWithStatOverrides(overrides) {
  return new Proxy(fs, {
    get(target, prop) {
      if (prop === 'statSync') {
        return (filePath, ...args) => {
          const stat = target.statSync(filePath, ...args);
          const override = overrides.get(filePath);
          if (!override) {
            return stat;
          }
          return new Proxy(stat, {
            get(innerTarget, innerProp) {
              if (Object.prototype.hasOwnProperty.call(override, innerProp)) {
                return override[innerProp];
              }
              const value = innerTarget[innerProp];
              return typeof value === 'function' ? value.bind(innerTarget) : value;
            }
          });
        };
      }

      const value = target[prop];
      return typeof value === 'function' ? value.bind(target) : value;
    }
  });
}

describe('Bridge discovery', () => {
  let root;

  beforeEach(() => {
    clearConnectionCache();
    root = makeTempRoot();
  });

  afterEach(() => {
    clearConnectionCache();
    cleanupTempRoot(root);
  });

  it('env var override takes priority', () => {
    const options = makeTestOptions(root, {
      env: { THUNDERBIRD_MCP_CONNECTION_FILE: path.join(root, 'env', 'connection.json') },
    });

    writeConnectionFile(path.join(root, 'tmp', 'thunderbird-mcp', 'connection.json'), {
      port: 20001,
      token: 'native-token',
    });
    writeConnectionFile(options.env.THUNDERBIRD_MCP_CONNECTION_FILE, {
      port: 20002,
      token: 'env-token',
    });

    const connInfo = readConnectionInfo(options);
    assert.deepStrictEqual(connInfo, {
      port: 20002,
      token: 'env-token',
      pid: process.pid,
    });
  });

  it('snap detection works from a mocked /proc tree', () => {
    const options = makeTestOptions(root, {
      platform: 'linux',
      procRoot: path.join(root, 'proc'),
    });

    fs.mkdirSync(path.join(options.homeDir, 'snap', 'thunderbird'), { recursive: true });
    fs.mkdirSync(path.join(options.procRoot, '4242'), { recursive: true });
    fs.writeFileSync(
      path.join(options.procRoot, '4242', 'cmdline'),
      'snap/thunderbird\0--some-flag',
      'utf8'
    );

    const snapTmpDir = path.join(root, 'snap-tmp');
    fs.writeFileSync(
      path.join(options.procRoot, '4242', 'environ'),
      `TMPDIR=${snapTmpDir}\0HOME=${options.homeDir}\0`,
      'utf8'
    );

    writeConnectionFile(path.join(snapTmpDir, 'thunderbird-mcp', 'connection.json'), {
      port: 20003,
      token: 'snap-token',
    });

    const connInfo = readConnectionInfo(options);
    assert.equal(connInfo.port, 20003);
    assert.equal(connInfo.token, 'snap-token');
  });

  it('flatpak scan finds a runtime connection file', () => {
    const options = makeTestOptions(root, {
      platform: 'linux',
      runtimeDir: path.join(root, 'runtime'),
    });

    const flatpakConnFile = path.join(
      options.runtimeDir,
      'app',
      'eu.betterbird.Betterbird',
      'thunderbird-mcp',
      'connection.json'
    );
    writeConnectionFile(flatpakConnFile, {
      port: 20004,
      token: 'flatpak-token',
    });

    const connInfo = readConnectionInfo(options);
    assert.equal(connInfo.port, 20004);
    assert.equal(connInfo.token, 'flatpak-token');
  });

  it('macOS scan finds current uid files and ignores other owners', () => {
    const currentUid = typeof process.getuid === 'function' ? process.getuid() : 1000;
    const darwinRoot = path.join(root, 'var', 'folders');
    const options = makeTestOptions(root, {
      platform: 'darwin',
      darwinFoldersRoot: darwinRoot,
      uid: currentUid,
    });

    const ownedConnFile = path.join(darwinRoot, 'aa', 'bb', 'T', 'thunderbird-mcp', 'connection.json');
    const foreignConnFile = path.join(darwinRoot, 'cc', 'dd', 'T', 'thunderbird-mcp', 'connection.json');

    writeConnectionFile(ownedConnFile, {
      port: 20005,
      token: 'owned-token',
    });
    writeConnectionFile(foreignConnFile, {
      port: 20006,
      token: 'foreign-token',
    });

    const statOverrides = new Map();
    statOverrides.set(foreignConnFile, { uid: currentUid + 1 });

    const connInfo = readConnectionInfo({
      ...options,
      fsImpl: makeFsWithStatOverrides(statOverrides),
    });

    assert.equal(connInfo.port, 20005);
    assert.equal(connInfo.token, 'owned-token');
  });

  it('re-resolves candidates on the next cache miss after a startup race', () => {
    const options = makeTestOptions(root, {
      platform: 'linux',
      runtimeDir: path.join(root, 'runtime'),
    });

    assert.equal(readConnectionInfo(options), null);

    const delayedConnFile = path.join(
      options.runtimeDir,
      'app',
      'org.mozilla.thunderbird',
      'thunderbird-mcp',
      'connection.json'
    );
    writeConnectionFile(delayedConnFile, {
      port: 20007,
      token: 'delayed-token',
    });

    const connInfo = readConnectionInfo(options);
    assert.equal(connInfo.port, 20007);
    assert.equal(connInfo.token, 'delayed-token');
  });

  it('reports useful discovery failures', () => {
    const options = makeTestOptions(root, {
      env: { THUNDERBIRD_MCP_CONNECTION_FILE: path.join(root, 'missing', 'connection.json') },
      platform: 'linux',
    });

    assert.equal(readConnectionInfo(options), null);
    assert.match(buildConnectionDiscoveryErrorMessage(), /THUNDERBIRD_MCP_CONNECTION_FILE/);
    assert.match(buildConnectionDiscoveryErrorMessage(), /file not found/);
  });
});
