const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const {
  buildConnectionDiscoveryErrorMessage,
  clearConnectionCache,
  detectWslGatewayIp,
  isWsl,
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

  it('snap detection ignores decoy processes with thunderbird only as a file arg', () => {
    const options = makeTestOptions(root, {
      platform: 'linux',
      procRoot: path.join(root, 'proc'),
    });

    fs.mkdirSync(path.join(options.homeDir, 'snap', 'thunderbird'), { recursive: true });

    // Decoy: a text editor opened on "thunderbird.txt". argv[0] is /usr/bin/vim,
    // argv[1] contains 'thunderbird' as a substring. Must NOT be picked up.
    const decoyPid = '9999';
    fs.mkdirSync(path.join(options.procRoot, decoyPid), { recursive: true });
    fs.writeFileSync(
      path.join(options.procRoot, decoyPid, 'cmdline'),
      '/usr/bin/vim\0/home/user/thunderbird.txt\0',
      'utf8'
    );
    const decoyTmpDir = path.join(root, 'decoy-tmp');
    fs.writeFileSync(
      path.join(options.procRoot, decoyPid, 'environ'),
      `TMPDIR=${decoyTmpDir}\0`,
      'utf8'
    );
    // If the decoy was picked up, the bridge would read this file and succeed.
    writeConnectionFile(path.join(decoyTmpDir, 'thunderbird-mcp', 'connection.json'), {
      port: 29999,
      token: 'attacker-token',
    });

    const connInfo = readConnectionInfo(options);
    // No real Thunderbird process in our mocked /proc -> discovery returns null.
    // Critically, the decoy's TMPDIR file is NOT selected.
    assert.equal(connInfo, null);
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

function makeMockFs(fileContents) {
  return {
    readFileSync(filePath, encoding) {
      if (Object.prototype.hasOwnProperty.call(fileContents, filePath)) {
        return fileContents[filePath];
      }
      const err = new Error('ENOENT: no such file or directory');
      err.code = 'ENOENT';
      throw err;
    },
  };
}

describe('WSL detection', () => {
  it('returns true when /proc/version contains WSL', () => {
    const mockFs = makeMockFs({
      '/proc/version': 'Linux version 5.15.0-microsoft-standard-WSL2',
    });
    assert.equal(isWsl({ fsImpl: mockFs }), true);
  });

  it('returns false for Azure Linux (has "microsoft" but no "WSL")', () => {
    const mockFs = makeMockFs({
      '/proc/version': 'Linux version 5.15.0-microsoft-standard-Azure',
    });
    assert.equal(isWsl({ fsImpl: mockFs }), false);
  });

  it('returns false for generic WSL string without "microsoft"', () => {
    const mockFs = makeMockFs({
      '/proc/version': 'Linux version 6.0-wsl-custom',
    });
    assert.equal(isWsl({ fsImpl: mockFs }), false);
  });

  it('returns false when /proc/version has no WSL markers', () => {
    const mockFs = makeMockFs({
      '/proc/version': 'Linux version 6.8.0-generic (Ubuntu)',
    });
    assert.equal(isWsl({ fsImpl: mockFs }), false);
  });

  it('returns false when /proc/version is missing', () => {
    const mockFs = makeMockFs({});
    assert.equal(isWsl({ fsImpl: mockFs }), false);
  });
});

describe('WSL gateway detection', () => {
  const wslVersion = 'Linux version 5.15.0-microsoft-standard-WSL2';

  it('returns gateway IP from resolv.conf when in WSL', () => {
    const mockFs = makeMockFs({
      '/proc/version': wslVersion,
      '/etc/resolv.conf': 'nameserver 172.20.100.1\nsearch localdomain\n',
    });
    const ip = detectWslGatewayIp({ fsImpl: mockFs });
    assert.equal(ip, '172.20.100.1');
  });

  it('returns null when not in WSL', () => {
    const mockFs = makeMockFs({
      '/proc/version': 'Linux version 6.8.0-generic (Ubuntu)',
    });
    const ip = detectWslGatewayIp({ fsImpl: mockFs });
    assert.equal(ip, null);
  });

  it('returns null when /proc/version is missing', () => {
    const mockFs = makeMockFs({});
    const ip = detectWslGatewayIp({ fsImpl: mockFs });
    assert.equal(ip, null);
  });

  it('returns null when resolv.conf has no nameserver line', () => {
    const mockFs = makeMockFs({
      '/proc/version': wslVersion,
      '/etc/resolv.conf': 'search localdomain\noptions timeout:2\n',
    });
    const ip = detectWslGatewayIp({ fsImpl: mockFs });
    assert.equal(ip, null);
  });

  it('returns null when resolv.conf is empty', () => {
    const mockFs = makeMockFs({
      '/proc/version': wslVersion,
      '/etc/resolv.conf': '',
    });
    const ip = detectWslGatewayIp({ fsImpl: mockFs });
    assert.equal(ip, null);
  });

  it('handles nameserver with trailing whitespace', () => {
    const mockFs = makeMockFs({
      '/proc/version': wslVersion,
      '/etc/resolv.conf': 'nameserver  192.168.1.1   \n',
    });
    const ip = detectWslGatewayIp({ fsImpl: mockFs });
    assert.equal(ip, '192.168.1.1');
  });

  it('uses first nameserver when multiple are present', () => {
    const mockFs = makeMockFs({
      '/proc/version': wslVersion,
      '/etc/resolv.conf': 'nameserver 10.0.0.1\nnameserver 10.0.0.2\n',
    });
    const ip = detectWslGatewayIp({ fsImpl: mockFs });
    assert.equal(ip, '10.0.0.1');
  });

  it('returns null on read error other than ENOENT', () => {
    const mockFs = {
      readFileSync() {
        const err = new Error('EACCES: permission denied');
        err.code = 'EACCES';
        throw err;
      },
    };
    const ip = detectWslGatewayIp({ fsImpl: mockFs });
    assert.equal(ip, null);
  });
});
