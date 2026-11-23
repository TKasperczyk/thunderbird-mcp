const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const { discoverConnectionInfo } = require('./helpers/bridge.cjs');
const { makeAllowedTempDir } = require('./helpers/fixture-dir.cjs');

const TOKEN = 'ab'.repeat(32);

// A Proxy around fs that reports overridden ownership/mode for given paths,
// mirroring the stat override helper used by mcp-bridge.test.cjs.
function makeFsWithStatOverrides(overrides) {
  const openedPaths = new Map();
  return new Proxy({ ...fs }, {
    get(target, prop) {
      if (prop === 'openSync') {
        return (filePath, flags, ...args) => {
          const fd = target.openSync(filePath, flags, ...args);
          openedPaths.set(fd, filePath);
          return fd;
        };
      }
      if (prop === 'closeSync') {
        return (fd) => {
          openedPaths.delete(fd);
          return target.closeSync(fd);
        };
      }
      if (prop === 'statSync' || prop === 'fstatSync' || prop === 'lstatSync') {
        return (fileOrFd, ...args) => {
          const stat = target[prop](fileOrFd, ...args);
          const key = typeof fileOrFd === 'string' ? fileOrFd : openedPaths.get(fileOrFd);
          const override = key != null ? overrides.get(key) : undefined;
          if (!override) return stat;
          // Proxy over the real stat: methods stay bound to the original so
          // isFile() keeps seeing the real file-type bits.
          return new Proxy(stat, {
            get(innerTarget, innerProp) {
              if (Object.prototype.hasOwnProperty.call(override, innerProp)) {
                return override[innerProp];
              }
              const value = innerTarget[innerProp];
              return typeof value === 'function' ? value.bind(innerTarget) : value;
            },
          });
        };
      }
      return target[prop];
    },
  });
}

// One mountinfo line for a 9p Windows drive (WSL2) mounted at `mountpoint`,
// plus the decoy lines that must NOT count as Windows drives.
function wslMountinfo(mountpoint) {
  return [
    `88 85 0:73 / ${mountpoint} rw,noatime - 9p C:\\134 rw,aname=drvfs;path=C:\\;uid=1234;gid=1234`,
    // WSLG driver share: 9p but aname=drivers, not a Windows drive.
    '90 85 0:75 / /mnt/wslg/drivers ro - 9p / ro,aname=drivers;fmask=222;dmask=222',
    // Ordinary local mount: must never be exempted.
    '91 85 0:76 / /home strict rw,relatime - ext4 /dev/sdd rw,discard',
  ].join('\n');
}

function makeContext(root, { file, mode, mountinfo, uid = 1234 }) {
  const homeDir = path.join(root, 'home');
  fs.mkdirSync(homeDir, { recursive: true });
  return {
    env: { THUNDERBIRD_MCP_CONNECTION_FILE: file },
    fsImpl: makeFsWithStatOverrides(new Map([
      [file, { uid, mode }],
      [path.dirname(file), { uid }],
    ])),
    homeDir,
    osImpl: { tmpdir: () => path.join(root, 'tmp'), homedir: () => homeDir },
    pathImpl: path,
    platform: 'linux',
    // Empty fixture proc: auto mountinfo reads fail, keeping strict checks.
    procRoot: path.join(root, 'proc'),
    processImpl: { env: {}, platform: 'linux', kill: () => {} },
    runtimeDir: path.join(root, 'runtime'),
    uid,
    mountinfo,
  };
}

function writeConnectionFile(filePath, { port, mode = 0o600 }) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, JSON.stringify({ port, token: TOKEN, pid: process.pid }), { mode });
}

describe('connection file on a WSL Windows drive mount', () => {
  let root;

  beforeEach(() => {
    root = makeAllowedTempDir('tb-mcp-wsl-mount-');
  });

  afterEach(() => {
    fs.rmSync(root, { recursive: true, force: true });
  });

  it('accepts a group/other-readable file pinned on a drvfs 9p mount (WSL2)', () => {
    const file = path.join(root, 'mnt', 'c', 'connection.json');
    writeConnectionFile(file, { port: 22000, mode: 0o644 });
    const options = makeContext(root, {
      file,
      mode: 0o644,
      mountinfo: wslMountinfo(path.join(root, 'mnt', 'c')),
    });

    const result = discoverConnectionInfo(options);
    assert.equal(result.candidates.length, 1, result.attempts.map(a => a.reason).join('; '));
    assert.equal(result.candidates[0].data.token, TOKEN);
  });

  it('still enforces 0600 when no Windows drive mount is visible', () => {
    const file = path.join(root, 'plain', 'connection.json');
    writeConnectionFile(file, { port: 22001, mode: 0o644 });
    const options = makeContext(root, {
      file,
      mode: 0o644,
      mountinfo: undefined,
    });

    const result = discoverConnectionInfo(options);
    assert.equal(result.candidates.length, 0);
    assert.match(result.attempts.find(a => a.path === file).reason, /owner-only/);
  });

  it('still enforces 0600 on 9p shares that are not Windows drives', () => {
    const file = path.join(root, 'wslg', 'drivers', 'connection.json');
    writeConnectionFile(file, { port: 22002, mode: 0o644 });
    const options = makeContext(root, {
      file,
      mode: 0o644,
      mountinfo: [
        '90 85 0:75 / /mnt/wslg/drivers ro - 9p / ro,aname=drivers;fmask=222;dmask=222',
      ].join('\n'),
    });

    const result = discoverConnectionInfo(options);
    assert.equal(result.candidates.length, 0);
    assert.match(result.attempts.find(a => a.path === file).reason, /owner-only/);
  });

  it('exemption matches the mountpoint exactly, not as a path prefix', () => {
    const file = path.join(root, 'mnt', 'c2', 'connection.json');
    writeConnectionFile(file, { port: 22003, mode: 0o644 });
    const options = makeContext(root, {
      file,
      mode: 0o644,
      mountinfo: wslMountinfo(path.join(root, 'mnt', 'c')),
    });

    const result = discoverConnectionInfo(options);
    assert.equal(result.candidates.length, 0);
    assert.match(result.attempts.find(a => a.path === file).reason, /owner-only/);
  });

  it('keeps accepting a properly owned 0600 file on a Windows drive mount', () => {
    const file = path.join(root, 'mnt', 'c', 'connection.json');
    writeConnectionFile(file, { port: 22004, mode: 0o600 });
    const options = makeContext(root, {
      file,
      mode: 0o600,
      mountinfo: wslMountinfo(path.join(root, 'mnt', 'c')),
    });

    const result = discoverConnectionInfo(options);
    assert.equal(result.candidates.length, 1);
    assert.equal(result.candidates[0].data.token, TOKEN);
  });
});
