const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const {
  advanceToNextCandidate,
  attachmentLimits,
  buildConnectionDiscoveryErrorMessage,
  compactToolResultJsonText,
  clearConnectionCache,
  discoverConnectionInfo,
  inlineAttachmentPaths,
  isSensitiveFilePath,
  isValidAuthToken,
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

describe('Auth token validation', () => {
  it('accepts 64 lowercase hex characters', () => {
    assert.equal(isValidAuthToken('a'.repeat(64)), true);
    assert.equal(isValidAuthToken('0123456789abcdef'.repeat(4)), true);
  });

  it('rejects non-generated token shapes', () => {
    assert.equal(isValidAuthToken(''), false);
    assert.equal(isValidAuthToken(' '.repeat(64)), false);
    assert.equal(isValidAuthToken('a'.repeat(63)), false);
    assert.equal(isValidAuthToken('a'.repeat(65)), false);
    assert.equal(isValidAuthToken('A'.repeat(64)), false);
    assert.equal(isValidAuthToken('g'.repeat(64)), false);
    assert.equal(isValidAuthToken(`${'a'.repeat(64)}\n`), false);
    assert.equal(isValidAuthToken(null), false);
  });
});

describe('Tool result serialization', () => {
  it('compacts JSON text content without mutating the original response', () => {
    const originalText = JSON.stringify([{ name: 'INBOX', unreadMessages: 3 }], null, 2);
    const response = {
      jsonrpc: '2.0',
      id: 2,
      result: {
        content: [{
          type: 'text',
          text: originalText,
        }],
      },
    };

    const compacted = compactToolResultJsonText(response);

    assert.notStrictEqual(compacted, response);
    assert.deepStrictEqual(
      JSON.parse(compacted.result.content[0].text),
      JSON.parse(originalText)
    );
    assert.equal(response.result.content[0].text, originalText);
    assert.equal(compacted.result.content[0].text.includes('\n'), false);
  });

  it('preserves image content blocks while compacting the leading JSON text', () => {
    const imageBlock = {
      type: 'image',
      data: 'iVBORw0KGgo=',
      mimeType: 'image/png',
    };
    const response = {
      jsonrpc: '2.0',
      id: 137,
      result: {
        content: [
          { type: 'text', text: JSON.stringify({ inlineImages: 1 }, null, 2) },
          imageBlock,
        ],
      },
    };

    const compacted = compactToolResultJsonText(response);

    assert.equal(compacted.result.content[0].text, '{"inlineImages":1}');
    assert.strictEqual(compacted.result.content[1], imageBlock);
    assert.strictEqual(response.result.content[1], imageBlock);
  });
});

describe('Bridge attachment path policy', () => {
  let root;

  beforeEach(() => {
    root = makeTempRoot();
  });

  afterEach(() => {
    cleanupTempRoot(root);
  });

  it('rejects sensitive paths before trying to inline them', async () => {
    const sensitivePath = path.join(root, '.ssh', 'id_rsa');
    const args = { attachments: [sensitivePath] };

    await assert.rejects(
      inlineAttachmentPaths(args),
      error => {
        assert.match(error.message, /Sensitive attachment path blocked/);
        assert.ok(error.message.includes(sensitivePath), error.message);
        return true;
      }
    );
    assert.deepEqual(args.attachments, [sensitivePath]);
  });

  it('uses the same normalized deny-list rules for Windows paths', () => {
    assert.equal(isSensitiveFilePath('C:\\Users\\alice\\.ssh\\id_ed25519'), true);
    assert.equal(isSensitiveFilePath('C:\\Users\\alice\\Downloads\\report.pdf'), false);
  });

  it('rejects attachment counts above the extension limit before filesystem access', async () => {
    const attachments = Array.from(
      { length: attachmentLimits.MAX_ATTACHMENTS_PER_MESSAGE + 1 },
      (_, index) => ({ name: `inline-${index}.txt`, base64: 'QQ==' })
    );

    await assert.rejects(
      inlineAttachmentPaths({ attachments }),
      new RegExp(
        `Attachment count ${attachments.length} exceeds the ` +
        `${attachmentLimits.MAX_ATTACHMENTS_PER_MESSAGE} attachment limit`
      )
    );
  });

  it('rejects an aggregate of path attachments above 50 MB before reading', async () => {
    const sizes = [18, 18, 15].map(mib => mib * 1024 * 1024);
    const attachments = sizes.map((size, index) => {
      const filePath = path.join(root, `aggregate-${index}.bin`);
      fs.writeFileSync(filePath, '');
      fs.truncateSync(filePath, size);
      return filePath;
    });
    const args = { attachments };

    await assert.rejects(
      inlineAttachmentPaths(args),
      error => {
        assert.match(error.message, /50 MB aggregate attachment limit/);
        assert.ok(error.message.includes(attachments[2]), error.message);
        return true;
      }
    );
    assert.deepEqual(args.attachments, attachments);
  });

  it('rejects oversized and non-regular files during preflight', async () => {
    const oversizedPath = path.join(root, 'oversized.bin');
    fs.writeFileSync(oversizedPath, '');
    fs.truncateSync(oversizedPath, attachmentLimits.MAX_ATTACHMENT_BYTES + 1);

    await assert.rejects(
      inlineAttachmentPaths({ attachments: [oversizedPath] }),
      error => error.message.includes(`Attachment too large: ${oversizedPath}`)
    );
    await assert.rejects(
      inlineAttachmentPaths({ attachments: [root] }),
      error => error.message.includes(`Attachment is not a regular file: ${root}`)
    );
  });

  it('rejects symlinked attachment paths', async (t) => {
    const targetPath = path.join(root, 'target.txt');
    const symlinkPath = path.join(root, 'attachment.txt');
    fs.writeFileSync(targetPath, 'safe attachment', 'utf8');
    try {
      fs.symlinkSync(targetPath, symlinkPath, 'file');
    } catch (error) {
      if (error.code === 'EPERM' || error.code === 'EACCES') {
        t.skip(`symlinks unavailable: ${error.code}`);
        return;
      }
      throw error;
    }

    await assert.rejects(
      inlineAttachmentPaths({ attachments: [symlinkPath] }),
      error => {
        assert.match(error.message, /symlink/);
        assert.ok(error.message.includes(symlinkPath), error.message);
        return true;
      }
    );
  });

  it('inlines allowed files and leaves inline objects untouched', async () => {
    const filePath = path.join(root, 'report.txt');
    const inline = { name: 'already-inline.txt', base64: 'QQ==' };
    fs.writeFileSync(filePath, 'hello', 'utf8');
    const args = { attachments: [filePath, inline] };

    await inlineAttachmentPaths(args);

    assert.deepEqual(args.attachments, [
      {
        name: 'report.txt',
        contentType: 'text/plain',
        base64: Buffer.from('hello').toString('base64'),
      },
      inline,
    ]);
  });
});

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
    // Pin a synthetic uid rather than process.getuid(). On Windows the real
    // fs.statSync reports uid=0 for every file regardless of the caller, so we
    // can't rely on stat.uid matching process.getuid() — both files are stat-
    // overridden below so the uid-filter logic is exercised on any platform.
    const currentUid = 1000;
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
    // Force both stat results: the owned file to the caller's uid and the
    // foreign one to a different uid, so the filter is tested independent of
    // what the host's real fs.statSync returns.
    statOverrides.set(ownedConnFile, { uid: currentUid });
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

  it('discoverConnectionInfo collects every valid candidate, not just the winner', () => {
    const options = makeTestOptions(root, {
      platform: 'linux',
      runtimeDir: path.join(root, 'runtime'),
    });

    // Native /tmp file (first group, winner)
    writeConnectionFile(path.join(root, 'tmp', 'thunderbird-mcp', 'connection.json'), {
      port: 20100,
      token: 'native',
    });

    // Flatpak runtime file (later group, also valid)
    writeConnectionFile(
      path.join(options.runtimeDir, 'app', 'org.mozilla.thunderbird', 'thunderbird-mcp', 'connection.json'),
      { port: 20101, token: 'flatpak' }
    );

    const result = discoverConnectionInfo(options);
    assert.ok(result.candidates.length >= 2, `expected >=2 candidates, got ${result.candidates.length}`);
    assert.equal(result.candidates[0].data.token, 'native');
    assert.ok(result.candidates.some(c => c.data.token === 'flatpak'));
  });

  it('advanceToNextCandidate walks the cached list, then returns null when exhausted', () => {
    const options = makeTestOptions(root, {
      platform: 'linux',
      runtimeDir: path.join(root, 'runtime'),
    });

    writeConnectionFile(path.join(root, 'tmp', 'thunderbird-mcp', 'connection.json'), {
      port: 20200,
      token: 'first',
    });
    writeConnectionFile(
      path.join(options.runtimeDir, 'app', 'org.mozilla.thunderbird', 'thunderbird-mcp', 'connection.json'),
      { port: 20201, token: 'second' }
    );

    const first = readConnectionInfo(options);
    assert.equal(first.token, 'first');

    const second = advanceToNextCandidate();
    assert.ok(second, 'should advance to a second candidate');
    assert.equal(second.token, 'second');

    const third = advanceToNextCandidate();
    assert.equal(third, null, 'should return null after the last candidate');
  });

  it('advanceToNextCandidate returns null when no cache exists', () => {
    clearConnectionCache();
    assert.equal(advanceToNextCandidate(), null);
  });

  it('hard-pinned env override does not collect autodiscovery candidates as fallbacks', () => {
    const options = makeTestOptions(root, {
      platform: 'linux',
      runtimeDir: path.join(root, 'runtime'),
      env: { THUNDERBIRD_MCP_CONNECTION_FILE: path.join(root, 'pinned', 'connection.json') },
    });

    writeConnectionFile(options.env.THUNDERBIRD_MCP_CONNECTION_FILE, {
      port: 20300,
      token: 'pinned',
    });
    writeConnectionFile(path.join(root, 'tmp', 'thunderbird-mcp', 'connection.json'), {
      port: 20301,
      token: 'should-not-appear',
    });

    const result = discoverConnectionInfo(options);
    assert.equal(result.candidates.length, 1);
    assert.equal(result.candidates[0].data.token, 'pinned');
  });
});
