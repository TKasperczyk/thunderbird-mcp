const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

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
} = require('./helpers/bridge.cjs');

function makeTempRoot() {
  // Windows' default temp directory is inside denied AppData. Keep allowed
  // attachment fixtures in a unique local workspace directory in that case.
  const base = isSensitiveFilePath(os.tmpdir()) ? path.resolve(__dirname, '..') : os.tmpdir();
  return fs.mkdtempSync(path.join(base, 'tb-mcp-bridge-'));
}

function cleanupTempRoot(root) {
  fs.rmSync(root, { recursive: true, force: true });
}

function writeConnectionFile(filePath, { port, token, pid = process.pid }) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, JSON.stringify({ port, token, pid }), { encoding: 'utf8', mode: 0o600 });
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
    fsImpl: overrides.fsImpl || makeFsWithStatOverrides(new Map()),
    homeDir,
    osImpl: overrides.osImpl || {
      tmpdir: () => tmpDir,
      homedir: () => homeDir,
    },
    pathImpl: overrides.pathImpl || path,
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
  const openedPaths = new Map();
  // Simulated POSIX discovery also runs on Windows, whose native fs lacks these
  // flags and reports synthetic modes. Emulate only those OS differences here.
  const absentFlags = (fs.constants.O_NOFOLLOW ? 0 : 0x20000000) |
    (fs.constants.O_NONBLOCK ? 0 : 0x40000000);
  return new Proxy({ ...fs }, {
    get(target, prop) {
      if (prop === 'constants') return {
        ...target.constants,
        O_NOFOLLOW: target.constants.O_NOFOLLOW || 0x20000000,
        O_NONBLOCK: target.constants.O_NONBLOCK || 0x40000000,
      };
      if (prop === 'openSync') {
        return (filePath, flags, ...args) => {
          const fd = target.openSync(filePath, flags & ~absentFlags, ...args);
          openedPaths.set(fd, filePath);
          return fd;
        };
      }
      if (prop === 'closeSync') {
        return fd => {
          openedPaths.delete(fd);
          return target.closeSync(fd);
        };
      }
      if (prop === 'statSync' || prop === 'fstatSync') {
        return (fileOrFd, ...args) => {
          const stat = target[prop](fileOrFd, ...args);
          const filePath = prop === 'fstatSync' ? openedPaths.get(fileOrFd) : fileOrFd;
          const override = {
            ...(process.platform === 'win32' ? { mode: 0o600 } : {}),
            ...overrides.get(filePath),
          };
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

  it('rejects network and device namespaces before any filesystem call', async (t) => {
    const touched = [];
    for (const method of ['realpathSync', 'statSync', 'lstatSync', 'existsSync', 'openSync', 'readFileSync']) {
      t.mock.method(fs, method, () => { touched.push(method); throw new Error('unexpected filesystem access'); });
    }
    for (const method of ['stat', 'lstat', 'open', 'readFile', 'realpath']) {
      t.mock.method(fs.promises, method, async () => { touched.push(method); throw new Error('unexpected filesystem access'); });
    }
    for (const filePath of ['\\\\server\\share\\file.txt', '//server/share/file.txt', '\\\\?\\C:\\report.txt', '\\\\.\\C:\\report.txt', '//?/C:/report.txt', '//./C:/report.txt']) {
      await assert.rejects(inlineAttachmentPaths({ attachments: [filePath] }), /Sensitive attachment path blocked/);
    }
    assert.deepEqual(touched, []);
  });

  it('rejects Windows streams and trailing dots/spaces before any filesystem access', async () => {
    const touched = [];
    const fsImpl = new Proxy({}, { get(_target, method) { touched.push(method); throw new Error('unexpected filesystem access'); } });
    const platform = 'win32';
    for (const file of [
      'C:\\Keys\\backup.pem::$DATA', 'C:\\Keys\\file.pem:stream',
      'C:/Keys/file.txt:stream', 'C:/Keys/secret.pem.', 'C:/Keys/secret.pem ',
      'C:/Keys. /report.txt', 'C:/Keys /report.txt',
      'file.pem:stream', '/Keys/secret.pem.',
    ]) {
      await assert.rejects(inlineAttachmentPaths({ attachments: [file] }, { platform, fsImpl }), /Sensitive attachment path blocked/);
    }
    assert.deepEqual(touched, []);
  });

  it('allows legal POSIX colon and trailing-dot/space names', { skip: process.platform === 'win32' }, async () => {
    for (const platform of ['linux', 'darwin']) {
      for (const name of ['a:b:c.txt', 'report.txt.', 'report.txt ', 'file.pem:stream']) {
        const file = path.join(root, name);
        fs.writeFileSync(file, 'POSIX attachment');
        const args = { attachments: [file] };
        await inlineAttachmentPaths(args, {
          platform,
          fsImpl: process.platform === 'linux' ? fs : { ...fs, readlinkSync: () => fs.realpathSync(file) },
        });
        assert.equal(args.attachments[0].base64, Buffer.from('POSIX attachment').toString('base64'));
      }
    }
  });

  it('does not treat literal POSIX backslashes as export-directory separators', { skip: process.platform === 'win32' }, async () => {
    const tmpDir = path.join(root, '.temporary');
    fs.mkdirSync(tmpDir);
    const options = makeTestOptions(root, { osImpl: { tmpdir: () => tmpDir } });
    for (const name of ['thunderbird-mcp\\message_1\\auth.json', 'thunderbird-mcp\\message_1/auth.json']) {
      const file = path.join(tmpDir, name);
      fs.mkdirSync(path.dirname(file), { recursive: true });
      fs.writeFileSync(file, 'not an exported attachment');
      await assert.rejects(inlineAttachmentPaths({ attachments: [file] }, options), /Sensitive attachment path blocked/);
    }
  });

  for (const platform of ['darwin', 'linux']) {
    it(`accepts ${platform} exports below a symlinked temp ancestor but rejects redirects below TmpD`, { skip: process.platform === 'win32' }, async () => {
      const canonicalTmp = path.join(root, 'private', 'var', 'folders', '.temporary');
      const lexicalTmp = path.join(root, 'var', 'folders', '.temporary');
      const exportDir = path.join(canonicalTmp, 'thunderbird-mcp');
      fs.mkdirSync(path.join(exportDir, 'message_1'), { recursive: true });
      fs.symlinkSync(path.join(root, 'private', 'var'), path.join(root, 'var'), 'dir');
      fs.writeFileSync(path.join(exportDir, 'message_1', 'report.pdf'), 'exported attachment');
      const options = makeTestOptions(root, { platform, osImpl: { tmpdir: () => lexicalTmp } });
      options.fsImpl = {
        ...fs,
        readlinkSync: fdPath => process.platform === 'linux'
          ? fs.readlinkSync(path.join('/proc/self/fd', path.basename(fdPath)))
          : fs.realpathSync(path.join(exportDir, 'message_1', 'report.pdf')),
      };
      for (const tmp of [lexicalTmp, canonicalTmp]) {
        const args = { attachments: [path.join(tmp, 'thunderbird-mcp', 'message_1', 'report.pdf')] };
        await inlineAttachmentPaths(args, options);
        assert.equal(args.attachments[0].base64, Buffer.from('exported attachment').toString('base64'));
      }
      // A symlink within the export subtree is different from a trusted root alias.
      fs.symlinkSync(path.join(exportDir, 'message_1'), path.join(exportDir, 'message_2'), 'dir');
      await assert.rejects(inlineAttachmentPaths({ attachments: [path.join(lexicalTmp, 'thunderbird-mcp/message_2/report.pdf')] }, options), /export path is redirected/);
      const elsewhere = path.join(canonicalTmp, 'elsewhere');
      fs.renameSync(exportDir, elsewhere);
      fs.symlinkSync(elsewhere, exportDir, 'dir');
      await assert.rejects(inlineAttachmentPaths({ attachments: [path.join(lexicalTmp, 'thunderbird-mcp/message_1/report.pdf')] }, options), /Sensitive attachment path blocked|export path is redirected/);
    });
  }

  it('rejects non-array attachment shapes before filesystem access', async () => {
    const fsImpl = new Proxy({}, { get() { assert.fail('unexpected filesystem access'); } });
    for (const attachments of ['["/tmp/report.txt"]', '[]', {}, 1, true]) {
      await assert.rejects(inlineAttachmentPaths({ attachments }, { fsImpl }), /attachments must be an array/);
    }
  });

  for (const destination of ['denied-directory', 'export-directory']) {
    it(`checks the Linux opened descriptor after a parent swap into ${destination}`, { skip: process.platform !== 'linux' }, async () => {
      const options = makeTestOptions(root);
      const parent = path.join(root, 'Documents');
      const denied = destination === 'denied-directory' ? path.join(root, '.secrets')
        : path.join(options.homeDir, '.var/app/net.thunderbird.Thunderbird/cache/tmp/thunderbird-mcp/message_1');
      fs.mkdirSync(path.dirname(denied), { recursive: true });
      const file = path.join(parent, 'report.txt');
      fs.mkdirSync(parent);
      fs.writeFileSync(file, 'do not read');
      let reads = 0;
      let closed = false;
      options.fsImpl = {
        ...fs,
        readlinkSync(fdPath) {
          assert.equal(path.dirname(fdPath), path.join(options.procRoot, 'self', 'fd'));
          return fs.readlinkSync(path.join('/proc/self/fd', path.basename(fdPath)));
        },
        promises: {
          ...fs.promises,
          async open(filePath, flags) {
            // Retain the inode and size while changing the parent after realpath.
            fs.renameSync(parent, denied);
            fs.symlinkSync(denied, parent, 'dir');
            const handle = await fs.promises.open(filePath, flags);
            return {
              fd: handle.fd,
              stat: () => handle.stat(),
              read() { reads++; assert.fail('refused descriptor must not be read'); },
              async close() { await handle.close(); closed = true; },
            };
          },
        },
      };
      await assert.rejects(inlineAttachmentPaths({ attachments: [file] }, options), destination === 'denied-directory'
        ? /Sensitive or unresolved opened attachment path blocked/ : /Attachment export path is redirected/);
      assert.equal(reads, 0);
      assert.equal(closed, true);
    });
  }

  it('fails closed and closes the file when Linux descriptor resolution fails', async () => {
    const file = path.join(root, 'report.txt');
    fs.writeFileSync(file, 'do not read');
    let closed = false;
    const options = makeTestOptions(root);
    options.fsImpl = {
      ...fs,
      readlinkSync() { throw new Error('proc unavailable'); },
      promises: {
        ...fs.promises,
        async open(filePath, flags) {
          const handle = await fs.promises.open(filePath, flags);
          return {
            fd: handle.fd,
            stat: () => handle.stat(),
            read() { assert.fail('unresolved descriptor must not be read'); },
            async close() { await handle.close(); closed = true; },
          };
        },
      },
    };
    await assert.rejects(inlineAttachmentPaths({ attachments: [file] }, options), error => {
      assert.match(error.message, /Attachment resolve opened file failed/);
      assert.match(error.cause.message, /proc unavailable/);
      return true;
    });
    assert.equal(closed, true);
  });

  for (const layout of ['windows-temp', 'flatpak-cache', 'flatpak-runtime']) {
    it(`reattaches exports from ${layout} without allowing other temp files`, async () => {
      const tmpDir = path.join(root, 'AppData', 'Local', 'Temp');
      const options = makeTestOptions(root, {
        platform: layout === 'windows-temp' ? 'win32' : 'linux',
        osImpl: { tmpdir: () => tmpDir },
      });
      const exportRoot = layout === 'windows-temp' ? path.join(tmpDir, 'thunderbird-mcp')
        : layout === 'flatpak-cache' ? path.join(options.homeDir, '.var/app/net.thunderbird.Thunderbird/cache/tmp/thunderbird-mcp')
          : path.join(options.runtimeDir, 'app/net.thunderbird.Thunderbird/thunderbird-mcp');
      const file = path.join(exportRoot, 'message_1', 'report.pdf');
      fs.mkdirSync(path.dirname(file), { recursive: true });
      fs.writeFileSync(file, 'exported attachment');
      // Simulate /proc/self/fd on any host; no process discovery is involved.
      options.fsImpl = { ...fs, readlinkSync: () => file };
      const args = { attachments: [file] };
      await inlineAttachmentPaths(args, options);
      assert.equal(args.attachments[0].base64, Buffer.from('exported attachment').toString('base64'));
      const refused = [
        path.join(exportRoot, 'connection.json'),
        path.join(exportRoot, 'message_1', 'secret.pem'),
        path.join(exportRoot, 'message_1', '.env'),
        path.join(exportRoot, 'message_1', 'connection.json'),
        path.join(exportRoot, 'message_1') + '/../connection.json',
        path.join(exportRoot, 'message_1') + '/../message_1/report.pdf',
      ];
      if (layout !== 'flatpak-runtime') {
        refused.push(path.join(exportRoot, 'other.txt'), path.join(exportRoot, 'attachments', 'report.pdf'));
      }
      for (const denied of refused) {
        await assert.rejects(inlineAttachmentPaths({ attachments: [denied] }, options), /Sensitive attachment path blocked/);
      }
    });
  }

  it('does not exempt exports from unrelated Flatpak applications', async () => {
    const options = makeTestOptions(root);
    const file = path.join(options.homeDir, '.var/app/org.example.App/cache/tmp/thunderbird-mcp/message_1/report.pdf');
    await assert.rejects(inlineAttachmentPaths({ attachments: [file] }, options), /Sensitive attachment path blocked/);
  });

  it('refuses symlinks into and out of an export directory', async (t) => {
    const options = makeTestOptions(root);
    const exportRoot = path.join(options.homeDir, '.var/app/net.thunderbird.Thunderbird/cache/tmp/thunderbird-mcp');
    const exported = path.join(exportRoot, 'message_1');
    const documents = path.join(root, 'Documents');
    const alias = path.join(root, 'export-alias');
    fs.mkdirSync(exported, { recursive: true });
    fs.mkdirSync(documents);
    fs.writeFileSync(path.join(exported, 'report.pdf'), 'exported attachment');
    fs.writeFileSync(path.join(documents, 'report.pdf'), 'outside export directory');
    try {
      const type = process.platform === 'win32' ? 'junction' : 'dir';
      fs.symlinkSync(documents, path.join(exportRoot, 'message_2'), type);
      fs.symlinkSync(exported, alias, type);
    } catch (error) {
      if (error.code === 'EPERM' || error.code === 'EACCES') return t.skip('symlinks unavailable');
      throw error;
    }
    for (const file of [path.join(exportRoot, 'message_2', 'report.pdf'), path.join(alias, 'report.pdf')]) {
      await assert.rejects(inlineAttachmentPaths({ attachments: [file] }, options), /Attachment export path is redirected/);
    }
  });

  it('rejects a symlinked or junction parent that resolves into a denied directory', async (t) => {
    const denied = path.join(root, '.secrets');
    fs.mkdirSync(denied);
    fs.writeFileSync(path.join(denied, 'report.pdf'), 'credential fixture');
    const alias = path.join(root, 'Documents');
    try {
      fs.symlinkSync(denied, alias, process.platform === 'win32' ? 'junction' : 'dir');
    } catch (error) {
      if (error.code === 'EPERM' || error.code === 'EACCES') return t.skip('symlinks unavailable');
      throw error;
    }
    await assert.rejects(inlineAttachmentPaths({ attachments: [path.join(alias, 'report.pdf')] }), /Sensitive attachment path blocked/);
  });

  it('lists every refused path and leaves a mixed request unchanged', async () => {
    const valid = path.join(root, 'report.pdf');
    const missing = path.join(root, 'missing.pdf');
    const denied = path.join(root, '.env');
    fs.writeFileSync(valid, 'safe');
    const args = { attachments: [valid, denied, missing] };
    await assert.rejects(inlineAttachmentPaths(args), error => {
      assert.ok(error.message.includes(denied));
      assert.ok(error.message.includes(missing));
      return true;
    });
    assert.deepEqual(args.attachments, [valid, denied, missing]);
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
      token: '01'.repeat(32),
    });
    writeConnectionFile(options.env.THUNDERBIRD_MCP_CONNECTION_FILE, {
      port: 20002,
      token: '02'.repeat(32),
    });

    const connInfo = readConnectionInfo(options);
    assert.deepStrictEqual(connInfo, {
      port: 20002,
      token: '02'.repeat(32),
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
      token: '03'.repeat(32),
    });

    const exePath = path.join(options.procRoot, '4242', 'exe');
    options.fsImpl = { ...options.fsImpl, realpathSync: file => {
      assert.equal(file, exePath);
      return '/snap/thunderbird/123/usr/lib/thunderbird/thunderbird';
    } };
    const connInfo = readConnectionInfo(options);
    assert.equal(connInfo.port, 20003);
    assert.equal(connInfo.token, '03'.repeat(32));
  });

  it('ignores spoofed Snap argv[0] and TMPDIR unless exe belongs to Thunderbird Snap', () => {
    const options = makeTestOptions(root);
    const procDir = path.join(options.procRoot, '4242');
    const decoyTmpDir = path.join(root, 'decoy-tmp');
    fs.mkdirSync(path.join(options.homeDir, 'snap', 'thunderbird'), { recursive: true });
    fs.mkdirSync(procDir, { recursive: true });
    fs.writeFileSync(path.join(procDir, 'cmdline'), 'thunderbird\0');
    fs.writeFileSync(path.join(procDir, 'environ'), `TMPDIR=${decoyTmpDir}\0`);
    writeConnectionFile(path.join(decoyTmpDir, 'thunderbird-mcp', 'connection.json'), { port: 29999, token: '04'.repeat(32) });
    for (const executable of ['/app/bin/thunderbird', '/usr/bin/thunderbird', '/snap/thunderbird-evil/123/thunderbird']) {
      let environmentReads = 0;
      const result = discoverConnectionInfo({ ...options, fsImpl: {
        ...options.fsImpl,
        realpathSync(file) { assert.equal(file, path.join(procDir, 'exe')); return executable; },
        readFileSync(file, ...args) {
          if (file === path.join(procDir, 'environ')) environmentReads++;
          return fs.readFileSync(file, ...args);
        },
      } });
      assert.equal(result.candidates.length, 0, executable);
      assert.equal(environmentReads, 0, executable);
    }
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
      token: '04'.repeat(32),
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
      token: '05'.repeat(32),
    });

    const connInfo = readConnectionInfo(options);
    assert.equal(connInfo.port, 20004);
    assert.equal(connInfo.token, '05'.repeat(32));
  });

  it('flatpak scan finds a connection file under ~/.var/app/*/cache/tmp', () => {
    const options = makeTestOptions(root, { platform: 'linux' });

    const flatpakConnFile = path.join(
      options.homeDir,
      '.var',
      'app',
      'net.thunderbird.Thunderbird',
      'cache',
      'tmp',
      'thunderbird-mcp',
      'connection.json'
    );
    writeConnectionFile(flatpakConnFile, {
      port: 20005,
      token: '06'.repeat(32),
    });

    const connInfo = readConnectionInfo(options);
    assert.equal(connInfo.port, 20005);
    assert.equal(connInfo.token, '06'.repeat(32));
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
      token: '07'.repeat(32),
    });
    writeConnectionFile(foreignConnFile, {
      port: 20006,
      token: '08'.repeat(32),
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
    assert.equal(connInfo.token, '07'.repeat(32));
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
      token: '09'.repeat(32),
    });

    const connInfo = readConnectionInfo(options);
    assert.equal(connInfo.port, 20007);
    assert.equal(connInfo.token, '09'.repeat(32));
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
      token: '0a'.repeat(32),
    });

    // Flatpak runtime file (later group, also valid)
    writeConnectionFile(
      path.join(options.runtimeDir, 'app', 'org.mozilla.thunderbird', 'thunderbird-mcp', 'connection.json'),
      { port: 20101, token: '0b'.repeat(32) }
    );

    const result = discoverConnectionInfo(options);
    assert.ok(result.candidates.length >= 2, `expected >=2 candidates, got ${result.candidates.length}`);
    assert.equal(result.candidates[0].data.token, '0a'.repeat(32));
    assert.ok(result.candidates.some(c => c.data.token === '0b'.repeat(32)));
  });

  it('advanceToNextCandidate walks the cached list, then returns null when exhausted', () => {
    const options = makeTestOptions(root, {
      platform: 'linux',
      runtimeDir: path.join(root, 'runtime'),
    });

    writeConnectionFile(path.join(root, 'tmp', 'thunderbird-mcp', 'connection.json'), {
      port: 20200,
      token: '0c'.repeat(32),
    });
    writeConnectionFile(
      path.join(options.runtimeDir, 'app', 'org.mozilla.thunderbird', 'thunderbird-mcp', 'connection.json'),
      { port: 20201, token: '0d'.repeat(32) }
    );

    const first = readConnectionInfo(options);
    assert.equal(first.token, '0c'.repeat(32));

    const second = advanceToNextCandidate();
    assert.ok(second, 'should advance to a second candidate');
    assert.equal(second.token, '0d'.repeat(32));

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
      env: { THUNDERBIRD_MCP_CONNECTION_FILE: path.join(root, '0e'.repeat(32), 'connection.json') },
    });

    writeConnectionFile(options.env.THUNDERBIRD_MCP_CONNECTION_FILE, {
      port: 20300,
      token: '0e'.repeat(32),
    });
    writeConnectionFile(path.join(root, 'tmp', 'thunderbird-mcp', 'connection.json'), {
      port: 20301,
      token: '0f'.repeat(32),
    });

    const result = discoverConnectionInfo(options);
    assert.equal(result.candidates.length, 1);
    assert.equal(result.candidates[0].data.token, '0e'.repeat(32));
  });
  for (const layout of ['native', 'pin', 'snap', 'flatpak-runtime', 'flatpak-cache', 'macOS']) {
    for (const [name, override, reason] of [
      ['foreign owner', { uid: 1235 }, /not owned by the current user/],
      ['group-readable', { mode: 0o640 }, /owner-only/],
      ['world-writable', { mode: 0o602 }, /owner-only/],
      ['valid owner-only', {}, /^ok$/],
    ]) {
      it(`${layout} validates ${name} on the opened descriptor`, () => {
        const options = makeTestOptions(root, { platform: layout === 'macOS' ? 'darwin' : 'linux', uid: 1234 });
        let file;
        switch (layout) {
          case 'native': file = path.join(root, 'tmp', 'thunderbird-mcp', 'connection.json'); break;
          case 'pin':
            file = path.join(root, 'pinned.json');
            options.env.THUNDERBIRD_MCP_CONNECTION_FILE = file;
            break;
          case 'snap':
            fs.mkdirSync(options.procRoot, { recursive: true });
            fs.mkdirSync(path.join(options.homeDir, 'snap', 'thunderbird'), { recursive: true });
            file = path.join(options.homeDir, 'Downloads', 'thunderbird.tmp', 'thunderbird-mcp', 'connection.json');
            break;
          case 'flatpak-runtime': file = path.join(options.runtimeDir, 'app', 'org.mozilla.thunderbird', 'thunderbird-mcp', 'connection.json'); break;
          case 'flatpak-cache': file = path.join(options.homeDir, '.var', 'app', 'net.thunderbird.Thunderbird', 'cache', 'tmp', 'thunderbird-mcp', 'connection.json'); break;
          case 'macOS': file = path.join(options.darwinFoldersRoot, 'aa', 'bb', 'T', 'thunderbird-mcp', 'connection.json'); break;
        }
        writeConnectionFile(file, { port: 21000, token: 'a'.repeat(64) });
        options.fsImpl = makeFsWithStatOverrides(new Map([[file, { uid: 1234, mode: 0o600, ...override }]]));
        const before = fs.readFileSync(file);
        const result = discoverConnectionInfo(options);
        assert.equal(result.candidates.length, name === 'valid owner-only' ? 1 : 0);
        assert.match(result.attempts.find(attempt => attempt.path === file).reason, reason);
        assert.deepEqual(fs.readFileSync(file), before, 'rejected files must remain untouched');
      });
    }
  }

  it('Windows skips synthetic uid/mode while still validating file type and size', () => {
    const file = path.join(root, 'pinned.json');
    writeConnectionFile(file, { port: 21000, token: 'a'.repeat(64) });
    const options = makeTestOptions(root, { platform: 'win32', uid: null, env: { THUNDERBIRD_MCP_CONNECTION_FILE: file } });
    for (const [override, accepted] of [
      [{ uid: 9999, mode: 0o777 }, true],
      [{ isFile: () => false }, false],
      [{ size: 4097 }, false],
    ]) {
      options.fsImpl = makeFsWithStatOverrides(new Map([[file, override]]));
      assert.equal(discoverConnectionInfo(options).candidates.length, accepted ? 1 : 0);
    }
  });

  for (const kind of ['directory', 'oversized', 'symlink', 'FIFO']) {
    it(`rejects a ${kind} connection pin without falling back`, (t) => {
      const file = path.join(root, 'pinned.json');
      const options = makeTestOptions(root, { env: { THUNDERBIRD_MCP_CONNECTION_FILE: file } });
      writeConnectionFile(path.join(root, 'tmp', 'thunderbird-mcp', 'connection.json'), { port: 21000, token: 'a'.repeat(64) });
      if (kind === 'directory') fs.mkdirSync(file);
      if (kind === 'oversized') fs.writeFileSync(file, ' '.repeat(4097), { mode: 0o600 });
      if (kind === 'FIFO') {
        if (process.platform === 'win32') return t.skip('POSIX FIFO');
        execFileSync('mkfifo', [file]);
      }
      if (kind === 'symlink') {
        if (process.platform === 'win32') return t.skip('POSIX O_NOFOLLOW');
        try { fs.symlinkSync(path.join(root, 'tmp', 'thunderbird-mcp', 'connection.json'), file, 'file'); }
        catch (error) {
          if (error.code === 'EPERM' || error.code === 'EACCES') return t.skip('symlinks unavailable');
          throw error;
        }
      }
      const result = discoverConnectionInfo(options);
      assert.equal(result.candidates.length, 0);
      assert.equal(result.attempts.length, 1);
      assert.notEqual(result.attempts[0].reason, 'ok');
    });
  }

  it('rejects invalid connection data before caching, even for explicit pins', () => {
    const file = path.join(root, 'pinned.json');
    const options = makeTestOptions(root, { env: { THUNDERBIRD_MCP_CONNECTION_FILE: file } });
    for (const data of [null, [], {}, { port: 21000, token: 'invalid' }, { port: '21000', token: 'a'.repeat(64) }, { port: 65536, token: 'a'.repeat(64) }, { port: 1.5, token: 'a'.repeat(64) }]) {
      fs.writeFileSync(file, JSON.stringify(data), { mode: 0o600 });
      assert.equal(readConnectionInfo(options), null);
      assert.equal(advanceToNextCandidate(), null);
    }
    writeConnectionFile(file, { port: 21000, token: 'a'.repeat(64) });
    assert.equal(readConnectionInfo(options).port, 21000);
  });

  for (const appId of ['org.mozilla.Thunderbird', 'org.mozilla.thunderbird', 'org.mozilla.thunderbird_esr', 'net.thunderbird.Thunderbird', 'eu.betterbird.Betterbird', 'org.example.Unrelated', 'org.mozilla.Thunderbird.attacker']) {
    for (const layout of ['runtime', 'cache']) {
      it(`Flatpak ${layout} ${appId} uses the exact application allow-list`, () => {
        const options = makeTestOptions(root, { runtimeDir: layout === 'cache' ? null : path.join(root, 'runtime') });
        const file = layout === 'runtime'
          ? path.join(options.runtimeDir, 'app', appId, 'thunderbird-mcp', 'connection.json')
          : path.join(options.homeDir, '.var', 'app', appId, 'cache', 'tmp', 'thunderbird-mcp', 'connection.json');
        writeConnectionFile(file, { port: 21000, token: 'a'.repeat(64) });
        const result = discoverConnectionInfo(options);
        const allowed = !['org.example.Unrelated', 'org.mozilla.Thunderbird.attacker'].includes(appId);
        assert.equal(result.candidates.length, allowed ? 1 : 0);
        if (!allowed) assert.ok(result.attempts.some(attempt => attempt.reason === 'application ID not allowed'));
      });
    }
  }

  it('does not reopen a candidate replaced after fstat', () => {
    const file = path.join(root, 'pinned.json');
    writeConnectionFile(file, { port: 21000, token: 'a'.repeat(64) });
    const baseFs = makeFsWithStatOverrides(new Map());
    const fsImpl = new Proxy(baseFs, { get(target, prop) {
      if (prop === 'fstatSync') return fd => {
        const stat = target.fstatSync(fd);
        fs.renameSync(file, path.join(root, 'original.json'));
        writeConnectionFile(file, { port: 21001, token: 'b'.repeat(64) });
        return stat;
      };
      return target[prop];
    }});
    const result = discoverConnectionInfo(makeTestOptions(root, { fsImpl, env: { THUNDERBIRD_MCP_CONNECTION_FILE: file } }));
    assert.equal(result.candidates[0].data.port, 21000);
    assert.equal(JSON.parse(fs.readFileSync(file, 'utf8')).port, 21001);
  });

  it('reads only the opened fd and closes it on success and all validation/read failures', { skip: process.platform === 'win32' }, () => {
    const file = path.join(root, 'pinned.json');
    writeConnectionFile(file, { port: 21000, token: 'a'.repeat(64) });
    for (const failure of [null, 'fstat', 'read', 'owner', 'size', 'json', 'growth']) {
      let opened;
      let opens = 0;
      let closed = 0;
      let reads = 0;
      const fsImpl = new Proxy(fs, { get(target, prop) {
        if (prop === 'openSync') return (filename, flags) => {
          assert.equal(filename, file);
          opens++;
          assert.ok(flags & fs.constants.O_NOFOLLOW);
          assert.ok(flags & fs.constants.O_NONBLOCK);
          opened = target.openSync(filename, flags);
          return opened;
        };
        if (prop === 'fstatSync') return fd => {
          assert.equal(fd, opened);
          if (failure === 'fstat') throw new Error('fstat failed');
          const stat = target.fstatSync(fd);
          if (failure === 'owner') stat.uid += 1;
          if (failure === 'size') stat.size = 5000;
          return stat;
        };
        if (prop === 'readFileSync') return () => assert.fail('path reopened');
        if (prop === 'readSync') return (fd, buffer, offset, length, position) => {
          reads++;
          assert.equal(fd, opened);
          if (failure === 'read') throw new Error('read failed');
          if (failure === 'growth') { buffer.fill(32, offset, offset + length); return length; }
          if (failure === 'json') { if (reads > 1) return 0; buffer[offset] = 123; return 1; }
          return target.readSync(fd, buffer, offset, length, position);
        };
        if (prop === 'closeSync') return fd => { assert.equal(fd, opened); closed++; target.closeSync(fd); };
        return target[prop];
      }});
      const result = discoverConnectionInfo(makeTestOptions(root, { fsImpl, env: { THUNDERBIRD_MCP_CONNECTION_FILE: file } }));
      assert.equal(result.candidates.length, failure ? 0 : 1, failure);
      assert.equal(opens, 1, failure);
      assert.equal(closed, 1, failure);
      if (failure === 'owner' || failure === 'size' || failure === 'fstat') assert.equal(reads, 0);
    }
  });

});
