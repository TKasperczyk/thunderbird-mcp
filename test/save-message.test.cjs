const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const {
  handleMessage,
  sanitizeSaveFilename,
  SAVE_MESSAGE_ENV,
  SAVE_MESSAGE_TOOL,
  SAVE_ROOT_ENV,
  saveMessageToDisk,
} = require('../mcp-bridge.cjs');

// The tool is opt-in, so every test that expects it to do anything has to turn
// it on explicitly. Passing env per call (rather than mutating process.env)
// keeps the tests order-independent.
const ENABLED = { [SAVE_MESSAGE_ENV]: '1' };

// Tracks temp dirs created during a test so afterEach can remove them.
let tempDirs = [];

function mkTemp() {
  // realpath because macOS hands out /var/... symlinks for /private/var, and
  // saveMessageToDisk resolves destDir before returning it.
  const dir = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'tb-mcp-save-')));
  tempDirs.push(dir);
  return dir;
}

// Wrap a tool's data object in the JSON-RPC/content envelope the bridge expects
// (the same shape forwardToThunderbird returns for a tools/call).
function makeEnvelope(data) {
  return {
    jsonrpc: '2.0',
    id: 'mock',
    result: { content: [{ type: 'text', text: JSON.stringify(data) }] }
  };
}

// saveMessageToDisk with the capability gate on by default.
function save(opts) {
  return saveMessageToDisk({ env: ENABLED, ...opts });
}

function modeOf(p) {
  return fs.lstatSync(p).mode & 0o777;
}

describe('saveMessageToDisk', () => {
  beforeEach(() => {
    tempDirs = [];
  });

  afterEach(() => {
    for (const dir of tempDirs) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
    tempDirs = [];
  });

  it('writes the .eml with the exact rawSource and a subject-derived filename', async () => {
    const dir = mkTemp();
    const raw = 'Subject: Hello World\r\nFrom: a@b.com\r\n\r\nBody line one\r\nBody line two\r\n';

    let seenArgs = null;
    const fetchMessage = async (sub) => {
      seenArgs = sub.params.arguments;
      return makeEnvelope({ rawSource: raw, subject: 'Hello World' });
    };

    const result = await save({
      args: { messageId: 'mid-1', folderPath: 'imap://x/INBOX', destDir: dir },
      fetchMessage
    });

    assert.equal(seenArgs.rawSource, true);
    assert.equal(seenArgs.messageId, 'mid-1');
    const expected = path.join(dir, 'Hello World.eml');
    assert.equal(result.emlPath, expected);
    assert.equal(fs.readFileSync(expected, 'latin1'), raw);
    assert.deepEqual(result.attachments, []);
    assert.deepEqual(result.skippedAttachments, []);
    assert.equal(result.destDir, dir);
  });

  it('names the file from the decoded subject field, not the raw Subject header', async () => {
    const dir = mkTemp();
    // Thunderbird hands over mime2DecodedSubject; the raw header still carries
    // the encoded word. The bridge must not re-derive it from the headers.
    const raw = 'Subject: =?UTF-8?B?R3LDvMOfZQ==?=\r\n\r\nbody';

    const fetchMessage = async () => makeEnvelope({ rawSource: raw, subject: 'Grüße' });
    const result = await save({
      args: { messageId: 'mid-2', folderPath: 'fp', destDir: dir },
      fetchMessage
    });

    assert.equal(result.emlPath, path.join(dir, 'Grüße.eml'));
    assert.ok(fs.existsSync(result.emlPath));
  });

  // Regression test for the UTF-8 write that expanded every byte above 0x7f.
  // rawSource is a Latin-1 byte string, so the bytes on disk must come back
  // out identical -- 0x80 0xff must not become c2 80 c3 bf.
  it('writes raw 8-bit bytes unchanged instead of re-encoding them as UTF-8', async () => {
    const dir = mkTemp();
    const raw = 'Subject: Binary\r\n\r\n' + Buffer.from([0x80, 0xff, 0xe9, 0x00, 0x7f]).toString('latin1');

    const fetchMessage = async () => makeEnvelope({ rawSource: raw, subject: 'Binary' });
    const result = await save({
      args: { messageId: 'mid-bin', folderPath: 'fp', destDir: dir },
      fetchMessage
    });

    const onDisk = fs.readFileSync(result.emlPath);
    assert.deepEqual(
      [...onDisk.subarray(-5)],
      [0x80, 0xff, 0xe9, 0x00, 0x7f]
    );
    assert.equal(onDisk.length, Buffer.byteLength(raw, 'latin1'));
  });

  it('honors emlFilename and appends .eml when absent', async () => {
    const dir = mkTemp();
    const raw = 'Subject: Ignored Subject\r\n\r\nbody';

    const fetchMessage = async () => makeEnvelope({ rawSource: raw, subject: 'Ignored Subject' });
    const result = await save({
      args: { messageId: 'mid-3', folderPath: 'fp', destDir: dir, emlFilename: 'custom-name' },
      fetchMessage
    });

    assert.equal(result.emlPath, path.join(dir, 'custom-name.eml'));
    assert.equal(fs.readFileSync(result.emlPath, 'latin1'), raw);
  });

  it('falls back to messageId when there is no subject', async () => {
    const dir = mkTemp();
    const raw = 'From: a@b.com\r\n\r\nbody without a subject';

    const fetchMessage = async () => makeEnvelope({ rawSource: raw });
    const result = await save({
      args: { messageId: 'fallback-id', folderPath: 'fp', destDir: dir },
      fetchMessage
    });

    assert.equal(result.emlPath, path.join(dir, 'fallback-id.eml'));
  });

  it('creates a missing nested destDir, private to the user', async () => {
    const dir = path.join(mkTemp(), 'a', 'b', 'c');
    const raw = 'Subject: Nested\r\n\r\nbody';

    const fetchMessage = async () => makeEnvelope({ rawSource: raw, subject: 'Nested' });
    const result = await save({
      args: { messageId: 'mid-4', folderPath: 'fp', destDir: dir },
      fetchMessage
    });

    assert.ok(fs.existsSync(dir));
    assert.equal(result.emlPath, path.join(dir, 'Nested.eml'));
    // Saved mail must not be group- or world-readable.
    assert.equal(modeOf(dir) & 0o077, 0);
    assert.equal(modeOf(result.emlPath) & 0o077, 0);
  });

  it('overwrite:false throws on an existing file, overwrite:true replaces it', async () => {
    const dir = mkTemp();
    const raw1 = 'Subject: Dup\r\n\r\nfirst';
    const raw2 = 'Subject: Dup\r\n\r\nsecond';

    const fetchFirst = async () => makeEnvelope({ rawSource: raw1, subject: 'Dup' });
    const first = await save({
      args: { messageId: 'mid-5', folderPath: 'fp', destDir: dir },
      fetchMessage: fetchFirst
    });
    assert.equal(fs.readFileSync(first.emlPath, 'latin1'), raw1);

    const fetchSecond = async () => makeEnvelope({ rawSource: raw2, subject: 'Dup' });
    await assert.rejects(
      save({
        args: { messageId: 'mid-5', folderPath: 'fp', destDir: dir },
        fetchMessage: fetchSecond
      }),
      (e) => {
        assert.match(e.message, /already exists/);
        // The underlying EEXIST is preserved rather than swallowed.
        assert.equal(e.cause && e.cause.code, 'EEXIST');
        return true;
      }
    );
    // The failed non-overwrite write must not have clobbered the original.
    assert.equal(fs.readFileSync(first.emlPath, 'latin1'), raw1);

    const replaced = await save({
      args: { messageId: 'mid-5', folderPath: 'fp', destDir: dir, overwrite: true },
      fetchMessage: fetchSecond
    });
    assert.equal(replaced.emlPath, first.emlPath);
    assert.equal(fs.readFileSync(replaced.emlPath, 'latin1'), raw2);
  });

  it('copies attachments into destDir by name, suffixes collisions, and skips path-less entries', async () => {
    const dir = mkTemp();
    const srcDir = mkTemp();
    const src1 = path.join(srcDir, 'temp-a.bin');
    const src2 = path.join(srcDir, 'temp-b.bin');
    fs.writeFileSync(src1, 'AAA');
    fs.writeFileSync(src2, 'BBB');

    const fetchMessage = async (sub) => {
      assert.equal(sub.params.arguments.saveAttachments, true);
      return makeEnvelope({
        attachments: [
          { name: 'report.pdf', filePath: src1, contentType: 'application/pdf', size: 3 },
          { name: 'report.pdf', filePath: src2, contentType: 'application/pdf', size: 3 },
          { name: 'inline-only.txt' }
        ]
      });
    };

    const result = await save({
      args: { messageId: 'mid-6', folderPath: 'fp', destDir: dir, saveEml: false, saveAttachments: true },
      fetchMessage
    });

    const p1 = path.join(dir, 'report.pdf');
    const p2 = path.join(dir, 'report (2).pdf');
    assert.deepEqual(result.attachments, [p1, p2]);
    assert.equal(fs.readFileSync(p1, 'utf8'), 'AAA');
    assert.equal(fs.readFileSync(p2, 'utf8'), 'BBB');
    assert.equal(modeOf(p1) & 0o077, 0);
    assert.deepEqual(result.skippedAttachments, ['inline-only.txt']);
    assert.equal(result.emlPath, null);
  });

  // Regression test for existsSync + copyFileSync: copyFileSync follows a
  // symlink at the destination and writes straight through it, so a planted
  // link turned "save my attachments here" into "overwrite that file there".
  it('never writes through a symlink planted at the destination', async () => {
    const dir = mkTemp();
    const srcDir = mkTemp();
    const outsideDir = mkTemp();
    const src = path.join(srcDir, 'temp.bin');
    const victim = path.join(outsideDir, 'victim.txt');
    fs.writeFileSync(src, 'ATTACKER BYTES');
    fs.writeFileSync(victim, 'ORIGINAL');
    fs.symlinkSync(victim, path.join(dir, 'report.pdf'));

    const fetchMessage = async () => makeEnvelope({
      attachments: [{ name: 'report.pdf', filePath: src }]
    });

    const result = await save({
      args: { messageId: 'mid-sym', folderPath: 'fp', destDir: dir, saveEml: false, saveAttachments: true },
      fetchMessage
    });

    // The link counts as a collision, so the copy lands on a suffixed name.
    assert.deepEqual(result.attachments, [path.join(dir, 'report (2).pdf')]);
    assert.equal(fs.readFileSync(victim, 'utf8'), 'ORIGINAL');
  });

  it('replaces the link, not the link target, when overwrite is on', async () => {
    const dir = mkTemp();
    const srcDir = mkTemp();
    const outsideDir = mkTemp();
    const src = path.join(srcDir, 'temp.bin');
    const victim = path.join(outsideDir, 'victim.txt');
    fs.writeFileSync(src, 'NEW BYTES');
    fs.writeFileSync(victim, 'ORIGINAL');
    const planted = path.join(dir, 'report.pdf');
    fs.symlinkSync(victim, planted);

    const fetchMessage = async () => makeEnvelope({
      attachments: [{ name: 'report.pdf', filePath: src }]
    });

    const result = await save({
      args: {
        messageId: 'mid-sym2', folderPath: 'fp', destDir: dir,
        saveEml: false, saveAttachments: true, overwrite: true
      },
      fetchMessage
    });

    assert.deepEqual(result.attachments, [planted]);
    assert.equal(fs.lstatSync(planted).isSymbolicLink(), false);
    assert.equal(fs.readFileSync(planted, 'utf8'), 'NEW BYTES');
    assert.equal(fs.readFileSync(victim, 'utf8'), 'ORIGINAL');
  });

  // overwrite:true used to bypass the used-name set entirely, so a second
  // attachment with the same name clobbered the first -- and an attachment
  // named like the .eml clobbered the message written moments earlier.
  it('overwrite:true still refuses to clobber files written in the same call', async () => {
    const dir = mkTemp();
    const srcDir = mkTemp();
    const src1 = path.join(srcDir, 'a.bin');
    const src2 = path.join(srcDir, 'b.bin');
    fs.writeFileSync(src1, 'FIRST');
    fs.writeFileSync(src2, 'SECOND');
    const raw = 'Subject: Combined\r\n\r\nmessage body';

    const fetchMessage = async (sub) => {
      if (sub.params.arguments.rawSource) {
        return makeEnvelope({ rawSource: raw, subject: 'Combined' });
      }
      return makeEnvelope({
        attachments: [
          // Same name as the .eml about to be written, then a self-collision.
          { name: 'Combined.eml', filePath: src1 },
          { name: 'dup.bin', filePath: src2 },
          { name: 'dup.bin', filePath: src1 }
        ]
      });
    };

    const result = await save({
      args: {
        messageId: 'mid-clobber', folderPath: 'fp', destDir: dir,
        saveAttachments: true, overwrite: true
      },
      fetchMessage
    });

    // The message survives its own call.
    assert.equal(fs.readFileSync(result.emlPath, 'latin1'), raw);
    assert.equal(result.emlPath, path.join(dir, 'Combined.eml'));
    // Each attachment got its own name.
    assert.deepEqual(result.attachments, [
      path.join(dir, 'Combined (2).eml'),
      path.join(dir, 'dup.bin'),
      path.join(dir, 'dup (2).bin')
    ]);
    assert.equal(fs.readFileSync(path.join(dir, 'dup.bin'), 'utf8'), 'SECOND');
    assert.equal(fs.readFileSync(path.join(dir, 'dup (2).bin'), 'utf8'), 'FIRST');
  });

  it('writes both the .eml and attachments in one call', async () => {
    const dir = mkTemp();
    const srcDir = mkTemp();
    const src1 = path.join(srcDir, 'temp.bin');
    fs.writeFileSync(src1, 'DATA');
    const raw = 'Subject: Combined\r\n\r\nbody';

    const fetchMessage = async (sub) => {
      if (sub.params.arguments.rawSource) {
        return makeEnvelope({ rawSource: raw, subject: 'Combined' });
      }
      return makeEnvelope({ attachments: [{ name: 'doc.txt', filePath: src1 }] });
    };

    const result = await save({
      args: { messageId: 'mid-7', folderPath: 'fp', destDir: dir, saveAttachments: true },
      fetchMessage
    });

    assert.equal(result.emlPath, path.join(dir, 'Combined.eml'));
    assert.equal(fs.readFileSync(result.emlPath, 'latin1'), raw);
    assert.deepEqual(result.attachments, [path.join(dir, 'doc.txt')]);
    assert.equal(fs.readFileSync(path.join(dir, 'doc.txt'), 'utf8'), 'DATA');
  });

  it('propagates a tool-level { error } from getMessage as a thrown Error', async () => {
    const dir = mkTemp();
    const fetchMessage = async () => makeEnvelope({ error: 'Message not found' });

    await assert.rejects(
      save({
        args: { messageId: 'missing', folderPath: 'fp', destDir: dir },
        fetchMessage
      }),
      /Message not found/
    );
  });

  it('throws when rawSource is missing (no offline copy)', async () => {
    const dir = mkTemp();
    const fetchMessage = async () => makeEnvelope({ id: 'mid', subject: 'x' });

    await assert.rejects(
      save({
        args: { messageId: 'mid-8', folderPath: 'fp', destDir: dir },
        fetchMessage
      }),
      /offline/i
    );
  });

  it('validates required string arguments', async () => {
    const dir = mkTemp();
    const fetchMessage = async () => makeEnvelope({ rawSource: 'Subject: x\r\n\r\ny', subject: 'x' });

    await assert.rejects(
      save({ args: { folderPath: 'fp', destDir: dir }, fetchMessage }),
      /messageId/
    );
    await assert.rejects(
      save({ args: { messageId: 'm', destDir: dir }, fetchMessage }),
      /folderPath/
    );
    await assert.rejects(
      save({ args: { messageId: 'm', folderPath: 'fp' }, fetchMessage }),
      /destDir/
    );
  });
});

describe('saveMessage capability gate', () => {
  beforeEach(() => {
    tempDirs = [];
  });

  afterEach(() => {
    for (const dir of tempDirs) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
    tempDirs = [];
  });

  it('refuses to write anything unless explicitly enabled', async () => {
    const dir = mkTemp();
    const fetchMessage = async () => makeEnvelope({ rawSource: 'Subject: x\r\n\r\ny', subject: 'x' });

    await assert.rejects(
      saveMessageToDisk({
        args: { messageId: 'm', folderPath: 'fp', destDir: dir },
        fetchMessage,
        env: {}
      }),
      new RegExp(SAVE_MESSAGE_ENV)
    );
    assert.deepEqual(fs.readdirSync(dir), []);
  });

  it('is absent from tools/list when disabled and present when enabled', async () => {
    const forward = async () => ({
      jsonrpc: '2.0',
      id: 1,
      result: { tools: [{ name: 'getMessage' }] }
    });
    const line = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list' });

    const off = await handleMessage(line, { forward, env: {} });
    assert.deepEqual(off.result.tools.map((t) => t.name), ['getMessage']);

    const on = await handleMessage(line, { forward, env: ENABLED });
    assert.deepEqual(on.result.tools.map((t) => t.name), ['getMessage', SAVE_MESSAGE_TOOL.name]);
  });

  it('returns a JSON-RPC error envelope for a disabled tools/call', async () => {
    const dir = mkTemp();
    const forward = async () => makeEnvelope({ rawSource: 'Subject: x\r\n\r\ny', subject: 'x' });
    const line = JSON.stringify({
      jsonrpc: '2.0',
      id: 7,
      method: 'tools/call',
      params: { name: 'saveMessage', arguments: { messageId: 'm', folderPath: 'fp', destDir: dir } }
    });

    const resp = await handleMessage(line, { forward, env: {} });
    assert.equal(resp.id, 7);
    assert.equal(resp.error.code, -32602);
    assert.match(resp.error.message, new RegExp(SAVE_MESSAGE_ENV));
    assert.equal(resp.result, undefined);
    assert.deepEqual(fs.readdirSync(dir), []);
  });

  it('handles an enabled tools/call through the handler and reports the path', async () => {
    const dir = mkTemp();
    const forward = async () => makeEnvelope({ rawSource: 'Subject: Via handler\r\n\r\nbody', subject: 'Via handler' });
    const line = JSON.stringify({
      jsonrpc: '2.0',
      id: 8,
      method: 'tools/call',
      params: { name: 'saveMessage', arguments: { messageId: 'm', folderPath: 'fp', destDir: dir } }
    });

    const resp = await handleMessage(line, { forward, env: ENABLED });
    assert.equal(resp.error, undefined);
    const data = JSON.parse(resp.result.content[0].text);
    assert.equal(data.emlPath, path.join(dir, 'Via handler.eml'));
    assert.ok(fs.existsSync(data.emlPath));
  });

  it('confines destDir to THUNDERBIRD_MCP_SAVE_ROOT when one is set', async () => {
    const root = mkTemp();
    const outside = mkTemp();
    const fetchMessage = async () => makeEnvelope({ rawSource: 'Subject: Rooted\r\n\r\nbody', subject: 'Rooted' });
    const env = { ...ENABLED, [SAVE_ROOT_ENV]: root };

    const inside = await saveMessageToDisk({
      args: { messageId: 'm', folderPath: 'fp', destDir: path.join(root, 'sub') },
      fetchMessage,
      env
    });
    assert.equal(inside.emlPath, path.join(root, 'sub', 'Rooted.eml'));

    await assert.rejects(
      saveMessageToDisk({
        args: { messageId: 'm', folderPath: 'fp', destDir: outside },
        fetchMessage,
        env
      }),
      new RegExp(SAVE_ROOT_ENV)
    );
    assert.deepEqual(fs.readdirSync(outside), []);
  });

  it('rejects a destDir that escapes the save root through a symlink', async () => {
    const root = mkTemp();
    const outside = mkTemp();
    const escape = path.join(root, 'escape');
    fs.symlinkSync(outside, escape);
    const fetchMessage = async () => makeEnvelope({ rawSource: 'Subject: x\r\n\r\ny', subject: 'x' });

    await assert.rejects(
      saveMessageToDisk({
        args: { messageId: 'm', folderPath: 'fp', destDir: escape },
        fetchMessage,
        env: { ...ENABLED, [SAVE_ROOT_ENV]: root }
      }),
      new RegExp(SAVE_ROOT_ENV)
    );
    assert.deepEqual(fs.readdirSync(outside), []);
  });

  it('rejects traversal out of the save root', async () => {
    const root = mkTemp();
    const fetchMessage = async () => makeEnvelope({ rawSource: 'Subject: x\r\n\r\ny', subject: 'x' });

    await assert.rejects(
      saveMessageToDisk({
        args: { messageId: 'm', folderPath: 'fp', destDir: path.join(root, '..') },
        fetchMessage,
        env: { ...ENABLED, [SAVE_ROOT_ENV]: root }
      }),
      new RegExp(SAVE_ROOT_ENV)
    );
  });
});

describe('sanitizeSaveFilename', () => {
  it('keeps ordinary names, including unicode and parentheses', () => {
    assert.equal(sanitizeSaveFilename('Rechnung (final) — Grüße.pdf'), 'Rechnung (final) — Grüße.pdf');
  });

  it('neutralizes path separators and control characters', () => {
    assert.equal(sanitizeSaveFilename('../../etc/passwd'), '.. .. etc passwd');
    assert.equal(sanitizeSaveFilename('a\u0000b'), 'a b');
    assert.equal(sanitizeSaveFilename('C:\\Windows\\system32'), 'C Windows system32');
  });

  it('strips the characters Windows forbids, including NTFS stream colons', () => {
    assert.equal(sanitizeSaveFilename('report:$DATA'), 'report $DATA');
    assert.equal(sanitizeSaveFilename('a<b>c"d|e?f*g'), 'a b c d e f g');
  });

  it('drops trailing dots and spaces that Windows silently discards', () => {
    assert.equal(sanitizeSaveFilename('report.pdf...'), 'report.pdf');
    assert.equal(sanitizeSaveFilename('report.pdf   '), 'report.pdf');
  });

  it('defuses reserved device names, with or without an extension', () => {
    assert.equal(sanitizeSaveFilename('CON'), '_CON');
    assert.equal(sanitizeSaveFilename('com1.txt'), '_com1.txt');
    assert.equal(sanitizeSaveFilename('LPT9.tar.gz'), '_LPT9.tar.gz');
    // Not reserved -- must be left alone.
    assert.equal(sanitizeSaveFilename('CONTRACT.pdf'), 'CONTRACT.pdf');
    assert.equal(sanitizeSaveFilename('COM10.txt'), 'COM10.txt');
  });

  it('refuses names that cannot be made safe', () => {
    assert.throws(() => sanitizeSaveFilename(''), /Cannot derive/);
    assert.throws(() => sanitizeSaveFilename('...'), /Cannot derive/);
    assert.throws(() => sanitizeSaveFilename('/'), /Cannot derive/);
    assert.throws(() => sanitizeSaveFilename(null), /Cannot derive/);
  });
});
