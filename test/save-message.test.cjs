const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const { saveMessageToDisk } = require('../mcp-bridge.cjs');

// Tracks temp dirs created during a test so afterEach can remove them.
let tempDirs = [];

function mkTemp() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'tb-mcp-save-'));
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
      return makeEnvelope({ rawSource: raw });
    };

    const result = await saveMessageToDisk({
      args: { messageId: 'mid-1', folderPath: 'imap://x/INBOX', destDir: dir },
      fetchMessage
    });

    assert.equal(seenArgs.rawSource, true);
    assert.equal(seenArgs.messageId, 'mid-1');
    const expected = path.join(dir, 'Hello World.eml');
    assert.equal(result.emlPath, expected);
    assert.equal(fs.readFileSync(expected, 'utf8'), raw);
    assert.deepEqual(result.attachments, []);
    assert.deepEqual(result.skippedAttachments, []);
    assert.equal(result.destDir, dir);
  });

  it('unfolds a multi-line Subject header when deriving the filename', async () => {
    const dir = mkTemp();
    const raw = 'From: a@b.com\r\nSubject: First part\r\n continued part\r\n\r\nbody';

    const fetchMessage = async () => makeEnvelope({ rawSource: raw });
    const result = await saveMessageToDisk({
      args: { messageId: 'mid-2', folderPath: 'fp', destDir: dir },
      fetchMessage
    });

    assert.equal(result.emlPath, path.join(dir, 'First part continued part.eml'));
    assert.ok(fs.existsSync(result.emlPath));
  });

  it('honors emlFilename and appends .eml when absent', async () => {
    const dir = mkTemp();
    const raw = 'Subject: Ignored Subject\r\n\r\nbody';

    const fetchMessage = async () => makeEnvelope({ rawSource: raw });
    const result = await saveMessageToDisk({
      args: { messageId: 'mid-3', folderPath: 'fp', destDir: dir, emlFilename: 'custom-name' },
      fetchMessage
    });

    assert.equal(result.emlPath, path.join(dir, 'custom-name.eml'));
    assert.equal(fs.readFileSync(result.emlPath, 'utf8'), raw);
  });

  it('falls back to messageId when there is no Subject header', async () => {
    const dir = mkTemp();
    const raw = 'From: a@b.com\r\n\r\nbody without a subject';

    const fetchMessage = async () => makeEnvelope({ rawSource: raw });
    const result = await saveMessageToDisk({
      args: { messageId: 'fallback-id', folderPath: 'fp', destDir: dir },
      fetchMessage
    });

    assert.equal(result.emlPath, path.join(dir, 'fallback-id.eml'));
  });

  it('creates a missing nested destDir', async () => {
    const dir = path.join(mkTemp(), 'a', 'b', 'c');
    const raw = 'Subject: Nested\r\n\r\nbody';

    const fetchMessage = async () => makeEnvelope({ rawSource: raw });
    const result = await saveMessageToDisk({
      args: { messageId: 'mid-4', folderPath: 'fp', destDir: dir },
      fetchMessage
    });

    assert.ok(fs.existsSync(dir));
    assert.equal(result.emlPath, path.join(dir, 'Nested.eml'));
    assert.ok(fs.existsSync(result.emlPath));
  });

  it('overwrite:false throws on an existing file, overwrite:true replaces it', async () => {
    const dir = mkTemp();
    const raw1 = 'Subject: Dup\r\n\r\nfirst';
    const raw2 = 'Subject: Dup\r\n\r\nsecond';

    const fetchFirst = async () => makeEnvelope({ rawSource: raw1 });
    const first = await saveMessageToDisk({
      args: { messageId: 'mid-5', folderPath: 'fp', destDir: dir },
      fetchMessage: fetchFirst
    });
    assert.equal(fs.readFileSync(first.emlPath, 'utf8'), raw1);

    const fetchSecond = async () => makeEnvelope({ rawSource: raw2 });
    await assert.rejects(
      saveMessageToDisk({
        args: { messageId: 'mid-5', folderPath: 'fp', destDir: dir },
        fetchMessage: fetchSecond
      }),
      /already exists/
    );
    // The failed non-overwrite write must not have clobbered the original.
    assert.equal(fs.readFileSync(first.emlPath, 'utf8'), raw1);

    const replaced = await saveMessageToDisk({
      args: { messageId: 'mid-5', folderPath: 'fp', destDir: dir, overwrite: true },
      fetchMessage: fetchSecond
    });
    assert.equal(replaced.emlPath, first.emlPath);
    assert.equal(fs.readFileSync(replaced.emlPath, 'utf8'), raw2);
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

    const result = await saveMessageToDisk({
      args: { messageId: 'mid-6', folderPath: 'fp', destDir: dir, saveEml: false, saveAttachments: true },
      fetchMessage
    });

    const p1 = path.join(dir, 'report.pdf');
    const p2 = path.join(dir, 'report (2).pdf');
    assert.deepEqual(result.attachments, [p1, p2]);
    assert.equal(fs.readFileSync(p1, 'utf8'), 'AAA');
    assert.equal(fs.readFileSync(p2, 'utf8'), 'BBB');
    assert.deepEqual(result.skippedAttachments, ['inline-only.txt']);
    assert.equal(result.emlPath, null);
  });

  it('writes both the .eml and attachments in one call', async () => {
    const dir = mkTemp();
    const srcDir = mkTemp();
    const src1 = path.join(srcDir, 'temp.bin');
    fs.writeFileSync(src1, 'DATA');
    const raw = 'Subject: Combined\r\n\r\nbody';

    const fetchMessage = async (sub) => {
      if (sub.params.arguments.rawSource) {
        return makeEnvelope({ rawSource: raw });
      }
      return makeEnvelope({ attachments: [{ name: 'doc.txt', filePath: src1 }] });
    };

    const result = await saveMessageToDisk({
      args: { messageId: 'mid-7', folderPath: 'fp', destDir: dir, saveAttachments: true },
      fetchMessage
    });

    assert.equal(result.emlPath, path.join(dir, 'Combined.eml'));
    assert.equal(fs.readFileSync(result.emlPath, 'utf8'), raw);
    assert.deepEqual(result.attachments, [path.join(dir, 'doc.txt')]);
    assert.equal(fs.readFileSync(path.join(dir, 'doc.txt'), 'utf8'), 'DATA');
  });

  it('propagates a tool-level { error } from getMessage as a thrown Error', async () => {
    const dir = mkTemp();
    const fetchMessage = async () => makeEnvelope({ error: 'Message not found' });

    await assert.rejects(
      saveMessageToDisk({
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
      saveMessageToDisk({
        args: { messageId: 'mid-8', folderPath: 'fp', destDir: dir },
        fetchMessage
      }),
      /offline/i
    );
  });

  it('validates required string arguments', async () => {
    const dir = mkTemp();
    const fetchMessage = async () => makeEnvelope({ rawSource: 'Subject: x\r\n\r\ny' });

    await assert.rejects(
      saveMessageToDisk({ args: { folderPath: 'fp', destDir: dir }, fetchMessage }),
      /messageId/
    );
    await assert.rejects(
      saveMessageToDisk({ args: { messageId: 'm', destDir: dir }, fetchMessage }),
      /folderPath/
    );
    await assert.rejects(
      saveMessageToDisk({ args: { messageId: 'm', folderPath: 'fp' }, fetchMessage }),
      /destDir/
    );
  });
});
