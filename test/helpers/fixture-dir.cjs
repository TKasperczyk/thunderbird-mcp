'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const { isSensitiveFilePath } = require('./bridge.cjs');

// Attachment fixtures must live where the attachment policy allows them.
// The temp directory can be refused (Windows AppData, or a TMPDIR under a
// dot-directory), and so can a checkout under a dot-directory, so try a few
// candidates and use the first one the policy accepts.
function makeAllowedTempDir(prefix) {
  const candidates = [os.tmpdir(), path.resolve(__dirname, '..', '..'), '/tmp', '/var/tmp', os.homedir()];
  for (const candidate of candidates) {
    let base;
    try { base = fs.realpathSync(candidate); } catch { continue; }
    if (isSensitiveFilePath(base, { windows: process.platform === 'win32' })) continue;
    try { return fs.mkdtempSync(path.join(base, prefix)); } catch { /* try the next candidate */ }
  }
  throw new Error(`No writable directory allowed by the attachment policy for ${prefix} fixtures`);
}

module.exports = { makeAllowedTempDir };
