/**
 * Unit tests for the fine-grained permission engine.
 *
 * The .sys.mjs module can't be required in plain Node because it expects to
 * be loaded via ChromeUtils.importESModule, but the factory itself is plain
 * JavaScript -- we extract the source text and eval it in a sandbox with a
 * mock Services / MailServices, then exercise the engine against fabricated
 * policy documents.
 */

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ---- load makePermissions out of the .sys.mjs file ----------------------

function loadPermissionsFactory() {
  const src = fs.readFileSync(
    path.join(__dirname, '..', 'extension', 'mcp_server', 'lib', 'permissions.sys.mjs'),
    'utf8'
  );
  // Strip the leading `export ` so we can eval the function declaration as
  // a plain expression in a sandbox. Same trick the existing tests use.
  const stripped = src.replace(/^export\s+function\s+makePermissions/m, 'function makePermissions');
  const sandbox = { console, makePermissions: undefined };
  vm.createContext(sandbox);
  vm.runInContext(stripped + '\n;this.makePermissions = makePermissions;', sandbox);
  return sandbox.makePermissions;
}

// ---- tiny mocks for Services + MailServices ----------------------------

function makeMockServices(initialPref) {
  const prefs = new Map();
  if (initialPref !== undefined) prefs.set('extensions.thunderbird-mcp.permissions', initialPref);
  return {
    prefs: {
      prefHasUserValue: (name) => prefs.has(name),
      getStringPref: (name, def) => prefs.has(name) ? prefs.get(name) : def,
      setStringPref: (name, val) => prefs.set(name, val),
      clearUserPref: (name) => prefs.delete(name),
    },
  };
}

function makeMockMailServices(folderToAccount) {
  return {
    folderLookup: {
      getFolderForURL: (uri) => {
        const accountKey = folderToAccount[uri];
        if (!accountKey) return null;
        return { server: { _key: accountKey } };
      },
    },
    accounts: {
      findAccountForServer: (server) => server && server._key ? { key: server._key } : null,
    },
  };
}

// ---- tests --------------------------------------------------------------

const makePermissions = loadPermissionsFactory();

const UNDISABLEABLE = new Set(['listAccounts', 'listFolders', 'getAccountAccess']);
const DEFAULTS_OFF = new Set([
  'sendMail', 'replyToMessage', 'forwardMessage',
  'deleteMessages', 'deleteContact', 'deleteEvent', 'deleteFilter',
  'deleteFolder', 'emptyTrash', 'emptyJunk',
]);

function makeEngine({ pref, folderMap = {} } = {}) {
  return makePermissions({
    Services: makeMockServices(pref),
    MailServices: makeMockMailServices(folderMap),
    PREF_PERMISSIONS: 'extensions.thunderbird-mcp.permissions',
    UNDISABLEABLE_TOOLS: UNDISABLEABLE,
    DEFAULT_DISABLED_TOOLS: DEFAULTS_OFF,
  });
}

describe('permissions engine: Level 0 (undisableable tools)', () => {
  it('listAccounts always allowed even with empty policy', () => {
    const e = makeEngine();
    assert.equal(e.evaluate('listAccounts', {}, false).allow, true);
  });
  it('listFolders allowed even when policy explicitly disables it', () => {
    const e = makeEngine({ pref: JSON.stringify({ tools: { listFolders: { enabled: false } } }) });
    assert.equal(e.evaluate('listFolders', {}, true).allow, true);
  });
});

describe('permissions engine: Level 1 (sandbox + global enable/disable)', () => {
  it('sendMail denied with no policy and baseAllowed=false (sandbox default)', () => {
    const e = makeEngine();
    const r = e.evaluate('sendMail', { to: 'x@y.z' }, false);
    assert.equal(r.allow, false);
    assert.match(r.reason, /off by default/);
  });
  it('sendMail allowed when policy explicitly enables it (overrides sandbox default)', () => {
    const e = makeEngine({ pref: JSON.stringify({ tools: { sendMail: { enabled: true } } }) });
    assert.equal(e.evaluate('sendMail', { to: 'x@y.z' }, false).allow, true);
  });
  it('searchMessages denied when policy disables it (overrides baseAllowed=true)', () => {
    const e = makeEngine({ pref: JSON.stringify({ tools: { searchMessages: { enabled: false } } }) });
    const r = e.evaluate('searchMessages', { query: 'foo' }, true);
    assert.equal(r.allow, false);
    assert.match(r.reason, /disabled by policy/);
  });
});

describe('permissions engine: Level 2 (per-tool constraints)', () => {
  it('alwaysReview forces skipReview to false on compose tools', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { sendMail: { enabled: true, alwaysReview: true } }
    }) });
    const r = e.evaluate('sendMail', { to: 'x@y.z', skipReview: true }, false);
    assert.equal(r.allow, true);
    assert.equal(r.args.skipReview, false);
  });
  it('maxBatchSize denies oversized bulk operations', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { deleteMessages: { enabled: true, maxBatchSize: 2 } }
    }) });
    const ok = e.evaluate('deleteMessages', { messageIds: ['1', '2'], folderPath: 'f' }, false);
    assert.equal(ok.allow, true);
    const bad = e.evaluate('deleteMessages', { messageIds: ['1', '2', '3'], folderPath: 'f' }, false);
    assert.equal(bad.allow, false);
    assert.match(bad.reason, /maxBatchSize=2/);
  });
  it('requireFolderArg denies tools without a folder argument', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { searchMessages: { enabled: true, requireFolderArg: true } }
    }) });
    const denied = e.evaluate('searchMessages', { query: 'foo' }, true);
    assert.equal(denied.allow, false);
    assert.match(denied.reason, /requires a folderPath/);
    const allowed = e.evaluate('searchMessages', { query: 'foo', folderPath: 'imap://x/INBOX' }, true);
    assert.equal(allowed.allow, true);
  });
});

describe('permissions engine: Level 3 (per-account + per-folder rules)', () => {
  it('per-account override denies for one account, allows for another', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { searchMessages: { enabled: true } },
      accounts: { account_a: { tools: { searchMessages: { enabled: false } } } }
    }) });
    assert.equal(e.evaluate('searchMessages', { query: 'foo', accountId: 'account_a' }, true).allow, false);
    assert.equal(e.evaluate('searchMessages', { query: 'foo', accountId: 'account_b' }, true).allow, true);
  });
  it('account default=deny blocks tools without explicit per-tool rule', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { searchMessages: { enabled: true } },
      accounts: { account_a: { default: 'deny' } }
    }) });
    assert.equal(e.evaluate('searchMessages', { query: 'foo', accountId: 'account_a' }, true).allow, false);
  });
  it('per-folder rule overrides per-account rule', () => {
    const e = makeEngine({
      pref: JSON.stringify({
        tools: { deleteMessages: { enabled: false } },
        folders: {
          'imap://acct/Trash': { tools: { deleteMessages: { enabled: true, maxBatchSize: 999 } } }
        }
      }),
      folderMap: { 'imap://acct/Trash': 'account_x', 'imap://acct/INBOX': 'account_x' },
    });
    assert.equal(e.evaluate('deleteMessages', { messageIds: ['1'], folderPath: 'imap://acct/INBOX' }, true).allow, false);
    assert.equal(e.evaluate('deleteMessages', { messageIds: ['1'], folderPath: 'imap://acct/Trash' }, true).allow, true);
  });
});

describe('permissions engine: Level 4 (argument whitelists)', () => {
  it('anyOf glob whitelist matches single recipient', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { sendMail: { enabled: true, argRules: { to: { anyOf: ['*@allowed.com'] } } } }
    }) });
    assert.equal(e.evaluate('sendMail', { to: 'x@allowed.com' }, false).allow, true);
    const r = e.evaluate('sendMail', { to: 'x@evil.com' }, false);
    assert.equal(r.allow, false);
    assert.match(r.reason, /not in allowed list/);
  });
  it('match regex applied to every element of array values', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { updateMessage: { enabled: true, argRules: { addTags: { match: '^\\$label[1-5]$' } } } }
    }) });
    assert.equal(e.evaluate('updateMessage', { folderPath: 'f', addTags: ['$label1', '$label3'] }, true).allow, true);
    const r = e.evaluate('updateMessage', { folderPath: 'f', addTags: ['$label1', 'malicious-tag'] }, true);
    assert.equal(r.allow, false);
    assert.match(r.reason, /does not match/);
  });
  it('maxLength caps array and string values', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { deleteMessages: { enabled: true, argRules: { messageIds: { maxLength: 3 } } } }
    }) });
    assert.equal(e.evaluate('deleteMessages', { messageIds: ['a', 'b'], folderPath: 'f' }, true).allow, true);
    assert.equal(e.evaluate('deleteMessages', { messageIds: ['a', 'b', 'c', 'd'], folderPath: 'f' }, true).allow, false);
  });
});

describe('permissions engine: policy I/O', () => {
  it('loadPolicy returns empty when pref unset', () => {
    const e = makeEngine();
    const p = e.loadPolicy();
    assert.equal(p.version, 1);
    assert.deepEqual(Object.keys(p.tools), []);
    assert.deepEqual(Object.keys(p.accounts), []);
    assert.deepEqual(Object.keys(p.folders), []);
  });
  it('loadPolicy normalizes a partial document', () => {
    const e = makeEngine({ pref: JSON.stringify({ tools: { sendMail: { enabled: true } } }) });
    const p = e.loadPolicy();
    assert.deepEqual(Object.keys(p.accounts), []);
    assert.deepEqual(Object.keys(p.folders), []);
    assert.equal(p.tools.sendMail.enabled, true);
  });
  it('loadPolicy returns empty on corrupt JSON instead of throwing', () => {
    const e = makeEngine({ pref: '{ not valid json' });
    const p = e.loadPolicy();
    assert.equal(p.version, 1);
    assert.deepEqual(Object.keys(p.tools), []);
  });
});
