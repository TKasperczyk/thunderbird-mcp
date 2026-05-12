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
  // permissions.sys.mjs now depends on makeJsonPrefStore from the shared
  // pref-store module -- load that too and expose it in the sandbox so
  // the factory can call it. Both files get their `export ` prefix
  // stripped so the function declarations evaluate as plain expressions.
  const storeSrc = fs.readFileSync(
    path.join(__dirname, '..', 'extension', 'mcp_server', 'lib', 'json-pref-store.sys.mjs'),
    'utf8'
  );
  const permSrc = fs.readFileSync(
    path.join(__dirname, '..', 'extension', 'mcp_server', 'lib', 'permissions.sys.mjs'),
    'utf8'
  );
  const sandbox = {
    console,
    structuredClone: globalThis.structuredClone || ((v) => JSON.parse(JSON.stringify(v))),
    makeJsonPrefStore: undefined,
    wrapApi: undefined,
    makePermissions: undefined,
  };
  vm.createContext(sandbox);
  const strippedStore = storeSrc
    .replace(/^export\s+function\s+makeJsonPrefStore/m, 'function makeJsonPrefStore')
    .replace(/^export\s+async\s+function\s+wrapApi/m, 'async function wrapApi');
  vm.runInContext(strippedStore + '\n;this.makeJsonPrefStore = makeJsonPrefStore;this.wrapApi = wrapApi;', sandbox);
  const strippedPerm = permSrc.replace(/^export\s+function\s+makePermissions/m, 'function makePermissions');
  vm.runInContext(strippedPerm + '\n;this.makePermissions = makePermissions;', sandbox);
  return { makePermissions: sandbox.makePermissions, makeJsonPrefStore: sandbox.makeJsonPrefStore };
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

function makeMockMailServices(folderToAccount, extraAccountKeys = []) {
  // Build a stable set of known account keys: every key referenced by any
  // folder mapping plus any explicitly listed extras (e.g. accounts that
  // appear in the policy but not in folderToAccount). The engine uses this
  // list to reject unknown accountId args; tests that exercise the
  // account-default branch must list the accounts they expect to validate.
  const knownKeys = new Set([
    ...Object.values(folderToAccount),
    ...extraAccountKeys,
  ]);
  const accountObjects = [...knownKeys].map((key) => ({ key }));
  return {
    folderLookup: {
      getFolderForURL: (uri) => {
        const accountKey = folderToAccount[uri];
        if (!accountKey) return null;
        return { server: { _key: accountKey } };
      },
    },
    accounts: {
      accounts: accountObjects,
      findAccountForServer: (server) => server && server._key ? { key: server._key } : null,
    },
  };
}

// ---- tests --------------------------------------------------------------

const { makePermissions, makeJsonPrefStore } = loadPermissionsFactory();

const UNDISABLEABLE = new Set(['listAccounts', 'listFolders', 'getAccountAccess']);
const DEFAULTS_OFF = new Set([
  'sendMail', 'replyToMessage', 'forwardMessage',
  'deleteMessages', 'deleteContact', 'deleteEvent', 'deleteFilter',
  'deleteFolder', 'emptyTrash', 'emptyJunk',
]);

function makeEngine({ pref, folderMap = {}, knownAccountKeys = [] } = {}) {
  return makePermissions({
    Services: makeMockServices(pref),
    MailServices: makeMockMailServices(folderMap, knownAccountKeys),
    PREF_PERMISSIONS: 'extensions.thunderbird-mcp.permissions',
    UNDISABLEABLE_TOOLS: UNDISABLEABLE,
    DEFAULT_DISABLED_TOOLS: DEFAULTS_OFF,
    makeJsonPrefStore,
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
    const e = makeEngine({
      pref: JSON.stringify({
        tools: { searchMessages: { enabled: true } },
        accounts: { account_a: { tools: { searchMessages: { enabled: false } } } }
      }),
      knownAccountKeys: ['account_a', 'account_b'],
    });
    assert.equal(e.evaluate('searchMessages', { query: 'foo', accountId: 'account_a' }, true).allow, false);
    assert.equal(e.evaluate('searchMessages', { query: 'foo', accountId: 'account_b' }, true).allow, true);
  });
  it('account default=deny blocks tools without explicit per-tool rule', () => {
    const e = makeEngine({
      pref: JSON.stringify({
        tools: { searchMessages: { enabled: true } },
        accounts: { account_a: { default: 'deny' } }
      }),
      knownAccountKeys: ['account_a'],
    });
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

describe('permissions engine: bypass-fix regression tests', () => {
  it('bulk_move_by_query clamps args.limit when omitted (CodeRabbit)', () => {
    // Caller omits `limit`; tool default is 1000 (per tools-schema.sys.mjs)
    // which would silently bypass maxBatchSize=50. Engine must clamp to cap.
    const e = makeEngine({ pref: JSON.stringify({
      tools: { bulk_move_by_query: { enabled: true, maxBatchSize: 50 } }
    }) });
    const r = e.evaluate(
      'bulk_move_by_query',
      { query: 'subject:foo', folderPath: 'imap://s/INBOX' },
      true
    );
    assert.equal(r.allow, true);
    assert.equal(r.args.limit, 50, 'limit must be clamped to maxBatchSize when omitted');
  });

  it('bulk_move_by_query clamps args.limit when over cap (CodeRabbit)', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { bulk_move_by_query: { enabled: true, maxBatchSize: 50 } }
    }) });
    const r = e.evaluate(
      'bulk_move_by_query',
      { query: 'subject:foo', folderPath: 'imap://s/INBOX', limit: 9999 },
      true
    );
    assert.equal(r.allow, true);
    assert.equal(r.args.limit, 50, 'limit must be clamped down to maxBatchSize');
  });

  it('bulk_move_by_query leaves args.limit alone when within cap (CodeRabbit)', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { bulk_move_by_query: { enabled: true, maxBatchSize: 50 } }
    }) });
    const r = e.evaluate(
      'bulk_move_by_query',
      { query: 'subject:foo', folderPath: 'imap://s/INBOX', limit: 10 },
      true
    );
    assert.equal(r.allow, true);
    assert.equal(r.args.limit, 10);
  });

  it('maxBatchSize denies oversized JSON-string-encoded array (Claude §2a)', () => {
    // deleteMessages JSON.parses string messageIds itself; the policy engine
    // must apply maxBatchSize to the parsed payload, not silently let the
    // stringified form through.
    const e = makeEngine({ pref: JSON.stringify({
      tools: { deleteMessages: { enabled: true, maxBatchSize: 3 } }
    }) });
    const overSized = JSON.stringify(['a', 'b', 'c', 'd', 'e']);
    const r = e.evaluate('deleteMessages', { messageIds: overSized, folderPath: 'f' }, false);
    assert.equal(r.allow, false);
    assert.match(r.reason, /maxBatchSize=3/);
  });

  it('maxBatchSize permits within-cap JSON-string-encoded array (Claude §2a)', () => {
    const e = makeEngine({ pref: JSON.stringify({
      tools: { deleteMessages: { enabled: true, maxBatchSize: 5 } }
    }) });
    const withinCap = JSON.stringify(['a', 'b', 'c']);
    const r = e.evaluate('deleteMessages', { messageIds: withinCap, folderPath: 'f' }, false);
    assert.equal(r.allow, true);
  });

  it('destination folder rule denies bulk_move into a denied folder (Claude §2b)', () => {
    // Per-folder rule on the destination must be honoured even when the
    // source folder rule allows the move.
    const e = makeEngine({
      pref: JSON.stringify({
        tools: { bulk_move_by_query: { enabled: true, maxBatchSize: 1000 } },
        folders: {
          'imap://acct/Archive': { tools: { bulk_move_by_query: { enabled: false } } },
        },
      }),
      folderMap: {
        'imap://acct/INBOX': 'account_x',
        'imap://acct/Archive': 'account_x',
      },
    });
    const r = e.evaluate(
      'bulk_move_by_query',
      { query: 'subject:x', folderPath: 'imap://acct/INBOX', to: 'imap://acct/Archive' },
      true
    );
    assert.equal(r.allow, false);
    assert.match(r.reason, /destination folder/);
  });

  it('destination folder rule allows bulk_move when destination is not denied (Claude §2b)', () => {
    const e = makeEngine({
      pref: JSON.stringify({
        tools: { bulk_move_by_query: { enabled: true, maxBatchSize: 1000 } },
        folders: {
          'imap://acct/Archive': { tools: { bulk_move_by_query: { enabled: false } } },
        },
      }),
      folderMap: {
        'imap://acct/INBOX': 'account_x',
        'imap://acct/Backups': 'account_x',
      },
    });
    const r = e.evaluate(
      'bulk_move_by_query',
      { query: 'subject:x', folderPath: 'imap://acct/INBOX', to: 'imap://acct/Backups' },
      true
    );
    assert.equal(r.allow, true);
  });

  it('unknown accountId is rejected up front (Claude §2c)', () => {
    // Without this check, account.default=deny would silently not apply
    // because policy.accounts["doesnotexist"] is undefined.
    const e = makeEngine({
      pref: JSON.stringify({
        tools: { searchMessages: { enabled: true } },
        accounts: { account_a: { default: 'deny' } },
      }),
      knownAccountKeys: ['account_a'],
    });
    const r = e.evaluate('searchMessages', { query: 'foo', accountId: 'doesnotexist' }, true);
    assert.equal(r.allow, false);
    assert.match(r.reason, /Unknown accountId/);
  });

  it('known accountId still resolves normally after the validation gate (Claude §2c)', () => {
    const e = makeEngine({
      pref: JSON.stringify({
        tools: { searchMessages: { enabled: true } },
      }),
      knownAccountKeys: ['account_a'],
    });
    assert.equal(e.evaluate('searchMessages', { query: 'foo', accountId: 'account_a' }, true).allow, true);
  });

  it('includeSubfolders=true is denied when per-folder rules exist for the tool (Claude §2d)', () => {
    // A per-folder deny on a subfolder must not be bypassable by recursing
    // from the parent. The pragmatic policy: block recursion outright as
    // soon as any per-folder rule exists for this tool.
    const e = makeEngine({
      pref: JSON.stringify({
        tools: { bulk_move_by_query: { enabled: true, maxBatchSize: 1000 } },
        folders: {
          'imap://acct/INBOX/Sensitive': { tools: { bulk_move_by_query: { enabled: false } } },
        },
      }),
      folderMap: {
        'imap://acct/INBOX': 'account_x',
        'imap://acct/INBOX/Sensitive': 'account_x',
      },
    });
    const r = e.evaluate(
      'bulk_move_by_query',
      { query: 'subject:x', folderPath: 'imap://acct/INBOX', includeSubfolders: true },
      true
    );
    assert.equal(r.allow, false);
    assert.match(r.reason, /includeSubfolders/);
  });

  it('includeSubfolders=true is allowed when no per-folder rules exist for the tool (Claude §2d)', () => {
    const e = makeEngine({
      pref: JSON.stringify({
        tools: { bulk_move_by_query: { enabled: true, maxBatchSize: 1000 } },
      }),
    });
    const r = e.evaluate(
      'bulk_move_by_query',
      { query: 'subject:x', folderPath: 'imap://acct/INBOX', includeSubfolders: true },
      true
    );
    assert.equal(r.allow, true);
  });

  it('catastrophic-backtracking regex in argRule.match is rejected at load (Claude §5 / CodeRabbit)', () => {
    // (a+)+ -- classic ReDoS shape. The engine must treat the rule as
    // invalid-skip and deny any value that hits it.
    const e = makeEngine({ pref: JSON.stringify({
      tools: { sendMail: { enabled: true, argRules: { subject: { match: '(a+)+$' } } } }
    }) });
    const r = e.evaluate('sendMail', { to: 'x@y.z', subject: 'aaaaaaaaa!' }, false);
    assert.equal(r.allow, false);
    assert.match(r.reason, /rejected at policy load|not whitelisted/);
  });

  it('overlong argRule.match (>200 chars) is rejected at load (CodeRabbit)', () => {
    const longPattern = 'a'.repeat(250);
    const e = makeEngine({ pref: JSON.stringify({
      tools: { sendMail: { enabled: true, argRules: { subject: { match: longPattern } } } }
    }) });
    const r = e.evaluate('sendMail', { to: 'x@y.z', subject: 'hi' }, false);
    assert.equal(r.allow, false);
    assert.match(r.reason, /rejected at policy load|not whitelisted/);
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
