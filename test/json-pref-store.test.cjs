/**
 * Unit tests for the reusable JSON-pref-store pattern in
 * lib/json-pref-store.sys.mjs. Mirrors the approach in permissions.test.cjs
 * (vm sandbox + strip export + mock Services).
 */

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function loadStore() {
  const src = fs.readFileSync(
    path.join(__dirname, '..', 'extension', 'mcp_server', 'lib', 'json-pref-store.sys.mjs'),
    'utf8'
  );
  const sandbox = {
    console,
    structuredClone: globalThis.structuredClone || ((v) => JSON.parse(JSON.stringify(v))),
    makeJsonPrefStore: undefined,
    wrapApi: undefined,
  };
  vm.createContext(sandbox);
  const stripped = src
    .replace(/^export\s+function\s+makeJsonPrefStore/m, 'function makeJsonPrefStore')
    .replace(/^export\s+async\s+function\s+wrapApi/m, 'async function wrapApi');
  vm.runInContext(stripped + '\n;this.makeJsonPrefStore = makeJsonPrefStore;this.wrapApi = wrapApi;', sandbox);
  return { makeJsonPrefStore: sandbox.makeJsonPrefStore, wrapApi: sandbox.wrapApi };
}

function makeMockServices(initial) {
  const prefs = new Map();
  if (initial !== undefined) prefs.set('test.pref', initial);
  return {
    prefs: {
      prefHasUserValue: (n) => prefs.has(n),
      getStringPref: (n, d) => prefs.has(n) ? prefs.get(n) : d,
      setStringPref: (n, v) => prefs.set(n, v),
      clearUserPref: (n) => prefs.delete(n),
    },
  };
}

// Silent logger for the intentional negative-path tests below (corrupt
// JSON, validator throws, wrapApi op throws). Without this the tests
// pollute CI output with the module's expected console.warn / .error
// lines -- TAP renders them as `# ...` comments above each assertion.
const silentLogger = { warn() {}, error() {}, log() {}, info() {} };

const { makeJsonPrefStore, wrapApi } = loadStore();

describe('makeJsonPrefStore: load', () => {
  it('returns cloned defaultValue when pref has no user value', () => {
    const defaultValue = { foo: [1, 2] };
    const store = makeJsonPrefStore({ Services: makeMockServices(), prefKey: 'test.pref', defaultValue });
    const a = store.load();
    const b = store.load();
    assert.deepEqual(a, { foo: [1, 2] });
    assert.notStrictEqual(a, defaultValue);     // not the same reference
    a.foo.push(999);
    assert.deepEqual(b.foo, [1, 2]);            // second load unaffected
  });

  it('returns parsed pref when present and valid', () => {
    const store = makeJsonPrefStore({
      Services: makeMockServices('{"a":1,"b":"x"}'),
      prefKey: 'test.pref',
      defaultValue: {},
    });
    const r = store.load();
    assert.equal(r.a, 1);
    assert.equal(r.b, 'x');
  });

  it('runs validate() on loaded pref', () => {
    const store = makeJsonPrefStore({
      Services: makeMockServices('{"n":7}'),
      prefKey: 'test.pref',
      defaultValue: {},
      validate: (raw) => ({ ...raw, validated: true }),
    });
    assert.deepEqual(store.load(), { n: 7, validated: true });
  });

  it('falls back to defaultValue on corrupt JSON without throwing', () => {
    const store = makeJsonPrefStore({
      Services: makeMockServices('{ not valid'),
      prefKey: 'test.pref',
      defaultValue: { fallback: true },
      logger: silentLogger,
    });
    assert.deepEqual(store.load(), { fallback: true });
  });

  it('falls back to defaultValue when validate() throws', () => {
    const store = makeJsonPrefStore({
      Services: makeMockServices('{}'),
      prefKey: 'test.pref',
      defaultValue: { safe: true },
      validate: () => { throw new Error('nope'); },
      logger: silentLogger,
    });
    assert.deepEqual(store.load(), { safe: true });
  });
});

describe('makeJsonPrefStore: save', () => {
  it('persists normalized value via validate()', () => {
    const svcs = makeMockServices();
    const store = makeJsonPrefStore({
      Services: svcs,
      prefKey: 'test.pref',
      defaultValue: {},
      validate: (v) => ({ ...v, stamp: 1 }),
    });
    const result = store.save({ x: 'y' });
    assert.deepEqual(result, { x: 'y', stamp: 1 });
    assert.equal(svcs.prefs.getStringPref('test.pref', ''), JSON.stringify({ x: 'y', stamp: 1 }));
  });

  it('throws when validate() rejects', () => {
    const store = makeJsonPrefStore({
      Services: makeMockServices(),
      prefKey: 'test.pref',
      defaultValue: {},
      validate: () => { throw new Error('reject'); },
    });
    assert.throws(() => store.save({}), /reject/);
  });
});

describe('makeJsonPrefStore: clear', () => {
  it('clears the pref silently even when unset', () => {
    const svcs = makeMockServices();
    const store = makeJsonPrefStore({ Services: svcs, prefKey: 'test.pref', defaultValue: {} });
    store.clear(); // no throw
    store.save({ a: 1 });
    store.clear();
    assert.equal(svcs.prefs.prefHasUserValue('test.pref'), false);
  });
});

describe('makeJsonPrefStore: contract', () => {
  it('requires Services and prefKey', () => {
    assert.throws(() => makeJsonPrefStore({}), /Services/);
    assert.throws(() => makeJsonPrefStore({ Services: makeMockServices() }), /prefKey/);
  });
});

describe('wrapApi', () => {
  it('returns op() result on success', async () => {
    const res = await wrapApi('foo', () => ({ ok: true }));
    assert.deepEqual(res, { ok: true });
  });
  it('awaits async op', async () => {
    const res = await wrapApi('foo', async () => ({ async: true }));
    assert.deepEqual(res, { async: true });
  });
  it('returns { error: ... } on throw with method prefix', async () => {
    const res = await wrapApi('setFoo', () => { throw new Error('boom'); }, { logger: silentLogger });
    assert.match(res.error, /^setFoo failed: boom$/);
  });
  it('returns { error: ... } on async rejection', async () => {
    const res = await wrapApi('setFoo', async () => { throw new Error('async boom'); }, { logger: silentLogger });
    assert.match(res.error, /^setFoo failed: async boom$/);
  });
});
