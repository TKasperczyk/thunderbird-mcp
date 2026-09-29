const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const apiSource = fs.readFileSync(
  path.resolve(__dirname, '../extension/mcp_server/api.js'),
  'utf8'
);

function getMarkedApiSnippet(startMarker, endMarker) {
  const start = apiSource.indexOf(startMarker);
  const end = apiSource.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0, `api.js marker missing: ${startMarker}`);
  assert.ok(end > start, `api.js marker missing: ${endMarker}`);
  return apiSource.slice(start, end);
}

function loadNormalizeDraftHeaders() {
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext([
    getMarkedApiSnippet('// BEGIN SAVE DRAFT HEADERS NORMALIZER', '// END SAVE DRAFT HEADERS NORMALIZER'),
    'this.normalizeDraftHeaders = normalizeDraftHeaders;',
  ].join('\n'), sandbox);
  return sandbox.normalizeDraftHeaders;
}

let normalizeDraftHeaders;
let normalizerLoadError;
try {
  normalizeDraftHeaders = loadNormalizeDraftHeaders();
} catch (error) {
  normalizerLoadError = error;
}

function normalizeHeaders(value) {
  assert.ifError(normalizerLoadError);
  return normalizeDraftHeaders(value);
}

function getSaveDraftToolDefinition() {
  const toolName = apiSource.indexOf('name: "saveDraft"');
  const start = apiSource.lastIndexOf('      {', toolName);
  const nextTool = apiSource.indexOf('\n      {\n        name:', toolName);
  assert.ok(toolName >= 0 && start >= 0 && nextTool > toolName,
    'production saveDraft tool definition missing');

  const sandbox = { MAX_ATTACHMENTS_PER_MESSAGE: 10 };
  vm.createContext(sandbox);
  vm.runInContext(`this.saveDraftTool = ${apiSource.slice(start, nextTool)};`, sandbox);
  return sandbox.saveDraftTool;
}

describe('saveDraft headers', () => {
  it('normalizes an object while preserving header order', () => {
    const headers = { 'X-First': 'one', Subject: 'two' };
    assert.deepEqual(normalizeHeaders(headers), {
      headers: [['X-First', 'one'], ['Subject', 'two']],
    });
  });

  it('parses a JSON object string', () => {
    assert.deepEqual(normalizeHeaders('{"X-Test":"value"}'), {
      headers: [['X-Test', 'value']],
    });
  });

  it('returns an error for invalid JSON', () => {
    const result = normalizeHeaders('{');
    assert.match(result.error, /^headers is a string but not valid JSON: /);
  });

  it('normalizes missing and blank headers to an empty list', () => {
    for (const value of [undefined, null, '', '  ']) {
      assert.deepEqual(normalizeHeaders(value), { headers: [] });
    }
  });

  it('rejects arrays, numbers, and other non-object JSON values', () => {
    for (const value of [[], 42, '[]', '42', 'true', 'null']) {
      assert.deepEqual(normalizeHeaders(value), {
        error: 'headers must be an object of {name: value}',
      });
    }
  });

  it('rejects header names containing spaces, colons, or no characters', () => {
    for (const name of ['X Bad', 'X:Bad', '']) {
      assert.deepEqual(normalizeHeaders({ [name]: 'value' }), {
        error: `Invalid header name: ${name}`,
      });
    }
  });

  it('rejects header values containing carriage returns or line feeds', () => {
    for (const value of ['before\rafter', 'before\nafter']) {
      assert.deepEqual(normalizeHeaders({ 'X-Test': value }), {
        error: 'Invalid header value for X-Test (must be a string without CR/LF)',
      });
    }
  });

  it('rejects non-string header values', () => {
    assert.deepEqual(normalizeHeaders({ 'X-Test': 42 }), {
      error: 'Invalid header value for X-Test (must be a string without CR/LF)',
    });
  });

  it('validates headers before saving and applies each header with setHeader', () => {
    const saveDraftStart = apiSource.indexOf('function saveDraft(');
    const saveDraftEnd = apiSource.indexOf('\n            /**', saveDraftStart);
    const saveDraftSource = apiSource.slice(saveDraftStart, saveDraftEnd);
    assert.match(saveDraftSource, /normalizeDraftHeaders\(headers\)/);
    assert.match(saveDraftSource, /composeFields\.setHeader\(name, value\)/);
    assert.ok(saveDraftSource.indexOf('normalizeDraftHeaders(headers)')
      < saveDraftSource.indexOf('sendMessageDirectly('));
  });

  it('passes args.headers from the dispatcher to saveDraft', () => {
    assert.match(apiSource,
      /case "saveDraft":\s*return await saveDraft\([^;]*args\.headers\);/);
  });

  it('exposes headers on the saveDraft tool definition', () => {
    const tool = getSaveDraftToolDefinition();
    const headers = tool.inputSchema.properties.headers;
    assert.ok(headers, 'saveDraft headers property missing');
    assert.equal(headers.oneOf[0].type, 'object');
    assert.equal(headers.oneOf[0].additionalProperties.type, 'string');
    assert.equal(headers.oneOf[1].type, 'string');
    assert.match(headers.description, /Extra RFC 5322 headers/);
    assert.match(headers.description, /X-Send-Later-At/);
    assert.match(headers.description, /X-Send-Later-Uuid/);
  });
});
