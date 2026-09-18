const { it } = require('node:test');
const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const path = require('node:path');
const {
  describeBase64Attachments,
  inlineOnlyAttachmentsEnabled,
  validateBase64Attachments,
} = require('../mcp-bridge.cjs');

it('requires explicit opt-in for remote attachment mode', () => {
  assert.equal(inlineOnlyAttachmentsEnabled({}), false);
  assert.equal(inlineOnlyAttachmentsEnabled({ THUNDERBIRD_MCP_INLINE_ATTACHMENTS_ONLY: 'false' }), false);
  assert.equal(inlineOnlyAttachmentsEnabled({ THUNDERBIRD_MCP_INLINE_ATTACHMENTS_ONLY: 'true' }), true);
});

it('advertises the enforced contract for all four outbound tools only', () => {
  const names = ['saveDraft', 'sendMail', 'replyToMessage', 'forwardMessage', 'getMessage'];
  const response = { result: { tools: names.map(name => ({
    name, description: 'Original description',
    inputSchema: { properties: { attachments: { type: 'array' } } },
  })) } };
  describeBase64Attachments(response);
  for (const tool of response.result.tools.slice(0, 4)) {
    assert.match(tool.description, /separate container/);
    assert.deepEqual(tool.inputSchema.properties.attachments.items.required, ['name', 'contentType', 'base64']);
  }
  assert.equal(response.result.tools[4].description, 'Original description');
});

it('accepts original file bytes and rejects paths, truncation and malformed objects', () => {
  const attachment = { name: 'image.png', contentType: 'image/png', base64: Buffer.from([0, 255, 10, 42]).toString('base64') };
  assert.doesNotThrow(() => validateBase64Attachments({ attachments: [attachment] }));
  assert.doesNotThrow(() => validateBase64Attachments({}));
  for (const entry of ['/tmp/thunderbird-mcp/image.png', { ...attachment, base64: 'AAA' },
    { ...attachment, base64: 'AA A' }, { ...attachment, name: '../image.png' },
    { ...attachment, base64: 'data:image/png;base64,AAAA' }]) {
    assert.throws(() => validateBase64Attachments({ attachments: [entry] }));
  }
});

it('rejects path calls through stdio before attempting Thunderbird access', () => {
  const names = ['saveDraft', 'sendMail', 'replyToMessage', 'forwardMessage'];
  const input = names.map((name, id) => JSON.stringify({
    jsonrpc: '2.0', id, method: 'tools/call',
    params: { name, arguments: { attachments: ['/tmp/not-an-agent-file.png'] } },
  })).join('\n') + '\n';
  const result = spawnSync(process.execPath, [path.resolve(__dirname, '../mcp-bridge.cjs')], {
    env: { ...process.env, THUNDERBIRD_MCP_INLINE_ATTACHMENTS_ONLY: 'true' },
    input, encoding: 'utf8', timeout: 5000,
  });
  assert.equal(result.status, 0, result.stderr);
  const responses = result.stdout.trim().split('\n').map(line => JSON.parse(line));
  assert.equal(responses.length, 4);
  for (const response of responses) {
    assert.equal(response.error.code, -32602);
    assert.match(response.error.message, /File paths are not supported/);
  }
});
