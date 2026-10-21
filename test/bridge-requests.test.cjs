'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const bridgePath = path.resolve(__dirname, '../mcp-bridge.cjs');
const source = fs.readFileSync(bridgePath, 'utf8');

function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `${name} markers missing`);
  return source.slice(start, end);
}

// Searches found production request/retry helpers in the bridge, but no HTTP
// request test harness. Exercise those helpers through their production markers.
function loadRequests({ candidates = 1, profileRoots = [], inlineAttachmentPaths = async () => {} } = {}) {
  const requests = [];
  let candidateIndex = 0;
  const runtime = {
    Buffer,
    setTimeout,
    process: { platform: process.platform },
    inlineAttachmentPaths,
    cachedCandidateList: Array.from({ length: candidates }, (_, index) => ({
      data: { port: 21000 + index, token: 'a'.repeat(64), profileRoots: profileRoots[index] },
    })),
    THUNDERBIRD_HOSTS: ['127.0.0.1'],
    CONNECTION_MAX_RETRIES: 0,
    CONNECTION_RETRY_DELAY_MS: 1000,
    readConnectionInfo: () => candidateIndex < candidates
      ? runtime.cachedCandidateList[candidateIndex].data
      : null,
    advanceToNextCandidate() {
      candidateIndex++;
      return runtime.readConnectionInfo();
    },
    clearConnectionCache() {},
    buildConnectionDiscoveryErrorMessage: () => 'Connection discovery failed',
    sanitizeJson: (data) => data,
    http: {
      request(options, onResponse) {
        const request = new EventEmitter();
        request.options = options;
        request.onResponse = onResponse;
        request.setTimeout = (timeoutMs, callback) => {
          request.timeoutMs = timeoutMs;
          request.onTimeout = callback;
        };
        request.write = (data) => { request.data = data; };
        request.end = () => { request.ended = true; };
        request.destroy = () => {
          request.destroyed = true;
          request.emit('error', new Error('socket destroyed'));
          request.emit('close');
        };
        requests.push(request);
        return request;
      },
    },
  };
  vm.createContext(runtime);
  vm.runInContext([
    snippet('BRIDGE ATTACHMENT POLICY'),
    snippet('BRIDGE HTTP REQUEST HELPERS'),
    snippet('BRIDGE FORWARDING'),
  ].join('\n'), runtime);
  return { runtime, requests };
}

function toolCall(name, args = {}) {
  return { jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } };
}

function connect(request, { reused = false, pending = false } = {}) {
  const socket = new EventEmitter();
  socket.connecting = !reused;
  request.emit('socket', socket);
  if (!reused && !pending) {
    socket.connecting = false;
    socket.emit('connect');
  }
  return socket;
}

function response(request, { statusCode = 200, complete = false } = {}) {
  const res = new EventEmitter();
  res.statusCode = statusCode;
  res.complete = complete;
  request.onResponse(res);
  return res;
}

function finishResponse(request, result = { jsonrpc: '2.0', id: 1, result: {} }) {
  const res = response(request, { complete: true });
  res.emit('data', Buffer.from(JSON.stringify(result)));
  res.emit('end');
  res.emit('close');
  request.emit('close');
}

const mailOperations = [
  ['sendMail', { skipReview: true }, /Sent folder and the Outbox/],
  ['replyToMessage', { skipReview: true }, /Sent folder and the Outbox/],
  ['forwardMessage', { skipReview: true }, /Sent folder and the Outbox/],
  ['sendMail', { skipReview: 'true' }, /Sent folder and the Outbox/],
  ['replyToMessage', { skipReview: 'true' }, /Sent folder and the Outbox/],
  ['forwardMessage', { skipReview: 'true' }, /Sent folder and the Outbox/],
  ['saveDraft', {}, /Drafts/],
  ['saveDraft', { replaceMessageId: '  ' }, /Drafts/],
  ['replyToMessage', { saveAsDraft: true }, /Drafts/],
  ['replyToMessage', { saveAsDraft: 'true' }, /Drafts/],
];

describe('bridge request outcomes', () => {
  it('preflights against verified profile metadata and preserves paths across candidate retries', async () => {
    const profile = process.platform === 'win32' ? 'D:/profile' : '/profile';
    const otherProfile = process.platform === 'win32' ? 'D:/other-profile' : '/other-profile';
    const conversions = [];
    const { runtime, requests } = loadRequests({
      candidates: 2,
      profileRoots: [[profile, '//server/share/profiles'], [null, otherProfile]],
      inlineAttachmentPaths: async (args, options) => {
        conversions.push({ attachments: [...args.attachments], roots: [...options.profileRoots] });
        args.attachments = [{ name: 'report.pdf', base64: 'QQ==' }];
      },
    });
    const message = toolCall('sendMail', { attachments: ['/documents/report.pdf'], profileRoots: ['/untrusted'] });
    const pending = runtime.forwardToThunderbird(message);
    await new Promise(setImmediate);
    assert.deepEqual(conversions[0].attachments, ['/documents/report.pdf']);
    assert.deepEqual(new Set(conversions[0].roots), new Set([profile + '/', otherProfile + '/']));
    connect(requests[0], { pending: true });
    requests[0].emit('error', Object.assign(new Error('refused'), { code: 'ECONNREFUSED' }));
    await new Promise(setImmediate);
    assert.equal(conversions.length, 2);
    assert.deepEqual(conversions[1].attachments, ['/documents/report.pdf']);
    assert.deepEqual(message.params.arguments.attachments, ['/documents/report.pdf']);
    assert.equal(JSON.parse(requests[1].data).params.arguments.attachments[0].base64, 'QQ==');
    finishResponse(requests[1]);
    await pending;
  });

  it('returns attachment refusals without sending an HTTP request', async () => {
    const { runtime, requests } = loadRequests({
      inlineAttachmentPaths: async () => { throw new Error('Sensitive attachment path blocked'); },
    });
    const result = await runtime.forwardToThunderbird(toolCall('saveDraft', { attachments: ['/profile/report.pdf'] }));
    assert.equal(result.error.code, -32602);
    assert.equal(requests.length, 0);
  });

  for (const [name, args, folderAdvice] of mailOperations) {
    it(`${name} ${JSON.stringify(args)} waits 150 seconds and reports an uncertain timeout`, async () => {
      const { runtime, requests } = loadRequests();
      const pending = runtime.forwardToThunderbird(toolCall(name, args));
      const request = requests[0];
      assert.equal(request.timeoutMs, 150000);
      assert.equal(request.options.headers.Authorization, `Bearer ${'a'.repeat(64)}`);
      connect(request);
      request.onTimeout();
      await assert.rejects(pending, (error) => {
        assert.match(error.message, /timed out after 150 seconds/);
        assert.match(error.message, /outcome is unknown/);
        assert.match(error.message, folderAdvice);
        assert.match(error.message, /before retrying/);
        assert.doesNotMatch(error.message, /socket destroyed/);
        return true;
      });
      assert.equal(request.destroyed, true);
      assert.equal(requests.length, 1);
    });
  }

  it("a replacing saveDraft outlasts the extension's save plus Drafts check (240 s)", async () => {
    const { runtime, requests } = loadRequests();
    const pending = runtime.forwardToThunderbird(toolCall('saveDraft', { replaceMessageId: 'old@example.test' }));
    connect(requests[0]);
    assert.equal(requests[0].timeoutMs, 240000);
    requests[0].onTimeout();
    await assert.rejects(pending, /timed out after 240 seconds/);
  });

  for (const message of [
    toolCall('listFolders'),
    toolCall('sendMail'),
    ...['sendMail', 'replyToMessage', 'forwardMessage'].flatMap((name) =>
      [false, 'false', 1, 'TRUE', ' true', 'true ', 'yes'].map((skipReview) => toolCall(name, { skipReview }))
    ),
    ...[false, 'false', 1, 'TRUE', ' true', 'true ', 'yes'].map((saveAsDraft) => toolCall('replyToMessage', { saveAsDraft })),
    toolCall('unrelatedTool', { skipReview: true }),
    { method: 'tools/list', params: { name: 'saveDraft' } },
  ]) {
    it(`keeps 30 seconds for ${JSON.stringify(message)}`, async () => {
      const { runtime, requests } = loadRequests();
      const pending = runtime.forwardToThunderbird(message);
      connect(requests[0]);
      assert.equal(requests[0].timeoutMs, 30000);
      requests[0].onTimeout();
      await assert.rejects(pending, (error) => {
        assert.match(error.message, /timed out after 30 seconds/);
        assert.doesNotMatch(error.message, /outcome is unknown|before retrying/);
        return true;
      });
    });
  }

  for (const [name, args, folderAdvice] of mailOperations) {
    for (const failure of ['request error', 'request close', 'response error', 'response aborted', 'response close']) {
      it(`${name} ${JSON.stringify(args)} reports unknown outcome after ${failure}`, async () => {
        const { runtime, requests } = loadRequests({ candidates: 2 });
        const pending = runtime.forwardToThunderbird(toolCall(name, args));
        const request = requests[0];
        connect(request);
        if (failure === 'request error') {
          // Even an ordinarily retryable error must not replay a dispatched send.
          request.emit('error', Object.assign(new Error('connection lost'), { code: 'ECONNREFUSED' }));
        } else if (failure === 'request close') {
          request.emit('close');
        } else {
          const res = response(request);
          res.emit('data', Buffer.from('{"result":'));
          res.emit(failure.slice('response '.length), new Error('response interrupted'));
        }
        await assert.rejects(pending, (error) => {
          assert.match(error.message, /outcome is unknown/);
          assert.match(error.message, folderAdvice);
          assert.match(error.message, /before retrying/);
          return true;
        });
        assert.equal(requests.length, 1, 'uncertain operations must not be retried');
      });
    }
  }

  it('recognizes a reused socket as already connected', async () => {
    const { runtime, requests } = loadRequests();
    const pending = runtime.forwardToThunderbird(toolCall('sendMail', { skipReview: true }));
    connect(requests[0], { reused: true });
    requests[0].emit('error', new Error('connection reset'));
    await assert.rejects(pending, /outcome is unknown/);
  });

  for (const failure of ['error', 'timeout', 'close']) {
    it(`keeps a never-opened connection ${failure} ordinary`, async () => {
      const { runtime, requests } = loadRequests();
      const pending = runtime.forwardToThunderbird(toolCall('saveDraft'));
      connect(requests[0], { pending: true });
      if (failure === 'error') requests[0].emit('error', Object.assign(new Error('refused'), { code: 'ECONNREFUSED' }));
      if (failure === 'timeout') requests[0].onTimeout();
      if (failure === 'close') requests[0].emit('close');
      await assert.rejects(pending, (error) => {
        assert.doesNotMatch(error.message, /outcome is unknown|before retrying/);
        return true;
      });
    });
  }

  it('retries another connection candidate only when the first never connected', async () => {
    const { runtime, requests } = loadRequests({ candidates: 2 });
    const pending = runtime.forwardToThunderbird(toolCall('sendMail', { skipReview: true }));
    connect(requests[0], { pending: true });
    requests[0].emit('error', Object.assign(new Error('refused'), { code: 'ECONNREFUSED' }));
    await new Promise(setImmediate);
    assert.equal(requests.length, 2);
    assert.equal(requests[1].timeoutMs, 150000);
    finishResponse(requests[1]);
    assert.equal((await pending).id, 1);
  });

  it('continues to retry an explicit authentication rejection', async () => {
    const { runtime, requests } = loadRequests({ candidates: 2 });
    const pending = runtime.forwardToThunderbird(toolCall('saveDraft'));
    connect(requests[0]);
    const rejected = response(requests[0], { statusCode: 403, complete: true });
    rejected.emit('end');
    requests[0].emit('close');
    await new Promise(setImmediate);
    assert.equal(requests.length, 2);
    finishResponse(requests[1]);
    assert.equal((await pending).id, 1);
  });

  it('settles success once despite normal close events and later errors', async () => {
    const { runtime, requests } = loadRequests();
    const pending = runtime.forwardToThunderbird(toolCall('saveDraft'));
    const expected = { jsonrpc: '2.0', id: 1, result: { saved: true } };
    finishResponse(requests[0], expected);
    requests[0].emit('error', new Error('late error'));
    assert.deepEqual(JSON.parse(JSON.stringify(await pending)), expected);
  });

  it('does not replace an uncertain timeout with a late successful response', async () => {
    const { runtime, requests } = loadRequests();
    const pending = runtime.forwardToThunderbird(toolCall('sendMail', { skipReview: true }));
    connect(requests[0]);
    requests[0].onTimeout();
    finishResponse(requests[0]);
    await assert.rejects(pending, /timed out.*outcome is unknown/);
  });
});

describe('bridge response envelopes', () => {
  for (const id of [1, 0, '1', '0', 'request-id', null]) {
    it(`accepts a matching ${JSON.stringify(id)} response id`, async () => {
      const { runtime, requests } = loadRequests();
      const pending = runtime.forwardToThunderbird({ ...toolCall('listFolders'), id });
      const expected = { jsonrpc: '2.0', id, result: {} };
      finishResponse(requests[0], expected);
      assert.deepEqual(JSON.parse(JSON.stringify(await pending)), expected);
    });
  }

  for (const result of [null, false, 0, '', []]) {
    it(`accepts a present result ${JSON.stringify(result)}`, async () => {
      const { runtime, requests } = loadRequests();
      const pending = runtime.forwardToThunderbird(toolCall('listFolders'));
      const expected = { jsonrpc: '2.0', id: 1, result };
      finishResponse(requests[0], expected);
      assert.deepEqual(JSON.parse(JSON.stringify(await pending)), expected);
    });
  }

  it('passes through a valid JSON-RPC error', async () => {
    const { runtime, requests } = loadRequests();
    const pending = runtime.forwardToThunderbird(toolCall('listFolders'));
    const expected = { jsonrpc: '2.0', id: 1, error: { code: -32602, message: 'Invalid params', data: {} } };
    finishResponse(requests[0], expected);
    assert.deepEqual(JSON.parse(JSON.stringify(await pending)), expected);
  });

  for (const [name, reply] of [
    ['null', null],
    ['array', [{ jsonrpc: '2.0', id: 1, result: {} }]],
    ['string', 'not a response'],
    ['number', 1],
    ['boolean', true],
    ['missing version', { id: 1, result: {} }],
    ['wrong version', { jsonrpc: '1.0', id: 1, result: {} }],
    ['missing id', { jsonrpc: '2.0', result: {} }],
    ['wrong id', { jsonrpc: '2.0', id: 2, result: {} }],
    ['string id for numeric request', { jsonrpc: '2.0', id: '1', result: {} }],
    ['null id for numeric request', { jsonrpc: '2.0', id: null, result: {} }],
    ['missing result and error', { jsonrpc: '2.0', id: 1 }],
    ['both result and error', { jsonrpc: '2.0', id: 1, result: {}, error: { code: -32603, message: 'Error' } }],
    ['null error', { jsonrpc: '2.0', id: 1, error: null }],
    ['array error', { jsonrpc: '2.0', id: 1, error: [] }],
    ['string error', { jsonrpc: '2.0', id: 1, error: 'Error' }],
    ['missing error code', { jsonrpc: '2.0', id: 1, error: { message: 'Error' } }],
    ['string error code', { jsonrpc: '2.0', id: 1, error: { code: '-32603', message: 'Error' } }],
    ['fractional error code', { jsonrpc: '2.0', id: 1, error: { code: 0.5, message: 'Error' } }],
    ['missing error message', { jsonrpc: '2.0', id: 1, error: { code: -32603 } }],
    ['numeric error message', { jsonrpc: '2.0', id: 1, error: { code: -32603, message: 1 } }],
  ]) {
    it(`rejects ${name} without replaying the request`, async () => {
      const { runtime, requests } = loadRequests({ candidates: 2 });
      const pending = runtime.forwardToThunderbird(toolCall('saveDraft'));
      finishResponse(requests[0], reply);
      await assert.rejects(pending, new RegExp(`port ${runtime.cachedCandidateList[0].data.port}.*something that is not Thunderbird MCP`));
      assert.equal(requests.length, 1);
    });
  }

  it('does not coerce a numeric response id to a string request id', async () => {
    const { runtime, requests } = loadRequests({ candidates: 2 });
    const pending = runtime.forwardToThunderbird({ ...toolCall('listFolders'), id: '0' });
    finishResponse(requests[0], { jsonrpc: '2.0', id: 0, result: {} });
    await assert.rejects(pending, new RegExp(`port ${runtime.cachedCandidateList[0].data.port}.*something that is not Thunderbird MCP`));
    assert.equal(requests.length, 1);
  });

  it('reports invalid JSON without logging the reply or trying another candidate', async () => {
    const { runtime, requests } = loadRequests({ candidates: 2 });
    const pending = runtime.forwardToThunderbird(toolCall('listFolders'));
    const res = response(requests[0], { complete: true });
    res.emit('data', Buffer.from('<html>private reply contents</html>'));
    res.emit('end');
    await assert.rejects(pending, (error) => {
      assert.match(error.message, new RegExp(`port ${runtime.cachedCandidateList[0].data.port}.*something that is not Thunderbird MCP.*invalid JSON`));
      assert.doesNotMatch(error.message, /private reply contents/);
      return true;
    });
    assert.equal(requests.length, 1);
  });

  it('validates the envelope after the JSON sanitization fallback', async () => {
    const { runtime, requests } = loadRequests();
    const expected = { jsonrpc: '2.0', id: 1, result: {} };
    runtime.sanitizeJson = (data) => {
      assert.equal(data, 'needs sanitizing');
      return JSON.stringify(expected);
    };
    const pending = runtime.forwardToThunderbird(toolCall('listFolders'));
    const res = response(requests[0], { complete: true });
    res.emit('data', Buffer.from('needs sanitizing'));
    res.emit('end');
    assert.deepEqual(JSON.parse(JSON.stringify(await pending)), expected);
  });

  it('rejects a malformed envelope from the JSON sanitization fallback', async () => {
    const { runtime, requests } = loadRequests({ candidates: 2 });
    runtime.sanitizeJson = () => JSON.stringify({ jsonrpc: '2.0', id: '1', result: {} });
    const pending = runtime.forwardToThunderbird(toolCall('listFolders'));
    const res = response(requests[0], { complete: true });
    res.emit('data', Buffer.from('needs sanitizing'));
    res.emit('end');
    await assert.rejects(pending, new RegExp(`port ${runtime.cachedCandidateList[0].data.port}.*something that is not Thunderbird MCP`));
    assert.equal(requests.length, 1);
  });

  it('names the selected port after a 403 advances to another candidate', async () => {
    const { runtime, requests } = loadRequests({ candidates: 3 });
    const pending = runtime.forwardToThunderbird(toolCall('listFolders'));
    const rejected = response(requests[0], { statusCode: 403, complete: true });
    rejected.emit('end');
    await new Promise(setImmediate);
    assert.equal(requests.length, 2);
    finishResponse(requests[1], { error: 'not a JSON-RPC response' });
    await assert.rejects(pending, new RegExp(`port ${runtime.cachedCandidateList[1].data.port}.*something that is not Thunderbird MCP`));
    assert.equal(requests.length, 2);
  });
});

describe('bridge loader startup', () => {
  for (const disabled of [false, true]) {
    it(`production startup ${disabled ? 'honors an explicit opt-out' : 'runs when require.main differs'}`, () => {
      let starts = 0;
      vm.runInNewContext(snippet('BRIDGE STARTUP'), {
        process: { env: disabled ? { THUNDERBIRD_MCP_NO_AUTOSTART: '1' } : {} },
        require: { main: {} },
        module: {},
        startBridge() { starts++; },
      });
      assert.equal(starts, disabled ? 0 : 1);
    });
  }

  for (const [mode, args] of [
    ['direct executable', [bridgePath]],
    ['bootstrap require', ['-e', `require(${JSON.stringify(bridgePath)})`]],
  ]) {
    it(`answers stdio from a ${mode}`, () => {
      const result = execFileSync(process.execPath, args, {
        input: '{"jsonrpc":"2.0","id":7,"method":"ping"}\n',
        env: { ...process.env, THUNDERBIRD_MCP_NO_AUTOSTART: '0' },
        timeout: 5000,
        encoding: 'utf8',
      });
      assert.deepEqual(JSON.parse(result), { jsonrpc: '2.0', id: 7, result: {} });
    });
  }

  it('can be required for exports without installing stdin handlers', () => {
    const result = execFileSync(process.execPath, ['-e', `
      const bridge = require(${JSON.stringify(bridgePath)});
      console.log(typeof bridge.startBridge + ':' + process.stdin.listenerCount('data'));
    `], {
      env: { ...process.env, THUNDERBIRD_MCP_NO_AUTOSTART: '1' },
      timeout: 5000,
      encoding: 'utf8',
    });
    assert.equal(result.trim(), 'function:0');
  });
});
