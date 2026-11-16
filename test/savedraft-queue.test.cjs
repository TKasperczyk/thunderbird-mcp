"use strict";

/**
 * saveDraft serializes concurrent draft saves through an async FIFO queue.
 *
 * Overlapping nsIMsgSend / IMAP APPEND calls contend on the Drafts folder and
 * provider rate limits. These tests verify FIFO ordering, non-overlapping
 * execution, and queue recovery after failures or thrown exceptions.
 */

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");

function snippet(name) {
  const startMarker = `// BEGIN ${name}`;
  const endMarker = `// END ${name}`;
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

const Ci = {
  nsIMsgComposeParams: Symbol("nsIMsgComposeParams"),
  nsIMsgCompFields: Symbol("nsIMsgCompFields"),
  nsIMsgCompType: { New: 0 },
  nsIMsgCompDeliverMode: { SaveAsDraft: 4 },
};

function loadSaveDraftHarness(sendImpl) {
  const identity = { key: "id1", draftFolder: "imap://user@example.com/Drafts" };
  const calls = [];
  const sandbox = {
    Ci,
    Cc: {
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => ({}) },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => ({ setHeader() {} }) },
    },
    setComposeIdentity: (p) => { p.identity = identity; return null; },
    resolveComposeFormat: () => ({ useHtml: false, format: 1 }),
    buildBodyWithSignature: body => body || "",
    filePathsToAttachDescs: (attachments) => {
      if (attachments === "THROW_ATTACH") throw new Error("attachment read error");
      return { descs: [], failed: [] };
    },
    getAccessibleFolder: uri => ({ folder: { URI: uri } }),
    sendMessageDirectly(composeFields, ...rest) {
      calls.push(composeFields.subject);
      return sendImpl(composeFields.subject, ...rest);
    },
    console,
  };

  vm.createContext(sandbox);
  vm.runInContext([
    snippet("DRAFT HELPERS"),
    snippet("SAVE DRAFT TOOL"),
    "this.saveDraft = saveDraft;",
    "this.enqueueSaveDraft = enqueueSaveDraft;",
  ].join("\n"), sandbox);

  return { saveDraft: sandbox.saveDraft, enqueueSaveDraft: sandbox.enqueueSaveDraft, calls };
}

describe("saveDraft queue serialization", () => {
  it("runs concurrent saveDraft calls sequentially in FIFO order", async () => {
    const deferreds = [];
    const active = { count: 0, maxConcurrent: 0 };
    const { saveDraft, calls } = loadSaveDraftHarness((subject) => {
      active.count++;
      active.maxConcurrent = Math.max(active.maxConcurrent, active.count);
      return new Promise((resolve) => {
        deferreds.push({ subject, resolve: (val) => { active.count--; resolve(val); } });
      });
    });

    const p1 = saveDraft("a@example.com", "First", "body1");
    const p2 = saveDraft("b@example.com", "Second", "body2");
    const p3 = saveDraft("c@example.com", "Third", "body3");

    await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(calls, ["First"], "second saveDraft must wait until first finishes");
    assert.equal(active.maxConcurrent, 1);

    deferreds[0].resolve({ success: true });
    const r1 = await p1;
    assert.equal(r1.success, true);

    await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(calls, ["First", "Second"]);
    assert.equal(active.maxConcurrent, 1);

    deferreds[1].resolve({ error: "Save failed (status: 0x80004005)" });
    const r2 = await p2;
    assert.match(r2.error, /Save failed/);

    await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(calls, ["First", "Second", "Third"]);
    assert.equal(active.maxConcurrent, 1);

    deferreds[2].resolve({ success: true });
    const r3 = await p3;
    assert.equal(r3.success, true);
  });

  it("recovers the queue when a prior task throws or rejects", async () => {
    const { saveDraft, enqueueSaveDraft, calls } = loadSaveDraftHarness(() => Promise.resolve({ success: true }));

    const rejected = enqueueSaveDraft(() => Promise.reject(new Error("boom")));
    const thrown = saveDraft("a@example.com", "Throws", "body", undefined, undefined, false, undefined, "THROW_ATTACH");
    const next = saveDraft("b@example.com", "AfterError", "body");

    await assert.rejects(rejected, /boom/);
    const thrownRes = await thrown;
    assert.match(thrownRes.error, /attachment read error/);

    const nextRes = await next;
    assert.equal(nextRes.success, true);
    assert.deepEqual(calls, ["AfterError"]);
  });
});
