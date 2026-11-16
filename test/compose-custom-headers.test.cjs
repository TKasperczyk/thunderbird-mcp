"use strict";

/**
 * The outbound tools (sendMail, dryRunCompose, replyToMessage, forwardMessage,
 * saveDraft) accept an optional `headers` list of extra X- headers. These tests
 * pin the shared validation: only X- names, no reserved Thunderbird or Send
 * Later headers, no duplicates, and single-line printable ASCII values.
 */

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { loadToolSchema } = require("./helpers/tool-schema.cjs");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");

function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function loadHelpers() {
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext([
    snippet("OUTBOUND ATTACHMENT LIMITS"),
    snippet("CUSTOM HEADER HELPERS"),
    "this.MAX_CUSTOM_HEADERS = MAX_CUSTOM_HEADERS;",
  ].join("\n"), sandbox);
  return sandbox;
}

const api = loadHelpers();
const header = (name, value = "yes") => ({ name, value });

describe("resolveCustomHeaders", () => {
  it("treats omitted and null headers as none", () => {
    for (const value of [undefined, null]) {
      const result = api.resolveCustomHeaders(value);
      assert.equal(result.error, undefined);
      assert.equal(result.headers.length, 0);
    }
  });

  it("keeps valid X- headers in order and spelling", () => {
    const input = [header("X-AI-Assisted", "Drafted with Claude, reviewed by sender"), header("x-ticket", "WOA-1234")];
    assert.deepEqual(JSON.parse(JSON.stringify(api.resolveCustomHeaders(input))), { headers: input });
  });

  for (const name of ["Bcc", "From", "Reply-To", "X-", "X-AI Assisted", "X-AI:Assisted", "X-é", "", 42]) {
    it(`rejects header name ${JSON.stringify(name)}`, () => {
      assert.match(api.resolveCustomHeaders([header(name)]).error, /must start with X-/);
    });
  }

  for (const name of ["X-Mozilla-Draft-Info", "x-mozilla-keys", "X-Send-Later-At", "X-Send-Later-Uuid", "X-Identity-Key", "X-Account-Key"]) {
    it(`rejects reserved header ${name}`, () => {
      assert.match(api.resolveCustomHeaders([header(name)]).error, /reserved/);
    });
  }

  it("rejects a name given twice, ignoring case", () => {
    assert.match(api.resolveCustomHeaders([header("X-AI"), header("x-ai")]).error, /more than once/);
  });

  for (const value of ["", "   ", "line\r\nBcc: spy@example.test", "tab\there", "café", 7]) {
    it(`rejects value ${JSON.stringify(value)}`, () => {
      assert.match(api.resolveCustomHeaders([header("X-AI", value)]).error, /printable ASCII/);
    });
  }

  it("rejects a header line longer than 998 characters", () => {
    assert.equal(api.resolveCustomHeaders([header("X-AI", "a".repeat(992))]).error, undefined);
    assert.match(api.resolveCustomHeaders([header("X-AI", "a".repeat(993))]).error, /longer than 998/);
  });

  it("rejects more than MAX_CUSTOM_HEADERS entries", () => {
    const many = Array.from({ length: api.MAX_CUSTOM_HEADERS + 1 }, (_, i) => header(`X-H${i}`));
    assert.match(api.resolveCustomHeaders(many).error, /at most/);
    assert.equal(api.resolveCustomHeaders(many.slice(1)).error, undefined);
  });

  it("rejects a non-array value", () => {
    assert.match(api.resolveCustomHeaders({ "X-AI": "yes" }).error, /must be an array/);
  });

  it("applies resolved headers with setHeader", () => {
    const set = [];
    api.applyCustomHeaders({ setHeader: (name, value) => set.push([name, value]) }, [header("X-AI", "yes")]);
    assert.deepEqual(set, [["X-AI", "yes"]]);
  });
});

describe("headers tool schema", () => {
  const tools = loadToolSchema().buildTools();
  for (const name of ["sendMail", "dryRunCompose", "replyToMessage", "forwardMessage", "saveDraft"]) {
    it(`${name} accepts a bounded headers list`, () => {
      const schema = tools.find(tool => tool.name === name).inputSchema.properties.headers;
      assert.equal(schema.type, "array");
      assert.equal(schema.maxItems, api.MAX_CUSTOM_HEADERS);
      assert.deepEqual([...schema.items.required], ["name", "value"]);
      assert.equal(schema.items.additionalProperties, false);
    });
  }
});
