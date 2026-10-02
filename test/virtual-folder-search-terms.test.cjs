"use strict";

/**
 * Tests for saved-search (virtual folder) term construction.
 *
 * nsIMsgSearchValue is a tagged union: only the member matching the
 * attribute's type may be written. Assigning .str to a status/numeric/date
 * attribute throws NS_ERROR_ILLEGAL_VALUE at runtime, which is invisible
 * until someone actually creates a saved search on e.g. hasAttachment.
 * These tests pin the per-attribute dispatch.
 */

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ATTACHMENT_FLAG = 0x10000000;

function makeStubTerm() {
  return {
    attrib: null,
    op: null,
    booleanAnd: null,
    arbitraryHeader: null,
    value: {
      attrib: null,
      str: undefined,
      status: undefined,
      priority: undefined,
      age: undefined,
      size: undefined,
      junkPercent: undefined,
      date: undefined,
    },
  };
}

function loadVirtualFolderHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN VIRTUAL FOLDER (SAVED SEARCH) HELPERS";
  const endMarker = "// END VIRTUAL FOLDER (SAVED SEARCH) HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "virtual folder helper start marker missing");
  assert.ok(end > start, "virtual folder helper end marker missing");

  const snippet = source.slice(start, end);
  const sandbox = {
    // Only buildSearchTerms is exercised here; the other helpers in the block
    // resolve their XPCOM dependencies lazily at call time, so they can be
    // declared without stubbing MailServices/ChromeUtils.
    Cc: {
      "@mozilla.org/messenger/searchSession;1": {
        createInstance: () => ({ createTerm: () => makeStubTerm() }),
      },
    },
    Ci: {
      nsIMsgSearchSession: {},
      nsMsgMessageFlags: { Attachment: ATTACHMENT_FLAG },
    },
    ATTRIB_MAP: {
      subject: 0, from: 1, body: 2, date: 3, priority: 4,
      status: 5, to: 6, cc: 7, toOrCc: 8, allAddresses: 9,
      ageInDays: 10, size: 11, tag: 12, hasAttachment: 13,
      junkStatus: 14, junkPercent: 15, otherHeader: 16,
    },
    OP_MAP: {
      contains: 0, doesntContain: 1, is: 2, isnt: 3, isEmpty: 4,
      isBefore: 5, isAfter: 6, beginsWith: 9, endsWith: 10,
      isGreaterThan: 13, isLessThan: 14, matches: 19, doesntMatch: 20,
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(`${snippet}
this.buildSearchTerms = buildSearchTerms;`, sandbox);
  return sandbox.buildSearchTerms;
}

const buildSearchTerms = loadVirtualFolderHelpers();

describe("buildSearchTerms - value union dispatch", () => {
  it("writes .str for string attributes and leaves numeric members unset", () => {
    const [term] = buildSearchTerms([
      { attrib: "from", op: "contains", value: "invoice@example.com" },
    ]);
    assert.equal(term.attrib, 1);
    assert.equal(term.op, 0);
    assert.equal(term.value.str, "invoice@example.com");
    assert.equal(term.value.status, undefined);
    assert.equal(term.value.size, undefined);
  });

  it("writes .status (not .str) for hasAttachment", () => {
    const [term] = buildSearchTerms([
      { attrib: "hasAttachment", op: "is", value: "true" },
    ]);
    assert.equal(term.value.status, ATTACHMENT_FLAG);
    assert.equal(term.value.str, undefined, ".str must not be set for a status attribute");
  });

  it("writes .size / .age / .priority / .junkPercent as numbers", () => {
    const [size, age, prio, junk] = buildSearchTerms([
      { attrib: "size", op: "isGreaterThan", value: "1024" },
      { attrib: "ageInDays", op: "isGreaterThan", value: "30" },
      { attrib: "priority", op: "is", value: "5" },
      { attrib: "junkPercent", op: "isGreaterThan", value: "90" },
    ]);
    assert.equal(size.value.size, 1024);
    assert.equal(age.value.age, 30);
    assert.equal(prio.value.priority, 5);
    assert.equal(junk.value.junkPercent, 90);
    assert.equal(size.value.str, undefined);
  });

  it("writes .date as PRTime microseconds", () => {
    const [term] = buildSearchTerms([
      { attrib: "date", op: "isAfter", value: "2026-01-01T00:00:00Z" },
    ]);
    assert.equal(term.value.date, Date.parse("2026-01-01T00:00:00Z") * 1000);
  });

  it("rejects an unparseable date rather than silently storing 0", () => {
    assert.throws(
      () => buildSearchTerms([{ attrib: "date", op: "isAfter", value: "not-a-date" }]),
      /Invalid date value/
    );
  });

  it("defaults booleanAnd to true and honours an explicit false", () => {
    const [a, b] = buildSearchTerms([
      { attrib: "subject", op: "contains", value: "Rechnung" },
      { attrib: "subject", op: "contains", value: "Invoice", booleanAnd: false },
    ]);
    assert.equal(a.booleanAnd, true);
    assert.equal(b.booleanAnd, false);
  });

  it("sets arbitraryHeader only when a header is supplied", () => {
    const [withHeader, without] = buildSearchTerms([
      { attrib: "otherHeader", op: "contains", value: "bulk", header: "X-Precedence" },
      { attrib: "subject", op: "contains", value: "hi" },
    ]);
    assert.equal(withHeader.arbitraryHeader, "X-Precedence");
    assert.equal(without.arbitraryHeader, null);
  });

  it("rejects attributes and operators outside the allow-list", () => {
    assert.throws(
      () => buildSearchTerms([{ attrib: "nope", op: "contains", value: "x" }]),
      /Unknown attribute/
    );
    assert.throws(
      () => buildSearchTerms([{ attrib: "subject", op: "nope", value: "x" }]),
      /Unknown operator/
    );
    // Raw enum values must not bypass the named allow-list.
    assert.throws(
      () => buildSearchTerms([{ attrib: 13, op: 2, value: "x" }]),
      /Unknown attribute/
    );
  });

  it("coerces a missing value to an empty string instead of throwing", () => {
    const [term] = buildSearchTerms([{ attrib: "subject", op: "isEmpty" }]);
    assert.equal(term.value.str, "");
  });
});
