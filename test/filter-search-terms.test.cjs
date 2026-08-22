"use strict";

/**
 * Tests for filter search term construction (createFilter / updateFilter).
 *
 * nsIMsgSearchValue is a tagged union: only the member matching the
 * attribute's type may be written. Assigning .str to a status/numeric/date
 * attribute throws NS_ERROR_ILLEGAL_VALUE at runtime, which stays invisible
 * until someone actually builds a filter on e.g. hasAttachment or size.
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

/** Minimal stand-in for nsIMsgFilter: mints terms and collects appended ones. */
function makeStubFilter() {
  const appended = [];
  return {
    appended,
    createTerm: () => makeStubTerm(),
    appendTerm: (t) => appended.push(t),
  };
}

function loadFilterSearchTermHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN FILTER SEARCH TERM HELPERS";
  const endMarker = "// END FILTER SEARCH TERM HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "filter search term helper start marker missing");
  assert.ok(end > start, "filter search term helper end marker missing");

  const sandbox = {
    Ci: { nsMsgMessageFlags: { Attachment: ATTACHMENT_FLAG } },
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
  vm.runInContext(`${source.slice(start, end)}
this.buildTerms = buildTerms;`, sandbox);
  return sandbox.buildTerms;
}

const buildTerms = loadFilterSearchTermHelpers();

function build(conditions) {
  const filter = makeStubFilter();
  buildTerms(filter, conditions);
  return filter.appended;
}

describe("buildTerms - value union dispatch", () => {
  it("writes .str for string attributes and leaves numeric members unset", () => {
    const [term] = build([{ attrib: "from", op: "contains", value: "news@example.com" }]);
    assert.equal(term.attrib, 1);
    assert.equal(term.op, 0);
    assert.equal(term.value.str, "news@example.com");
    assert.equal(term.value.status, undefined);
    assert.equal(term.value.size, undefined);
  });

  it("writes .status (not .str) for hasAttachment", () => {
    const [term] = build([{ attrib: "hasAttachment", op: "is", value: "true" }]);
    assert.equal(term.value.status, ATTACHMENT_FLAG);
    assert.equal(term.value.str, undefined, ".str must not be set for a status attribute");
  });

  it("writes .size / .age / .priority / .junkPercent as numbers", () => {
    const [size, age, prio, junk] = build([
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
    const [term] = build([{ attrib: "date", op: "isAfter", value: "2026-01-01T00:00:00Z" }]);
    assert.equal(term.value.date, Date.parse("2026-01-01T00:00:00Z") * 1000);
  });

  it("rejects an unparseable date rather than silently storing 0", () => {
    assert.throws(
      () => build([{ attrib: "date", op: "isAfter", value: "not-a-date" }]),
      /Invalid date value/
    );
  });

  it("appends every term to the filter in order", () => {
    const terms = build([
      { attrib: "subject", op: "contains", value: "one" },
      { attrib: "subject", op: "contains", value: "two" },
    ]);
    assert.equal(terms.length, 2);
    assert.equal(terms[0].value.str, "one");
    assert.equal(terms[1].value.str, "two");
  });

  it("defaults booleanAnd to true and honours an explicit false", () => {
    const [a, b] = build([
      { attrib: "subject", op: "contains", value: "a" },
      { attrib: "subject", op: "contains", value: "b", booleanAnd: false },
    ]);
    assert.equal(a.booleanAnd, true);
    assert.equal(b.booleanAnd, false);
  });

  it("sets arbitraryHeader only when a header is supplied", () => {
    const [withHeader, without] = build([
      { attrib: "otherHeader", op: "contains", value: "bulk", header: "X-Precedence" },
      { attrib: "subject", op: "contains", value: "hi" },
    ]);
    assert.equal(withHeader.arbitraryHeader, "X-Precedence");
    assert.equal(without.arbitraryHeader, null);
  });

  it("rejects attributes and operators outside the allow-list", () => {
    assert.throws(() => build([{ attrib: "nope", op: "contains", value: "x" }]), /Unknown attribute/);
    assert.throws(() => build([{ attrib: "subject", op: "nope", value: "x" }]), /Unknown operator/);
    // Raw enum values must not bypass the named allow-list.
    assert.throws(() => build([{ attrib: 13, op: 2, value: "x" }]), /Unknown attribute/);
  });

  it("coerces a missing value to an empty string instead of throwing", () => {
    const [term] = build([{ attrib: "subject", op: "isEmpty" }]);
    assert.equal(term.value.str, "");
  });
});
