const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

// These tests read the real ATTRIB_MAP / ACTION_MAP out of api.js rather than
// re-declaring them, so the constants cannot silently drift from Thunderbird's
// IDL. api.js itself runs in XPCOM and cannot be required directly.
const API_SOURCE = fs.readFileSync(
  path.join(__dirname, "..", "extension", "mcp_server", "api.js"),
  "utf8"
);

function extractObjectLiteral(name) {
  const declaration = `const ${name} = {`;
  const start = API_SOURCE.indexOf(declaration);
  assert.notEqual(start, -1, `${name} declaration not found in api.js`);
  const open = API_SOURCE.indexOf("{", start);
  let depth = 0;
  let end = -1;
  for (let i = open; i < API_SOURCE.length; i++) {
    if (API_SOURCE[i] === "{") depth++;
    else if (API_SOURCE[i] === "}") {
      depth--;
      if (depth === 0) { end = i; break; }
    }
  }
  assert.notEqual(end, -1, `Unbalanced braces while reading ${name}`);
  // The literal maps names to integers only -- safe to evaluate.
  return Function(`"use strict"; return (${API_SOURCE.slice(open, end + 1)});`)();
}

// Authoritative values from mailnews/search/public/nsMsgSearchCore.idl
const NS_MSG_SEARCH_ATTRIB = {
  subject: 0, from: 1, body: 2, date: 3, priority: 4,
  status: 5, to: 6, cc: 7, toOrCc: 8, allAddresses: 9,
  ageInDays: 12, size: 14, anyText: 15, tag: 16,
  hasAttachment: 44, junkStatus: 45, junkPercent: 46,
  otherHeader: 52,
};

// Authoritative values from mailnews/search/public/nsMsgFilterCore.idl
const NS_MSG_FILTER_ACTION = {
  moveToFolder: 1, changePriority: 2, delete: 3,
  markRead: 4, killThread: 5, watchThread: 6,
  markFlagged: 7, reply: 9, forward: 10,
  stopExecution: 11, deleteFromServer: 12, leaveOnServer: 13,
  junkScore: 14, fetchBody: 15, copyToFolder: 16,
  addTag: 17, killSubthread: 18, markUnread: 19,
};

describe("filter enum values match Thunderbird's IDL", () => {
  it("ATTRIB_MAP matches nsMsgSearchAttrib", () => {
    assert.deepStrictEqual(extractObjectLiteral("ATTRIB_MAP"), NS_MSG_SEARCH_ATTRIB);
  });

  it("ACTION_MAP matches nsMsgFilterAction", () => {
    assert.deepStrictEqual(extractObjectLiteral("ACTION_MAP"), NS_MSG_FILTER_ACTION);
  });

  it("does not map stopExecution onto DeleteFromPop3Server", () => {
    // Regression guard: stopExecution was 0x0C (12), which is
    // DeleteFromPop3Server -- a destructive action on POP3 accounts.
    const actions = extractObjectLiteral("ACTION_MAP");
    assert.equal(actions.stopExecution, 11);
    assert.notEqual(actions.stopExecution, actions.deleteFromServer);
  });

  it("keeps the attribute enum sparse where the IDL is sparse", () => {
    // Regression guard: the old map assumed a dense 0..16 run, which silently
    // aliased ageInDays->Location, tag->AgeInDays, otherHeader->Keywords.
    const attribs = extractObjectLiteral("ATTRIB_MAP");
    assert.equal(attribs.ageInDays, 12);
    assert.equal(attribs.tag, 16);
    assert.equal(attribs.hasAttachment, 44);
    assert.equal(attribs.otherHeader, 52);
  });

  it("routes every numeric attribute to a non-str value member", () => {
    const attribs = extractObjectLiteral("ATTRIB_MAP");
    const numericFields = extractObjectLiteral("NUMERIC_VALUE_FIELDS");
    for (const name of ["priority", "status", "ageInDays", "size", "hasAttachment", "junkStatus", "junkPercent"]) {
      assert.ok(
        Object.prototype.hasOwnProperty.call(numericFields, attribs[name]),
        `${name} (attrib ${attribs[name]}) has no typed value member`
      );
    }
    // date is handled on its own path, not via NUMERIC_VALUE_FIELDS
    assert.equal(Object.prototype.hasOwnProperty.call(numericFields, attribs.date), false);
  });
});

// Mirrors setSearchValue from api.js. nsIMsgSearchValue is a tagged union:
// writing .str for a date/numeric attribute throws NS_ERROR_ILLEGAL_VALUE.
const NUMERIC_VALUE_FIELDS = {
  4: "priority", 5: "status", 12: "age", 14: "size",
  44: "status", 45: "junkStatus", 46: "junkPercent",
};

function parseFilterDate(raw) {
  if (typeof raw === "string") {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(raw.trim());
    if (m) {
      return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getTime();
    }
  }
  return Date.parse(raw);
}

const MSG_FLAG_ATTACHMENT = 0x10000000;

function setSearchValue(value, attrib, raw) {
  if (attrib === 3) {
    const parsed = parseFilterDate(raw);
    if (!Number.isFinite(parsed)) {
      throw new Error(`Invalid date value: ${JSON.stringify(raw)}`);
    }
    value.date = parsed * 1000;
    return;
  }
  if (attrib === 44) {
    value.status = MSG_FLAG_ATTACHMENT;
    return;
  }
  const field = NUMERIC_VALUE_FIELDS[attrib];
  if (field) {
    const num = typeof raw === "number" ? raw : parseInt(raw, 10);
    if (!Number.isFinite(num)) {
      throw new Error(`Invalid numeric value: ${JSON.stringify(raw)}`);
    }
    value[field] = num;
    return;
  }
  value.str = raw == null ? "" : String(raw);
}

describe("setSearchValue routing", () => {
  it("writes date attributes as PRTime microseconds, never as str", () => {
    const value = {};
    setSearchValue(value, 3, "2024-01-01");
    assert.equal(value.date, new Date(2024, 0, 1).getTime() * 1000);
    assert.equal("str" in value, false);
  });

  it("reads a date-only value as a local calendar day, not UTC midnight", () => {
    // Thunderbird renders filter dates in local time. Parsing as UTC shifted
    // the stored day back by one in any timezone west of UTC, so "2024-01-01"
    // was written to msgFilterRules.dat as 31-Dec-2023.
    const value = {};
    setSearchValue(value, 3, "2024-01-01");
    const stored = new Date(value.date / 1000);
    assert.equal(stored.getFullYear(), 2024);
    assert.equal(stored.getMonth(), 0);
    assert.equal(stored.getDate(), 1);
  });

  it("still honours an explicit timestamp with timezone information", () => {
    const value = {};
    setSearchValue(value, 3, "2024-01-01T12:00:00Z");
    assert.equal(value.date, Date.parse("2024-01-01T12:00:00Z") * 1000);
  });

  it("writes ageInDays to .age as a number", () => {
    const value = {};
    setSearchValue(value, 12, "952");
    assert.strictEqual(value.age, 952);
    assert.equal("str" in value, false);
  });

  it("writes size and junkPercent to their own members", () => {
    const size = {};
    setSearchValue(size, 14, "2048");
    assert.strictEqual(size.size, 2048);

    const junk = {};
    setSearchValue(junk, 46, 80);
    assert.strictEqual(junk.junkPercent, 80);
  });

  it("writes the fixed Attachment flag for hasAttachment, ignoring the raw value", () => {
    // The operator (is/isnt) carries has-vs-hasn't. Writing a caller-supplied
    // number here yields a filter that silently matches nothing.
    const value = {};
    setSearchValue(value, 44, "whatever");
    assert.strictEqual(value.status, 0x10000000);

    const ignored = {};
    setSearchValue(ignored, 44, "1");
    assert.strictEqual(ignored.status, 0x10000000);
  });

  it("still writes string attributes to .str", () => {
    const value = {};
    setSearchValue(value, 1, "example.com");
    assert.strictEqual(value.str, "example.com");
  });

  it("coerces a null string value to an empty string", () => {
    const value = {};
    setSearchValue(value, 0, null);
    assert.strictEqual(value.str, "");
  });

  it("rejects unparseable dates and numbers instead of writing garbage", () => {
    assert.throws(() => setSearchValue({}, 3, "not-a-date"), /Invalid date value/);
    assert.throws(() => setSearchValue({}, 12, "soon"), /Invalid numeric value/);
  });
});
