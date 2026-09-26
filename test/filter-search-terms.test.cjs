"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

// The real nsMsgSearchAttrib enum, from
// comm-central/mailnews/search/public/nsMsgSearchCore.idl. Note the gaps: the
// enum is not contiguous past AllAddresses(9), which is what the original
// ATTRIB_MAP got wrong.
const ATTRIB = {
  Custom: -2, Default: -1,
  Subject: 0, Sender: 1, Body: 2, Date: 3, Priority: 4, MsgStatus: 5,
  To: 6, CC: 7, ToOrCC: 8, AllAddresses: 9, Location: 10, MessageKey: 11,
  AgeInDays: 12, FolderInfo: 13, Size: 14, AnyText: 15, Keywords: 16,
  HasAttachmentStatus: 44, JunkStatus: 45, JunkPercent: 46, JunkScoreOrigin: 47,
  HdrProperty: 49, FolderFlag: 50, Uint32HdrProperty: 51, OtherHeader: 52,
};

const ATTACHMENT_FLAG = 0x10000000; // nsMsgMessageFlags.Attachment

// nsMsgMessageFlags (nsMsgMessageFlags.idl) -- the bits the filter UI offers
// as "status", plus the attachment flag hasAttachment stores.
const MESSAGE_FLAGS = {
  Read: 0x1, Replied: 0x2, Marked: 0x4, Forwarded: 0x1000, New: 0x10000,
  Attachment: ATTACHMENT_FLAG,
};

// nsMsgPriority (MailNewsTypes2.idl).
const PRIORITY = { notSet: 0, none: 1, lowest: 2, low: 3, normal: 4, high: 5, highest: 6, Default: 4 };

// The real nsMsgSearchOp enum (nsMsgSearchCore.idl), including the
// kNumMsgSearchOperators sentinel that must NOT become an operator.
// The real nsMsgFilterAction enum (nsMsgFilterCore.idl). Note the hole at 8:
// Label existed only up to TB 102.
const ACTIONS = {
  Custom: -1, None: 0, MoveToFolder: 1, ChangePriority: 2, Delete: 3,
  MarkRead: 4, KillThread: 5, WatchThread: 6, MarkFlagged: 7, Reply: 9,
  Forward: 10, StopExecution: 11, DeleteFromPop3Server: 12,
  LeaveOnPop3Server: 13, JunkScore: 14, FetchBodyFromPop3Server: 15,
  CopyToFolder: 16, AddTag: 17, KillSubthread: 18, MarkUnread: 19,
};

const OPS = {
  Contains: 0, DoesntContain: 1, Is: 2, Isnt: 3, IsEmpty: 4,
  IsBefore: 5, IsAfter: 6, IsHigherThan: 7, IsLowerThan: 8,
  BeginsWith: 9, EndsWith: 10, SoundsLike: 11, LdapDwim: 12,
  IsGreaterThan: 13, IsLessThan: 14, NameCompletion: 15,
  IsInAB: 16, IsntInAB: 17, IsntEmpty: 18, Matches: 19, DoesntMatch: 20,
  kNumMsgSearchOperators: 21,
};

// Which nsIMsgSearchValue accessor is legal for which attribute. Mirrors
// IS_STRING_ATTRIBUTE in nsMsgSearchCore.idl and Thunderbird's
// searchWidgets.js save()/updateDisplay(); anything not listed uses .str.
const LEGAL_ACCESSOR = {
  [ATTRIB.Priority]: "priority",
  [ATTRIB.MsgStatus]: "status",
  [ATTRIB.Date]: "date",
  [ATTRIB.AgeInDays]: "age",
  [ATTRIB.Size]: "size",
  [ATTRIB.JunkStatus]: "junkStatus",
  [ATTRIB.JunkPercent]: "junkPercent",
  [ATTRIB.HasAttachmentStatus]: "status",
};

// The real extension context exposes Ci as a wrapper that answers named
// property access but reports no own keys -- Object.keys/entries come back
// empty. Model that exactly, so an enumeration-based implementation cannot
// pass these tests again.
function nonEnumerable(constants) {
  return new Proxy({}, {
    get: (_t, name) => constants[name],
    has: (_t, name) => name in constants,
    ownKeys: () => [],
    getOwnPropertyDescriptor: () => undefined,
  });
}

function makeCi(overrides = {}) {
  const attribs = { ...ATTRIB, ...(overrides.attribs || {}) };
  for (const name of overrides.removeAttribs || []) delete attribs[name];
  return {
    nsMsgSearchAttrib: nonEnumerable(attribs),
    nsMsgSearchOp: nonEnumerable({ ...OPS }),
    nsMsgFilterAction: nonEnumerable({ ...ACTIONS, ...(overrides.actions || {}) }),
    nsMsgMessageFlags: nonEnumerable({ ...MESSAGE_FLAGS }),
    nsMsgPriority: nonEnumerable({ ...PRIORITY }),
  };
}

let customHeadersPref = null; // value of mailnews.customHeaders for the sandbox

function loadFilterHelpers({ ci = makeCi() } = {}) {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN FILTER SEARCH TERM HELPERS";
  const endMarker = "// END FILTER SEARCH TERM HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "filter search term helper start marker missing");
  assert.ok(end > start, "filter search term helper end marker missing");

  const sandbox = ci === null ? {} : { Ci: ci };
  sandbox.Services = {
    prefs: { getCharPref: (_name, fallback) => customHeadersPref ?? fallback },
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(start, end)}
this.ATTRIB_MAP = ATTRIB_MAP;
this.ATTRIB_NAMES = ATTRIB_NAMES;
this.FILTER_ATTRIBUTES = FILTER_ATTRIBUTES;
this.OP_MAP = OP_MAP;
this.setSearchValue = setSearchValue;
this.getSearchValue = getSearchValue;
this.buildTerms = buildTerms;
this.buildActions = buildActions;
this.copySearchTerms = copySearchTerms;
this.copyActions = copyActions;
this.serializeSearchTerm = serializeSearchTerm;
this.serializeRuleAction = serializeRuleAction;
this.FILTER_ATTRIB_DESCRIPTION = FILTER_ATTRIB_DESCRIPTION;
this.FILTER_OP_DESCRIPTION = FILTER_OP_DESCRIPTION;
this.FILTER_VALUE_DESCRIPTION = FILTER_VALUE_DESCRIPTION;
this.FILTER_HEADER_DESCRIPTION = FILTER_HEADER_DESCRIPTION;
this.ACTION_MAP = ACTION_MAP;
this.ACTION_SPECS = ACTION_SPECS;
this.FILTER_ACTION_TYPE_DESCRIPTION = FILTER_ACTION_TYPE_DESCRIPTION;
this.FILTER_ACTION_VALUE_DESCRIPTION = FILTER_ACTION_VALUE_DESCRIPTION;`,
    sandbox
  );
  return sandbox;
}

// An nsIMsgSearchValue stand-in that enforces the union rule the real XPCOM
// object enforces: "accessing these will throw an exception if the above
// attribute does not match the type!"
function makeSearchValue() {
  const state = { attrib: undefined, stored: undefined };
  const value = {
    get attrib() { return state.attrib; },
    set attrib(v) { state.attrib = v; },
  };
  const check = (accessor) => {
    const legal = LEGAL_ACCESSOR[state.attrib] || "str";
    if (accessor !== legal) {
      throw new Error(
        `Component returned failure code: 0x80070057 (NS_ERROR_ILLEGAL_VALUE) [nsIMsgSearchValue.${accessor}]`
      );
    }
  };
  for (const accessor of ["str", "priority", "date", "status", "size", "age", "junkStatus", "junkPercent"]) {
    Object.defineProperty(value, accessor, {
      enumerable: true,
      get() { check(accessor); return state.stored; },
      set(v) { check(accessor); state.stored = v; },
    });
  }
  return value;
}

// Reads a search value through whichever accessor its attribute allows.
function storedValue(term) {
  return term.value[LEGAL_ACCESSOR[term.attrib] || "str"];
}

function makeSearchTerm() {
  return {
    attrib: undefined,
    op: undefined,
    booleanAnd: undefined,
    arbitraryHeader: "",
    hdrProperty: "",
    customId: "",
    beginsGrouping: false,
    endsGrouping: false,
    matchAll: false,
    value: makeSearchValue(),
  };
}

// An nsIMsgRuleAction stand-in with the real accessor rules (nsMsgFilter.cpp):
// targetFolderUri, priority and junkScore throw NS_ERROR_ILLEGAL_VALUE unless
// the action's type owns them; strValue and customId are untyped.
function makeRuleAction() {
  const state = { type: undefined, targetFolderUri: "", priority: undefined, junkScore: undefined };
  const action = { strValue: "", customId: "" };
  Object.defineProperty(action, "type", {
    enumerable: true,
    get() { return state.type; },
    set(v) { state.type = v; },
  });
  const typed = (member, owners, validate = () => {}) => {
    const guard = () => {
      if (!owners.includes(state.type)) {
        throw new Error(`Component returned failure code: 0x80070057 (NS_ERROR_ILLEGAL_VALUE) [nsIMsgRuleAction.${member}]`);
      }
    };
    Object.defineProperty(action, member, {
      enumerable: true,
      get() { guard(); return state[member]; },
      set(v) { guard(); validate(v); state[member] = v; },
    });
  };
  typed("targetFolderUri", [ACTIONS.MoveToFolder, ACTIONS.CopyToFolder]);
  typed("priority", [ACTIONS.ChangePriority]);
  typed("junkScore", [ACTIONS.JunkScore], (v) => {
    if (v < 0 || v > 100) throw new Error("NS_ERROR_ILLEGAL_VALUE [nsIMsgRuleAction.junkScore]");
  });
  return action;
}

function makeFilter() {
  const terms = [];
  const actions = [];
  return {
    searchTerms: terms,
    createTerm: makeSearchTerm,
    appendTerm(term) { terms.push(term); },
    createAction: makeRuleAction,
    appendAction(action) { actions.push(action); },
    get actionCount() { return actions.length; },
    getActionAt(i) { return actions[i]; },
  };
}

function buildOne(helpers, cond) {
  const filter = makeFilter();
  helpers.buildTerms(filter, [cond]);
  assert.equal(filter.searchTerms.length, 1);
  return filter.searchTerms[0];
}

function buildOneAction(helpers, act, options) {
  const filter = makeFilter();
  helpers.buildActions(filter, [act], options);
  assert.equal(filter.actionCount, 1);
  return filter.getActionAt(0);
}

const localMidnightMicros = (year, month, day) => new Date(year, month - 1, day).getTime() * 1000;

describe("ATTRIB_MAP matches the real nsMsgSearchAttrib enum", () => {
  const expected = {
    subject: ATTRIB.Subject,
    from: ATTRIB.Sender,
    body: ATTRIB.Body,
    date: ATTRIB.Date,
    priority: ATTRIB.Priority,
    status: ATTRIB.MsgStatus,
    to: ATTRIB.To,
    cc: ATTRIB.CC,
    toOrCc: ATTRIB.ToOrCC,
    allAddresses: ATTRIB.AllAddresses,
    ageInDays: ATTRIB.AgeInDays,
    size: ATTRIB.Size,
    tag: ATTRIB.Keywords,
    hasAttachment: ATTRIB.HasAttachmentStatus,
    junkStatus: ATTRIB.JunkStatus,
    junkPercent: ATTRIB.JunkPercent,
    otherHeader: ATTRIB.OtherHeader,
  };

  // The helpers live in their own vm realm, so spread the map into a plain
  // object of this realm before comparing.
  const attribMapOf = (options) => ({ ...loadFilterHelpers(options).ATTRIB_MAP });

  it("resolves every attribute from Ci.nsMsgSearchAttrib", () => {
    assert.deepEqual(attribMapOf(), expected);
  });

  it("drops attributes the running Thunderbird does not define", () => {
    const helpers = loadFilterHelpers({ ci: makeCi({ removeAttribs: ["HasAttachmentStatus"] }) });
    assert.equal(helpers.ATTRIB_MAP.hasAttachment, undefined);
    assert.ok(!helpers.FILTER_ATTRIB_DESCRIPTION.includes("hasAttachment"));
    assert.ok(!helpers.FILTER_VALUE_DESCRIPTION.includes("hasAttachment"));
    assert.throws(
      () => helpers.buildTerms(makeFilter(), [{ attrib: "hasAttachment", op: "is", value: "" }]),
      /Unknown attribute/
    );
    // Everything else is unaffected.
    assert.equal(helpers.ATTRIB_MAP.ageInDays, ATTRIB.AgeInDays);
  });

  it("refuses instead of inventing a vocabulary without XPCOM", () => {
    // There are no fallback ids on purpose: if the search interfaces are
    // missing, nsIMsgSearchTerm and the filter list are missing too, so
    // correct ids would only describe something nothing can execute.
    const helpers = loadFilterHelpers({ ci: null });
    assert.deepEqual({ ...helpers.ATTRIB_MAP }, {});
    assert.deepEqual({ ...helpers.OP_MAP }, {});
    assert.match(helpers.FILTER_ATTRIB_DESCRIPTION, /unavailable/);
    assert.match(helpers.FILTER_OP_DESCRIPTION, /unavailable/);
    // And the failure names its cause rather than blaming each attribute.
    assert.throws(
      () => helpers.buildTerms(makeFilter(), [{ attrib: "subject", op: "contains", value: "x" }]),
      /did not expose nsMsgSearchAttrib/
    );
  });

  it("says nothing about availability when Thunderbird answered", () => {
    const helpers = loadFilterHelpers();
    assert.ok(!helpers.FILTER_ATTRIB_DESCRIPTION.includes("unavailable"));
    assert.ok(!helpers.FILTER_OP_DESCRIPTION.includes("unavailable"));
  });

  it("resolves through named access only -- enumeration yields nothing", () => {
    // Regression guard for the bug this cost a real Thunderbird round to find:
    // Object.keys(Ci.nsMsgSearchAttrib) is empty in the extension context, so
    // an enumeration-based implementation produced an empty vocabulary while
    // looking like it worked.
    const ci = makeCi();
    assert.deepEqual(Object.keys(ci.nsMsgSearchAttrib), []);
    assert.equal(ci.nsMsgSearchAttrib.AgeInDays, ATTRIB.AgeInDays);
    const helpers = loadFilterHelpers({ ci });
    assert.equal(helpers.ATTRIB_MAP.ageInDays, ATTRIB.AgeInDays);
    assert.equal(Object.keys({ ...helpers.OP_MAP }).length, 21);
  });

  it("never reuses an attribute id for two names", () => {
    const { ATTRIB_MAP, ATTRIB_NAMES } = loadFilterHelpers();
    assert.equal(Object.keys(ATTRIB_NAMES).length, Object.keys(ATTRIB_MAP).length);
  });

  it("reports a UI-created AgeInDays term as ageInDays, not tag", () => {
    // The original symptom: a filter built in the Thunderbird UI read back as
    // {"attrib":"tag","op":"isGreaterThan","value":""}.
    const { ATTRIB_NAMES } = loadFilterHelpers();
    assert.equal(ATTRIB_NAMES[ATTRIB.AgeInDays], "ageInDays");
    assert.equal(ATTRIB_NAMES[ATTRIB.Keywords], "tag");
  });
});

describe("buildTerms writes the value member the attribute actually requires", () => {
  const helpers = loadFilterHelpers();

  it("stores ageInDays via .age", () => {
    const term = buildOne(helpers, { attrib: "ageInDays", op: "isGreaterThan", value: "3" });
    assert.equal(term.attrib, ATTRIB.AgeInDays);
    assert.equal(term.op, helpers.OP_MAP.isGreaterThan);
    assert.equal(term.value.age, 3);
  });

  it("stores date via .date as PRTime microseconds", () => {
    const term = buildOne(helpers, { attrib: "date", op: "isBefore", value: "2026-01-02T03:04:05.000Z" });
    assert.equal(term.value.date, Date.parse("2026-01-02T03:04:05.000Z") * 1000);
  });

  it("stores size via .size", () => {
    assert.equal(buildOne(helpers, { attrib: "size", op: "isGreaterThan", value: "1024" }).value.size, 1024);
  });

  it("stores priority via .priority and status via .status", () => {
    assert.equal(buildOne(helpers, { attrib: "priority", op: "isHigherThan", value: "4" }).value.priority, 4);
    assert.equal(buildOne(helpers, { attrib: "status", op: "is", value: "2" }).value.status, 2);
  });

  it("stores junkPercent via .junkPercent", () => {
    assert.equal(buildOne(helpers, { attrib: "junkPercent", op: "isGreaterThan", value: "90" }).value.junkPercent, 90);
  });

  it("stores junkStatus via .junkStatus and accepts names", () => {
    assert.equal(buildOne(helpers, { attrib: "junkStatus", op: "is", value: "junk" }).value.junkStatus, 2);
    assert.equal(buildOne(helpers, { attrib: "junkStatus", op: "is", value: "good" }).value.junkStatus, 1);
    assert.equal(buildOne(helpers, { attrib: "junkStatus", op: "is", value: "2" }).value.junkStatus, 2);
  });

  it("stores hasAttachment as the attachment flag in .status", () => {
    const term = buildOne(helpers, { attrib: "hasAttachment", op: "is", value: "" });
    assert.equal(term.attrib, ATTRIB.HasAttachmentStatus);
    assert.equal(term.value.status, ATTACHMENT_FLAG);
  });

  it("stores text attributes via .str", () => {
    assert.equal(buildOne(helpers, { attrib: "subject", op: "contains", value: "invoice" }).value.str, "invoice");
    assert.equal(buildOne(helpers, { attrib: "tag", op: "is", value: "$label1" }).value.str, "$label1");
    assert.equal(buildOne(helpers, { attrib: "from", op: "is", value: "" }).value.str, "");
  });

  it("defaults booleanAnd to true and honours an explicit false", () => {
    assert.equal(buildOne(helpers, { attrib: "subject", op: "contains", value: "x" }).booleanAnd, true);
    assert.equal(
      buildOne(helpers, { attrib: "subject", op: "contains", value: "x", booleanAnd: false }).booleanAnd,
      false
    );
  });

  it("rejects non-numeric values for numeric attributes", () => {
    assert.throws(
      () => buildOne(helpers, { attrib: "ageInDays", op: "isGreaterThan", value: "soon" }),
      /Condition value for "ageInDays" must be a non-negative integer/
    );
    assert.throws(
      () => buildOne(helpers, { attrib: "date", op: "isBefore", value: "not-a-date" }),
      /must be YYYY-MM-DD \(a local calendar day\) or an ISO-8601 date-time/
    );
  });

  it("keeps rejecting unknown attributes and operators", () => {
    assert.throws(() => buildOne(helpers, { attrib: "44", op: "is", value: "x" }), /Unknown attribute/);
    assert.throws(() => buildOne(helpers, { attrib: "subject", op: "pwn", value: "x" }), /Unknown operator/);
  });

  it("explains why a custom term cannot be created", () => {
    assert.throws(
      () => buildOne(helpers, { attrib: "custom", op: "is", value: "x" }),
      /custom search term needs a customId/
    );
  });
});

describe("date conditions", () => {
  const helpers = loadFilterHelpers();
  const dateOf = (value) => buildOne(helpers, { attrib: "date", op: "isBefore", value }).value.date;

  it("reads a date-only value as a local calendar day, not UTC midnight", () => {
    // Thunderbird stores and shows filter dates in local time. Date.parse
    // would read "2026-01-01" as UTC midnight, which is 31-Dec-2025 anywhere
    // west of UTC -- the off-by-one @ncrosty58 reported on #175.
    const stored = dateOf("2026-01-01");
    assert.equal(stored, localMidnightMicros(2026, 1, 1));
    const local = new Date(stored / 1000);
    assert.deepEqual(
      [local.getFullYear(), local.getMonth() + 1, local.getDate(), local.getHours()],
      [2026, 1, 1, 0]
    );
  });

  it("keeps the instant of a zoned date-time and reads an unzoned one as local", () => {
    assert.equal(dateOf("2026-01-02T03:04:05Z"), Date.parse("2026-01-02T03:04:05Z") * 1000);
    assert.equal(dateOf("2026-01-02T03:04:05+02:00"), Date.parse("2026-01-02T03:04:05+02:00") * 1000);
    assert.equal(dateOf("2026-01-02T03:04"), new Date(2026, 0, 2, 3, 4).getTime() * 1000);
  });

  it("rejects bare numbers instead of taking them as epoch milliseconds", () => {
    // "2026" used to match the epoch-ms branch and was saved as 01-Jan-1970.
    for (const value of ["2026", "1767322800000", "2026-01", "20260101"]) {
      assert.throws(() => dateOf(value), /must be YYYY-MM-DD/, `accepted ${value}`);
    }
  });

  it("rejects a calendar day that does not exist", () => {
    assert.throws(() => dateOf("2026-02-30"), /must be YYYY-MM-DD/);
    assert.throws(() => dateOf("2026-13-01"), /must be YYYY-MM-DD/);
  });

  it("reads a local-midnight date back as the day it was written", () => {
    const term = buildOne(helpers, { attrib: "date", op: "isBefore", value: "2026-01-01" });
    assert.equal(helpers.getSearchValue(term.value, term.attrib), "2026-01-01");
  });

  it("reads a date with a time of day back as an ISO-8601 instant", () => {
    const term = buildOne(helpers, { attrib: "date", op: "isBefore", value: "2026-01-02T03:04:05.000Z" });
    assert.equal(helpers.getSearchValue(term.value, term.attrib), "2026-01-02T03:04:05.000Z");
  });
});

describe("condition values are validated strictly", () => {
  const helpers = loadFilterHelpers();
  const reject = (cond, pattern) => assert.throws(() => buildOne(helpers, cond), pattern, JSON.stringify(cond));

  it("does not let parseInt truncate", () => {
    // parseInt("30abc") is 30 and parseInt("1.5") is 1.
    reject({ attrib: "ageInDays", op: "isGreaterThan", value: "30abc" }, /must be a non-negative integer/);
    reject({ attrib: "ageInDays", op: "isGreaterThan", value: "1.5" }, /must be a non-negative integer/);
    reject({ attrib: "size", op: "isGreaterThan", value: "1e3" }, /must be a non-negative integer/);
    assert.equal(buildOne(helpers, { attrib: "ageInDays", op: "isGreaterThan", value: " 30 " }).value.age, 30);
  });

  it("rejects negative sizes and ages", () => {
    // nsIMsgSearchValue.size is unsigned: -5 was stored as 4294967291.
    reject({ attrib: "size", op: "isGreaterThan", value: "-5" }, /must be a non-negative integer \(KB\)/);
    reject({ attrib: "ageInDays", op: "isGreaterThan", value: "-1" }, /must be a non-negative integer \(days\)/);
    assert.equal(buildOne(helpers, { attrib: "size", op: "isGreaterThan", value: "0" }).value.size, 0);
  });

  it("bounds junkPercent to 0..100", () => {
    reject({ attrib: "junkPercent", op: "isGreaterThan", value: "101" }, /from 0 to 100/);
    assert.equal(buildOne(helpers, { attrib: "junkPercent", op: "isGreaterThan", value: "100" }).value.junkPercent, 100);
  });

  it("accepts only the nsMsgPriority levels for priority", () => {
    reject({ attrib: "priority", op: "isHigherThan", value: "1" }, /from 2 to 6/);
    reject({ attrib: "priority", op: "isHigherThan", value: "7" }, /from 2 to 6/);
    reject({ attrib: "priority", op: "isHigherThan", value: "High" }, /2=lowest, 3=low, 4=normal, 5=high, 6=highest/);
    assert.equal(buildOne(helpers, { attrib: "priority", op: "isHigherThan", value: "6" }).value.priority, 6);
  });

  it("requires a non-zero flag bitmask for status", () => {
    reject({ attrib: "status", op: "is", value: "0" }, /message-flag bitmask/);
    reject({ attrib: "status", op: "is", value: "replied" }, /1=read, 2=replied, 4=flagged, 4096=forwarded, 65536=new/);
    assert.equal(buildOne(helpers, { attrib: "status", op: "is", value: "65536" }).value.status, 0x10000);
  });

  it("bounds junkStatus to the three nsMsgJunkStatus values", () => {
    reject({ attrib: "junkStatus", op: "is", value: "3" }, /junk, good or unclassified/);
    reject({ attrib: "junkStatus", op: "is", value: "spam" }, /junk, good or unclassified/);
  });

  it("refuses a value for hasAttachment -- the operator carries the meaning", () => {
    // Thunderbird ignores the value: is + "false" was persisted as is,true.
    reject({ attrib: "hasAttachment", op: "is", value: "false" }, /hasAttachment takes no value/);
    reject({ attrib: "hasAttachment", op: "is", value: "true" }, /hasAttachment takes no value/);
    assert.equal(buildOne(helpers, { attrib: "hasAttachment", op: "isnt", value: "" }).value.status, ATTACHMENT_FLAG);
    assert.equal(buildOne(helpers, { attrib: "hasAttachment", op: "isnt" }).value.status, ATTACHMENT_FLAG);
  });
});

describe("otherHeader requires its header name", () => {
  const helpers = loadFilterHelpers();

  it("uses OtherHeader+1, never OtherHeader itself", () => {
    // Thunderbird treats OtherHeader(52) as the UI "Customize..." placeholder
    // and serialises a term left at it with an EMPTY attribute name, which
    // silently breaks the filter on reload. Real header terms start at 53.
    customHeadersPref = null;
    const term = buildOne(helpers, {
      attrib: "otherHeader", op: "contains", value: "bulk", header: "X-Mailer",
    });
    assert.equal(term.attrib, ATTRIB.OtherHeader + 1);
    assert.equal(term.arbitraryHeader, "X-Mailer");
    assert.equal(term.value.str, "bulk");
  });

  it("offsets by the header's index in mailnews.customHeaders", () => {
    customHeadersPref = "X-Spam-Flag:X-Mailer:X-Priority";
    try {
      const term = buildOne(helpers, {
        attrib: "otherHeader", op: "contains", value: "bulk", header: "x-mailer",
      });
      assert.equal(term.attrib, ATTRIB.OtherHeader + 1 + 1);
    } finally {
      customHeadersPref = null;
    }
  });

  it("rejects a malformed header name", () => {
    assert.throws(
      () => buildOne(helpers, { attrib: "otherHeader", op: "contains", value: "x", header: "bad header" }),
      /Invalid header name/
    );
  });

  it("reads an arbitrary-header term back as otherHeader", () => {
    const spec = helpers.ATTRIB_NAMES[ATTRIB.OtherHeader + 3];
    assert.equal(spec, undefined, "53+ is deliberately not in ATTRIB_NAMES");
    // getSearchValue must still treat it as a text attribute.
    assert.equal(
      helpers.getSearchValue({ str: "bulk" }, ATTRIB.OtherHeader + 3),
      "bulk"
    );
  });

  it("rejects otherHeader without a header name", () => {
    assert.throws(
      () => buildOne(helpers, { attrib: "otherHeader", op: "contains", value: "bulk" }),
      /requires a "header" name/
    );
  });

  it("rejects a header name on any other attribute", () => {
    assert.throws(
      () => buildOne(helpers, { attrib: "subject", op: "contains", value: "x", header: "X-Mailer" }),
      /not valid for attrib "subject"/
    );
  });
});

describe("getSearchValue reads back what buildTerms wrote", () => {
  const helpers = loadFilterHelpers();
  const roundTrip = (cond) => {
    const term = buildOne(helpers, cond);
    return helpers.getSearchValue(term.value, term.attrib);
  };

  it("round-trips every typed attribute", () => {
    assert.equal(roundTrip({ attrib: "ageInDays", op: "isGreaterThan", value: "3" }), "3");
    assert.equal(roundTrip({ attrib: "size", op: "isGreaterThan", value: "1024" }), "1024");
    assert.equal(roundTrip({ attrib: "priority", op: "isHigherThan", value: "4" }), "4");
    assert.equal(roundTrip({ attrib: "junkPercent", op: "isGreaterThan", value: "90" }), "90");
    assert.equal(roundTrip({ attrib: "junkStatus", op: "is", value: "junk" }), "junk");
    assert.equal(roundTrip({ attrib: "date", op: "isBefore", value: "2026-01-01" }), "2026-01-01");
    assert.equal(
      roundTrip({ attrib: "date", op: "isBefore", value: "2026-01-02T03:04:05.000Z" }),
      "2026-01-02T03:04:05.000Z"
    );
    assert.equal(roundTrip({ attrib: "subject", op: "contains", value: "invoice" }), "invoice");
    assert.equal(roundTrip({ attrib: "tag", op: "is", value: "$label1" }), "$label1");
  });

  it("reports no value for hasAttachment -- the operator carries the meaning", () => {
    assert.equal(roundTrip({ attrib: "hasAttachment", op: "is", value: "" }), "");
  });

  it("degrades to an empty string instead of throwing on an unreadable value", () => {
    const hostile = { get str() { throw new Error("NS_ERROR_ILLEGAL_VALUE"); } };
    assert.equal(helpers.getSearchValue(hostile, ATTRIB.Subject), "");
  });
});

describe("the tool schema text is generated from the attribute table", () => {
  const helpers = loadFilterHelpers();

  it("lists every attribute the tools actually accept", () => {
    const names = Object.keys({ ...helpers.ATTRIB_MAP });
    for (const name of names) {
      assert.ok(
        helpers.FILTER_ATTRIB_DESCRIPTION.includes(name),
        `attrib description does not mention ${name}`
      );
    }
    // ...and nothing beyond them.
    const listed = helpers.FILTER_ATTRIB_DESCRIPTION.split(": ")[1].split(", ");
    assert.deepEqual(listed.slice().sort(), names.slice().sort());
  });

  it("documents a value format for every attribute", () => {
    for (const name of Object.keys({ ...helpers.ATTRIB_MAP })) {
      assert.ok(
        new RegExp(`\\b${name.replace("$", "\\$")}\\b`).test(helpers.FILTER_VALUE_DESCRIPTION),
        `value description does not mention ${name}`
      );
    }
  });

  it("states units and value meanings, resolved from Thunderbird's own enums", () => {
    // A caller could not tell that Normal is 4, replied is 2, or that size
    // is in kilobytes; the hints now say so, with the numbers taken from
    // Ci.nsMsgPriority / Ci.nsMsgMessageFlags rather than typed in.
    const d = helpers.FILTER_VALUE_DESCRIPTION;
    assert.match(d, /size: a non-negative integer \(KB\)/);
    assert.match(d, /ageInDays: a non-negative integer \(days\)/);
    assert.match(d, /priority: an integer from 2 to 6 \(2=lowest, 3=low, 4=normal, 5=high, 6=highest\)/);
    assert.match(d, /status: a message-flag bitmask \(1=read, 2=replied, 4=flagged, 4096=forwarded, 65536=new\)/);
    assert.match(d, /date: YYYY-MM-DD \(a local calendar day\) or an ISO-8601 date-time/);
    assert.match(d, /junkPercent: an integer from 0 to 100/);
    assert.match(d, /hasAttachment: no value/);
  });

  it("groups attributes that share a value format", () => {
    assert.match(helpers.FILTER_VALUE_DESCRIPTION, /subject\/from\/body\/to\/cc\/toOrCc\/allAddresses\/otherHeader: text/);
  });

  it("names otherHeader as the attribute that requires a header", () => {
    assert.match(helpers.FILTER_HEADER_DESCRIPTION, /Required when attrib is otherHeader/);
  });
});

describe("OP_MAP is resolved from the live nsMsgSearchOp interface", () => {
  const helpers = loadFilterHelpers();

  it("lowers the first letter of every IDL constant name", () => {
    assert.deepEqual({ ...helpers.OP_MAP }, {
      contains: 0, doesntContain: 1, is: 2, isnt: 3, isEmpty: 4,
      isBefore: 5, isAfter: 6, isHigherThan: 7, isLowerThan: 8,
      beginsWith: 9, endsWith: 10, soundsLike: 11, ldapDwim: 12,
      isGreaterThan: 13, isLessThan: 14, nameCompletion: 15,
      isInAB: 16, isntInAB: 17, isntEmpty: 18, matches: 19, doesntMatch: 20,
    });
  });

  it("never exposes a sentinel constant as an operator", () => {
    for (const name of Object.keys({ ...helpers.OP_MAP })) {
      assert.ok(!/^kNum/.test(name), `sentinel leaked into OP_MAP: ${name}`);
    }
  });

  it("describes exactly the operators the tools accept", () => {
    const listed = helpers.FILTER_OP_DESCRIPTION.split(": ")[1].split(", ");
    assert.deepEqual(listed.slice().sort(), Object.keys({ ...helpers.OP_MAP }).sort());
  });
});

describe("version compatibility", () => {
  const helpers = loadFilterHelpers();

  it("reports a clear error when the union member does not exist", () => {
    // TB 115 removed the "label" member from nsIMsgSearchValue; if that ever
    // happens to a member we write, the error must name member and attribute
    // instead of surfacing an opaque XPCOM failure.
    const value = { attrib: undefined, str: "" }; // no .age member
    assert.throws(
      () => helpers.setSearchValue(value, ATTRIB.AgeInDays, "3"),
      /no "age" member \(needed for attribute "ageInDays"\)/
    );
  });
});

describe("ACTION_MAP matches the real nsMsgFilterAction enum", () => {
  const helpers = loadFilterHelpers();

  it("maps every action to the id Thunderbird actually uses", () => {
    // The old table invented an ordering and numbered it 1..21, so only
    // moveToFolder and addTag were right. Every row below is a value that
    // used to point at a different action entirely.
    assert.deepEqual({ ...helpers.ACTION_MAP }, {
      moveToFolder: ACTIONS.MoveToFolder,
      copyToFolder: ACTIONS.CopyToFolder,
      changePriority: ACTIONS.ChangePriority,
      junkScore: ACTIONS.JunkScore,
      addTag: ACTIONS.AddTag,
      reply: ACTIONS.Reply,
      forward: ACTIONS.Forward,
      delete: ACTIONS.Delete,
      markRead: ACTIONS.MarkRead,
      markUnread: ACTIONS.MarkUnread,
      markFlagged: ACTIONS.MarkFlagged,
      killThread: ACTIONS.KillThread,
      killSubthread: ACTIONS.KillSubthread,
      watchThread: ACTIONS.WatchThread,
      stopExecution: ACTIONS.StopExecution,
      deleteFromServer: ACTIONS.DeleteFromPop3Server,
      leaveOnServer: ACTIONS.LeaveOnPop3Server,
      fetchBody: ACTIONS.FetchBodyFromPop3Server,
      // label is absent: removed from Thunderbird in 115.
    });
  });

  it("does not confuse markRead with killThread", () => {
    // The concrete symptom found on Thunderbird 153: a filter asked to mark
    // read was persisted as action="Ignore thread".
    assert.notEqual(helpers.ACTION_MAP.markRead, ACTIONS.KillThread);
    assert.equal(helpers.ACTION_MAP.markRead, ACTIONS.MarkRead);
    assert.equal(helpers.ACTION_SPECS[ACTIONS.KillThread].action, "killThread");
  });

  it("offers label only where Thunderbird still has it", () => {
    const withLabel = loadFilterHelpers({ ci: makeCi({ actions: { Label: 8 } }) });
    assert.equal(withLabel.ACTION_MAP.label, 8);
    assert.equal(helpers.ACTION_MAP.label, undefined);
    assert.ok(!helpers.FILTER_ACTION_TYPE_DESCRIPTION.includes("label"));
  });

  it("drops the invented action names", () => {
    // deleteBody never existed in any nsMsgFilterAction; its old value 0x12
    // was KillSubthread, which is now exposed under its real name.
    assert.equal(helpers.ACTION_MAP.deleteBody, undefined);
    assert.equal(helpers.ACTION_MAP.killSubthread, ACTIONS.KillSubthread);
    // custom needs a customId we do not expose, and its old value 0x15 was
    // not the real Custom(-1) either.
    assert.equal(helpers.ACTION_MAP.custom, undefined);
  });

  it("refuses instead of inventing action ids without XPCOM", () => {
    const bare = loadFilterHelpers({ ci: null });
    assert.deepEqual({ ...bare.ACTION_MAP }, {});
    assert.match(bare.FILTER_ACTION_TYPE_DESCRIPTION, /unavailable/);
  });

  it("describes exactly the actions the tools accept", () => {
    const listed = helpers.FILTER_ACTION_TYPE_DESCRIPTION.split(": ")[1].split(", ");
    assert.deepEqual(listed.slice().sort(), Object.keys({ ...helpers.ACTION_MAP }).sort());
  });

  it("documents which actions take a value, which do not, and what the values mean", () => {
    const d = helpers.FILTER_ACTION_VALUE_DESCRIPTION;
    assert.match(d, /required for every action that takes one/);
    assert.match(d, /moveToFolder\/copyToFolder: a folder URI/);
    assert.match(d, /changePriority: an integer from 2 to 6 \(2=lowest, 3=low, 4=normal, 5=high, 6=highest\)/);
    assert.match(d, /junkScore: an integer from 0 \(not junk\) to 100 \(junk\)/);
    assert.match(d, /delete\/markRead\/markUnread\/markFlagged\/killThread\/killSubthread\/watchThread\/stopExecution\/deleteFromServer\/leaveOnServer\/fetchBody: no value/);
  });
});

describe("buildActions writes the member the action type owns", () => {
  const helpers = loadFilterHelpers();
  const reject = (act, pattern, options) =>
    assert.throws(() => buildOneAction(helpers, act, options), pattern, JSON.stringify(act));

  it("writes typed members through the table", () => {
    const move = buildOneAction(helpers, { type: "moveToFolder", value: "imap://a/Inbox/x" });
    assert.equal(move.type, ACTIONS.MoveToFolder);
    assert.equal(move.targetFolderUri, "imap://a/Inbox/x");
    assert.equal(buildOneAction(helpers, { type: "changePriority", value: "6" }).priority, 6);
    assert.equal(buildOneAction(helpers, { type: "junkScore", value: "100" }).junkScore, 100);
    assert.equal(buildOneAction(helpers, { type: "addTag", value: "$label1" }).strValue, "$label1");
    assert.equal(buildOneAction(helpers, { type: "forward", value: "a@example.com" }).strValue, "a@example.com");
  });

  it("appends valueless actions without touching any member", () => {
    const action = buildOneAction(helpers, { type: "markRead" });
    assert.equal(action.type, ACTIONS.MarkRead);
    assert.equal(action.strValue, "");
  });

  it("requires a value for every action that takes one", () => {
    // Thunderbird saves "Move to folder" with no folder and the filter then
    // silently does nothing.
    reject({ type: "moveToFolder" }, /Action "moveToFolder" requires a value: a folder URI/);
    reject({ type: "moveToFolder", value: "  " }, /requires a value/);
    reject({ type: "forward" }, /Action "forward" requires a value: an email address/);
    reject({ type: "changePriority" }, /requires a value/);
  });

  it("rejects a value on an action that takes none", () => {
    reject({ type: "markRead", value: "yes" }, /Action "markRead" does not take a value/);
  });

  it("validates action values as action values, with the same strictness as conditions", () => {
    // The message used to say "Condition value" for an action.
    reject({ type: "changePriority", value: "99" }, /^Error: Action value for "changePriority" must be an integer from 2 to 6/);
    reject({ type: "changePriority", value: "High" }, /Action value for "changePriority"/);
    reject({ type: "junkScore", value: "beaucoup" }, /Action value for "junkScore" must be an integer from 0 \(not junk\) to 100 \(junk\)/);
    reject({ type: "junkScore", value: "101" }, /Action value for "junkScore"/);
    reject({ type: "junkScore", value: "5.5" }, /Action value for "junkScore"/);
  });

  it("lets the caller refuse a move/copy target folder", () => {
    const seen = [];
    const checkTargetFolder = (uri) => { seen.push(uri); return uri.includes("secret") ? { error: "no" } : { folder: {} }; };
    reject({ type: "copyToFolder", value: "imap://a/secret" }, /Filter target folder not accessible: imap:\/\/a\/secret/, { checkTargetFolder });
    const ok = buildOneAction(helpers, { type: "copyToFolder", value: "imap://a/ok" }, { checkTargetFolder });
    assert.equal(ok.targetFolderUri, "imap://a/ok");
    assert.deepEqual(seen, ["imap://a/secret", "imap://a/ok"]);
  });

  it("refuses unknown and custom actions with a reason", () => {
    reject({ type: "deleteBody" }, /Unknown action type: deleteBody/);
    reject({ type: "label", value: "1" }, /Unknown action type: label/);
    reject({ type: "custom", value: "x" }, /custom action needs a customId/);
  });
});

describe("reading filters back", () => {
  const helpers = loadFilterHelpers();

  it("reports a custom search term by name with its customId", () => {
    // Thunderbird persists a custom term as "<customId>,<op>,<value>"; the
    // customId is the only thing that identifies it. It used to read back
    // as attrib "-2" with nothing else.
    const term = makeSearchTerm();
    term.attrib = ATTRIB.Custom;
    term.op = OPS.Is;
    term.booleanAnd = false;
    term.customId = "quickfilter@example.org#hasSticker";
    term.value.attrib = ATTRIB.Custom;
    term.value.str = "yes";
    assert.deepEqual({ ...helpers.serializeSearchTerm(term) }, {
      attrib: "custom",
      op: "is",
      booleanAnd: false,
      value: "yes",
      customId: "quickfilter@example.org#hasSticker",
    });
  });

  it("reports hdrProperty on terms that read a header property", () => {
    const term = makeSearchTerm();
    term.attrib = ATTRIB.HdrProperty;
    term.op = OPS.Contains;
    term.booleanAnd = true;
    term.hdrProperty = "x-custom";
    term.value.attrib = ATTRIB.HdrProperty;
    term.value.str = "v";
    const out = helpers.serializeSearchTerm(term);
    assert.equal(out.hdrProperty, "x-custom");
    assert.equal(out.value, "v");
  });

  it("omits customId and hdrProperty on ordinary terms", () => {
    const term = buildOne(helpers, { attrib: "subject", op: "contains", value: "x" });
    assert.deepEqual({ ...helpers.serializeSearchTerm(term) }, {
      attrib: "subject", op: "contains", booleanAnd: true, value: "x",
    });
  });

  it("reports a Custom action by name with its customId and value", () => {
    const action = makeRuleAction();
    action.type = ACTIONS.Custom;
    action.customId = "addon@example.org#archive";
    action.strValue = "2026";
    assert.deepEqual({ ...helpers.serializeRuleAction(action) }, {
      type: "custom", value: "2026", customId: "addon@example.org#archive",
    });
  });

  it("reports typed action values and no value for valueless actions", () => {
    assert.deepEqual(
      { ...helpers.serializeRuleAction(buildOneAction(helpers, { type: "changePriority", value: "6" })) },
      { type: "changePriority", value: "6" }
    );
    assert.deepEqual(
      { ...helpers.serializeRuleAction(buildOneAction(helpers, { type: "moveToFolder", value: "imap://a/b" })) },
      { type: "moveToFolder", value: "imap://a/b" }
    );
    assert.deepEqual(
      { ...helpers.serializeRuleAction(buildOneAction(helpers, { type: "markRead" })) },
      { type: "markRead" }
    );
  });
});

describe("copySearchTerms keeps every condition exactly", () => {
  const helpers = loadFilterHelpers();

  // The lab's 13-condition rule: every value type Thunderbird's filter UI can
  // produce, plus the kinds this API cannot create but must not damage.
  function makeSourceFilter() {
    const source = makeFilter();
    helpers.buildTerms(source, [
      { attrib: "subject", op: "contains", value: "invoice", booleanAnd: false },
      { attrib: "from", op: "contains", value: "boss@example.com", booleanAnd: false },
      { attrib: "date", op: "isBefore", value: "2026-01-01", booleanAnd: false },
      { attrib: "priority", op: "isHigherThan", value: "4", booleanAnd: false },
      { attrib: "status", op: "is", value: "2", booleanAnd: false },
      { attrib: "tag", op: "contains", value: "$label1", booleanAnd: false },
      { attrib: "otherHeader", op: "contains", value: "bulk", header: "x-mailer", booleanAnd: false },
      { attrib: "ageInDays", op: "isGreaterThan", value: "30", booleanAnd: false },
      { attrib: "size", op: "isGreaterThan", value: "1024", booleanAnd: false },
      { attrib: "hasAttachment", op: "is", value: "", booleanAnd: false },
      { attrib: "junkStatus", op: "is", value: "2", booleanAnd: false },
      { attrib: "junkPercent", op: "isGreaterThan", value: "90", booleanAnd: false },
    ]);
    // Terms the UI can make but this API does not model.
    const origin = makeSearchTerm();
    origin.attrib = ATTRIB.JunkScoreOrigin;
    origin.op = OPS.Is;
    origin.booleanAnd = false;
    origin.value.attrib = ATTRIB.JunkScoreOrigin;
    origin.value.str = "plugin";
    source.appendTerm(origin);
    const custom = makeSearchTerm();
    custom.attrib = ATTRIB.Custom;
    custom.op = OPS.Is;
    custom.booleanAnd = false;
    custom.customId = "quickfilter@example.org#hasSticker";
    custom.beginsGrouping = true;
    custom.endsGrouping = true;
    custom.value.attrib = ATTRIB.Custom;
    custom.value.str = "yes";
    source.appendTerm(custom);
    return source;
  }

  it("copies each value through the member its attribute owns", () => {
    // The old copy read .str for everything and .date for attrib 3, then
    // swallowed the NS_ERROR_ILLEGAL_VALUE that every other typed attribute
    // throws -- so priority, status, age, size, junkStatus and junkPercent
    // were silently reset to 0 on every updateFilter.
    const source = makeSourceFilter();
    const target = makeFilter();
    const copied = helpers.copySearchTerms(source, target);
    assert.equal(copied, source.searchTerms.length);
    assert.equal(target.searchTerms.length, source.searchTerms.length);
    source.searchTerms.forEach((from, i) => {
      const to = target.searchTerms[i];
      assert.equal(to.attrib, from.attrib, `attrib of term ${i}`);
      assert.equal(to.op, from.op, `op of term ${i}`);
      assert.equal(to.booleanAnd, from.booleanAnd, `booleanAnd of term ${i}`);
      assert.equal(to.value.attrib, from.attrib, `value.attrib of term ${i}`);
      assert.equal(storedValue(to), storedValue(from), `value of term ${i}`);
    });
    assert.equal(target.searchTerms[7].value.age, 30);
    assert.equal(target.searchTerms[8].value.size, 1024);
    assert.equal(target.searchTerms[2].value.date, localMidnightMicros(2026, 1, 1));
  });

  it("carries arbitraryHeader, customId and grouping over", () => {
    const source = makeSourceFilter();
    const target = makeFilter();
    helpers.copySearchTerms(source, target);
    const header = target.searchTerms[6];
    assert.equal(header.attrib, ATTRIB.OtherHeader + 1);
    assert.equal(header.arbitraryHeader, "x-mailer");
    const custom = target.searchTerms[13];
    assert.equal(custom.attrib, ATTRIB.Custom);
    assert.equal(custom.customId, "quickfilter@example.org#hasSticker");
    assert.equal(custom.beginsGrouping, true);
    assert.equal(custom.endsGrouping, true);
    assert.equal(custom.value.str, "yes");
  });

  it("carries hdrProperty and matchAll over", () => {
    const source = makeFilter();
    const term = makeSearchTerm();
    term.attrib = ATTRIB.HdrProperty;
    term.op = OPS.Contains;
    term.booleanAnd = true;
    term.hdrProperty = "x-custom";
    term.value.attrib = ATTRIB.HdrProperty;
    term.value.str = "v";
    source.appendTerm(term);
    const all = makeSearchTerm();
    all.attrib = ATTRIB.Subject;
    all.op = OPS.Contains;
    all.booleanAnd = true;
    all.matchAll = true;
    all.value.attrib = ATTRIB.Subject;
    all.value.str = "";
    source.appendTerm(all);
    const target = makeFilter();
    helpers.copySearchTerms(source, target);
    assert.equal(target.searchTerms[0].hdrProperty, "x-custom");
    assert.equal(target.searchTerms[1].matchAll, true);
  });

  it("propagates a failure instead of leaving the default value in place", () => {
    const source = makeFilter();
    helpers.buildTerms(source, [{ attrib: "ageInDays", op: "isGreaterThan", value: "30" }]);
    source.searchTerms[0].value = { attrib: ATTRIB.AgeInDays, get age() { throw new Error("NS_ERROR_FAILURE"); } };
    assert.throws(() => helpers.copySearchTerms(source, makeFilter()), /NS_ERROR_FAILURE/);
  });
});

describe("copyActions keeps every action exactly", () => {
  const helpers = loadFilterHelpers();

  function makeSourceFilter() {
    const source = makeFilter();
    helpers.buildActions(source, [
      { type: "moveToFolder", value: "imap://a/Inbox/x" },
      { type: "changePriority", value: "6" },
      { type: "junkScore", value: "100" },
      { type: "addTag", value: "$label1" },
      { type: "markRead" },
    ]);
    const custom = makeRuleAction();
    custom.type = ACTIONS.Custom;
    custom.customId = "addon@example.org#archive";
    custom.strValue = "2026";
    source.appendAction(custom);
    return source;
  }

  it("copies the member each type owns, plus customId", () => {
    const source = makeSourceFilter();
    const target = makeFilter();
    assert.equal(helpers.copyActions(source, target), 6);
    const types = [];
    for (let i = 0; i < target.actionCount; i++) types.push(target.getActionAt(i).type);
    assert.deepEqual(types, [
      ACTIONS.MoveToFolder, ACTIONS.ChangePriority, ACTIONS.JunkScore,
      ACTIONS.AddTag, ACTIONS.MarkRead, ACTIONS.Custom,
    ]);
    assert.equal(target.getActionAt(0).targetFolderUri, "imap://a/Inbox/x");
    assert.equal(target.getActionAt(1).priority, 6);
    assert.equal(target.getActionAt(2).junkScore, 100);
    assert.equal(target.getActionAt(3).strValue, "$label1");
    assert.equal(target.getActionAt(4).strValue, "");
    // A Custom action loses its identity without customId: it used to read
    // back as -1 and be written as action="Custom" with no id.
    assert.equal(target.getActionAt(5).customId, "addon@example.org#archive");
    assert.equal(target.getActionAt(5).strValue, "2026");
  });

  it("propagates a failure instead of skipping the action", () => {
    const source = makeSourceFilter();
    source.getActionAt = () => { throw new Error("NS_ERROR_FAILURE"); };
    assert.throws(() => helpers.copyActions(source, makeFilter()), /NS_ERROR_FAILURE/);
  });
});
