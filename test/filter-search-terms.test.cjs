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
// Thunderbird's searchWidgets.js save()/updateDisplay(); anything not listed
// uses .str.
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
    nsMsgMessageFlags: { Attachment: ATTACHMENT_FLAG },
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

function makeFilter() {
  const terms = [];
  return {
    searchTerms: terms,
    createTerm() {
      return {
        attrib: undefined,
        op: undefined,
        booleanAnd: undefined,
        arbitraryHeader: "",
        value: makeSearchValue(),
      };
    },
    appendTerm(term) { terms.push(term); },
  };
}

function buildOne(helpers, cond) {
  const filter = makeFilter();
  helpers.buildTerms(filter, [cond]);
  assert.equal(filter.searchTerms.length, 1);
  return filter.searchTerms[0];
}

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

  it("accepts epoch milliseconds for date", () => {
    const term = buildOne(helpers, { attrib: "date", op: "isAfter", value: "1767322800000" });
    assert.equal(term.value.date, 1767322800000 * 1000);
  });

  it("stores size via .size", () => {
    assert.equal(buildOne(helpers, { attrib: "size", op: "isGreaterThan", value: "1024" }).value.size, 1024);
  });

  it("stores priority via .priority and status via .status", () => {
    assert.equal(buildOne(helpers, { attrib: "priority", op: "isHigherThan", value: "4" }).value.priority, 4);
    assert.equal(buildOne(helpers, { attrib: "status", op: "is", value: "1" }).value.status, 1);
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
      /must be an integer/
    );
    assert.throws(
      () => buildOne(helpers, { attrib: "date", op: "isBefore", value: "not-a-date" }),
      /must be an ISO-8601 date or epoch ms/
    );
  });

  it("keeps rejecting unknown attributes and operators", () => {
    assert.throws(() => buildOne(helpers, { attrib: "44", op: "is", value: "x" }), /Unknown attribute/);
    assert.throws(() => buildOne(helpers, { attrib: "subject", op: "pwn", value: "x" }), /Unknown operator/);
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

  it("groups attributes that share a value format", () => {
    // ageInDays/size/priority/status/junkPercent are all plain integers and
    // must therefore appear as one group, not five.
    assert.match(helpers.FILTER_VALUE_DESCRIPTION, /priority\/status\/ageInDays\/size\/junkPercent: an integer/);
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

  it("documents which actions take a value and which do not", () => {
    const d = helpers.FILTER_ACTION_VALUE_DESCRIPTION;
    assert.match(d, /moveToFolder\/copyToFolder: a folder URI/);
    assert.match(d, /changePriority\/junkScore: an integer/);
    assert.match(d, /no value/);
    assert.ok(/markRead/.test(d.split("no value")[0].split(";").pop() + "no value"));
  });
});
