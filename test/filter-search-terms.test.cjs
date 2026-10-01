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

const LEGACY_ATTRIB = { Label: 48 };
const LEGACY_ACTIONS = { Label: 8 };
const FILTER_TYPES = {
  None: 0, InboxRule: 1, InboxJavaScript: 2, Inbox: 3,
  NewsRule: 4, NewsJavaScript: 8, News: 12, Incoming: 15,
  Manual: 16, PostPlugin: 32, PostOutgoing: 64, Archive: 128, Periodic: 256,
  All: 31,
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
  [ATTRIB.FolderFlag]: "status",
  [ATTRIB.Uint32HdrProperty]: "status",
  [LEGACY_ATTRIB.Label]: "label",
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
    nsMsgFilterType: nonEnumerable({ ...FILTER_TYPES, ...(overrides.filterTypes || {}) }),
    nsMsgMessageFlags: nonEnumerable({ ...MESSAGE_FLAGS }),
    nsMsgPriority: nonEnumerable({ ...PRIORITY }),
  };
}

let customHeadersPref = null; // value of mailnews.customHeaders for the sandbox

function loadFilterHelpers({ ci = makeCi(), prefs = {}, globals = {}, handlers = false, preferences = false } = {}) {
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
    prefs: {
      getCharPref: (_name, fallback) => customHeadersPref ?? fallback,
      getBoolPref: (_name, fallback) => fallback,
      ...prefs,
    },
  };
  Object.assign(sandbox, globals);
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
  if (handlers) {
    const handlerStart = source.indexOf("// BEGIN FILTER TOOL HANDLERS");
    const handlerEnd = source.indexOf("// END FILTER TOOL HANDLERS", handlerStart);
    assert.ok(handlerStart >= 0, "filter tool handler start marker missing");
    assert.ok(handlerEnd > handlerStart, "filter tool handler end marker missing");
    vm.runInContext(`${source.slice(handlerStart, handlerEnd)}
this.listFilters = listFilters;
this.createFilter = createFilter;
this.updateFilter = updateFilter;
this.deleteFilter = deleteFilter;
this.reorderFilters = reorderFilters;
this.applyFilters = applyFilters;`, sandbox);
  }
  if (preferences) {
    const prefStart = source.indexOf("// BEGIN FILTER SEND PREFERENCE METHODS");
    const prefEnd = source.indexOf("// END FILTER SEND PREFERENCE METHODS", prefStart);
    assert.ok(prefStart >= 0, "filter send preference start marker missing");
    assert.ok(prefEnd > prefStart, "filter send preference end marker missing");
    sandbox.preferenceAPI = vm.runInContext(`({${source.slice(prefStart, prefEnd)}})`, sandbox);
  }
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
  for (const accessor of ["str", "priority", "date", "status", "size", "age", "junkStatus", "junkPercent", "label"]) {
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
  typed("label", [LEGACY_ACTIONS.Label], (v) => {
    if (!Number.isInteger(v) || v < 0 || v > 5) throw new Error("NS_ERROR_ILLEGAL_VALUE [nsIMsgRuleAction.label]");
  });
  typed("junkScore", [ACTIONS.JunkScore], (v) => {
    if (v < 0 || v > 100) throw new Error("NS_ERROR_ILLEGAL_VALUE [nsIMsgRuleAction.junkScore]");
  });
  return action;
}

function makeFilter(name = "") {
  const terms = [];
  const actions = [];
  return {
    filterName: name,
    filterDesc: "",
    enabled: true,
    filterType: FILTER_TYPES.InboxRule | FILTER_TYPES.Manual,
    temporary: false,
    unparseable: false,
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

// The helper fixtures above model typed values. Handler tests additionally
// need native list identity, persistence failures and the submitted list.
function makeFilterHarness({
  ci = makeCi(), allowSend, prefError = false,
  getAccessibleFolder = (uri) => ({ folder: { URI: uri } }),
  addressBooksRestricted = false,
} = {}) {
  const prefReads = [];
  const submissions = [];
  const temporaryLists = [];
  const folder = { URI: "imap://account/Inbox" };
  function makeList() {
    const filters = [];
    const mutations = [];
    return {
      filters,
      mutations,
      loggingEnabled: false,
      logStream: null,
      saveAttempts: 0,
      saveError: false,
      get filterCount() { return filters.length; },
      createFilter: makeFilter,
      getFilterAt(index) {
        assert.ok(index >= 0 && index < filters.length, `invalid getFilterAt index: ${index}`);
        return filters[index];
      },
      insertFilterAt(index, filter) {
        assert.ok(index >= 0 && index <= filters.length, `invalid insertFilterAt index: ${index}`);
        mutations.push("insert");
        filters.splice(index, 0, filter);
      },
      setFilterAt(index, filter) {
        assert.ok(index >= 0 && index < filters.length, `invalid setFilterAt index: ${index}`);
        mutations.push("set");
        filters[index] = filter;
      },
      removeFilterAt(index) {
        assert.ok(index >= 0 && index < filters.length, `invalid removeFilterAt index: ${index}`);
        mutations.push("remove");
        filters.splice(index, 1);
      },
      saveToDefaultFile() {
        this.saveAttempts++;
        if (this.saveError) throw new Error("filter save failed");
      },
    };
  }
  const filterList = makeList();
  const account = {
    key: "account",
    incomingServer: {
      canHaveFilters: true,
      prettyName: "Test account",
      rootFolder: folder,
      getFilterList: () => filterList,
    },
  };
  const api = loadFilterHelpers({
    ci,
    handlers: true,
    prefs: {
      getBoolPref(name, fallback) {
        prefReads.push(name);
        if (prefError) throw new Error("preference unreadable");
        return allowSend === undefined ? fallback : allowSend;
      },
    },
    globals: {
      isAccountAllowed: (id) => id === account.key,
      isAddressBookAccessRestricted: () => addressBooksRestricted,
      getAccessibleAccounts: () => [account],
      getAccessibleFolder,
      MailServices: {
        accounts: { getAccount: (id) => id === account.key ? account : null },
        filters: {
          getTempFilterList() {
            const list = makeList();
            temporaryLists.push(list);
            return list;
          },
          applyFiltersToFolders(list, folders) {
            submissions.push({ list, filters: [...list.filters], folders: [...folders] });
          },
        },
      },
    },
  });
  return {
    api, ci, filterList, submissions, temporaryLists, prefReads, folder,
    seed({ name = "Existing", conditions, actions, ...metadata } = {}) {
      const filter = makeFilter(name);
      Object.assign(filter, metadata);
      api.buildTerms(filter, conditions || [{ attrib: "subject", op: "contains", value: "invoice" }]);
      api.buildActions(filter, actions || [{ type: "markRead" }]);
      filterList.filters.push(filter);
      return filter;
    },
    snapshot() {
      return JSON.stringify(api.listFilters(account.key));
    },
  };
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
    const term = buildOne(helpers, { attrib: "date", op: "isBefore", value: "2026-01-02" });
    assert.equal(term.value.date, localMidnightMicros(2026, 1, 2));
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
      /must be YYYY-MM-DD \(a local calendar day\); date-times are not accepted/
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

describe("numeric search values fit their native fields", () => {
  const cases = [
    { attrib: "size", min: 0, max: 4294967295 },
    { attrib: "ageInDays", min: 0, max: 2147483647 },
    { attrib: "status", min: 1, max: 4294967295 },
    { attrib: "priority", min: 2, max: 6 },
    { attrib: "junkPercent", min: 0, max: 100 },
    { attrib: "junkStatus", min: 0, max: 2 },
  ];

  for (const { attrib, min, max } of cases) {
    it(`${attrib} accepts its boundaries and rejects values outside them`, () => {
      const helpers = loadFilterHelpers();
      for (const value of [min, max]) {
        const term = buildOne(helpers, { attrib, op: "is", value: String(value) });
        assert.equal(storedValue(term), value);
        const copy = makeFilter();
        helpers.copySearchTerms({ searchTerms: [term] }, copy);
        assert.equal(storedValue(copy.searchTerms[0]), value);
      }
      for (const value of [min - 1, max + 1, 4294967296, Number.MAX_SAFE_INTEGER]) {
        assert.throws(
          () => buildOne(helpers, { attrib, op: "is", value: String(value) }),
          /Condition value.*must be/,
          `${attrib} accepted ${value}`
        );
      }
    });

    it(`${attrib} overflow rejects create/update before mutation or save`, () => {
      const h = makeFilterHarness();
      const original = h.seed();
      const before = h.snapshot();
      const conditions = [{ attrib, op: "is", value: String(max + 1) }];
      const created = h.api.createFilter("account", "Overflow", true, undefined,
        conditions, [{ type: "markRead" }]);
      assert.match(created.error, /Condition value.*must be/);
      const updated = h.api.updateFilter("account", 0, "Renamed", false, undefined, conditions);
      assert.match(updated.error, /Condition value.*must be/);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.filters[0], original);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.deepEqual(h.filterList.mutations, []);
    });
  }

  it("keeps priority within signed 32-bit bounds when level constants are unavailable", () => {
    const ci = makeCi();
    delete ci.nsMsgPriority;
    const helpers = loadFilterHelpers({ ci });
    for (const value of [-2147483648, 2147483647]) {
      assert.equal(buildOne(helpers, { attrib: "priority", op: "is", value: String(value) }).value.priority, value);
    }
    for (const value of [-2147483649, 2147483648, 4294967296]) {
      assert.throws(
        () => buildOne(helpers, { attrib: "priority", op: "is", value: String(value) }),
        /signed 32-bit integer/
      );
    }
  });

  it("bounds the typed setters for folder flags, uint32 properties and legacy labels", () => {
    const ci = makeCi({ attribs: LEGACY_ATTRIB });
    const helpers = loadFilterHelpers({ ci });
    for (const attrib of [ci.nsMsgSearchAttrib.FolderFlag, ci.nsMsgSearchAttrib.Uint32HdrProperty, ci.nsMsgSearchAttrib.Label]) {
      const value = makeSearchValue();
      value.attrib = attrib;
      const member = LEGAL_ACCESSOR[attrib];
      for (const boundary of [0, 4294967295]) {
        helpers.setSearchValue(value, attrib, String(boundary));
        assert.equal(value[member], boundary);
      }
      for (const invalid of [-1, 4294967296]) {
        assert.throws(() => helpers.setSearchValue(value, attrib, String(invalid)), /must be/);
        assert.equal(value[member], 4294967295);
      }
    }
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

  it("refuses zoned and local date-times, including midnight", () => {
    for (const value of [
      "2026-01-02T00:00:00Z", "2026-01-02T03:04:05.000Z",
      "2026-01-02T03:04:05+02:00", "2026-01-02T03:04:05-07:00",
      "2026-01-02T03:04", "2026-01-02 03:04:05",
    ]) {
      assert.throws(() => dateOf(value), /must be YYYY-MM-DD .*date-times are not accepted/, `accepted ${value}`);
    }
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

  it("can still read an existing native date with a time of day", () => {
    const term = buildOne(helpers, { attrib: "date", op: "isBefore", value: "2026-01-02" });
    term.value.date = Date.parse("2026-01-02T03:04:05.000Z") * 1000;
    assert.equal(helpers.getSearchValue(term.value, term.attrib), "2026-01-02T03:04:05.000Z");
  });

  it("rejects date-times on create and update before changing or saving the list", () => {
    for (const operation of ["create", "update"]) {
      const h = makeFilterHarness();
      const original = h.seed();
      const before = h.snapshot();
      const conditions = [{ attrib: "date", op: "isBefore", value: "2026-01-01T00:00:00Z" }];
      const result = operation === "create"
        ? h.api.createFilter("account", "New", true, undefined, conditions, [{ type: "markRead" }])
        : h.api.updateFilter("account", 0, "Renamed", false, undefined, conditions);
      assert.match(result.error, /date-times are not accepted/);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.filters[0], original);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.deepEqual(h.filterList.mutations, []);
    }
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
    assert.match(d, /size: a non-negative integer \(KB\), at most 4294967295/);
    assert.match(d, /ageInDays: a non-negative integer \(days\), at most 2147483647/);
    assert.match(d, /priority: an integer from 2 to 6 \(2=lowest, 3=low, 4=normal, 5=high, 6=highest\)/);
    assert.match(d, /status: a message-flag bitmask from 1 to 4294967295 \(1=read, 2=replied, 4=flagged, 4096=forwarded, 65536=new\)/);
    assert.match(d, /date: YYYY-MM-DD \(a local calendar day\); date-times are not accepted/);
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
    // Custom actions are not part of the creation vocabulary.
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
    const checkTargetFolder = (uri) => { seen.push(uri); return uri.includes("secret") ? { error: "no" } : { folder: { URI: uri } }; };
    reject({ type: "copyToFolder", value: "imap://a/secret" }, /Filter target folder not accessible: imap:\/\/a\/secret/, { checkTargetFolder });
    const ok = buildOneAction(helpers, { type: "copyToFolder", value: "imap://a/ok" }, { checkTargetFolder });
    assert.equal(ok.targetFolderUri, "imap://a/ok");
    assert.deepEqual(seen, ["imap://a/secret", "imap://a/ok"]);
  });

  for (const type of ["moveToFolder", "copyToFolder"]) {
    it(`writes the resolved URI to the native ${type} action`, () => {
      const requested = "imap://user@example.invalid/Project Work";
      const canonicalURI = "imap://user%40example.invalid/Project%20Work";
      const filter = makeFilter();
      const native = makeRuleAction();
      filter.createAction = () => new Proxy(native, {
        set(target, member, value) {
          if (member === "targetFolderUri" && value !== canonicalURI) {
            throw new Error("NS_ERROR_INVALID_ARG: targetFolderUri is not canonical");
          }
          return Reflect.set(target, member, value);
        },
      });
      helpers.buildActions(filter, [{ type, value: requested }], {
        checkTargetFolder(uri) {
          assert.equal(uri, requested);
          return { folder: { URI: canonicalURI } };
        },
      });
      assert.equal(filter.actionCount, 1);
      assert.equal(filter.getActionAt(0).targetFolderUri, canonicalURI);
    });

    it(`refuses ${type} when resolution supplies no usable folder URI`, () => {
      for (const resolved of [undefined, null, {}, { folder: {} }, { folder: { URI: "" } }, { folder: { URI: "  " } }]) {
        const filter = makeFilter();
        assert.throws(() => helpers.buildActions(filter, [{ type, value: "imap://a/b" }], {
          checkTargetFolder: () => resolved,
        }), /Filter target folder not accessible/);
        assert.equal(filter.actionCount, 0);
      }
    });
  }

  it("refuses unknown and custom actions with a reason", () => {
    reject({ type: "deleteBody" }, /Unknown action type: deleteBody/);
    reject({ type: "label", value: "1" }, /Unknown action type: label/);
    reject({ type: "custom", value: "x" }, /Custom filter actions are unsupported/);
  });
});

describe("reading filters back", () => {
  const helpers = loadFilterHelpers();

  it("recognizes ALL before reading condition members that do not exist", () => {
    const term = {
      matchAll: true,
      booleanAnd: false,
      get attrib() { throw new Error("ALL has no attribute"); },
      get op() { throw new Error("ALL has no operator"); },
      get value() { throw new Error("ALL has no value"); },
    };
    assert.deepEqual({ ...helpers.serializeSearchTerm(term) }, { matchAll: true, booleanAnd: false });
  });

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

describe("listFilters identifies rules matching all messages", () => {
  it("does not report an empty native rule as matching all messages", () => {
    // nsMsgSearchOfflineMail::MatchTerms returns !Filtering for an empty list:
    // https://searchfox.org/comm-central/source/mailnews/search/src/nsMsgLocalSearch.cpp
    const h = makeFilterHarness();
    h.seed({ conditions: [] });
    const listed = h.api.listFilters("account")[0].filters[0];
    assert.notEqual(listed.matchAll, true);
    assert.equal(listed.terms.length, 0);
  });

  it("reports native ALL as matchAll without exposing a bogus condition", () => {
    const h = makeFilterHarness();
    const native = h.seed({ conditions: [] });
    native.appendTerm({
      matchAll: true,
      get attrib() { throw new Error("ALL has no attribute"); },
      get op() { throw new Error("ALL has no operator"); },
      get value() { throw new Error("ALL has no value"); },
    });
    const listed = h.api.listFilters("account")[0].filters[0];
    assert.equal(listed.matchAll, true);
    assert.equal(listed.terms.length, 0);
    assert.equal(native.searchTerms.length, 1);
    assert.equal(native.searchTerms[0].matchAll, true);
  });

  for (const booleanAnd of [true, false]) {
    for (const allFirst of [false, true]) {
      it(`retains compound conditions and their operator, AND=${booleanAnd}, ALL first=${allFirst}`, () => {
        const h = makeFilterHarness();
        const condition = { attrib: "subject", op: "contains", value: "invoice", booleanAnd };
        const native = h.seed({ conditions: [condition] });
        const all = {
          matchAll: true, booleanAnd,
          get attrib() { throw new Error("ALL has no attribute"); },
          get op() { throw new Error("ALL has no operator"); },
          get value() { throw new Error("ALL has no value"); },
        };
        if (allFirst) native.searchTerms.unshift(all);
        else native.appendTerm(all);
        const listed = h.api.listFilters("account")[0].filters[0];
        assert.notEqual(listed.matchAll, true);
        const allTerm = { matchAll: true, booleanAnd };
        assert.deepEqual(Array.from(listed.terms, term => ({ ...term })),
          allFirst ? [allTerm, condition] : [condition, allTerm]);
        assert.equal(native.searchTerms.length, 2);
      });
    }
  }

  it("only collapses a lone ALL, leaving multiple native terms explicit", () => {
    const h = makeFilterHarness();
    const native = h.seed({ conditions: [] });
    native.appendTerm({ matchAll: true, booleanAnd: true });
    native.appendTerm({ matchAll: true, booleanAnd: false });
    const listed = h.api.listFilters("account")[0].filters[0];
    assert.notEqual(listed.matchAll, true);
    assert.deepEqual(Array.from(listed.terms, term => ({ ...term })), [
      { matchAll: true, booleanAnd: true }, { matchAll: true, booleanAnd: false },
    ]);
  });

  it("keeps real conditions, including an empty subject value", () => {
    const h = makeFilterHarness();
    h.seed({ conditions: [{ attrib: "subject", op: "isEmpty", value: "" }] });
    const listed = h.api.listFilters("account")[0].filters[0];
    assert.notEqual(listed.matchAll, true);
    assert.equal(listed.terms.length, 1);
    assert.deepEqual({ ...listed.terms[0] }, {
      attrib: "subject", op: "isEmpty", booleanAnd: true, value: "",
    });
  });

  it("preserves native ALL on update and reports replacement conditions normally", () => {
    const h = makeFilterHarness();
    const native = h.seed({ conditions: [{ attrib: "subject", op: "contains", value: "" }] });
    native.searchTerms[0].matchAll = true;
    const renamed = h.api.updateFilter("account", 0, "All messages");
    assert.equal(renamed.success, true, renamed.error);
    assert.equal(renamed.filter.matchAll, true);
    assert.equal(renamed.filter.terms.length, 0);
    assert.equal(h.filterList.filters[0].searchTerms[0].matchAll, true);

    const replaced = h.api.updateFilter("account", 0, undefined, undefined, undefined,
      [{ attrib: "from", op: "contains", value: "sender@example.invalid" }]);
    assert.equal(replaced.success, true, replaced.error);
    assert.notEqual(replaced.filter.matchAll, true);
    assert.deepEqual({ ...replaced.filter.terms[0] }, {
      attrib: "from", op: "contains", booleanAnd: true, value: "sender@example.invalid",
    });
  });

  it("does not infer matchAll when native terms cannot be read", () => {
    for (const firstTerm of [undefined, { matchAll: true }]) {
      const h = makeFilterHarness();
      const native = h.seed();
      native.searchTerms = {
        *[Symbol.iterator]() {
          if (firstTerm) yield firstTerm;
          throw new Error("native terms unavailable");
        },
      };
      const listed = h.api.listFilters("account")[0].filters[0];
      assert.notEqual(listed.matchAll, true);
    }
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

describe("filter handlers preserve native values", () => {
  it("creates and reads typed conditions, then preserves them when replacing actions", () => {
    const h = makeFilterHarness();
    const conditions = [
      { attrib: "date", op: "isBefore", value: "2026-01-01" },
      { attrib: "priority", op: "is", value: "4" },
      { attrib: "status", op: "is", value: "2" },
      { attrib: "ageInDays", op: "isGreaterThan", value: "30" },
      { attrib: "size", op: "isGreaterThan", value: "1024" },
      { attrib: "junkStatus", op: "is", value: "2" },
      { attrib: "junkPercent", op: "isGreaterThan", value: "90" },
      { attrib: "hasAttachment", op: "is", value: "" },
      { attrib: "otherHeader", op: "contains", value: "bulk", header: "X-Mailer" },
    ];
    const created = h.api.createFilter("account", "Typed", true, undefined, conditions, [{ type: "markRead" }]);
    assert.equal(created.success, true, created.error);
    const before = JSON.stringify(h.api.listFilters("account")[0].filters[0].terms);
    const values = h.filterList.filters[0].searchTerms.map(storedValue);

    const updated = h.api.updateFilter("account", 0, undefined, undefined, undefined, undefined, [{ type: "addTag", value: "$label2" }]);
    assert.equal(updated.success, true, updated.error);
    assert.equal(JSON.stringify(updated.filter.terms), before);
    assert.deepEqual(h.filterList.filters[0].searchTerms.map(storedValue), values);
    assert.equal(updated.filter.terms[0].value, "2026-01-01");
    assert.equal(updated.filter.terms[8].header, "X-Mailer");
  });

  it("keeps typed action members and filter metadata when replacing conditions", () => {
    const h = makeFilterHarness();
    const original = h.seed({
      filterDesc: 'A "quoted" description',
      temporary: true,
      actions: [
        { type: "copyToFolder", value: "imap://account/Archive" },
        { type: "changePriority", value: "6" },
        { type: "junkScore", value: "100" },
        { type: "addTag", value: "$label1" },
      ],
    });
    const before = JSON.stringify(h.api.listFilters("account")[0].filters[0].actions);
    const result = h.api.updateFilter("account", 0, undefined, undefined, undefined, [{ attrib: "from", op: "contains", value: "sender@example.invalid" }]);
    assert.equal(result.success, true, result.error);
    assert.equal(JSON.stringify(result.filter.actions), before);
    assert.equal(h.filterList.filters[0].filterDesc, original.filterDesc);
    assert.equal(h.filterList.filters[0].temporary, true);
    assert.deepEqual(h.filterList.mutations, ["set"]);
  });

  it("copies uncommon and legacy typed terms without exposing them for creation", () => {
    const ci = makeCi({ attribs: LEGACY_ATTRIB, actions: LEGACY_ACTIONS });
    const h = makeFilterHarness({ ci });
    const original = h.seed();
    for (const [attrib, member, value] of [
      [ci.nsMsgSearchAttrib.FolderFlag, "status", 4096],
      [ci.nsMsgSearchAttrib.Uint32HdrProperty, "status", 23],
      [ci.nsMsgSearchAttrib.Label, "label", 3],
    ]) {
      const term = makeSearchTerm();
      term.attrib = attrib;
      term.op = ci.nsMsgSearchOp.Is;
      term.booleanAnd = false;
      term.hdrProperty = "custom-property";
      term.beginsGrouping = true;
      term.endsGrouping = true;
      term.value.attrib = attrib;
      term.value[member] = value;
      original.appendTerm(term);
    }
    const custom = makeSearchTerm();
    custom.attrib = ci.nsMsgSearchAttrib.Custom;
    custom.op = ci.nsMsgSearchOp.Contains;
    custom.customId = "test@example.invalid#condition";
    custom.matchAll = true;
    custom.value.attrib = custom.attrib;
    custom.value.str = "kept";
    original.appendTerm(custom);

    const result = h.api.updateFilter("account", 0, undefined, undefined, undefined, undefined, [{ type: "markFlagged" }]);
    assert.equal(result.success, true, result.error);
    const copied = h.filterList.filters[0].searchTerms;
    assert.deepEqual(copied.map(storedValue), original.searchTerms.map(storedValue));
    for (const term of copied.slice(1, 4)) {
      assert.equal(term.hdrProperty, "custom-property");
      assert.equal(term.beginsGrouping, true);
      assert.equal(term.endsGrouping, true);
      assert.equal(term.booleanAnd, false);
    }
    assert.equal(copied[4].customId, custom.customId);
    assert.equal(copied[4].matchAll, true);
    for (const attrib of ["folderFlag", "uint32HdrProperty", "label"]) {
      assert.equal(Object.hasOwn(h.api.ATTRIB_MAP, attrib), false);
      assert.throws(() => h.api.buildTerms(makeFilter(), [{ attrib, op: "is", value: "1" }]));
    }
  });

  it("writes, reads and copies Thunderbird 102 labels through the label member", () => {
    const h = makeFilterHarness({ ci: makeCi({ actions: LEGACY_ACTIONS, attribs: LEGACY_ATTRIB }) });
    const created = h.api.createFilter("account", "Label", true, undefined,
      [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type: "label", value: "5" }]);
    assert.equal(created.success, true, created.error);
    assert.equal(h.filterList.filters[0].getActionAt(0).label, 5);
    assert.equal(h.api.listFilters("account")[0].filters[0].actions[0].value, "5");
    const copied = h.api.updateFilter("account", 0, undefined, undefined, undefined, [{ attrib: "subject", op: "contains", value: "receipt" }]);
    assert.equal(copied.success, true, copied.error);
    assert.equal(h.filterList.filters[0].getActionAt(0).label, 5);
    assert.equal(h.filterList.filters[0].getActionAt(0).strValue, "");
    for (const value of ["-1", "6", "1.5", "label"]) {
      assert.throws(() => buildOneAction(h.api, { type: "label", value }));
    }
    assert.equal(buildOneAction(h.api, { type: "label", value: "0" }).label, 0);
  });
});

describe("persisted filter text is validated before changing the list", () => {
  const forbidden = [...Array.from({ length: 32 }, (_, i) => String.fromCharCode(i)), "\x7f", "\\"];
  const fields = [
    ["name", (request, text) => { request.name = text; }],
    ["condition value", (request, text) => { request.conditions[0].value = text; }],
    ["condition header", (request, text) => {
      request.conditions[0] = { attrib: "otherHeader", op: "contains", value: "bulk", header: text };
    }],
    ...[
      ["moveToFolder", "imap://account/Archive"], ["copyToFolder", "imap://account/Archive"],
      ["addTag", "tag"], ["reply", "mailbox://templates?number=1"], ["forward", "target@example.invalid"],
    ].map(([type, value]) => [
      `${type} value`, (request, text) => { request.actions = [{ type, value: text }]; }, value,
    ]),
    ...[
      ["priority", "4"], ["status", "1"], ["ageInDays", "1"], ["size", "1"],
      ["junkStatus", "1"], ["junkPercent", "1"], ["date", "2026-01-01"], ["hasAttachment", ""],
    ].map(([attrib, value]) => [
      `${attrib} condition value`, (request, text) => { request.conditions = [{ attrib, op: "is", value: text }]; }, value,
    ]),
    ...["changePriority", "junkScore", "label"].map((type) => [
      `${type} action value`, (request, text) => { request.actions = [{ type, value: text }]; }, "4",
    ]),
  ];
  for (const [field, change, value = "1"] of fields) {
    it(`rejects every forbidden character in ${field} on create and update`, () => {
      const h = makeFilterHarness({ allowSend: true, ci: makeCi({ actions: LEGACY_ACTIONS }) });
      const original = h.seed();
      const before = h.snapshot();
      for (const character of forbidden) {
        for (const operation of ["create", "update"]) {
          const request = {
            name: "Updated",
            conditions: [{ attrib: "subject", op: "contains", value: "invoice" }],
            actions: [{ type: "markRead" }],
          };
          change(request, `${value}${character}`);
          const result = operation === "create"
            ? h.api.createFilter("account", request.name, false, undefined, request.conditions, request.actions)
            : h.api.updateFilter("account", 0, request.name, false, undefined, request.conditions, request.actions);
          assert.ok(result.error, `${operation} ${field}: ${JSON.stringify(character)}`);
          assert.equal(h.snapshot(), before);
          assert.equal(h.filterList.filters[0], original);
          assert.deepEqual(h.filterList.mutations, []);
          assert.equal(h.filterList.saveAttempts, 0);
        }
      }
    });
  }

  it("preserves quotes, Unicode and Unicode separators in valid text", () => {
    const h = makeFilterHarness();
    const text = 'Zażółć "gęślą" 日本語\u2028next\u2029last';
    const result = h.api.createFilter("account", text, true, undefined,
      [{ attrib: "subject", op: "contains", value: text }], [{ type: "addTag", value: text }]);
    assert.equal(result.success, true, result.error);
    const stored = h.api.listFilters("account")[0].filters[0];
    assert.equal(stored.name, text);
    assert.equal(stored.terms[0].value, text);
    assert.equal(stored.actions[0].value, text);
  });

  it("validates copied text even when only the name changes", () => {
    const h = makeFilterHarness();
    const original = h.seed();
    original.searchTerms[0].value.str = "old\\value";
    const before = h.snapshot();
    const result = h.api.updateFilter("account", 0, "Renamed");
    assert.ok(result.error);
    assert.equal(h.snapshot(), before);
    assert.equal(h.filterList.filters[0], original);
    assert.equal(h.filterList.saveAttempts, 0);
  });

  it("validates preserved descriptions, headers and identifiers before committing an update", () => {
    for (const change of [
      (filter) => { filter.filterDesc = "old\ndescription"; },
      (filter) => { filter.searchTerms[0].arbitraryHeader = "old\nheader"; },
      (filter) => { filter.searchTerms[0].hdrProperty = "old\nproperty"; },
      (filter) => { filter.searchTerms[0].customId = "old\ncondition"; },
      (filter) => { filter.getActionAt(0).strValue = "old\nvalue"; },
      (filter) => { filter.getActionAt(0).customId = "old\naction"; },
    ]) {
      const h = makeFilterHarness();
      const original = h.seed();
      change(original);
      const before = h.snapshot();
      const result = h.api.updateFilter("account", 0, "Renamed");
      assert.ok(result.error);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.filters[0], original);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.deepEqual(h.filterList.mutations, []);
    }
  });
});

describe("filter sending preference methods", () => {
  it("stores explicit boolean choices and reads them in a fresh production context", async () => {
    const stored = new Map();
    const writes = [];
    const prefs = {
      getBoolPref: (name, fallback) => stored.has(name) ? stored.get(name) : fallback,
      setBoolPref(name, value) { stored.set(name, value); writes.push([name, value]); },
    };
    const first = loadFilterHelpers({ prefs, preferences: true }).preferenceAPI;
    assert.equal((await first.getAllowFilterSendActions()).allowFilterSendActions, false);
    assert.equal((await first.setAllowFilterSendActions(true)).success, true);
    const second = loadFilterHelpers({ prefs, preferences: true }).preferenceAPI;
    assert.equal((await second.getAllowFilterSendActions()).allowFilterSendActions, true);
    assert.equal((await second.setAllowFilterSendActions(false)).success, true);
    assert.equal((await first.getAllowFilterSendActions()).allowFilterSendActions, false);
    assert.deepEqual(writes, [
      ["extensions.thunderbird-mcp.allowFilterSendActions", true],
      ["extensions.thunderbird-mcp.allowFilterSendActions", false],
    ]);
  });

  it("rejects non-boolean writes and reads an unavailable preference as false", async () => {
    const writes = [];
    const api = loadFilterHelpers({
      preferences: true,
      prefs: {
        getBoolPref() { throw new Error("preference unavailable"); },
        setBoolPref: (...args) => writes.push(args),
      },
    }).preferenceAPI;
    assert.equal((await api.getAllowFilterSendActions()).allowFilterSendActions, false);
    for (const value of [undefined, null, 0, 1, "true", "false", {}, []]) {
      assert.ok((await api.setAllowFilterSendActions(value)).error);
    }
    assert.deepEqual(writes, []);
  });
});

describe("filter sending requires an explicit preference", () => {
  for (const [name, options] of [
    ["default", {}], ["false", { allowSend: false }],
    ["unreadable", { prefError: true }], ["string", { allowSend: "true" }], ["number", { allowSend: 1 }],
  ]) {
    it(`blocks forward and reply with the ${name} preference`, () => {
      for (const type of ["forward", "reply"]) {
        for (const enabled of [true, false]) {
          const h = makeFilterHarness(options);
          const result = h.api.createFilter("account", "Sending", enabled, undefined,
            [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type, value: "target@example.invalid" }]);
          assert.ok(result.error, `${type}, enabled=${enabled}`);
          assert.equal(h.filterList.filterCount, 0);
          assert.equal(h.filterList.saveAttempts, 0);
          assert.ok(h.prefReads.length > 0);
          assert.ok(h.prefReads.every((pref) => pref === "extensions.thunderbird-mcp.allowFilterSendActions"));
        }
      }
    });
  }

  it("uses resolved native action identities and leaves stopExecution available", () => {
    const ci = makeCi({ actions: { Forward: 110, Reply: 109, StopExecution: 111 } });
    for (const type of ["forward", "reply", "stopExecution"]) {
      const h = makeFilterHarness({ ci });
      const action = type === "stopExecution" ? { type } : { type, value: "target@example.invalid" };
      const result = h.api.createFilter("account", type, true, undefined,
        [{ attrib: "subject", op: "contains", value: "invoice" }], [action]);
      if (type === "stopExecution") {
        assert.equal(result.success, true, result.error);
        assert.equal(h.filterList.filters[0].getActionAt(0).type, ci.nsMsgFilterAction.StopExecution);
      } else {
        assert.ok(result.error);
        assert.equal(h.filterList.filterCount, 0);
      }
    }
  });

  it("refuses writes and execution if a required native action constant is unavailable", () => {
    for (const constant of ["Forward", "Reply", "Custom"]) {
      const h = makeFilterHarness({ allowSend: true, ci: makeCi({ actions: { [constant]: undefined } }) });
      const original = h.seed();
      const before = h.snapshot();
      const created = h.api.createFilter("account", "New", true, undefined,
        [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type: "markRead" }]);
      assert.match(created.error, /native action constants are unavailable/);
      assert.match(h.api.updateFilter("account", 0, "Renamed").error, /native action constants are unavailable/);
      assert.match(h.api.applyFilters("account", h.folder.URI).error, /native action constants are unavailable/);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.filters[0], original);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.deepEqual(h.filterList.mutations, []);
      assert.equal(h.submissions.length, 0);
    }
  });

  it("allows forward and reply creation only with the boolean true preference", () => {
    const h = makeFilterHarness({ allowSend: true });
    for (const type of ["forward", "reply"]) {
      const result = h.api.createFilter("account", type, true, undefined,
        [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type, value: "target@example.invalid" }]);
      assert.equal(result.success, true, result.error);
    }
    assert.equal(h.filterList.filterCount, 2);
  });

  it("checks the resulting existing rule for name, enable, type and condition updates", () => {
    const updates = [
      ["Renamed"],
      [undefined, true],
      [undefined, undefined, FILTER_TYPES.Manual],
      [undefined, undefined, undefined, [{ attrib: "subject", op: "contains", value: "broader" }]],
      ["Renamed", false],
      ["Invalid\nname", false],
    ];
    for (const type of ["forward", "reply"]) {
      for (const update of updates) {
        const h = makeFilterHarness();
        const original = h.seed({ enabled: update[1] !== true, actions: [{ type, value: "target@example.invalid" }] });
        const before = h.snapshot();
        const result = h.api.updateFilter("account", 0, ...update);
        assert.ok(result.error, `${type}: ${JSON.stringify(update)}`);
        assert.equal(h.snapshot(), before);
        assert.equal(h.filterList.filters[0], original);
        assert.equal(h.filterList.saveAttempts, 0);
        assert.deepEqual(h.filterList.mutations, []);
      }
    }
  });

  it("permits a pure disable and safe action replacement on an existing sending rule", () => {
    const h = makeFilterHarness();
    h.seed({ actions: [{ type: "forward", value: "target@example.invalid" }] });
    const disabled = h.api.updateFilter("account", 0, undefined, false);
    assert.equal(disabled.success, true, disabled.error);
    assert.equal(h.filterList.filters[0].enabled, false);
    assert.equal(h.filterList.filters[0].getActionAt(0).type, h.ci.nsMsgFilterAction.Forward);
    const renamed = h.api.updateFilter("account", 0, "Still blocked");
    assert.ok(renamed.error);
    const replaced = h.api.updateFilter("account", 0, undefined, undefined, undefined, undefined, [{ type: "markRead" }]);
    assert.equal(replaced.success, true, replaced.error);
    assert.equal(h.filterList.filters[0].getActionAt(0).type, h.ci.nsMsgFilterAction.MarkRead);
  });

  it("allows copying sending actions in an update after explicit opt-in", () => {
    const h = makeFilterHarness({ allowSend: true });
    h.seed({ actions: [{ type: "reply", value: "mailbox://templates?number=1" }] });
    const result = h.api.updateFilter("account", 0, "Renamed");
    assert.equal(result.success, true, result.error);
    assert.equal(h.filterList.filters[0].getActionAt(0).type, h.ci.nsMsgFilterAction.Reply);
    assert.equal(h.filterList.filters[0].getActionAt(0).strValue, "mailbox://templates?number=1");
  });

  it("checks replacement sending actions before changing a safe existing rule", () => {
    const h = makeFilterHarness();
    const original = h.seed();
    const before = h.snapshot();
    const result = h.api.updateFilter("account", 0, "Renamed", false, undefined, undefined,
      [{ type: "forward", value: "target@example.invalid" }]);
    assert.ok(result.error);
    assert.equal(h.snapshot(), before);
    assert.equal(h.filterList.filters[0], original);
    assert.equal(h.filterList.saveAttempts, 0);
  });
});

describe("Move/Copy filter writes resolve canonical destinations", () => {
  const conditions = [{ attrib: "subject", op: "contains", value: "invoice" }];
  const requested = "imap://user@example.invalid/Project Work";
  const canonicalURI = "imap://user%40example.invalid/Project%20Work";

  for (const type of ["moveToFolder", "copyToFolder"]) {
    for (const operation of ["create", "update"]) {
      it(`stores the canonical ${type} URI on ${operation}`, () => {
        const lookups = [];
        const h = makeFilterHarness({
          getAccessibleFolder(uri) {
            lookups.push(uri);
            assert.ok(uri === requested || uri === canonicalURI);
            return { folder: { URI: canonicalURI } };
          },
        });
        if (operation === "update") h.seed();
        const actions = [{ type, value: requested }];
        const result = operation === "create"
          ? h.api.createFilter("account", "New", true, undefined, conditions, actions)
          : h.api.updateFilter("account", 0, "Renamed", undefined, undefined, undefined, actions);
        assert.equal(result.success, true, result.error);
        assert.equal(lookups[0], requested);
        assert.equal(h.filterList.filters[0].getActionAt(0).targetFolderUri, canonicalURI);
        assert.equal(h.api.listFilters("account")[0].filters[0].actions[0].value, canonicalURI);
        assert.equal(h.filterList.saveAttempts, 1);
      });

      it(`leaves the list unchanged when ${operation} cannot resolve its ${type} target`, () => {
        for (const failure of ["restricted", "missing", "invalid-uri", "lookup-error"]) {
          const h = makeFilterHarness({
            getAccessibleFolder() {
              if (failure === "lookup-error") throw new Error("Folder lookup failed");
              if (failure === "invalid-uri") return { folder: {} };
              return { error: failure === "restricted" ? "Account not accessible" : "Folder not found" };
            },
          });
          const original = h.seed();
          const before = h.snapshot();
          const actions = [{ type, value: requested }];
          const result = operation === "create"
            ? h.api.createFilter("account", "New", true, undefined, conditions, actions)
            : h.api.updateFilter("account", 0, "Renamed", false, undefined, undefined, actions);
          assert.match(result.error, /Filter target folder not accessible|Folder lookup failed/);
          assert.equal(h.snapshot(), before);
          assert.equal(h.filterList.filters[0], original);
          assert.equal(h.filterList.saveAttempts, 0);
          assert.deepEqual(h.filterList.mutations, []);
        }
      });
    }
  }
});

describe("retained Move/Copy actions respect current destination access", () => {
  const inaccessible = [
    "imap://restricted/Archive",
    "imap://account/Missing",
    "imap://account/LookupFailure",
  ];
  const accessible = "imap://account/Archive";
  const getAccessibleFolder = (uri) => {
    if (uri === inaccessible[0]) return { error: "Account not accessible" };
    if (uri === inaccessible[1]) return { error: "Folder not found" };
    if (uri === inaccessible[2]) throw new Error("Folder lookup failed");
    return { folder: { URI: uri } };
  };

  for (const type of ["moveToFolder", "copyToFolder"]) {
    it(`rejects updates retaining an inaccessible ${type} without changing the live rule`, () => {
      for (const value of inaccessible) {
        for (const update of [
          [],
          ["Renamed"],
          [undefined, true],
          [undefined, undefined, FILTER_TYPES.Manual],
          [undefined, undefined, undefined, [{ attrib: "subject", op: "contains", value: "changed" }]],
          ["Renamed", false],
        ]) {
          const h = makeFilterHarness({ getAccessibleFolder });
          const original = h.seed({ enabled: false, actions: [{ type, value }] });
          const before = h.snapshot();
          const result = h.api.updateFilter("account", 0, ...update);
          assert.match(result.error, /Filter target folder not accessible/, `${value}: ${JSON.stringify(update)}`);
          assert.equal(h.snapshot(), before);
          assert.equal(h.filterList.filters[0], original);
          assert.equal(h.filterList.saveAttempts, 0);
          assert.deepEqual(h.filterList.mutations, []);
        }
      }
    });

    it(`permits pure-disable and delete recovery for an inaccessible ${type}`, () => {
      for (const value of inaccessible) {
        const h = makeFilterHarness({ getAccessibleFolder });
        h.seed({ actions: [{ type, value }] });
        const result = h.api.updateFilter("account", 0, undefined, false);
        assert.equal(result.success, true, result.error);
        assert.equal(h.filterList.filters[0].enabled, false);
        assert.equal(h.filterList.filters[0].getActionAt(0).targetFolderUri, value);
        assert.equal(h.api.deleteFilter("account", 0).success, true);
        assert.equal(h.filterList.filterCount, 0);
      }
      const h = makeFilterHarness({ getAccessibleFolder });
      h.seed({ actions: [{ type, value: inaccessible[0] }] });
      assert.equal(h.api.deleteFilter("account", 0).success, true);
    });

    it(`checks replacement ${type} destinations and permits repairing an existing rule`, () => {
      const h = makeFilterHarness({ getAccessibleFolder });
      const original = h.seed({ actions: [{ type, value: inaccessible[0] }] });
      const before = h.snapshot();
      const rejected = h.api.updateFilter("account", 0, undefined, false, undefined, undefined,
        [{ type, value: inaccessible[1] }]);
      assert.match(rejected.error, /Filter target folder not accessible/);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.filters[0], original);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.deepEqual(h.filterList.mutations, []);
      const repaired = h.api.updateFilter("account", 0, undefined, undefined, undefined, undefined,
        [{ type, value: accessible }]);
      assert.equal(repaired.success, true, repaired.error);
      assert.equal(h.filterList.filters[0].getActionAt(0).targetFolderUri, accessible);
      const renamed = h.api.updateFilter("account", 0, "Accessible retained destination");
      assert.equal(renamed.success, true, renamed.error);
    });
  }

  it("allows removing an inaccessible destination from the resulting rule", () => {
    const h = makeFilterHarness({ getAccessibleFolder });
    h.seed({ actions: [{ type: "moveToFolder", value: inaccessible[0] }] });
    const result = h.api.updateFilter("account", 0, undefined, undefined, undefined, undefined, [{ type: "markRead" }]);
    assert.equal(result.success, true, result.error);
    assert.equal(h.filterList.filters[0].actionCount, 1);
    assert.equal(h.filterList.filters[0].getActionAt(0).type, h.ci.nsMsgFilterAction.MarkRead);
  });

  it("skips entire rules with inaccessible destinations regardless of the sending preference", () => {
    for (const allowSend of [false, true]) {
      const h = makeFilterHarness({ allowSend, getAccessibleFolder });
      const expectedSkipped = [];
      for (const type of ["moveToFolder", "copyToFolder"]) {
        for (const value of inaccessible) {
          const name = `${type}: ${value}`;
          h.seed({ name, actions: [{ type: "markRead" }, { type, value }] });
          expectedSkipped.push({ name, reason: "inaccessible-destination" });
        }
        h.seed({ name: type, actions: [{ type, value: accessible }] });
      }
      const before = h.snapshot();
      const result = h.api.applyFilters("account", h.folder.URI);
      assert.equal(result.success, true, result.error);
      assert.equal(result.submittedFilters, 2);
      assert.deepEqual(Array.from(result.submitted), ["moveToFolder", "copyToFolder"]);
      assert.deepEqual(Array.from(result.skipped, (entry) => ({ ...entry })), expectedSkipped);
      assert.equal(h.submissions.length, 1);
      assert.deepEqual(h.submissions[0].filters.map((filter) => filter.filterName), ["moveToFolder", "copyToFolder"]);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.deepEqual(h.filterList.mutations, []);
    }
  });

  it("makes no native submission when every destination is inaccessible", () => {
    const h = makeFilterHarness({ getAccessibleFolder });
    h.seed({ actions: [{ type: "moveToFolder", value: inaccessible[0] }] });
    const result = h.api.applyFilters("account", h.folder.URI);
    assert.equal(result.success, true, result.error);
    assert.equal(result.submittedFilters, 0);
    assert.deepEqual(Array.from(result.skipped, (entry) => ({ ...entry })), [
      { name: "Existing", reason: "inaccessible-destination" },
    ]);
    assert.equal(h.submissions.length, 0);
  });
});

describe("custom filter actions remain unsupported by mutations", () => {
  it("rejects custom creation and existing custom updates even with sending allowed", () => {
    for (const allowSend of [false, true]) {
      const h = makeFilterHarness({ allowSend });
      const created = h.api.createFilter("account", "Custom", true, undefined,
        [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type: "custom", value: "value" }]);
      assert.ok(created.error);
      const original = h.seed();
      const custom = makeRuleAction();
      custom.type = h.ci.nsMsgFilterAction.Custom;
      custom.customId = "test@example.invalid#action";
      custom.strValue = "value";
      original.appendAction(custom);
      const before = h.snapshot();
      assert.equal(h.api.listFilters("account")[0].filters[0].actions[1].customId, custom.customId);
      for (const update of [["Renamed"], [undefined, false]]) {
        const result = h.api.updateFilter("account", 0, ...update);
        assert.ok(result.error);
        assert.equal(h.snapshot(), before);
        assert.equal(h.filterList.filters[0], original);
        assert.equal(h.filterList.saveAttempts, 0);
      }
      const deleted = h.api.deleteFilter("account", 0);
      assert.equal(deleted.success, true, deleted.error);
      assert.equal(h.filterList.filterCount, 0);
    }
  });

  it("allows replacing an existing custom action with a supported safe action", () => {
    const h = makeFilterHarness();
    const original = h.seed();
    const custom = makeRuleAction();
    custom.type = h.ci.nsMsgFilterAction.Custom;
    custom.customId = "test@example.invalid#action";
    original.appendAction(custom);
    const result = h.api.updateFilter("account", 0, undefined, undefined, undefined, undefined, [{ type: "markRead" }]);
    assert.equal(result.success, true, result.error);
    assert.equal(h.filterList.filters[0].actionCount, 1);
  });
});

describe("applyFilters submits only eligible native filters", () => {
  it("separates enabled manual filters from disabled, non-manual, sending and unparseable rules", () => {
    const h = makeFilterHarness();
    h.filterList.loggingEnabled = true;
    h.filterList.logStream = { name: "test log" };
    h.seed({ name: "Disabled", enabled: false });
    h.seed({ name: "Outgoing", filterType: h.ci.nsMsgFilterType.PostOutgoing });
    h.seed({ name: "Unparseable", unparseable: true });
    h.seed({ name: "Forward", actions: [{ type: "forward", value: "target@example.invalid" }] });
    h.seed({ name: "Reply", actions: [{ type: "reply", value: "mailbox://templates?number=1" }] });
    h.seed({ name: "Safe", filterType: h.ci.nsMsgFilterType.Manual });
    h.seed({ name: "Stop", actions: [{ type: "stopExecution" }] });
    const before = h.snapshot();
    const result = h.api.applyFilters("account", h.folder.URI);
    assert.equal(result.success, true, result.error);
    assert.equal(result.submittedFilters, 2);
    assert.deepEqual(Array.from(result.submitted), ["Safe", "Stop"]);
    assert.deepEqual(Array.from(result.skipped, (entry) => ({ ...entry })), [
      { name: "Disabled", reason: "disabled" },
      { name: "Outgoing", reason: "non-manual" },
      { name: "Unparseable", reason: "unparseable" },
      { name: "Forward", reason: "sending" },
      { name: "Reply", reason: "sending" },
    ]);
    assert.equal(Object.hasOwn(result, "enabledFilters"), false);
    assert.equal(h.submissions.length, 1);
    assert.notEqual(h.submissions[0].list, h.filterList);
    assert.equal(h.submissions[0].list.loggingEnabled, true);
    assert.equal(h.submissions[0].list.logStream, h.filterList.logStream);
    assert.deepEqual(h.submissions[0].filters.map((filter) => filter.filterName), ["Safe", "Stop"]);
    assert.equal(h.submissions[0].folders[0].URI, h.folder.URI);
    assert.equal(h.snapshot(), before);
    assert.equal(h.filterList.saveAttempts, 0);
    assert.deepEqual(h.filterList.mutations, []);
  });

  it("does not call the native service when no rule is eligible", () => {
    const h = makeFilterHarness();
    h.seed({ enabled: false });
    h.seed({ name: "Sending", actions: [{ type: "forward", value: "target@example.invalid" }] });
    const result = h.api.applyFilters("account", h.folder.URI);
    assert.equal(result.success, true, result.error);
    assert.equal(result.submittedFilters, 0);
    assert.deepEqual(Array.from(result.submitted), []);
    assert.equal(h.submissions.length, 0);
  });

  it("submits manual sending rules only after explicit opt-in", () => {
    const h = makeFilterHarness({ allowSend: true });
    h.seed({ name: "Forward", actions: [{ type: "forward", value: "target@example.invalid" }] });
    h.seed({ name: "Reply", actions: [{ type: "reply", value: "mailbox://templates?number=1" }] });
    h.seed({ name: "Disabled", enabled: false, actions: [{ type: "forward", value: "target@example.invalid" }] });
    const result = h.api.applyFilters("account", h.folder.URI);
    assert.equal(result.success, true, result.error);
    assert.equal(result.submittedFilters, 2);
    assert.deepEqual(Array.from(result.submitted), ["Forward", "Reply"]);
    assert.deepEqual(h.submissions[0].filters.map((filter) => filter.filterName), ["Forward", "Reply"]);
  });

  it("rejects an eligible Custom action before submitting any filters", () => {
    for (const allowSend of [false, true]) {
      const h = makeFilterHarness({ allowSend });
      h.seed({ name: "Safe first" });
      const filter = h.seed({ name: "Custom and sending", actions: [{ type: "forward", value: "target@example.invalid" }] });
      const custom = makeRuleAction();
      custom.type = h.ci.nsMsgFilterAction.Custom;
      custom.customId = "test@example.invalid#action";
      filter.appendAction(custom);
      const before = h.snapshot();
      const result = h.api.applyFilters("account", h.folder.URI);
      assert.match(result.error, /custom/i);
      assert.equal(h.submissions.length, 0);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.saveAttempts, 0);
    }
  });

  it("does not treat disabled or non-manual Custom rules as eligible", () => {
    const h = makeFilterHarness();
    for (const metadata of [{ name: "Disabled", enabled: false }, { name: "Outgoing", filterType: h.ci.nsMsgFilterType.PostOutgoing }]) {
      const filter = h.seed(metadata);
      const action = makeRuleAction();
      action.type = h.ci.nsMsgFilterAction.Custom;
      action.customId = "test@example.invalid#action";
      filter.appendAction(action);
    }
    h.seed({ name: "Safe" });
    const result = h.api.applyFilters("account", h.folder.URI);
    assert.equal(result.success, true, result.error);
    assert.deepEqual(Array.from(result.submitted), ["Safe"]);
    assert.deepEqual(Array.from(result.skipped, (entry) => ({ ...entry })), [
      { name: "Disabled", reason: "disabled" }, { name: "Outgoing", reason: "non-manual" },
    ]);
  });
});

describe("filter type values are validated before native mutation", () => {
  for (const value of [0, -1, 17.5, 2147483648, 4294967295, 4294967296, 4294967313, Number.MAX_SAFE_INTEGER]) {
    it(`rejects out-of-range type ${value} in create/update without allocation, mutation or save`, () => {
      for (const type of [value, String(value)]) {
        const h = makeFilterHarness();
        const original = h.seed({ enabled: false });
        const before = h.snapshot();
        h.filterList.createFilter = () => assert.fail("unexpected native filter allocation");
        const created = h.api.createFilter("account", "Overflow", true, type,
          [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type: "markRead" }]);
        assert.match(created.error, /type must be a positive signed 32-bit integer/);
        const updated = h.api.updateFilter("account", 0, "Renamed", true, type);
        assert.match(updated.error, /type must be a positive signed 32-bit integer/);
        assert.equal(h.snapshot(), before);
        assert.equal(h.filterList.filters[0], original);
        assert.equal(h.filterList.saveAttempts, 0);
        assert.deepEqual(h.filterList.mutations, []);
      }
    });
  }

  it("rejects unknown bits within the signed native range before mutation or save", () => {
    for (const value of [512, 529, 1073741824, 2147483647]) {
      for (const type of [value, String(value)]) {
        const h = makeFilterHarness();
        const original = h.seed();
        const before = h.snapshot();
        h.filterList.createFilter = () => assert.fail("unexpected native filter allocation");
        const created = h.api.createFilter("account", "Unknown type", true, type,
          [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type: "markRead" }]);
        assert.match(created.error, /unknown or unavailable nsMsgFilterType bits/);
        const updated = h.api.updateFilter("account", 0, "Renamed", false, type);
        assert.match(updated.error, /unknown or unavailable nsMsgFilterType bits/);
        assert.equal(h.snapshot(), before);
        assert.equal(h.filterList.filters[0], original);
        assert.equal(h.filterList.saveAttempts, 0);
        assert.deepEqual(h.filterList.mutations, []);
      }
    }
  });

  it("accepts every known flag and combinations, including flags omitted from All", () => {
    for (const value of [...Object.values(FILTER_TYPES).filter(Boolean), 17, 511]) {
      for (const type of [value, String(value)]) {
        const h = makeFilterHarness();
        const created = h.api.createFilter("account", "Known type", true, type,
          [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type: "markRead" }]);
        assert.equal(created.success, true, created.error);
        assert.equal(h.filterList.filters[0].filterType, value);
        const updated = h.api.updateFilter("account", 0, undefined, undefined, type);
        assert.equal(updated.success, true, updated.error);
        assert.equal(h.filterList.filters[0].filterType, value);
      }
    }
  });

  it("uses named runtime flags and rejects flags absent from this Thunderbird", () => {
    for (const periodic of [512, undefined]) {
      const h = makeFilterHarness({ ci: makeCi({ filterTypes: { Periodic: periodic } }) });
      const original = h.seed();
      const before = h.snapshot();
      const created = h.api.createFilter("account", "Unavailable type", true, 256,
        [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type: "markRead" }]);
      assert.match(created.error, /unknown or unavailable nsMsgFilterType bits/);
      assert.match(h.api.updateFilter("account", 0, undefined, undefined, 256).error,
        /unknown or unavailable nsMsgFilterType bits/);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.filters[0], original);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.deepEqual(h.filterList.mutations, []);
      if (periodic !== undefined) {
        const result = h.api.updateFilter("account", 0, undefined, undefined, periodic);
        assert.equal(result.success, true, result.error);
        assert.equal(h.filterList.filters[0].filterType, periodic);
      }
    }
  });

  it("keeps the default type and retains the existing type when omitted", () => {
    const h = makeFilterHarness();
    assert.equal(h.api.createFilter("account", "Default", true, undefined,
      [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type: "markRead" }]).success, true);
    assert.equal(h.filterList.filters[0].filterType, 17);
    assert.equal(h.api.updateFilter("account", 0, "Renamed").success, true);
    assert.equal(h.filterList.filters[0].filterType, 17);
  });
});

describe("filter writes preserve the live rule when validation or saving fails", () => {
  it("does not rename, enable or change type before replacement values validate", () => {
    const h = makeFilterHarness();
    const original = h.seed({ enabled: false });
    const before = h.snapshot();
    const result = h.api.updateFilter("account", 0, "Renamed", true, h.ci.nsMsgFilterType.PostOutgoing,
      undefined, [{ type: "changePriority", value: "99" }]);
    assert.ok(result.error);
    assert.equal(h.snapshot(), before);
    assert.equal(h.filterList.filters[0], original);
    assert.equal(h.filterList.saveAttempts, 0);
    assert.deepEqual(h.filterList.mutations, []);
  });

  it("leaves the live filter unchanged when an existing typed value cannot be copied", () => {
    const h = makeFilterHarness();
    const original = h.seed({ conditions: [{ attrib: "ageInDays", op: "isGreaterThan", value: "30" }] });
    original.searchTerms[0].value = { attrib: h.ci.nsMsgSearchAttrib.AgeInDays, get age() { throw new Error("value unavailable"); } };
    const before = h.snapshot();
    const result = h.api.updateFilter("account", 0, "Renamed", false, undefined, undefined, [{ type: "markFlagged" }]);
    assert.ok(result.error);
    assert.equal(h.snapshot(), before);
    assert.equal(h.filterList.filters[0], original);
    assert.equal(h.filterList.saveAttempts, 0);
    assert.deepEqual(h.filterList.mutations, []);
  });

  for (const operation of ["create", "update", "delete", "reorder"]) {
    it(`restores the original list after ${operation} cannot be saved`, () => {
      const h = makeFilterHarness();
      h.seed({ name: "First" });
      h.seed({ name: "Second" });
      const originals = [...h.filterList.filters];
      const before = h.snapshot();
      h.filterList.saveError = true;
      let result;
      if (operation === "create") {
        result = h.api.createFilter("account", "New", true, undefined,
          [{ attrib: "subject", op: "contains", value: "invoice" }], [{ type: "markRead" }], 0);
      } else if (operation === "update") {
        result = h.api.updateFilter("account", 0, "Renamed", false);
      } else if (operation === "delete") {
        result = h.api.deleteFilter("account", 0);
      } else {
        result = h.api.reorderFilters("account", 0, 1);
      }
      assert.match(result.error, /filter save failed/);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.filters.length, originals.length);
      originals.forEach((filter, index) => assert.equal(h.filterList.filters[index], filter));
      assert.equal(h.filterList.saveAttempts, 1);
    });
  }
});

describe("reorderFilters uses final destination indices", () => {
  for (const [from, to, names] of [
    [0, 0, ["A", "B", "C"]], [0, 1, ["B", "A", "C"]], [0, 2, ["B", "C", "A"]],
    [1, 0, ["B", "A", "C"]], [1, 1, ["A", "B", "C"]], [1, 2, ["A", "C", "B"]],
    [2, 0, ["C", "A", "B"]], [2, 1, ["A", "C", "B"]], [2, 2, ["A", "B", "C"]],
  ]) {
    it(`moves ${from} to ${to}`, () => {
      const h = makeFilterHarness();
      for (const name of ["A", "B", "C"]) h.seed({ name });
      const result = h.api.reorderFilters("account", from, to);
      assert.equal(result.success, true, result.error);
      assert.deepEqual(h.filterList.filters.map((filter) => filter.filterName), names);
    });
  }

  for (const [from, to] of [[-1, 0], [0, -1], [3, 0], [0, 3], [1.5, 0], [0, 1.5], ["bad", 0], [0, "bad"]]) {
    it(`rejects indices ${from}, ${to} without mutation`, () => {
      const h = makeFilterHarness();
      for (const name of ["A", "B", "C"]) h.seed({ name });
      const before = h.snapshot();
      const result = h.api.reorderFilters("account", from, to);
      assert.ok(result.error);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.deepEqual(h.filterList.mutations, []);
    });
  }
});

describe("Move/Copy filter actions never target the Outbox", () => {
  const OUTBOX_FLAG = 0x800; // nsMsgFolderFlags.Queue
  const conditions = [{ attrib: "subject", op: "contains", value: "invoice" }];
  const outbox = "mailbox://nobody@Local%20Folders/Unsent%20Messages";
  const nested = `${outbox}/Nested`;
  const legacyOutbox = "imap://account/LegacyOutbox";
  const archive = "imap://account/Archive";
  const getAccessibleFolder = (uri) => {
    if (uri === outbox) return { folder: { URI: uri, isSpecialFolder: (flag) => flag === OUTBOX_FLAG } };
    if (uri === nested) {
      return { folder: { URI: uri, isSpecialFolder: (flag, ancestors) => ancestors === true && flag === OUTBOX_FLAG } };
    }
    if (uri === legacyOutbox) return { folder: { URI: uri, getFlag: (flag) => flag === OUTBOX_FLAG } };
    return { folder: { URI: uri, isSpecialFolder: () => false } };
  };
  const outboxTargets = [outbox, nested, legacyOutbox];

  for (const type of ["moveToFolder", "copyToFolder"]) {
    for (const operation of ["create", "update"]) {
      it(`refuses an Outbox ${type} destination on ${operation} without changing the list`, () => {
        for (const value of outboxTargets) {
          const h = makeFilterHarness({ getAccessibleFolder });
          const original = h.seed();
          const before = h.snapshot();
          const actions = [{ type: "markRead" }, { type, value }];
          const result = operation === "create"
            ? h.api.createFilter("account", "New", true, undefined, conditions, actions)
            : h.api.updateFilter("account", 0, undefined, undefined, undefined, undefined, actions);
          assert.match(result.error, /Outbox \(Unsent Messages\) cannot be used as a move or copy destination/, value);
          assert.equal(h.snapshot(), before);
          assert.equal(h.filterList.filters[0], original);
          assert.equal(h.filterList.saveAttempts, 0);
          assert.deepEqual(h.filterList.mutations, []);
        }
      });
    }

    it(`refuses updates retaining an Outbox ${type} but permits disabling and repairing it`, () => {
      const h = makeFilterHarness({ getAccessibleFolder });
      h.seed({ actions: [{ type, value: outbox }] });
      const before = h.snapshot();
      const renamed = h.api.updateFilter("account", 0, "Renamed");
      assert.match(renamed.error, /Outbox/);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.saveAttempts, 0);
      const disabled = h.api.updateFilter("account", 0, undefined, false);
      assert.equal(disabled.success, true, disabled.error);
      const repaired = h.api.updateFilter("account", 0, undefined, true, undefined, undefined, [{ type, value: archive }]);
      assert.equal(repaired.success, true, repaired.error);
      assert.equal(h.filterList.filters[0].getActionAt(0).targetFolderUri, archive);
    });
  }

  it("keeps ordinary destinations working", () => {
    const h = makeFilterHarness({ getAccessibleFolder });
    const result = h.api.createFilter("account", "Archive", true, undefined, conditions, [{ type: "moveToFolder", value: archive }]);
    assert.equal(result.success, true, result.error);
  });

  it("skips existing rules that would place messages in the Outbox", () => {
    const h = makeFilterHarness({ getAccessibleFolder, allowSend: true });
    h.seed({ name: "Move to Outbox", actions: [{ type: "moveToFolder", value: outbox }] });
    h.seed({ name: "Copy to nested", actions: [{ type: "markRead" }, { type: "copyToFolder", value: nested }] });
    h.seed({ name: "Archive", actions: [{ type: "moveToFolder", value: archive }] });
    const before = h.snapshot();
    const result = h.api.applyFilters("account", h.folder.URI);
    assert.equal(result.success, true, result.error);
    assert.deepEqual(Array.from(result.submitted), ["Archive"]);
    assert.deepEqual(Array.from(result.skipped, (entry) => ({ ...entry })), [
      { name: "Move to Outbox", reason: "outbox-destination" },
      { name: "Copy to nested", reason: "outbox-destination" },
    ]);
    assert.deepEqual(h.submissions[0].filters.map((filter) => filter.filterName), ["Archive"]);
    assert.equal(h.snapshot(), before);
  });

  it("treats a folder whose role cannot be read as the Outbox", () => {
    const h = makeFilterHarness({
      getAccessibleFolder: (uri) => ({ folder: { URI: uri, isSpecialFolder() { throw new Error("unavailable"); } } }),
    });
    const result = h.api.createFilter("account", "New", true, undefined, conditions, [{ type: "moveToFolder", value: archive }]);
    assert.match(result.error, /Outbox/);
    assert.equal(h.filterList.saveAttempts, 0);
  });
});

describe("address book filter conditions follow address book access", () => {
  const addressBook = "jsaddrbook://abook.sqlite";
  const actions = [{ type: "markRead" }];
  const plain = [{ attrib: "subject", op: "contains", value: "invoice" }];
  const abConditions = (op) => [{ attrib: "subject", op: "contains", value: "invoice" }, { attrib: "from", op, value: addressBook }];

  for (const op of ["isInAB", "isntInAB"]) {
    it(`refuses ${op} conditions on create and update while address books are restricted`, () => {
      const h = makeFilterHarness({ addressBooksRestricted: true });
      const original = h.seed();
      const before = h.snapshot();
      const created = h.api.createFilter("account", "New", true, undefined, abConditions(op), actions);
      assert.match(created.error, /Address book not accessible/);
      const updated = h.api.updateFilter("account", 0, undefined, undefined, undefined, abConditions(op));
      assert.match(updated.error, /Address book not accessible/);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.filters[0], original);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.deepEqual(h.filterList.mutations, []);
    });

    it(`accepts ${op} conditions when address books are accessible`, () => {
      const h = makeFilterHarness({ addressBooksRestricted: false });
      const created = h.api.createFilter("account", "New", true, undefined, abConditions(op), actions);
      assert.equal(created.success, true, created.error);
      assert.equal(h.api.listFilters("account")[0].filters[0].terms[1].op, op);
    });

    it(`refuses updates retaining ${op} while restricted but permits disabling and replacing it`, () => {
      const h = makeFilterHarness({ addressBooksRestricted: true });
      h.seed({ conditions: abConditions(op) });
      const before = h.snapshot();
      assert.match(h.api.updateFilter("account", 0, "Renamed").error, /Address book not accessible/);
      assert.equal(h.snapshot(), before);
      assert.equal(h.filterList.saveAttempts, 0);
      assert.equal(h.api.updateFilter("account", 0, undefined, false).success, true);
      const replaced = h.api.updateFilter("account", 0, undefined, true, undefined, plain);
      assert.equal(replaced.success, true, replaced.error);
    });
  }

  it("skips existing address book rules only while address books are restricted", () => {
    for (const addressBooksRestricted of [true, false]) {
      const h = makeFilterHarness({ addressBooksRestricted });
      h.seed({ name: "Known senders", conditions: abConditions("isInAB") });
      h.seed({ name: "Unknown senders", conditions: abConditions("isntInAB") });
      h.seed({ name: "Plain", conditions: plain });
      const result = h.api.applyFilters("account", h.folder.URI);
      assert.equal(result.success, true, result.error);
      if (addressBooksRestricted) {
        assert.deepEqual(Array.from(result.submitted), ["Plain"]);
        assert.deepEqual(Array.from(result.skipped, (entry) => ({ ...entry })), [
          { name: "Known senders", reason: "inaccessible-address-book" },
          { name: "Unknown senders", reason: "inaccessible-address-book" },
        ]);
      } else {
        assert.deepEqual(Array.from(result.submitted), ["Known senders", "Unknown senders", "Plain"]);
        assert.deepEqual(Array.from(result.skipped), []);
      }
    }
  });
});
