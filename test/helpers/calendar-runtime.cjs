"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../../extension/mcp_server/api.js"), "utf8");

// Reuse the production-marker/VM pattern from privacy-access.test.cjs. Only
// Thunderbird objects are mocked; handlers, schemas and preference gates run
// from api.js rather than from copies of their implementation.
function snippet(startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0, `api.js marker missing: ${startMarker}`);
  assert.ok(end > start, `api.js marker missing: ${endMarker}`);
  return source.slice(start, end);
}

function constant(name) {
  const match = source.match(new RegExp(`^\\s*const ${name} = [^\\n]+;`, "m"));
  assert.ok(match, `api.js constant missing: ${name}`);
  return match[0];
}

class CalendarDuration {
  constructor(seconds) { this.inSeconds = seconds; }
  get isNegative() { return this.inSeconds < 0; }
  set isNegative(value) { this.inSeconds = Math.abs(this.inSeconds) * (value ? -1 : 1); }
  clone() { return new CalendarDuration(this.inSeconds); }
}

class CalendarDate {
  constructor(value = "2026-09-28T10:00:00Z", timezone = { tzid: "UTC" }) {
    this.nativeTime = new Date(value).getTime() * 1000;
    this.timezone = timezone;
    this.isDate = false;
  }

  resetTo(year, month, day, hour, minute, second, timezone) {
    this.nativeTime = Date.UTC(year, month, day, hour, minute, second) * 1000;
    this.timezone = timezone;
  }

  compare(other) { return Math.sign(this.nativeTime - other.nativeTime); }
  clone() { return Object.assign(new CalendarDate(), this); }
  subtractDate(other) { return new CalendarDuration((this.nativeTime - other.nativeTime) / 1000000); }
  addDuration(duration) { this.nativeTime += duration.inSeconds * 1000000; }
  get year() { return new Date(this.nativeTime / 1000).getUTCFullYear(); }
  get month() { return new Date(this.nativeTime / 1000).getUTCMonth(); }
  get day() { return new Date(this.nativeTime / 1000).getUTCDate(); }
  set day(day) {
    const date = new Date(this.nativeTime / 1000);
    date.setUTCDate(day);
    this.nativeTime = date.getTime() * 1000;
  }
  get second() { return new Date(this.nativeTime / 1000).getUTCSeconds(); }
  set second(second) {
    const date = new Date(this.nativeTime / 1000);
    date.setUTCSeconds(second);
    this.nativeTime = date.getTime() * 1000;
  }
}

class CalendarAttendee {
  constructor(fields = {}) { Object.assign(this, fields); }
  getProperty(name) { return this._properties?.get(name) ?? null; }
  setProperty(name, value) {
    this._properties ??= new Map();
    this._properties.set(name, value);
  }
  clone() {
    const clone = Object.assign(new CalendarAttendee(), this);
    if (this._properties) clone._properties = new Map(this._properties);
    return clone;
  }
}

class CalendarItem {
  constructor() {
    this.id = "event-1";
    this.title = "Existing event";
    this.properties = new Map();
    this.attendees = [];
    this.categories = [];
    this.recurrenceInfo = null;
  }

  get parentItem() { return this._parentItem || this; }
  set parentItem(item) { this._parentItem = item; }
  get recurrenceStartDate() { return this.startDate; }
  get duration() { return this.endDate.subtractDate(this.startDate); }
  getProperty(name) { return this.properties.get(name) ?? null; }
  setProperty(name, value) { this.properties.set(name, value); }
  deleteProperty(name) { this.properties.delete(name); }
  getCategories() { return this.categories.slice(); }
  setCategories(categories) { this.categories = Array.from(categories); }
  getAttendees() { return this.attendees.slice(); }
  addAttendee(attendee) { this.attendees.push(attendee); }
  removeAttendee(attendee) { this.attendees = this.attendees.filter(value => value !== attendee); }

  clone() {
    const clone = Object.assign(new CalendarItem(), this);
    clone.properties = new Map(this.properties);
    clone.attendees = this.attendees.slice();
    clone.categories = this.categories.slice();
    for (const name of ["startDate", "endDate", "dueDate", "recurrenceId"]) {
      if (this[name]) clone[name] = this[name].clone();
    }
    if (this.recurrenceInfo) {
      clone.recurrenceInfo = this.recurrenceInfo.clone();
      clone.recurrenceInfo.item = clone;
    }
    return clone;
  }
}

class RecurrenceInfo {
  constructor() {
    this.rules = [];
    this.occurrences = [];
    this.excluded = [];
    this.lookups = [];
    this.exceptions = [];
    this.generationCalls = [];
    this.ruleCalls = [];
  }

  appendRecurrenceItem(rule) { this.rules.push(rule); }
  getRecurrenceItems() { return this.rules.slice(); }
  setRecurrenceItems(items) { this.rules = Array.from(items); this.materialized = true; }
  getExceptionIds() { return this.exceptions.map(item => item.recurrenceId); }
  getExceptionFor(date) {
    return this.exceptions.find(item => item.recurrenceId.compare(date) === 0) || null;
  }
  getOccurrenceFor(date) {
    this.lookups.push(date);
    const existing = this.occurrences.find(item => item.recurrenceId.compare(date) === 0);
    if (existing) return existing;
    // Thunderbird creates a proxy for arbitrary dates here, even when a date
    // is not in the rule or has been excluded. This is not a membership test.
    const proxy = this.item.clone();
    proxy.recurrenceInfo = null;
    proxy.parentItem = this.item;
    proxy.recurrenceId = date.clone();
    proxy.startDate = date.clone();
    proxy.endDate = date.clone();
    proxy.endDate.nativeTime += this.item.endDate.nativeTime - this.item.startDate.nativeTime;
    return proxy;
  }
  getOccurrences(start, end, maxCount = 0) {
    const call = { start, end, maxCount, generated: 0 };
    this.generationCalls.push(call);
    if (this.dense && !this.materialized) {
      // RecurrenceInfo drops maxCount when forwarding a bounded date range to
      // RRULEs. Guard the mock rather than materializing billions of items.
      throw new Error("Unbounded native recurrenceInfo RRULE generation");
    }
    const candidates = this.materialized
      ? this.rules.filter(rule => rule.date && !rule.isNegative).map(rule => {
        const existing = this.getExceptionFor(rule.date)
          || this.occurrences.find(item => item.recurrenceId.compare(rule.date) === 0);
        if (existing) return existing;
        const occurrence = this.item.clone();
        occurrence.recurrenceInfo = null;
        occurrence.parentItem = this.item;
        const offset = rule.date.nativeTime - occurrence.startDate.nativeTime;
        occurrence.startDate.nativeTime += offset;
        occurrence.endDate.nativeTime += offset;
        occurrence.recurrenceId = rule.date.clone();
        return occurrence;
      }).concat(this.exceptions, this.getOccurrenceFor(this.item.startDate))
      : this.occurrences;
    const excludedDates = this.excluded.concat(this.rules.filter(rule => rule.date && rule.isNegative).map(rule => rule.date));
    const result = [];
    const seen = new Set();
    for (const item of candidates.slice().sort((a, b) => a.startDate.compare(b.startDate))) {
      if (item.endDate.compare(start) <= 0 || (end && item.startDate.compare(end) >= 0)
          || excludedDates.some(date => item.recurrenceId.compare(date) === 0)
          || seen.has(item.recurrenceId.nativeTime)) continue;
      call.generated++;
      result.push(item);
      seen.add(item.recurrenceId.nativeTime);
      if (maxCount > 0 && result.length >= maxCount) break;
    }
    return result;
  }
  removeOccurrenceAt(date) { this.excluded.push(date); }
  clone() {
    return Object.assign(new RecurrenceInfo(), this, {
      rules: this.rules.slice(), occurrences: this.occurrences.slice(),
      excluded: this.excluded.slice(), lookups: this.lookups.slice(),
      exceptions: this.exceptions.slice(),
    });
  }
}

function makeRecurrenceRule(info, fields) {
  return {
    isFinite: false, isNegative: false, ...fields,
    getOccurrences(dtstart, start, end, maxCount) {
      const call = { dtstart, start, end, maxCount, generated: 0 };
      info.ruleCalls.push(call);
      if (!end || !Number.isInteger(maxCount) || maxCount <= 0) {
        throw new Error("Unbounded native RRULE generation");
      }
      const result = [];
      if (info.dense) {
        const step = info.dense.intervalSeconds * 1000000;
        const firstIndex = Math.max(0, Math.ceil((start.nativeTime - dtstart.nativeTime) / step));
        for (let index = firstIndex; index < info.dense.total && result.length < maxCount; index++) {
          const date = dtstart.clone();
          date.nativeTime += index * step;
          if (date.compare(end) >= 0) break;
          result.push(date);
          call.generated++;
        }
      } else {
        for (const occurrence of info.occurrences) {
          const date = occurrence.recurrenceId;
          if (date.compare(start) < 0 || date.compare(end) >= 0) continue;
          result.push(date.clone());
          call.generated++;
          if (result.length >= maxCount) break;
        }
      }
      return result;
    },
  };
}

function plain(value) { return JSON.parse(JSON.stringify(value)); }

function loadCalendarRuntime({
  blockSkipReview, unreadableSkipReview = false, allowedAccounts,
  allowAllCalendars = false, parser = "cal", parserFailure,
  identity = { email: "organizer@example.com", fullName: "Organizer" },
  identityKey = "identity-1", unreadableIdentity = false, resolvedIdentity, organizerId,
} = {}) {
  const prefValues = new Map();
  const prefTypes = { PREF_INVALID: 0, PREF_STRING: 32, PREF_BOOL: 128 };
  const calls = { reads: [], adds: [], modifies: [], deletes: [], dialogs: [], calendarProperties: [], nativeRules: [] };
  const items = new Map();
  const preferences = {
    getPrefType(name) {
      if (!prefValues.has(name)) return prefTypes.PREF_INVALID;
      return typeof prefValues.get(name) === "string" ? prefTypes.PREF_STRING : prefTypes.PREF_BOOL;
    },
    getStringPref(name, fallback = "") { return prefValues.get(name) ?? fallback; },
    getBoolPref(name, fallback) {
      if (unreadableSkipReview && name.endsWith(".blockSkipReview")) throw new Error("Unreadable preference");
      return prefValues.has(name) ? prefValues.get(name) : fallback;
    },
    getIntPref(_name, fallback) { return fallback; },
  };
  const calendarProperties = new Map([["imip.identity.key", identityKey]]);
  if (resolvedIdentity !== undefined) calendarProperties.set("imip.identity", resolvedIdentity);
  if (organizerId !== undefined) calendarProperties.set("organizerId", organizerId);
  const calendar = {
    id: "calendar-1", name: "Test calendar", type: "storage", readOnly: false,
    getProperty(name) { return calendarProperties.get(name) ?? null; },
    setProperty(name, value) {
      calls.calendarProperties.push({ name, value });
      calendarProperties.set(name, value);
    },
    async getItem(id) { calls.reads.push({ kind: "item", id }); return items.get(id) || null; },
    async getItemsAsArray(_filter, _count, start, end) {
      calls.reads.push({ kind: "range", start, end });
      return [...items.values()].filter(item => !start || !item.startDate
        || (item.startDate.compare(start) >= 0 && item.startDate.compare(end) < 0));
    },
    async addItem(item) { calls.adds.push(item); return item; },
    async modifyItem(item, previous) { calls.modifies.push({ item, previous }); return item; },
    async deleteItem(item) { calls.deletes.push(item); },
  };
  const window = {
    openDialog(_uri, _name, _features, args) { calls.dialogs.push({ kind: "event", ...args }); },
    createTodoWithDialog(targetCalendar, dueDate, _unused, todo) {
      calls.dialogs.push({ kind: "task", calendar: targetCalendar, dueDate, todo });
    },
  };
  const timezone = { tzid: "Europe/Warsaw" };
  const floating = { tzid: "floating" };
  function nativeRule(line, implementation) {
    calls.nativeRules.push({ line, implementation });
    if (parserFailure === "throw") throw new Error("Native recurrence parser refused rule");
    return {
      icalString: line,
      type: parserFailure === "missing-type" ? "" : /(?:^|[:;])FREQ=([^;]+)/i.exec(line)?.[1].toUpperCase(),
    };
  }
  const cal = {
    manager: { getCalendars() { calls.reads.push({ kind: "calendars" }); return [calendar]; } },
    dtz: {
      defaultTimezone: timezone, floating,
      jsDateToDateTime(date, tz) { return new CalendarDate(date, tz); },
    },
    createDateTime() { return new CalendarDate(); },
    createRecurrenceDate() { return { date: null, isNegative: false, isFinite: true }; },
    category: { fromPrefs: () => ["Work", "Personal"] },
    itip: { getInvitedAttendee: item => item.invitedAttendee || null },
  };
  if (parser === "cal") cal.createRecurrenceRule = line => nativeRule(line, "cal");
  const Cc = {
    "@mozilla.org/calendar/recurrence-info;1": { createInstance: () => new RecurrenceInfo() },
    "@mozilla.org/calendar/recurrence-rule;1": {
      createInstance() {
        return {
          set icalString(line) {
            const parsed = nativeRule(line, "xpcom");
            this.type = parsed.type;
            this._line = parsed.icalString;
          },
          get icalString() { return this._line; },
        };
      },
    },
  };
  const sandbox = {
    cal, CalEvent: CalendarItem, CalTodo: CalendarItem, CalAttendee: CalendarAttendee, Cc,
    Ci: { nsIPrefBranch: prefTypes },
    Services: { prefs: preferences, wm: { getMostRecentWindow: () => window } },
    MailServices: { accounts: { getIdentity() {
      if (unreadableIdentity) throw new Error("Unreadable calendar identity");
      return identity;
    } } },
    console: { warn() {}, error() {} },
  };
  const names = [
    "PREF_ALLOWED_ACCOUNTS", "PREF_DISABLED_TOOLS", "PREF_BLOCK_SKIPREVIEW",
    "PREF_ALLOW_ALL_CALENDARS", "PREF_ALLOW_ALL_ADDRESS_BOOKS", "PREF_GET_MESSAGES_LIMIT",
    "DEFAULT_GET_MESSAGES_LIMIT", "MAX_GET_MESSAGES_LIMIT", "UNDISABLEABLE_TOOLS",
  ];
  vm.createContext(sandbox);
  vm.runInContext([
    ...names.map(constant),
    snippet("// BEGIN CONTACT FIELD CONSTANTS", "// END CONTACT FIELD CONSTANTS"),
    snippet("// BEGIN OUTBOUND ATTACHMENT LIMITS", "// END OUTBOUND ATTACHMENT LIMITS"),
    snippet("function normalizeGetMessagesLimit(", "// BEGIN TOOL SCHEMA BUILDER"),
    snippet("// BEGIN FILTER SEARCH TERM HELPERS", "// END FILTER SEARCH TERM HELPERS"),
    snippet("// BEGIN PRIVACY PREFERENCE HELPERS", "// END PRIVACY PREFERENCE HELPERS"),
    snippet("// BEGIN SERVER ACCESS HELPERS", "// END SERVER ACCESS HELPERS"),
    snippet("// BEGIN MCP TEXT SANITIZATION", "// END MCP TEXT SANITIZATION"),
    snippet("// BEGIN TOOL SCHEMA BUILDER", "// END TOOL SCHEMA BUILDER"),
    snippet("// BEGIN TOOL SCHEMA VALIDATOR", "// END TOOL SCHEMA VALIDATOR"),
    snippet("// BEGIN CALENDAR TOOLS", "// END CALENDAR TOOLS"),
    snippet("// BEGIN TOOL CALL DISPATCH", "// END TOOL CALL DISPATCH"),
    "this.prefNames = { PREF_ALLOWED_ACCOUNTS, PREF_BLOCK_SKIPREVIEW, PREF_ALLOW_ALL_CALENDARS };",
    "Object.assign(this, { callTool, validateToolArgs, coerceToolArgs, buildTools, normalizeRRule, extractRRuleFromItem, formatEvent, formatAttendee, formatOrganizer });",
  ].join("\n"), sandbox);
  if (blockSkipReview !== undefined) prefValues.set(sandbox.prefNames.PREF_BLOCK_SKIPREVIEW, blockSkipReview);
  if (allowedAccounts !== undefined) {
    prefValues.set(sandbox.prefNames.PREF_ALLOWED_ACCOUNTS,
      typeof allowedAccounts === "string" ? allowedAccounts : JSON.stringify(allowedAccounts));
  }
  prefValues.set(sandbox.prefNames.PREF_ALLOW_ALL_CALENDARS, allowAllCalendars);

  function date(value, allDay = false) {
    const result = new CalendarDate(value, allDay ? floating : timezone);
    result.isDate = allDay;
    return result;
  }
  function makeEvent({ start = "2026-09-28T10:00:00Z", end = "2026-09-28T11:00:00Z", allDay = false, ...fields } = {}) {
    return Object.assign(new CalendarItem(), { calendar, startDate: date(start, allDay), endDate: date(end, allDay) }, fields);
  }
  function seedEvent(fields) {
    const item = makeEvent(fields);
    items.set(item.id, item);
    return item;
  }
  function seedSeries({ allDay = false, ...fields } = {}) {
    const master = seedEvent({
      start: allDay ? "2026-01-05" : "2026-01-05T10:00:00Z",
      end: allDay ? "2026-01-06" : "2026-01-05T11:00:00Z", allDay, ...fields,
    });
    const recurrenceInfo = new RecurrenceInfo();
    recurrenceInfo.item = master;
    recurrenceInfo.rules.push(makeRecurrenceRule(recurrenceInfo, { icalString: "RRULE:FREQ=WEEKLY\r\n", type: "WEEKLY" }));
    master.recurrenceInfo = recurrenceInfo;
    for (const day of ["2026-09-28", "2026-10-05"]) {
      const occurrence = makeEvent({
        id: master.id, allDay,
        start: allDay ? day : `${day}T10:00:00Z`,
        end: allDay ? `${day}T00:00:00Z` : `${day}T11:00:00Z`,
        attendees: master.attendees.slice(), organizer: master.organizer,
      });
      if (allDay) occurrence.endDate.nativeTime += 86400000 * 1000;
      occurrence.parentItem = master;
      occurrence.recurrenceId = occurrence.startDate.clone();
      recurrenceInfo.occurrences.push(occurrence);
    }
    return { master, occurrences: recurrenceInfo.occurrences };
  }
  function seedDenseSeries({ total = 1000000000, intervalSeconds = 1, ...fields } = {}) {
    const master = seedEvent({
      start: "2026-09-28T10:00:00Z", end: "2026-09-28T10:00:01Z", ...fields,
    });
    master.recurrenceInfo = new RecurrenceInfo();
    master.recurrenceInfo.item = master;
    master.recurrenceInfo.rules.push(makeRecurrenceRule(master.recurrenceInfo, { icalString: "RRULE:FREQ=SECONDLY", type: "SECONDLY" }));
    master.recurrenceInfo.dense = { total, intervalSeconds };
    return master;
  }
  return {
    ...sandbox, calls, items, calendar, prefValues, date, makeEvent, seedEvent, seedSeries, seedDenseSeries,
    makeAttendee: fields => new CalendarAttendee(fields),
    async invoke(name, args) {
      sandbox.coerceToolArgs(name, args);
      const errors = sandbox.validateToolArgs(name, args);
      if (errors.length) return { error: errors.join("; ") };
      return sandbox.callTool(name, args);
    },
  };
}

module.exports = { loadCalendarRuntime, plain };
