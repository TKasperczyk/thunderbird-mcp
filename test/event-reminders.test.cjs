"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { loadCalendarRuntime } = require("./helpers/calendar-runtime.cjs");

// Event `reminders` (whole minutes before the start, one display alarm each):
// set on createEvent, replaced on updateEvent, read back by listEvents. The
// helpers (normalizeReminders, buildReminderAlarm, normalizeReminderUpdate,
// applyReminderChanges, reminderMinutesOf) run from production api.js under
// node:vm with Thunderbird classes stubbed, so these tests cannot drift from the
// real code; the tool schemas and the createEvent/updateEvent/listEvents wiring
// are checked as source contracts.
const api = fs.readFileSync(
  path.resolve(__dirname, "../extension/mcp_server/api.js"),
  "utf8"
);

function extract(startMarker, endMarker) {
  const start = api.indexOf(startMarker);
  const end = api.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0, `api.js marker missing: ${startMarker}`);
  assert.ok(end > start, `api.js marker missing: ${endMarker}`);
  return api.slice(start, end);
}

function loadNormalizeReminders() {
  const source = extract("const MAX_REMINDERS =", "function buildReminderAlarm(");
  const sandbox = {};
  vm.runInNewContext(`${source}\nthis.normalizeReminders = normalizeReminders;`, sandbox);
  // Results are built in the VM realm; copy them into this realm so strict
  // deep-equality does not fail on differing Array/Object prototypes.
  return (input) => JSON.parse(JSON.stringify(sandbox.normalizeReminders(input)));
}

describe("normalizeReminders", () => {
  const normalizeReminders = loadNormalizeReminders();

  it("treats omitted or null as no reminders", () => {
    assert.deepEqual(normalizeReminders(undefined), { minutes: [] });
    assert.deepEqual(normalizeReminders(null), { minutes: [] });
    assert.deepEqual(normalizeReminders([]), { minutes: [] });
  });

  it("accepts minutes before start, including 0 (at event time)", () => {
    assert.deepEqual(normalizeReminders([0, 1440]).minutes, [1440, 0]);
  });

  it("de-duplicates and orders largest first", () => {
    assert.deepEqual(normalizeReminders([0, 4320, 1440, 1440, 0]).minutes, [4320, 1440, 0]);
  });

  it("rejects a non-array", () => {
    assert.match(normalizeReminders("1440").error, /array/);
    assert.match(normalizeReminders(1440).error, /array/);
  });

  it("rejects non-integer, negative, string and NaN entries, naming the index", () => {
    assert.match(normalizeReminders([30, 1.5]).error, /reminders\[1\]/);
    assert.match(normalizeReminders([-1]).error, /reminders\[0\]/);
    assert.match(normalizeReminders(["30"]).error, /reminders\[0\]/);
    assert.match(normalizeReminders([NaN]).error, /reminders\[0\]/);
  });

  it("caps the count and the lead time", () => {
    assert.match(normalizeReminders(Array.from({ length: 11 }, (_, i) => i)).error, /at most 10/);
    assert.match(normalizeReminders([40321]).error, /40320/);
    assert.deepEqual(normalizeReminders([40320]).minutes, [40320]);
  });
});

describe("createEvent reminders wiring (source contract)", () => {
  const createEventSource = extract("async function createEvent(", "function normalizeEventStatus(");

  it("declares the reminders tool parameter as an array of integers", () => {
    const i = api.indexOf('name: "createEvent"');
    const j = api.indexOf('name: "listEvents"', i);
    assert.match(api.slice(i, j), /reminders:\s*\{\s*type:\s*"array",\s*items:\s*\{\s*type:\s*"integer"/);
  });

  it("passes args.reminders from the dispatcher into createEvent", () => {
    assert.match(api, /return await createEvent\([^)]*args\.reminders\)/);
  });

  it("validates reminders and adds one alarm per entry", () => {
    assert.match(createEventSource, /normalizeReminders\(reminders\)/);
    assert.match(createEventSource, /event\.addAlarm\(/);
  });

});

describe("buildReminderAlarm (production code, Thunderbird classes stubbed)", () => {
  const START = 0;
  class StubAlarm {}
  const sandbox = {
    CalAlarm: StubAlarm,
    Ci: { calIAlarm: { ALARM_RELATED_START: START } },
    cal: { createDuration: () => ({ inSeconds: 0 }) },
  };
  vm.runInNewContext(
    `${extract("function buildReminderAlarm(", "function normalizeEventStatus(")}\nthis.buildReminderAlarm = buildReminderAlarm;`,
    sandbox
  );

  it("builds a DISPLAY alarm titled after the event, relative to the start", () => {
    const alarm = sandbox.buildReminderAlarm("Team standup", 1440);
    assert.ok(alarm instanceof StubAlarm);
    assert.equal(alarm.action, "DISPLAY");
    assert.equal(alarm.description, "Team standup");
    assert.equal(alarm.related, START);
  });

  it("offsets the trigger before the start by the requested minutes", () => {
    assert.equal(sandbox.buildReminderAlarm("x", 1440).offset.inSeconds, -86400);
    assert.equal(sandbox.buildReminderAlarm("x", 30).offset.inSeconds, -1800);
  });

  it("fires at the start for 0 minutes", () => {
    assert.equal(sandbox.buildReminderAlarm("x", 0).offset.inSeconds, 0);
  });
});

// calIAlarm.ALARM_RELATED_* values, as Thunderbird defines them.
const RELATED = { ABSOLUTE: 0, START: 1, END: 2 };

function loadReminderHelpers() {
  class StubAlarm {}
  const sandbox = {
    CalAlarm: StubAlarm,
    Ci: { calIAlarm: { ALARM_RELATED_ABSOLUTE: RELATED.ABSOLUTE, ALARM_RELATED_START: RELATED.START, ALARM_RELATED_END: RELATED.END } },
    cal: { createDuration: () => ({ inSeconds: 0 }) },
  };
  vm.runInNewContext(
    `${extract("const MAX_REMINDERS =", "function normalizeEventStatus(")}
     this.normalizeReminderUpdate = normalizeReminderUpdate;
     this.applyReminderChanges = applyReminderChanges;
     this.reminderMinutesOf = reminderMinutesOf;`,
    sandbox
  );
  return sandbox;
}

// A stand-in event item that records the alarm calls made on it.
function stubItem(title, alarms = []) {
  return {
    title,
    alarms: [...alarms],
    clearCalls: 0,
    clearAlarms() { this.clearCalls += 1; this.alarms = []; },
    addAlarm(alarm) { this.alarms.push(alarm); },
    getAlarms() { return [...this.alarms]; },
  };
}

describe("normalizeReminderUpdate (updateEvent: omitted/null keeps, array replaces, [] clears)", () => {
  const { normalizeReminderUpdate } = loadReminderHelpers();
  const run = (input) => JSON.parse(JSON.stringify(normalizeReminderUpdate(input)));

  it("leaves existing reminders alone when omitted or null", () => {
    assert.deepEqual(run(undefined), { minutes: null });
    assert.deepEqual(run(null), { minutes: null });
  });

  it("treats an empty array as an explicit request to remove every reminder", () => {
    assert.deepEqual(run([]), { minutes: [] });
  });

  it("uses the same rules as createEvent for a non-empty list", () => {
    assert.deepEqual(run([30, 1440, 30]), { minutes: [1440, 30] });
    assert.match(run("1440").error, /array/);
    assert.match(run([40321]).error, /40320/);
    assert.match(run([1.5]).error, /reminders\[0\]/);
  });
});

describe("applyReminderChanges (production code, item stubbed)", () => {
  const { applyReminderChanges } = loadReminderHelpers();

  it("does nothing when reminders were not supplied", () => {
    const item = stubItem("Review", ["existing"]);
    const changes = ["title"];
    applyReminderChanges(item, null, changes);
    assert.equal(item.clearCalls, 0);
    assert.deepEqual(item.alarms, ["existing"]);
    assert.deepEqual([...changes], ["title"]);
  });

  it("replaces every existing alarm with the requested ones, titled after the event", () => {
    const item = stubItem("Review", ["old-1", "old-2"]);
    const changes = ["title"];
    applyReminderChanges(item, [1440, 0], changes);
    assert.equal(item.clearCalls, 1);
    assert.equal(item.alarms.length, 2);
    assert.deepEqual(item.alarms.map((a) => a.offset.inSeconds), [-86400, 0]);
    assert.deepEqual(item.alarms.map((a) => a.description), ["Review", "Review"]);
    assert.deepEqual([...changes], ["title", "reminders"]);
  });

  it("removes every alarm for an empty list", () => {
    const item = stubItem("Review", ["old-1"]);
    const changes = [];
    applyReminderChanges(item, [], changes);
    assert.equal(item.clearCalls, 1);
    assert.deepEqual(item.alarms, []);
    assert.deepEqual([...changes], ["reminders"]);
  });
});

describe("reminderMinutesOf (what listEvents reports; production code, alarms stubbed)", () => {
  const { reminderMinutesOf } = loadReminderHelpers();
  const startAlarm = (seconds) => ({ related: RELATED.START, offset: { inSeconds: seconds } });

  it("reports start-relative alarms as minutes before the start, largest first", () => {
    const item = stubItem("x", [startAlarm(0), startAlarm(-86400), startAlarm(-1800)]);
    assert.deepEqual([...reminderMinutesOf(item)], [1440, 30, 0]);
  });

  it("reports nothing for an event with no alarms of its own", () => {
    assert.deepEqual([...reminderMinutesOf(stubItem("x"))], []);
  });

  it("drops duplicates and returns a plain 0, never -0", () => {
    const minutes = [...reminderMinutesOf(stubItem("x", [startAlarm(0), startAlarm(-0), startAlarm(-3600), startAlarm(-3600)]))];
    assert.deepEqual(minutes, [60, 0]);
    assert.ok(Object.is(minutes[1], 0));
  });

  it("skips alarms that cannot be expressed as minutes before the start", () => {
    const item = stubItem("x", [
      { related: RELATED.END, offset: { inSeconds: -600 } },
      { related: RELATED.ABSOLUTE, offset: null },
      startAlarm(300),
      startAlarm(-600),
    ]);
    assert.deepEqual([...reminderMinutesOf(item)], [10]);
  });
});

describe("updateEvent / listEvents reminders wiring (source contract)", () => {
  const updateEventSource = extract("async function updateEvent(", "async function deleteEvent(");

  it("declares reminders on the updateEvent tool as a replace-the-list array", () => {
    const i = api.indexOf('name: "updateEvent"');
    const j = api.indexOf('name: "deleteEvent"', i);
    const schema = api.slice(i, j);
    assert.match(schema, /reminders:\s*\{\s*type:\s*"array",\s*items:\s*\{\s*type:\s*"integer"/);
    assert.match(schema, /Replace the full list[^"]*\[\] removes all[^"]*omitted\/null (preserves|leaves)/);
  });

  it("passes args.reminders from the dispatcher into updateEvent", () => {
    assert.match(api, /return await updateEvent\([^)]*args\.reminders\)/);
  });

  it("validates before touching anything and applies on both the series and the occurrence path", () => {
    assert.match(updateEventSource, /normalizeReminderUpdate\(reminders\)/);
    assert.equal(updateEventSource.split("applyReminderChanges(").length - 1, 2);
  });

  it("returns reminders from listEvents and says so in the tool description", () => {
    assert.match(api, /reminders:\s*reminderMinutesOf\(item\)/);
    const i = api.indexOf('name: "listEvents"');
    const j = api.indexOf('name: "updateEvent"', i);
    assert.match(api.slice(i, j), /reminders/);
  });
});

// Reminder writes are event writes: the same Block skipReview guard that makes
// meetings with other attendees read-only must apply when only `reminders` is
// supplied, because calendar servers can email updates for any change.
describe("updateEvent reminders obey the Block skipReview attendee guard (production runtime)", () => {
  const updateArgs = { eventId: "event-1", calendarId: "calendar-1" };
  const recurrenceId = "2026-09-28T10:00:00.000Z";
  const otherAttendee = { id: "mailto:old@example.com", participationStatus: "ACCEPTED", role: "REQ-PARTICIPANT" };

  for (const [label, options] of [
    ["default", {}], ["enabled", { blockSkipReview: true }], ["unreadable", { unreadableSkipReview: true }],
  ]) {
    for (const occurrence of [false, true]) {
      for (const reminders of [[60], []]) {
        it(`rejects reminders ${JSON.stringify(reminders)} on ${occurrence ? "an occurrence" : "a series"} with other attendees while ${label}`, async () => {
          const runtime = loadCalendarRuntime(options);
          const { master } = runtime.seedSeries({ attendees: [otherAttendee] });
          const result = await runtime.invoke("updateEvent", {
            ...updateArgs, reminders, ...(occurrence ? { recurrenceId } : {}),
          });
          assert.match(result.error, /Block skipReview/i);
          assert.equal(runtime.calls.modifies.length, 0);
          assert.deepEqual(master.alarms, []);
        });
      }
    }
  }

  for (const occurrence of [false, true]) {
    it(`allows reminders on ${occurrence ? "an occurrence" : "a series"} with other attendees once the setting is off`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      runtime.seedSeries({ attendees: [otherAttendee] });
      const result = await runtime.invoke("updateEvent", {
        ...updateArgs, reminders: [60], ...(occurrence ? { recurrenceId } : {}),
      });
      assert.equal(result.success, true, result.error);
      assert.equal(runtime.calls.modifies.length, 1);
    });

    it(`allows reminders on ${occurrence ? "an occurrence" : "a series"} with no other attendees while the setting is on`, async () => {
      const runtime = loadCalendarRuntime();
      runtime.seedSeries({ attendees: [] });
      const result = await runtime.invoke("updateEvent", {
        ...updateArgs, reminders: [60], ...(occurrence ? { recurrenceId } : {}),
      });
      assert.equal(result.success, true, result.error);
      assert.equal(runtime.calls.modifies.length, 1);
    });
  }
});
