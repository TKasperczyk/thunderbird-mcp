"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { loadCalendarRuntime, plain } = require("./helpers/calendar-runtime.cjs");

const ids = { eventId: "event-1", calendarId: "calendar-1" };
const recurrenceId = "2026-09-28T10:00:00.000Z";

describe("Production occurrence updates and deletion", () => {
  for (const value of [undefined, null]) {
    it(`deletes a whole series when recurrenceId is ${String(value)}`, async () => {
      const runtime = loadCalendarRuntime();
      const { master } = runtime.seedSeries();
      const result = await runtime.invoke("deleteEvent", { ...ids, recurrenceId: value });
      assert.equal(result.success, true, result.error);
      assert.equal(result.mode, undefined);
      assert.match(result.warning, /entire series/i);
      assert.equal(runtime.calls.deletes[0], master);
      assert.equal(runtime.calls.modifies.length, 0);
      assert.equal(master.recurrenceInfo.lookups.length, 0);
    });

    it(`updates the master and permits a recurrence change when recurrenceId is ${String(value)}`, async () => {
      const runtime = loadCalendarRuntime();
      const { master } = runtime.seedSeries();
      const result = await runtime.invoke("updateEvent", {
        ...ids, recurrenceId: value, recurrence: "FREQ=DAILY;COUNT=3", title: "Changed series",
      });
      assert.equal(result.success, true, result.error);
      assert.equal(result.mode, undefined);
      const write = runtime.calls.modifies[0];
      assert.equal(write.previous, master);
      assert.equal(write.item.title, "Changed series");
      assert.equal(write.item.recurrenceInfo.rules[0].type, "DAILY");
      assert.equal(master.recurrenceInfo.lookups.length, 0);
    });
  }

  it("writes a modified occurrence rather than overwriting the master", async () => {
    const runtime = loadCalendarRuntime();
    const { master, occurrences } = runtime.seedSeries();
    const result = await runtime.invoke("updateEvent", {
      ...ids, recurrenceId, title: "Just this meeting", location: "Room 2",
      status: "tentative", categories: ["Work"],
    });
    assert.equal(result.success, true, result.error);
    assert.equal(result.mode, "occurrence");
    const write = runtime.calls.modifies[0];
    assert.equal(write.previous, occurrences[0]);
    assert.equal(write.item.parentItem, master);
    assert.equal(write.item.title, "Just this meeting");
    assert.equal(write.item.getProperty("LOCATION"), "Room 2");
    assert.equal(write.item.getProperty("STATUS"), "TENTATIVE");
    assert.deepEqual(write.item.getCategories(), ["Work"]);
    assert.equal(master.title, "Existing event");
    assert.equal(master.recurrenceInfo.lookups[0].timezone, master.startDate.timezone);
  });

  it("excludes a matching occurrence by modifying a clone of the series", async () => {
    const runtime = loadCalendarRuntime();
    const { master } = runtime.seedSeries();
    const result = await runtime.invoke("deleteEvent", { ...ids, recurrenceId });
    assert.equal(result.success, true, result.error);
    assert.equal(result.mode, "occurrence");
    assert.equal(result.recurrenceId, recurrenceId);
    const write = runtime.calls.modifies[0];
    assert.equal(write.previous, master);
    assert.notEqual(write.item, master);
    assert.equal(write.item.recurrenceInfo.excluded.length, 1);
    assert.equal(write.item.recurrenceInfo.excluded[0].timezone, master.startDate.timezone);
    assert.equal(master.recurrenceInfo.excluded.length, 0);
    assert.equal(runtime.calls.deletes.length, 0);
  });

  for (const name of ["updateEvent", "deleteEvent"]) {
    it(`${name} rejects a missing occurrence or invalid ID before persistence`, async () => {
      const runtime = loadCalendarRuntime();
      runtime.seedSeries();
      for (const id of ["2026-09-29T10:00:00Z", "not-a-date"]) {
        const result = await runtime.invoke(name, {
          ...ids, recurrenceId: id, ...(name === "updateEvent" ? { title: "Changed" } : {}),
        });
        assert.match(result.error, /occurrence|recurrenceId/i);
      }
      assert.equal(runtime.calls.modifies.length, 0);
      assert.equal(runtime.calls.deletes.length, 0);
    });

    it(`${name} rejects an already excluded occurrence even though getOccurrenceFor returns a proxy`, async () => {
      const runtime = loadCalendarRuntime();
      const { master } = runtime.seedSeries();
      master.recurrenceInfo.excluded.push(runtime.date(recurrenceId));
      const result = await runtime.invoke(name, {
        ...ids, recurrenceId, ...(name === "updateEvent" ? { title: "Changed" } : {}),
      });
      assert.match(result.error, /occurrence/i);
      assert.equal(runtime.calls.modifies.length, 0);
      assert.equal(runtime.calls.deletes.length, 0);
      assert.equal(master.recurrenceInfo.excluded.length, 1);
    });

    it(`${name} locates a moved exception by its original recurrence ID`, async () => {
      const runtime = loadCalendarRuntime();
      const { occurrences } = runtime.seedSeries();
      occurrences[0].startDate = runtime.date("2026-09-29T13:00:00Z");
      occurrences[0].endDate = runtime.date("2026-09-29T14:00:00Z");
      const result = await runtime.invoke(name, {
        ...ids, recurrenceId, ...(name === "updateEvent" ? { title: "Moved meeting" } : {}),
      });
      assert.equal(result.success, true, result.error);
      assert.equal(result.mode, "occurrence");
      const write = runtime.calls.modifies[0];
      if (name === "updateEvent") {
        assert.equal(write.previous, occurrences[0]);
        assert.equal(write.item.startDate.compare(runtime.date("2026-09-29T13:00:00Z")), 0);
        assert.equal(write.item.recurrenceId.compare(runtime.date(recurrenceId)), 0);
      } else {
        assert.equal(write.item.recurrenceInfo.excluded[0].compare(runtime.date(recurrenceId)), 0);
      }
      assert.equal(runtime.calls.deletes.length, 0);
    });

    it(`${name} rejects an off-schedule ID even when a different occurrence overlaps its time`, async () => {
      const runtime = loadCalendarRuntime();
      const { occurrences } = runtime.seedSeries();
      occurrences[0].endDate = runtime.date("2026-09-30T11:00:00Z");
      const result = await runtime.invoke(name, {
        ...ids, recurrenceId: "2026-09-29T10:00:00Z",
        ...(name === "updateEvent" ? { title: "Should not change" } : {}),
      });
      assert.match(result.error, /occurrence/i);
      assert.equal(runtime.calls.modifies.length, 0);
      assert.equal(runtime.calls.deletes.length, 0);
    });

    it(`${name} rejects recurrenceId on a non-recurring event`, async () => {
      const runtime = loadCalendarRuntime();
      runtime.seedEvent();
      const result = await runtime.invoke(name, {
        ...ids, recurrenceId, ...(name === "updateEvent" ? { title: "Changed" } : {}),
      });
      assert.match(result.error, /not recurring/i);
      assert.equal(runtime.calls.modifies.length, 0);
      assert.equal(runtime.calls.deletes.length, 0);
    });
  }

  it("rejects recurrence changes combined with a real recurrenceId", async () => {
    const runtime = loadCalendarRuntime();
    runtime.seedSeries();
    for (const recurrence of ["FREQ=DAILY", "", null]) {
      const result = await runtime.invoke("updateEvent", { ...ids, recurrenceId, recurrence });
      assert.match(result.error, /Cannot combine recurrence and recurrenceId/);
    }
    assert.equal(runtime.calls.modifies.length, 0);
  });

  it("keeps direct lookup fallback working when deleting a series with null recurrenceId", async () => {
    const runtime = loadCalendarRuntime();
    const { master } = runtime.seedSeries();
    runtime.calendar.getItem = undefined;
    const result = await runtime.invoke("deleteEvent", { ...ids, recurrenceId: null });
    assert.equal(result.success, true, result.error);
    assert.equal(runtime.calls.deletes[0], master);
    assert.ok(runtime.calls.reads.some(call => call.kind === "range"));
  });
});

describe("Calendar date compatibility uses production create/update handlers", () => {
  it("preserves previously accepted all-day RFC2822 dates in createEvent", async () => {
    const runtime = loadCalendarRuntime({ blockSkipReview: false });
    const startDate = "Mon, 28 Sep 2026 12:00:00 GMT";
    const endDate = "Tue, 29 Sep 2026 12:00:00 GMT";
    const result = await runtime.invoke("createEvent", {
      title: "All day", calendarId: "calendar-1", allDay: true, startDate, endDate, skipReview: true,
    });
    assert.equal(result.success, true, result.error);
    const event = runtime.calls.adds[0];
    const expectedStart = new Date(startDate);
    const expectedEnd = new Date(endDate);
    assert.equal(event.startDate.day, expectedStart.getDate());
    assert.equal(event.endDate.day, expectedEnd.getDate());
    assert.equal(event.startDate.isDate, true);
  });

  for (const occurrence of [false, true]) {
    it(`accepts legacy all-day dates on ${occurrence ? "occurrence" : "series"} updates`, async () => {
      const runtime = loadCalendarRuntime();
      runtime.seedSeries({ allDay: true });
      const startDate = "Mon, 28 Sep 2026 12:00:00 GMT";
      const endDate = "Tue, 29 Sep 2026 12:00:00 GMT";
      const result = await runtime.invoke("updateEvent", {
        ...ids, startDate, endDate,
        ...(occurrence ? { recurrenceId: "2026-09-28T00:00:00.000Z" } : {}),
      });
      assert.equal(result.success, true, result.error);
      const event = runtime.calls.modifies[0].item;
      const expected = new Date(startDate);
      const expectedEnd = new Date(endDate);
      expectedEnd.setDate(expectedEnd.getDate() + 1);
      assert.deepEqual([event.startDate.year, event.startDate.month, event.startDate.day],
        [expected.getFullYear(), expected.getMonth(), expected.getDate()]);
      assert.deepEqual([event.endDate.year, event.endDate.month, event.endDate.day],
        [expectedEnd.getFullYear(), expectedEnd.getMonth(), expectedEnd.getDate()]);
    });
  }

  it("keeps all-day ISO dates and recurrence IDs stable west of UTC", async () => {
    const previousTimezone = process.env.TZ;
    process.env.TZ = "America/Los_Angeles";
    try {
      const runtime = loadCalendarRuntime();
      const { master } = runtime.seedSeries({ allDay: true });
      const result = await runtime.invoke("updateEvent", {
        ...ids, recurrenceId: "2026-09-28T00:00:00.000Z",
        startDate: "2026-09-28T00:00:00.000Z", endDate: "2026-09-30",
      });
      assert.equal(result.success, true, result.error);
      const event = runtime.calls.modifies[0].item;
      assert.deepEqual([event.startDate.year, event.startDate.month, event.startDate.day], [2026, 8, 28]);
      assert.deepEqual([event.endDate.year, event.endDate.month, event.endDate.day], [2026, 9, 1]);
      assert.equal(master.recurrenceInfo.lookups[0].day, 28);
      assert.equal(master.recurrenceInfo.lookups[0].isDate, true);
      const deletion = await runtime.invoke("deleteEvent", { ...ids, recurrenceId: "2026-09-28T00:00:00.000Z" });
      assert.equal(deletion.success, true, deletion.error);
      assert.equal(runtime.calls.modifies[1].item.recurrenceInfo.excluded[0].day, 28);
    } finally {
      if (previousTimezone === undefined) delete process.env.TZ;
      else process.env.TZ = previousTimezone;
    }
  });

  it("rejects invalid all-day dates and preserves the stored event", async () => {
    const runtime = loadCalendarRuntime();
    const { master } = runtime.seedSeries({ allDay: true });
    for (const args of [{ startDate: "not-a-date" }, { endDate: "not-a-date" }]) {
      const result = await runtime.invoke("updateEvent", { ...ids, ...args });
      assert.match(result.error, /Invalid (startDate|endDate)/);
    }
    assert.equal(runtime.calls.modifies.length, 0);
    assert.equal(master.startDate.day, 5);
  });

  for (const name of ["updateEvent", "deleteEvent"]) {
    it(`${name} rejects malformed and rolled-over all-day recurrence IDs`, async () => {
      const runtime = loadCalendarRuntime();
      runtime.seedSeries({ allDay: true });
      for (const recurrenceId of ["2026-09-28not-a-date", "2026-02-30"]) {
        const result = await runtime.invoke(name, {
          ...ids, recurrenceId, ...(name === "updateEvent" ? { title: "Changed" } : {}),
        });
        assert.match(result.error, /Invalid recurrenceId/, recurrenceId);
      }
      assert.equal(runtime.calls.modifies.length, 0);
      assert.equal(runtime.calls.deletes.length, 0);
    });
  }
});

describe("Calendar access restrictions cover the actual new handler paths", () => {
  const requests = [
    ["listCalendars", {}], ["listCategories", {}],
    ["listEvents", { calendarId: "calendar-1", startDate: "2026-09-28" }],
    ["listTasks", { calendarId: "calendar-1" }],
    ["createEvent", { title: "Private invite", startDate: "2026-09-28T10:00:00Z", calendarId: "calendar-1", attendees: [{ email: "alice@example.com" }], skipReview: true }],
    ["updateEvent", { ...ids, title: "Occurrence", recurrenceId }],
    ["updateEvent", { ...ids, attendees: [{ email: "alice@example.com" }] }],
    ["updateEvent", { ...ids, recurrenceId, attendees: [] }],
    ["deleteEvent", { ...ids, recurrenceId }],
    ["deleteEvent", { ...ids, recurrenceId: null }],
    ["createTask", { title: "Private task", calendarId: "calendar-1", skipReview: true }],
    ["updateTask", { taskId: "event-1", calendarId: "calendar-1", title: "Private task" }],
  ];

  for (const allowedAccounts of [["account-1"], "{"]) {
    it(`blocks all calendar paths before reading or writing under ${JSON.stringify(allowedAccounts)}`, async () => {
      const runtime = loadCalendarRuntime({ allowedAccounts, blockSkipReview: false });
      runtime.seedSeries();
      for (const [name, args] of requests) {
        const result = await runtime.invoke(name, { ...args });
        assert.match(result.error, /Account restrictions block/, name);
        assert.match(result.error, /Allow all calendars/, name);
      }
      assert.equal(runtime.calls.reads.length, 0);
      assert.equal(runtime.calls.adds.length, 0);
      assert.equal(runtime.calls.modifies.length, 0);
      assert.equal(runtime.calls.deletes.length, 0);
      assert.equal(runtime.calls.dialogs.length, 0);
      assert.equal(runtime.calls.calendarProperties.length, 0);
    });
  }

  it("allows occurrence edits and deletions after Allow all calendars is enabled", async () => {
    const runtime = loadCalendarRuntime({
      allowedAccounts: ["account-1"], allowAllCalendars: true, blockSkipReview: false,
    });
    runtime.seedSeries();
    const updated = await runtime.invoke("updateEvent", { ...ids, recurrenceId, attendees: [{ email: "alice@example.com" }] });
    assert.equal(updated.success, true, updated.error);
    assert.equal(updated.mode, "occurrence");
    const deleted = await runtime.invoke("deleteEvent", { ...ids, recurrenceId });
    assert.equal(deleted.success, true, deleted.error);
    assert.equal(runtime.calls.modifies.length, 2);
  });

  it("calendar access opt-in does not bypass the independent attendee safety preference", async () => {
    const runtime = loadCalendarRuntime({ allowedAccounts: ["account-1"], allowAllCalendars: true });
    runtime.seedSeries({ attendees: [{ id: "mailto:guest@example.com" }] });
    const result = await runtime.invoke("updateEvent", { ...ids, recurrenceId, attendees: [] });
    assert.match(result.error, /skipReview/i);
    assert.equal(runtime.calls.modifies.length, 0);
  });
});

describe("Production event/task review defaults and tool schemas", () => {
  for (const [name, args] of [
    ["createEvent", { title: "Review me", startDate: "2026-09-28T10:00:00Z", calendarId: "calendar-1" }],
    ["createTask", { title: "Review me", dueDate: "2026-09-28", calendarId: "calendar-1" }],
  ]) {
    it(`${name} keeps default review and only honors skipReview after opt-out`, async () => {
      const blocked = loadCalendarRuntime();
      const review = await blocked.invoke(name, { ...args });
      assert.equal(review.success, true, review.error);
      assert.equal(blocked.calls.dialogs.length, 1);
      const refused = await blocked.invoke(name, { ...args, skipReview: true });
      assert.match(refused.error, /skipReview/);
      assert.equal(blocked.calls.adds.length, 0);
      const enabled = loadCalendarRuntime({ blockSkipReview: false });
      const direct = await enabled.invoke(name, { ...args, skipReview: true });
      assert.equal(direct.success, true, direct.error);
      assert.equal(enabled.calls.adds.length, 1);
      assert.equal(enabled.calls.dialogs.length, 0);
    });
  }

  it("validates recurrenceId types while retaining the null-as-unset API", () => {
    const runtime = loadCalendarRuntime();
    for (const name of ["updateEvent", "deleteEvent"]) {
      for (const id of [undefined, null, recurrenceId]) {
        assert.deepEqual(plain(runtime.validateToolArgs(name, { ...ids, recurrenceId: id })), []);
      }
      for (const id of [123, true, {}, []]) {
        assert.ok(runtime.validateToolArgs(name, { ...ids, recurrenceId: id }).length > 0);
      }
    }
    for (const recurrence of [null, "", "FREQ=WEEKLY"]) {
      assert.equal(runtime.validateToolArgs("updateEvent", { ...ids, recurrence }).length, 0);
    }
  });
});
