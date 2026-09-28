"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { loadCalendarRuntime, plain } = require("./helpers/calendar-runtime.cjs");

const range = { startDate: "2026-09-25T00:00:00Z", endDate: "2026-10-10T00:00:00Z" };
const eventArgs = {
  title: "Visible event", startDate: "2026-10-01T10:00:00Z", endDate: "2026-10-01T11:00:00Z",
  skipReview: true,
};

describe("Calendar provider contract", () => {
  for (const queryAPI of ["array", "stream"]) {
    it(`reads only events and applies task completion filters through the ${queryAPI} API`, async () => {
      const runtime = loadCalendarRuntime();
      if (queryAPI === "stream") runtime.calendar.getItemsAsArray = undefined;
      const event = runtime.seedEvent({ id: "event" });
      runtime.seedTask({ id: "pending", isCompleted: false });
      runtime.seedTask({ id: "done", isCompleted: true });
      const events = await runtime.invoke("listEvents", { calendarId: runtime.calendar.id, ...range });
      assert.ok(Array.isArray(events), events.error);
      assert.deepEqual(plain(events.map(item => item.id)), [event.id]);
      for (const [completed, expected] of [[undefined, ["done", "pending"]], [false, ["pending"]], [true, ["done"]]]) {
        const tasks = await runtime.invoke("listTasks", { calendarId: runtime.calendar.id, completed });
        assert.ok(Array.isArray(tasks), tasks.error);
        assert.deepEqual(plain(tasks.map(item => item.id).sort()), expected);
      }
      const reads = runtime.calls.reads.filter(call => call.kind === "range");
      assert.ok(reads.every(call => call.count === 0), "zero means an unbounded provider query");
      assert.ok(reads.every(call => !(call.filter & runtime.Ci.calICalendar.ITEM_FILTER_CLASS_OCCURRENCES)),
        "recurrence expansion must remain bounded by the application");
    });

    it(`finds events and tasks by scanning the ${queryAPI} API when direct lookup is unavailable`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      if (queryAPI === "stream") runtime.calendar.getItemsAsArray = undefined;
      runtime.calendar.getItem = undefined;
      const event = runtime.seedEvent({ id: "event" });
      const task = runtime.seedTask({ id: "task", isCompleted: true });
      const updated = await runtime.invoke("updateEvent", { calendarId: runtime.calendar.id, eventId: event.id, title: "Updated" });
      assert.equal(updated.success, true, updated.error);
      const deleted = await runtime.invoke("deleteEvent", { calendarId: runtime.calendar.id, eventId: event.id });
      assert.equal(deleted.success, true, deleted.error);
      const updatedTask = await runtime.invoke("updateTask", { calendarId: runtime.calendar.id, taskId: task.id, title: "Updated task" });
      assert.equal(updatedTask.success, true, updatedTask.error);
      assert.equal(runtime.calls.modifies[0].previous, event);
      assert.equal(runtime.calls.deletes[0], event);
      assert.equal(runtime.calls.modifies[1].previous, task);
    });

    it(`reopens completed tasks and updates completion filters through the ${queryAPI} API`, async () => {
      for (const [changes, percentComplete, status] of [
        [{ completed: false }, 0, "NEEDS-ACTION"],
        [{ percentComplete: 50 }, 50, "IN-PROCESS"],
      ]) {
        const runtime = loadCalendarRuntime();
        if (queryAPI === "stream") runtime.calendar.getItemsAsArray = undefined;
        const task = runtime.seedTask({ isCompleted: true });
        const completedDate = task.completedDate;
        const result = await runtime.invoke("updateTask", {
          calendarId: runtime.calendar.id, taskId: task.id, ...changes,
        });
        assert.equal(result.success, true, result.error);
        assert.equal(result.task.completed, false);
        assert.equal(result.task.completedDate, null);
        assert.equal(result.task.percentComplete, percentComplete);
        const written = runtime.calls.modifies[0].item;
        assert.equal(written.isCompleted, false);
        assert.equal(written.status, status);
        assert.equal(task.isCompleted, true, "reopening the clone must not alter the original");
        assert.equal(task.completedDate, completedDate);
        assert.equal(task.status, "COMPLETED");
        assert.equal(task.percentComplete, 100);
        // The fixture records writes; expose the written item for subsequent reads.
        runtime.items.set(task.id, written);
        const pending = await runtime.invoke("listTasks", { completed: false });
        assert.deepEqual(plain(pending.map(item => item.id)), [task.id]);
        assert.deepEqual(plain(await runtime.invoke("listTasks", { completed: true })), []);
      }
    });

    it(`reports a disabled calendar instead of treating the native ${queryAPI} result as empty`, async () => {
      const runtime = loadCalendarRuntime();
      const event = runtime.seedEvent();
      runtime.calendar.setProperty("disabled", true);
      // Thunderbird storage accepts writes and getItem even while getItems is disabled.
      assert.equal(await runtime.calendar.getItem(event.id), event);
      assert.deepEqual(await runtime.calendar.getItemsAsArray(8, 0, null, null), []);
      if (queryAPI === "stream") runtime.calendar.getItemsAsArray = undefined;
      for (const tool of ["listEvents", "listTasks"]) {
        for (const args of [{ calendarId: runtime.calendar.id, ...(tool === "listEvents" ? range : {}) }, {}]) {
          const result = await runtime.invoke(tool, args);
          assert.match(result.error, /disabled/i, `${tool} must distinguish disabled from empty`);
          assert.match(result.error, /enable.*Thunderbird/i);
        }
      }
      const calendars = await runtime.invoke("listCalendars", {});
      assert.equal(calendars[0].disabled, true);
      assert.equal(runtime.calendar.getProperty("disabled"), true, "reads must not enable the calendar");
    });
  }

  it("blocks writes to an explicitly selected disabled calendar before provider lookup or mutation", async () => {
    const runtime = loadCalendarRuntime({ blockSkipReview: false });
    const event = runtime.seedEvent();
    const task = runtime.seedTask();
    runtime.calendar.setProperty("disabled", true);
    const operations = [
      ["createEvent", eventArgs], ["createTask", { title: "Hidden task", skipReview: true }],
      ["updateEvent", { eventId: event.id, title: "Changed" }],
      ["deleteEvent", { eventId: event.id }],
      ["updateTask", { taskId: task.id, title: "Changed" }],
    ];
    for (const [name, args] of operations) {
      const result = await runtime.invoke(name, { ...args, calendarId: runtime.calendar.id });
      assert.match(result.error, /disabled/i, name);
    }
    assert.equal(runtime.calls.adds.length, 0);
    assert.equal(runtime.calls.modifies.length, 0);
    assert.equal(runtime.calls.deletes.length, 0);
    assert.equal(runtime.calls.dialogs.length, 0);
    assert.ok(runtime.calls.reads.every(call => call.kind === "calendars"));
    assert.equal(runtime.calendar.getProperty("disabled"), true);
  });

  it("skips disabled calendars when listing or choosing a default writable calendar", async () => {
    const runtime = loadCalendarRuntime({ blockSkipReview: false });
    runtime.calendar.setProperty("disabled", true);
    runtime.seedEvent({ id: "hidden" });
    const enabled = loadCalendarRuntime();
    enabled.calendar.id = "enabled-calendar";
    enabled.seedEvent({ id: "visible" });
    enabled.seedTask({ id: "task" });
    runtime.cal.manager.getCalendars = () => [runtime.calendar, enabled.calendar];
    const events = await runtime.invoke("listEvents", range);
    assert.deepEqual(plain(events.map(item => item.id)), ["visible"]);
    const tasks = await runtime.invoke("listTasks", {});
    assert.deepEqual(plain(tasks.map(item => item.id)), ["task"]);
    assert.equal((await runtime.invoke("createEvent", eventArgs)).success, true);
    assert.equal((await runtime.invoke("createTask", { title: "Visible task", skipReview: true })).success, true);
    assert.equal(runtime.calls.adds.length, 0);
    assert.equal(enabled.calls.adds.length, 2);
    assert.equal(runtime.calendar.getProperty("disabled"), true);
    const calendars = await runtime.invoke("listCalendars", {});
    assert.deepEqual(plain(calendars.map(item => item.disabled)), [true, false]);
  });
});
