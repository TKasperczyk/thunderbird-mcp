"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { loadCalendarRuntime } = require("./helpers/calendar-runtime.cjs");

const range = {
  calendarId: "calendar-1", startDate: "2026-09-28T10:00:00Z", endDate: "2036-09-28T10:00:00Z",
};

describe("Production listEvents bounds native recurrence generation", () => {
  for (const [maxResults, expectedLimit] of [[3, 3], [50000, 500], [-1, 1], [3.5, 3]]) {
    it(`bounds the RRULE generator itself for maxResults=${maxResults}`, async () => {
      const runtime = loadCalendarRuntime();
      const master = runtime.seedDenseSeries();
      const originalRules = master.recurrenceInfo.rules.slice();
      const result = await runtime.invoke("listEvents", { ...range, maxResults });
      assert.ok(Array.isArray(result), result.error);
      assert.equal(result.length, expectedLimit);
      assert.ok(result.every(event => !event.recurrenceNotExpanded));
      assert.equal(master.recurrenceInfo.ruleCalls.length, 1);
      const call = master.recurrenceInfo.ruleCalls[0];
      assert.equal(call.maxCount, expectedLimit + 1);
      assert.equal(call.generated, expectedLimit + 1);
      assert.ok(call.maxCount > 0 && call.maxCount <= 501);
      assert.ok(call.end, "the native recurrence rule requires a range end");
      assert.equal(master.recurrenceInfo.generationCalls.length, 1);
      assert.ok(master.recurrenceInfo.generationCalls[0].maxCount > 0);
      assert.ok(master.recurrenceInfo.generationCalls[0].maxCount <= 501);
      assert.deepEqual(master.recurrenceInfo.rules, originalRules, "materialization must use a clone");
      assert.equal(master.recurrenceInfo.materialized, undefined);
    });
  }

  it("keeps the plain array contract through JSON serialization when a finite series fits", async () => {
    const runtime = loadCalendarRuntime();
    const master = runtime.seedDenseSeries({ total: 2 });
    const result = await runtime.invoke("listEvents", { ...range, maxResults: 2 });
    assert.ok(Array.isArray(result), result.error);
    const serialized = JSON.parse(JSON.stringify(result));
    assert.ok(Array.isArray(serialized));
    assert.equal(serialized.length, 2);
    assert.deepEqual(Object.keys(result), ["0", "1"]);
    assert.ok(serialized.every(event => !event.recurrenceNotExpanded));
    assert.equal(master.recurrenceInfo.ruleCalls[0].maxCount, 3);
    assert.equal(master.recurrenceInfo.ruleCalls[0].generated, 2);
  });

  it("keeps generation bounded even when EXDATEs hide all generated candidates", async () => {
    const runtime = loadCalendarRuntime();
    const master = runtime.seedDenseSeries();
    for (let index = 0; index < 4; index++) {
      const excluded = runtime.cal.createRecurrenceDate();
      excluded.date = runtime.date("2026-09-28T10:00:00Z");
      excluded.date.second += index;
      excluded.isNegative = true;
      master.recurrenceInfo.appendRecurrenceItem(excluded);
    }
    const result = await runtime.invoke("listEvents", { ...range, maxResults: 3 });
    assert.ok(Array.isArray(result), result.error);
    assert.equal(result.length, 0);
    assert.equal(master.recurrenceInfo.ruleCalls[0].generated, 4);
  });

  it("shares one hard candidate budget across many dense series", async () => {
    const runtime = loadCalendarRuntime();
    const masters = Array.from({ length: 20 }, (_, index) => runtime.seedDenseSeries({ id: `series-${index}` }));
    const result = await runtime.invoke("listEvents", { ...range, maxResults: 500 });
    assert.equal(result.length, 500, result.error);
    const calls = masters.flatMap(master => master.recurrenceInfo.ruleCalls);
    assert.ok(calls.length > 0);
    assert.ok(calls.every(call => Number.isInteger(call.maxCount) && call.maxCount > 0 && call.maxCount <= 501));
    assert.ok(calls.reduce((total, call) => total + call.generated, 0) <= 5000);
    assert.ok(masters.some(master => master.recurrenceInfo.ruleCalls.length === 0), "stop generating when the global budget is exhausted");
    for (const master of masters.filter(item => item.recurrenceInfo.ruleCalls.length === 0)) {
      const entries = result.filter(event => event.id === master.id);
      assert.equal(entries.length, 1, "an unexpanded series is represented once by its master");
      assert.equal(entries[0].recurrenceNotExpanded, true);
      assert.equal(entries[0].recurrenceId, undefined);
    }
  });

  it("shares the per-series candidate budget across multiple RRULEs", async () => {
    const runtime = loadCalendarRuntime();
    const master = runtime.seedDenseSeries({ total: 2 });
    master.recurrenceInfo.appendRecurrenceItem({ ...master.recurrenceInfo.rules[0] });
    master.recurrenceInfo.appendRecurrenceItem({ ...master.recurrenceInfo.rules[0] });
    const result = await runtime.invoke("listEvents", { ...range, maxResults: 3 });
    assert.ok(Array.isArray(result), result.error);
    assert.ok(result.every(event => !event.recurrenceNotExpanded));
    const calls = master.recurrenceInfo.ruleCalls;
    assert.equal(calls.reduce((total, call) => total + call.generated, 0), 4);
    assert.equal(calls.length, 2, "do not generate the remaining rule once the series budget is spent");
  });

  it("charges failed generation attempts to the shared candidate budget", async () => {
    const runtime = loadCalendarRuntime();
    const masters = Array.from({ length: 20 }, (_, index) => {
      const master = runtime.seedDenseSeries({ id: `failing-series-${index}` });
      const rule = master.recurrenceInfo.rules[0];
      const generate = rule.getOccurrences;
      rule.getOccurrences = (...args) => {
        generate(...args);
        throw new Error("Native failure after generating candidates");
      };
      return master;
    });
    const result = await runtime.invoke("listEvents", { ...range, maxResults: 500 });
    assert.equal(result.length, masters.length, result.error);
    assert.equal(new Set(result.map(event => event.id)).size, masters.length);
    assert.ok(result.every(event => event.recurrenceNotExpanded === true && event.recurrenceId === undefined));
    const calls = masters.flatMap(master => master.recurrenceInfo.ruleCalls);
    assert.ok(calls.length > 0);
    assert.ok(calls.reduce((total, call) => total + call.generated, 0) <= 5000);
    assert.ok(masters.some(master => master.recurrenceInfo.ruleCalls.length === 0));
  });

  it("preserves DTSTART even when the recurrence rule does not generate that date", async () => {
    const runtime = loadCalendarRuntime();
    const { master } = runtime.seedSeries({ start: "2026-09-28T09:00:00Z", end: "2026-09-28T10:00:00Z" });
    master.recurrenceInfo.rules[0].icalString = "RRULE:FREQ=WEEKLY;BYHOUR=10";
    const result = await runtime.invoke("listEvents", {
      ...range, startDate: "2026-09-28T09:00:00Z", endDate: "2026-09-28T11:00:00Z", maxResults: 5,
    });
    assert.equal(result.length, 2, result.error);
    assert.equal(master.recurrenceInfo.ruleCalls[0].generated, 1);
    assert.equal(result[0].startDate, "2026-09-28T09:00:00.000Z");
    assert.equal(result[1].startDate, "2026-09-28T10:00:00.000Z");
  });

  it("preserves stored RDATE, EXDATE, and moved exceptions while replacing rules on a clone", async () => {
    const runtime = loadCalendarRuntime();
    const { master, occurrences } = runtime.seedSeries();
    const extra = runtime.cal.createRecurrenceDate();
    extra.date = runtime.date("2026-10-01T10:00:00Z");
    const excluded = runtime.cal.createRecurrenceDate();
    excluded.date = occurrences[0].recurrenceId.clone();
    excluded.isNegative = true;
    master.recurrenceInfo.appendRecurrenceItem(extra);
    master.recurrenceInfo.appendRecurrenceItem(excluded);
    occurrences[1].startDate = runtime.date("2026-10-06T13:00:00Z");
    occurrences[1].endDate = runtime.date("2026-10-06T14:00:00Z");
    master.recurrenceInfo.exceptions.push(occurrences[1]);
    const originalRules = master.recurrenceInfo.rules.slice();
    const result = await runtime.invoke("listEvents", {
      ...range, endDate: "2026-10-10T00:00:00Z", maxResults: 10,
    });
    assert.equal(result.length, 2, result.error);
    assert.equal(result[0].startDate, "2026-10-01T10:00:00.000Z");
    assert.equal(result[1].startDate, "2026-10-06T13:00:00.000Z");
    assert.equal(result[1].recurrenceId, "2026-10-05T10:00:00.000Z");
    assert.deepEqual(master.recurrenceInfo.rules, originalRules);
    assert.equal(master.recurrenceInfo.exceptions[0], occurrences[1]);
  });

  it("includes an occurrence that starts before the query but overlaps its beginning", async () => {
    const runtime = loadCalendarRuntime();
    runtime.seedSeries();
    const result = await runtime.invoke("listEvents", {
      ...range, startDate: "2026-09-28T10:30:00Z", endDate: "2026-09-28T10:45:00Z", maxResults: 5,
    });
    assert.equal(result.length, 1, result.error);
    assert.equal(result[0].startDate, "2026-09-28T10:00:00.000Z");
  });

  it("keeps a moved exception even when its original recurrence date is outside the query", async () => {
    const runtime = loadCalendarRuntime();
    const { master, occurrences } = runtime.seedSeries();
    occurrences[0].startDate = runtime.date("2026-10-06T13:00:00Z");
    occurrences[0].endDate = runtime.date("2026-10-06T14:00:00Z");
    master.recurrenceInfo.exceptions.push(occurrences[0]);
    const result = await runtime.invoke("listEvents", {
      ...range, startDate: "2026-10-06T00:00:00Z", endDate: "2026-10-07T00:00:00Z", maxResults: 5,
    });
    assert.equal(result.length, 1, result.error);
    assert.equal(result[0].recurrenceId, "2026-09-28T10:00:00.000Z");
  });

  it("returns an EXRULE series master once with a flag, without expanding exclusion rules", async () => {
    const runtime = loadCalendarRuntime();
    const { master } = runtime.seedSeries();
    master.recurrenceInfo.appendRecurrenceItem({
      ...master.recurrenceInfo.rules[0], isNegative: true, icalString: "EXRULE:FREQ=SECONDLY",
    });
    runtime.calendar.getItemsAsArray = async () => [master, master];
    const result = await runtime.invoke("listEvents", { ...range, maxResults: 3 });
    assert.ok(Array.isArray(result), result.error);
    assert.equal(result.length, 1);
    assert.equal(result[0].id, master.id);
    assert.equal(result[0].isRecurring, true);
    assert.equal(result[0].recurrenceNotExpanded, true);
    assert.equal(result[0].recurrenceId, undefined);
    assert.equal(result[0].startDate, "2026-01-05T10:00:00.000Z", "keep the master's original dates even outside the query");
    assert.equal(master.recurrenceNotExpanded, undefined, "do not mutate the stored master");
    assert.equal(master.recurrenceInfo.ruleCalls.length, 0);
    assert.equal(master.recurrenceInfo.generationCalls.length, 0);
  });

  it("counts unexpanded series masters toward maxResults", async () => {
    const runtime = loadCalendarRuntime();
    const masters = Array.from({ length: 4 }, (_, index) => {
      const { master } = runtime.seedSeries({ id: `excluded-series-${index}` });
      master.recurrenceInfo.appendRecurrenceItem({
        ...master.recurrenceInfo.rules[0], isNegative: true, icalString: "EXRULE:FREQ=DAILY",
      });
      return master;
    });
    const result = await runtime.invoke("listEvents", { ...range, maxResults: 2 });
    assert.ok(Array.isArray(result), result.error);
    assert.equal(result.length, 2);
    assert.ok(result.every(event => event.recurrenceNotExpanded === true));
    assert.ok(masters.every(master => master.recurrenceInfo.ruleCalls.length === 0));
  });

  it("caps the array when only non-recurring events exceed the output limit", async () => {
    const runtime = loadCalendarRuntime();
    for (let index = 0; index < 3; index++) runtime.seedEvent({ id: `event-${index}` });
    const result = await runtime.invoke("listEvents", { ...range, maxResults: 2 });
    assert.ok(Array.isArray(result), result.error);
    assert.equal(result.length, 2);
    assert.ok(result.every(event => !event.recurrenceNotExpanded));
  });

  it("returns a complete empty result when there are no events", async () => {
    const runtime = loadCalendarRuntime();
    const result = await runtime.invoke("listEvents", range);
    assert.ok(Array.isArray(result), result.error);
    assert.equal(result.length, 0);
  });
});
