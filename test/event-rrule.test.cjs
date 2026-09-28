"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { loadCalendarRuntime } = require("./helpers/calendar-runtime.cjs");

const createArgs = {
  title: "Recurring meeting", startDate: "2026-09-28T10:00:00Z",
  calendarId: "calendar-1", skipReview: true,
};
const updateArgs = { eventId: "event-1", calendarId: "calendar-1" };

function assertNotPersisted(runtime) {
  assert.equal(runtime.calls.adds.length, 0);
  assert.equal(runtime.calls.modifies.length, 0);
  assert.equal(runtime.calls.dialogs.length, 0);
}

describe("Production RRULE normalization", () => {
  it("accepts an optional case-insensitive prefix and surrounding spaces", () => {
    const runtime = loadCalendarRuntime();
    for (const rule of ["FREQ=WEEKLY;BYDAY=MO", "RRULE:FREQ=WEEKLY;BYDAY=MO", "  rrule: FREQ=WEEKLY;BYDAY=MO  "]) {
      assert.equal(runtime.normalizeRRule(rule), "RRULE:FREQ=WEEKLY;BYDAY=MO");
    }
    assert.equal(runtime.normalizeRRule("BYDAY=MO;FREQ=WEEKLY"), "RRULE:BYDAY=MO;FREQ=WEEKLY");
  });

  it("accepts the supported frequencies", () => {
    const runtime = loadCalendarRuntime();
    for (const frequency of ["HOURLY", "DAILY", "WEEKLY", "MONTHLY", "YEARLY"]) {
      assert.equal(runtime.normalizeRRule(`FREQ=${frequency}`), `RRULE:FREQ=${frequency}`);
    }
  });

  it("rejects missing/unsupported frequencies, duplicate parts and malformed syntax", () => {
    const runtime = loadCalendarRuntime();
    for (const rule of [
      "", " ", "RRULE:", "COUNT=10", "FREQ=", "FREQ=INVALID",
      "FREQ=SECONDLY", "freq=minutely", "FREQ=DAILY;FREQ=WEEKLY",
      "FREQ=DAILY;COUNT=2;COUNT=3", "FREQ=DAILY;BROKEN",
      "FREQ=DAILY;COUNT=", "FREQ=DAILY;;COUNT=2",
    ]) {
      assert.throws(() => runtime.normalizeRRule(rule), undefined, `accepted ${JSON.stringify(rule)}`);
    }
  });

  it("rejects every C0/C1 control before trimming, including leading/trailing CRLF", () => {
    const runtime = loadCalendarRuntime();
    const controls = [...Array.from({ length: 32 }, (_, index) => index),
      ...Array.from({ length: 33 }, (_, index) => 127 + index)];
    for (const code of controls) {
      const control = String.fromCharCode(code);
      for (const rule of [`${control}FREQ=DAILY`, `FREQ=DAILY${control}`, `FREQ=DAILY;${control}COUNT=2`]) {
        assert.throws(() => runtime.normalizeRRule(rule), /control/i, `accepted control ${code}`);
      }
    }
    assert.throws(() => runtime.normalizeRRule("FREQ=DAILY\r\nATTENDEE:mailto:injected@example.com"), /control/i);
  });
});

describe("Production calendar writes parse recurrence before persistence", () => {
  for (const parser of ["cal", "xpcom"]) {
    it(`uses Thunderbird's ${parser} recurrence parser to create a recurring event`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false, parser });
      const result = await runtime.invoke("createEvent", { ...createArgs, recurrence: "FREQ=WEEKLY;BYDAY=MO" });
      assert.equal(result.success, true, result.error);
      assert.equal(runtime.calls.nativeRules.length, 1);
      assert.equal(runtime.calls.nativeRules[0].implementation, parser);
      assert.equal(runtime.calls.nativeRules[0].line, "RRULE:FREQ=WEEKLY;BYDAY=MO");
      assert.equal(runtime.calls.adds[0].recurrenceInfo.rules[0].type, "WEEKLY");
    });

    for (const parserFailure of ["throw", "missing-type"]) {
      for (const name of ["createEvent", "updateEvent"]) {
        it(`${name} rejects ${parser} parser ${parserFailure} without writing an item`, async () => {
          const runtime = loadCalendarRuntime({ blockSkipReview: false, parser, parserFailure });
          const item = runtime.seedEvent();
          const result = await runtime.invoke(name, {
            ...(name === "createEvent" ? createArgs : updateArgs), recurrence: "FREQ=WEEKLY;BYDAY=MO",
          });
          assert.match(result.error, /recurrence|rule/i);
          assert.equal(runtime.calls.nativeRules.length, 1);
          assertNotPersisted(runtime);
          assert.equal(item.recurrenceInfo, null);
        });
      }
    }
  }

  for (const name of ["createEvent", "updateEvent"]) {
    it(`${name} rejects injection before the recurrence parser runs`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      runtime.seedEvent();
      const result = await runtime.invoke(name, {
        ...(name === "createEvent" ? createArgs : updateArgs),
        recurrence: "FREQ=DAILY\r\nATTENDEE:mailto:injected@example.com",
      });
      assert.match(result.error, /control/i);
      assert.equal(runtime.calls.nativeRules.length, 0);
      assertNotPersisted(runtime);
    });
  }

  it("refuses hourly recurrence on an all-day event", async () => {
    const runtime = loadCalendarRuntime({ blockSkipReview: false });
    const result = await runtime.invoke("createEvent", {
      ...createArgs, startDate: "2026-09-28", allDay: true, recurrence: "FREQ=HOURLY",
    });
    assert.match(result.error, /all-day/i);
    assertNotPersisted(runtime);
  });

  it("replaces the master rule and explicitly discards old recurrence exceptions", async () => {
    const runtime = loadCalendarRuntime({ blockSkipReview: false });
    const { master } = runtime.seedSeries();
    master.recurrenceInfo.excluded.push(runtime.date("2026-09-28T10:00:00Z"));
    const result = await runtime.invoke("updateEvent", { ...updateArgs, recurrence: "FREQ=DAILY;COUNT=3" });
    assert.equal(result.success, true, result.error);
    const written = runtime.calls.modifies[0].item;
    assert.equal(written.recurrenceInfo.rules[0].icalString, "RRULE:FREQ=DAILY;COUNT=3");
    assert.equal(written.recurrenceInfo.excluded.length, 0);
    assert.equal(master.recurrenceInfo.excluded.length, 1);
    assert.match(result.warning, /series/i);
  });

  for (const recurrence of [null, ""]) {
    it(`clears recurrence with ${JSON.stringify(recurrence)} and removes the series warning`, async () => {
      const runtime = loadCalendarRuntime();
      const { master } = runtime.seedSeries();
      const result = await runtime.invoke("updateEvent", { ...updateArgs, recurrence });
      assert.equal(result.success, true, result.error);
      assert.equal(runtime.calls.modifies[0].item.recurrenceInfo, null);
      assert.equal(result.warning, undefined);
      assert.ok(master.recurrenceInfo, "clearing a clone must leave the stored item alone");
    });
  }
});

describe("Production recurrence extraction", () => {
  it("reads a parent's RRULE on occurrence proxies and skips EXDATE/RDATE", () => {
    const runtime = loadCalendarRuntime();
    const { master, occurrences } = runtime.seedSeries();
    master.recurrenceInfo.rules.unshift({ icalString: "EXDATE:20260928T100000Z\r\n" });
    assert.equal(runtime.extractRRuleFromItem(occurrences[0]), "FREQ=WEEKLY");
    master.recurrenceInfo.rules = [{ icalString: "RDATE:20260928T100000Z" }];
    assert.equal(runtime.extractRRuleFromItem(master), null);
  });

  it("returns null for missing recurrence and provider failures", () => {
    const runtime = loadCalendarRuntime();
    assert.equal(runtime.extractRRuleFromItem({}), null);
    const item = runtime.makeEvent({ recurrenceInfo: { getRecurrenceItems() { throw new Error("provider failure"); } } });
    assert.equal(runtime.extractRRuleFromItem(item), null);
  });
});
