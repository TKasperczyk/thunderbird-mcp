"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

// Mirrors the RRULE handling in api.js (normalizeRRule / extractRRuleFromItem).
// XPCOM is unavailable in node:test, so this covers only the pure string rules
// and the parentItem fallback -- not calIRecurrenceInfo itself.

function normalizeRRule(recurrence) {
  const body = String(recurrence).trim().replace(/^rrule:/i, "").trim();
  if (!body) throw new Error("Empty recurrence rule");
  const m = /(?:^|;)FREQ=([^;]*)/i.exec(body);
  if (!m) throw new Error("Recurrence rule must contain a FREQ part (e.g. FREQ=WEEKLY)");
  const freq = m[1].toUpperCase();
  if (freq === "SECONDLY" || freq === "MINUTELY") {
    throw new Error(`FREQ=${freq} is not supported: Thunderbird's recurrence engine generates no occurrences for it`);
  }
  return "RRULE:" + body;
}

function extractRRuleFromItem(item) {
  const rinfo = (item.parentItem || item).recurrenceInfo;
  if (!rinfo) return null;
  try {
    const rules = rinfo.getRecurrenceItems();
    for (const r of rules) {
      if (r && typeof r.icalString === "string") {
        const line = r.icalString.replace(/\r?\n$/, "");
        if (line.startsWith("RRULE:")) return line.slice("RRULE:".length);
      }
    }
  } catch { /* ignore */ }
  return null;
}

function mockRecurrenceInfo(icalStrings) {
  return {
    getRecurrenceItems() {
      return icalStrings.map(s => ({ icalString: s }));
    },
  };
}

describe("normalizeRRule", () => {
  it("adds the RRULE: prefix when missing", () => {
    assert.equal(normalizeRRule("FREQ=DAILY;COUNT=10"), "RRULE:FREQ=DAILY;COUNT=10");
  });

  it("keeps an existing RRULE: prefix without doubling it", () => {
    assert.equal(normalizeRRule("RRULE:FREQ=DAILY"), "RRULE:FREQ=DAILY");
  });

  it("accepts a lowercase rrule: prefix without producing RRULE:rrule:", () => {
    assert.equal(normalizeRRule("rrule:FREQ=DAILY"), "RRULE:FREQ=DAILY");
  });

  it("trims surrounding whitespace", () => {
    assert.equal(normalizeRRule("  RRULE:FREQ=WEEKLY;BYDAY=MO  "), "RRULE:FREQ=WEEKLY;BYDAY=MO");
    assert.equal(normalizeRRule("rrule: FREQ=WEEKLY"), "RRULE:FREQ=WEEKLY");
  });

  it("rejects FREQ=SECONDLY and FREQ=MINUTELY", () => {
    assert.throws(() => normalizeRRule("FREQ=SECONDLY"), /SECONDLY/);
    assert.throws(() => normalizeRRule("RRULE:FREQ=MINUTELY;INTERVAL=5"), /MINUTELY/);
    assert.throws(() => normalizeRRule("freq=minutely"), /MINUTELY/);
  });

  it("accepts every expandable frequency", () => {
    for (const freq of ["HOURLY", "DAILY", "WEEKLY", "MONTHLY", "YEARLY"]) {
      assert.equal(normalizeRRule(`FREQ=${freq}`), `RRULE:FREQ=${freq}`);
    }
  });

  it("only matches FREQ as a property name, not inside another value", () => {
    // UNTIL value contains no FREQ; a BYDAY rule with FREQ later still parses.
    assert.equal(
      normalizeRRule("BYDAY=MO;FREQ=WEEKLY"),
      "RRULE:BYDAY=MO;FREQ=WEEKLY",
    );
  });

  it("rejects empty and prefix-only rules", () => {
    assert.throws(() => normalizeRRule(""), /Empty/);
    assert.throws(() => normalizeRRule("   "), /Empty/);
    assert.throws(() => normalizeRRule("RRULE:"), /Empty/);
  });

  it("rejects rules without a FREQ part", () => {
    assert.throws(() => normalizeRRule("BYDAY=MO,WE,FR"), /FREQ/);
    assert.throws(() => normalizeRRule("RRULE:COUNT=10;INTERVAL=2"), /FREQ/);
  });
});

describe("extractRRuleFromItem", () => {
  it("returns the RRULE body without prefix", () => {
    const item = { recurrenceInfo: mockRecurrenceInfo(["RRULE:FREQ=WEEKLY;BYDAY=MO,TU"]) };
    assert.equal(extractRRuleFromItem(item), "FREQ=WEEKLY;BYDAY=MO,TU");
  });

  it("trims the trailing CRLF an icalString may carry", () => {
    const item = { recurrenceInfo: mockRecurrenceInfo(["RRULE:FREQ=DAILY;COUNT=3\r\n"]) };
    assert.equal(extractRRuleFromItem(item), "FREQ=DAILY;COUNT=3");
  });

  it("reads the rule from parentItem on occurrence proxies", () => {
    const master = { recurrenceInfo: mockRecurrenceInfo(["RRULE:FREQ=MONTHLY"]) };
    const occurrence = { parentItem: master, recurrenceInfo: null };
    assert.equal(extractRRuleFromItem(occurrence), "FREQ=MONTHLY");
  });

  it("skips non-RRULE recurrence items (EXDATE, RDATE)", () => {
    const item = {
      recurrenceInfo: mockRecurrenceInfo([
        "EXDATE:20260101T100000Z\r\n",
        "RRULE:FREQ=YEARLY\r\n",
      ]),
    };
    assert.equal(extractRRuleFromItem(item), "FREQ=YEARLY");
  });

  it("returns null for RDATE-only recurrences", () => {
    const item = { recurrenceInfo: mockRecurrenceInfo(["RDATE:20260101T100000Z"]) };
    assert.equal(extractRRuleFromItem(item), null);
  });

  it("returns null for non-recurring items", () => {
    assert.equal(extractRRuleFromItem({ recurrenceInfo: null }), null);
    assert.equal(extractRRuleFromItem({}), null);
  });

  it("degrades to null when the provider throws", () => {
    const item = { recurrenceInfo: { getRecurrenceItems() { throw new Error("quirk"); } } };
    assert.equal(extractRRuleFromItem(item), null);
  });
});
