"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

// Mirrors formatEvent's organizer/attendees/myParticipationStatus extraction.
// XPCOM is unavailable in node:test, so this covers only the pure formatting
// rules (field reduction, guards, degradation to empty values).
function formatAttendee(att) {
  if (!att) return null;
  return {
    id: att.id || "",
    commonName: att.commonName || "",
    participationStatus: att.participationStatus || "",
    role: att.role || "",
  };
}

// The organizer is not an attendee: PARTSTAT and ROLE don't apply.
function formatOrganizer(org) {
  if (!org) return null;
  return {
    id: org.id || "",
    commonName: org.commonName || "",
  };
}

function applyAttendeeFields(result, item, calendar, calItip) {
  try {
    result.organizer = formatOrganizer(item.organizer);
  } catch { result.organizer = null; }
  try {
    const atts = typeof item.getAttendees === "function" ? item.getAttendees() : [];
    result.attendees = (atts || []).map(formatAttendee).filter(Boolean);
  } catch { result.attendees = []; }
  try {
    const me = calItip && typeof calItip.getInvitedAttendee === "function"
      ? calItip.getInvitedAttendee(item, calendar)
      : null;
    result.myParticipationStatus = me ? (me.participationStatus || "") : "";
  } catch { result.myParticipationStatus = ""; }
  return result;
}

describe("formatAttendee", () => {
  it("reduces an attendee to id, commonName, participationStatus, role", () => {
    assert.deepEqual(
      formatAttendee({
        id: "mailto:alice@example.com",
        commonName: "Alice",
        participationStatus: "ACCEPTED",
        role: "REQ-PARTICIPANT",
        icalString: "ATTENDEE;...",
      }),
      {
        id: "mailto:alice@example.com",
        commonName: "Alice",
        participationStatus: "ACCEPTED",
        role: "REQ-PARTICIPANT",
      },
    );
  });

  it("degrades missing fields to empty strings", () => {
    assert.deepEqual(formatAttendee({}), {
      id: "",
      commonName: "",
      participationStatus: "",
      role: "",
    });
  });

  it("returns null for a null attendee", () => {
    assert.equal(formatAttendee(null), null);
  });
});

describe("formatOrganizer", () => {
  it("reduces the organizer to id and commonName only", () => {
    assert.deepEqual(
      formatOrganizer({
        id: "mailto:boss@example.com",
        commonName: "Boss",
        participationStatus: "ACCEPTED",
        role: "CHAIR",
      }),
      { id: "mailto:boss@example.com", commonName: "Boss" },
    );
  });

  it("returns null when the event has no organizer", () => {
    assert.equal(formatOrganizer(null), null);
  });
});

describe("formatEvent attendee fields", () => {
  const calendar = { id: "cal1" };

  it("exposes organizer, attendees and myParticipationStatus", () => {
    const item = {
      organizer: { id: "mailto:boss@example.com", commonName: "Boss" },
      getAttendees() {
        return [
          { id: "mailto:me@example.com", commonName: "Me", participationStatus: "DECLINED", role: "REQ-PARTICIPANT" },
          null,
        ];
      },
    };
    const calItip = {
      getInvitedAttendee(it, cal2) {
        assert.equal(it, item);
        assert.equal(cal2, calendar);
        return { participationStatus: "DECLINED" };
      },
    };

    const result = applyAttendeeFields({}, item, calendar, calItip);

    assert.deepEqual(result.organizer, { id: "mailto:boss@example.com", commonName: "Boss" });
    assert.deepEqual(result.attendees, [
      { id: "mailto:me@example.com", commonName: "Me", participationStatus: "DECLINED", role: "REQ-PARTICIPANT" },
    ]);
    assert.equal(result.myParticipationStatus, "DECLINED");
  });

  it("degrades gracefully when the provider exposes none of it", () => {
    const result = applyAttendeeFields({}, {}, calendar, null);

    assert.equal(result.organizer, null);
    assert.deepEqual(result.attendees, []);
    assert.equal(result.myParticipationStatus, "");
  });

  it("degrades each field independently when accessors throw", () => {
    const item = {
      get organizer() { throw new Error("provider quirk"); },
      getAttendees() { throw new Error("provider quirk"); },
    };
    const calItip = {
      getInvitedAttendee() { throw new Error("provider quirk"); },
    };

    const result = applyAttendeeFields({}, item, calendar, calItip);

    assert.equal(result.organizer, null);
    assert.deepEqual(result.attendees, []);
    assert.equal(result.myParticipationStatus, "");
  });
});
