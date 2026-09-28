"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { loadCalendarRuntime, plain } = require("./helpers/calendar-runtime.cjs");

const createArgs = { title: "Meeting", startDate: "2026-09-28T10:00:00Z", calendarId: "calendar-1" };
const updateArgs = { eventId: "event-1", calendarId: "calendar-1" };
const attendee = { email: "alice@example.com", name: "Alice", role: "optional" };
const existingAttendee = { id: "mailto:old@example.com", participationStatus: "ACCEPTED", role: "REQ-PARTICIPANT" };
const recurrenceId = "2026-09-28T10:00:00.000Z";

function assertNoWrites(runtime) {
  assert.equal(runtime.calls.adds.length, 0);
  assert.equal(runtime.calls.modifies.length, 0);
  assert.equal(runtime.calls.deletes.length, 0);
  assert.equal(runtime.calls.dialogs.length, 0);
  assert.equal(runtime.calls.calendarProperties.length, 0);
}

describe("Attendee writes use the production Block skipReview preference", () => {
  for (const [label, options] of [
    ["default", {}], ["enabled", { blockSkipReview: true }],
    ["unreadable", { unreadableSkipReview: true }],
  ]) {
    it(`rejects createEvent attendees while ${label}, even with a review dialog requested`, async () => {
      const runtime = loadCalendarRuntime(options);
      for (const skipReview of [undefined, false, true]) {
        const result = await runtime.invoke("createEvent", { ...createArgs, attendees: [attendee], skipReview });
        assert.match(result.error, /skipReview/i);
        assertNoWrites(runtime);
      }
    });

    for (const occurrence of [false, true]) {
      it(`rejects replacing/clearing attendees on ${occurrence ? "an occurrence" : "a series"} while ${label}`, async () => {
        const runtime = loadCalendarRuntime(options);
        runtime.seedSeries({ attendees: [existingAttendee] });
        for (const attendees of [[attendee], []]) {
          const result = await runtime.invoke("updateEvent", {
            ...updateArgs, title: "Changed title", attendees,
            ...(occurrence ? { recurrenceId } : {}),
          });
          assert.match(result.error, /skipReview/i);
          assertNoWrites(runtime);
        }
      });
    }
  }

  it("allows event review without attendees, including null and an empty list", async () => {
    const runtime = loadCalendarRuntime();
    for (const attendees of [undefined, null, []]) {
      const result = await runtime.invoke("createEvent", { ...createArgs, attendees });
      assert.equal(result.success, true, result.error);
    }
    assert.equal(runtime.calls.dialogs.length, 3);
    assert.equal(runtime.calls.adds.length, 0);
  });

  for (const occurrence of [false, true]) {
    it(`preserves attendees on ${occurrence ? "occurrence" : "series"} updates when omitted or null with the preference off`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      runtime.seedSeries({ attendees: [existingAttendee] });
      for (const attendees of [undefined, null]) {
        const result = await runtime.invoke("updateEvent", {
          ...updateArgs, title: "Rename only", attendees,
          ...(occurrence ? { recurrenceId } : {}),
        });
        assert.equal(result.success, true, result.error);
        assert.deepEqual(plain(result.updated), ["title"]);
        assert.deepEqual(runtime.calls.modifies.at(-1).item.getAttendees(), [existingAttendee]);
      }
      assert.equal(runtime.calls.calendarProperties.length, 0);
    });

    it(`replaces attendees on ${occurrence ? "an occurrence" : "a series"} when the preference is off`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      const { master, occurrences } = runtime.seedSeries({ attendees: [existingAttendee] });
      const result = await runtime.invoke("updateEvent", {
        ...updateArgs, attendees: [attendee, { email: "mailto:bob@example.com" }],
        ...(occurrence ? { recurrenceId } : {}),
      });
      assert.equal(result.success, true, result.error);
      assert.ok(result.updated.includes("attendees"));
      const write = runtime.calls.modifies[0];
      assert.equal(write.previous, occurrence ? occurrences[0] : master);
      assert.deepEqual(plain(write.item.getAttendees()), [
        { id: "mailto:alice@example.com", commonName: "Alice", role: "OPT-PARTICIPANT", participationStatus: "NEEDS-ACTION" },
        { id: "mailto:bob@example.com", role: "REQ-PARTICIPANT", participationStatus: "NEEDS-ACTION" },
      ]);
      assert.equal(write.item.organizer.id, "mailto:organizer@example.com");
      assert.equal(runtime.calendar.getProperty("organizerId"), "mailto:organizer@example.com");
      assert.deepEqual(master.getAttendees(), [existingAttendee], "the stored master must not be mutated before modifyItem");
    });

    it(`clears attendees on ${occurrence ? "an occurrence" : "a series"} only with [] and the preference off`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      runtime.seedSeries({ attendees: [existingAttendee] });
      const result = await runtime.invoke("updateEvent", {
        ...updateArgs, attendees: [], ...(occurrence ? { recurrenceId } : {}),
      });
      assert.equal(result.success, true, result.error);
      assert.deepEqual(runtime.calls.modifies[0].item.getAttendees(), []);
    });
  }

  for (const skipReview of [false, true]) {
    it(`creates attendees and an organizer with the preference off (skipReview=${skipReview})`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      const result = await runtime.invoke("createEvent", { ...createArgs, skipReview, attendees: [attendee] });
      assert.equal(result.success, true, result.error);
      const event = skipReview ? runtime.calls.adds[0] : runtime.calls.dialogs[0].calendarEvent;
      assert.equal(event.getAttendees()[0].id, "mailto:alice@example.com");
      assert.equal(event.organizer.id, "mailto:organizer@example.com");
      assert.equal(runtime.calls.adds.length, skipReview ? 1 : 0);
    });
  }
});

describe("Meetings are read-only while the production Block skipReview guard is active", () => {
  for (const [label, options] of [
    ["default", {}], ["enabled", { blockSkipReview: true }],
    ["unreadable", { unreadableSkipReview: true }],
  ]) {
    for (const occurrence of [false, true]) {
      it(`rejects ${occurrence ? "occurrence" : "series"} updates with attendees omitted or null while ${label}`, async () => {
        const runtime = loadCalendarRuntime(options);
        runtime.seedSeries({ attendees: [existingAttendee] });
        const changes = [
          { title: "New title" }, { description: "New description" },
          { startDate: "2026-09-28T12:00:00Z", endDate: "2026-09-28T13:00:00Z" },
          ...(!occurrence ? [{ recurrence: "FREQ=DAILY;COUNT=3" }, { recurrence: null }] : []),
        ];
        for (const change of changes) {
          for (const attendees of [undefined, null]) {
            const result = await runtime.invoke("updateEvent", {
              ...updateArgs, ...change, attendees, ...(occurrence ? { recurrenceId } : {}),
            });
            assert.match(result.error, /Block skipReview/i, JSON.stringify(change));
            assert.match(result.error, /read.only|update|modif|calendar.*server/i);
            assertNoWrites(runtime);
          }
        }
      });

      it(`rejects ${occurrence ? "occurrence" : "series"} deletion while ${label}`, async () => {
        const runtime = loadCalendarRuntime(options);
        const { master } = runtime.seedSeries({ attendees: [existingAttendee] });
        const result = await runtime.invoke("deleteEvent", {
          ...updateArgs, ...(occurrence ? { recurrenceId } : {}),
        });
        assert.match(result.error, /Block skipReview/i);
        assertNoWrites(runtime);
        assert.equal(master.recurrenceInfo.excluded.length, 0);
      });
    }
  }

  for (const name of ["updateEvent", "deleteEvent"]) {
    for (const occurrence of [false, true]) {
      it(`${name} allows ${occurrence ? "occurrence" : "series"} writes when the preference is off`, async () => {
        const runtime = loadCalendarRuntime({ blockSkipReview: false });
        runtime.seedSeries({ attendees: [existingAttendee] });
        const result = await runtime.invoke(name, {
          ...updateArgs, ...(name === "updateEvent" ? { description: "Changed without attendees" } : {}),
          ...(occurrence ? { recurrenceId } : {}),
        });
        assert.equal(result.success, true, result.error);
        assert.equal(runtime.calls.modifies.length + runtime.calls.deletes.length, 1);
      });

      it(`${name} permits ${occurrence ? "occurrence" : "series"} writes with no attendees or only the calendar user`, async () => {
        for (const attendees of [[], [{ id: "MAILTO:ORGANIZER@EXAMPLE.COM", participationStatus: "ACCEPTED" }]]) {
          const runtime = loadCalendarRuntime();
          runtime.seedSeries({ attendees });
          const result = await runtime.invoke(name, {
            ...updateArgs, ...(name === "updateEvent" ? { title: "Personal event" } : {}),
            ...(occurrence ? { recurrenceId } : {}),
          });
          assert.equal(result.success, true, result.error);
          assert.equal(runtime.calls.modifies.length + runtime.calls.deletes.length, 1);
        }
      });
    }

    it(`${name} checks series attendees even when the selected occurrence has none`, async () => {
      const runtime = loadCalendarRuntime();
      const { occurrences } = runtime.seedSeries({ attendees: [existingAttendee] });
      occurrences[0].attendees = [];
      const result = await runtime.invoke(name, {
        ...updateArgs, recurrenceId, ...(name === "updateEvent" ? { title: "Changed" } : {}),
      });
      assert.match(result.error, /Block skipReview/i);
      assertNoWrites(runtime);
    });

    for (const occurrence of [false, true]) {
      it(`${name} checks attendee-bearing exceptions for ${occurrence ? "occurrence" : "whole-series"} writes`, async () => {
        const runtime = loadCalendarRuntime();
        const { master, occurrences } = runtime.seedSeries();
        occurrences[0].attendees = [existingAttendee];
        master.recurrenceInfo.exceptions.push(occurrences[0]);
        const result = await runtime.invoke(name, {
          ...updateArgs, ...(name === "updateEvent" ? { title: "Changed" } : {}),
          ...(occurrence ? { recurrenceId } : {}),
        });
        assert.match(result.error, /Block skipReview/i);
        assertNoWrites(runtime);
      });

      it(`${name} fails closed on unreadable ${occurrence ? "occurrence" : "series"} attendees`, async () => {
        const runtime = loadCalendarRuntime();
        const { master, occurrences } = runtime.seedSeries();
        (occurrence ? occurrences[0] : master).getAttendees = () => { throw new Error("Provider unavailable"); };
        const result = await runtime.invoke(name, {
          ...updateArgs, ...(name === "updateEvent" ? { title: "Changed" } : {}),
          ...(occurrence ? { recurrenceId } : {}),
        });
        assert.match(result.error, /Block skipReview/i);
        assertNoWrites(runtime);
      });
    }

    it(`${name} fails closed if exception attendee inspection is unavailable`, async () => {
      const runtime = loadCalendarRuntime();
      const { master } = runtime.seedSeries();
      master.recurrenceInfo.getExceptionIds = () => { throw new Error("Provider unavailable"); };
      const result = await runtime.invoke(name, {
        ...updateArgs, ...(name === "updateEvent" ? { title: "Changed" } : {}),
      });
      assert.match(result.error, /Block skipReview/i);
      assertNoWrites(runtime);
    });

    it(`${name} gives the setting error when an attendee identifier cannot be read`, async () => {
      const runtime = loadCalendarRuntime();
      const unreadable = {};
      Object.defineProperty(unreadable, "id", { get() { throw new Error("Unreadable attendee ID"); } });
      runtime.seedEvent({ attendees: [unreadable] });
      const result = await runtime.invoke(name, {
        ...updateArgs, ...(name === "updateEvent" ? { title: "Changed" } : {}),
      });
      assert.match(result.error, /Block skipReview/i);
      assertNoWrites(runtime);
    });

    it(`${name} does not mistake a provider's invited attendee for the calendar's own identity`, async () => {
      const runtime = loadCalendarRuntime();
      runtime.seedEvent({ attendees: [existingAttendee], invitedAttendee: existingAttendee });
      const result = await runtime.invoke(name, {
        ...updateArgs, ...(name === "updateEvent" ? { title: "Changed" } : {}),
      });
      assert.match(result.error, /Block skipReview/i);
      assertNoWrites(runtime);
    });

    for (const options of [{ identity: null }, { identityKey: null }, { unreadableIdentity: true }]) {
      it(`${name} requires a known calendar identity to exempt self-only attendees (${JSON.stringify(options)})`, async () => {
        const runtime = loadCalendarRuntime(options);
        runtime.seedEvent({ attendees: [{ id: "mailto:organizer@example.com" }] });
        const result = await runtime.invoke(name, {
          ...updateArgs, ...(name === "updateEvent" ? { title: "Changed" } : {}),
        });
        assert.match(result.error, /Block skipReview/i);
        assertNoWrites(runtime);
      });
    }
  }

  for (const name of ["updateEvent", "deleteEvent"]) {
    for (const lookup of ["getItem", "scan"]) {
      it(`${name} checks series attendees when ${lookup} returns a personal exception first`, async () => {
        const runtime = loadCalendarRuntime();
        const { master, occurrences } = runtime.seedSeries();
        master.attendees = [existingAttendee];
        runtime.calendar.getItem = async () => lookup === "getItem" ? occurrences[0] : null;
        runtime.calendar.getItemsAsArray = async () => [occurrences[0], master];
        const result = await runtime.invoke(name, {
          ...updateArgs, ...(name === "updateEvent" ? { title: "Changed" } : {}),
        });
        assert.match(result.error, /Block skipReview/i);
        assertNoWrites(runtime);
      });
    }
  }

  it("allows a selected personal occurrence even if another exception has attendees", async () => {
    const runtime = loadCalendarRuntime();
    const { master, occurrences } = runtime.seedSeries();
    occurrences[1].attendees = [existingAttendee];
    master.recurrenceInfo.exceptions.push(occurrences[1]);
    const result = await runtime.invoke("updateEvent", { ...updateArgs, recurrenceId, title: "Personal occurrence" });
    assert.equal(result.success, true, result.error);
    assert.equal(runtime.calls.modifies.length, 1);
  });

  it("allows [] on an already empty event while the preference is on", async () => {
    const runtime = loadCalendarRuntime();
    runtime.seedEvent();
    const result = await runtime.invoke("updateEvent", { ...updateArgs, attendees: [] });
    assert.equal(result.success, true, result.error);
    assert.deepEqual(runtime.calls.modifies[0].item.getAttendees(), []);
  });

  it("keeps attendee-free events writable without a calendar identity", async () => {
    for (const name of ["updateEvent", "deleteEvent"]) {
      const runtime = loadCalendarRuntime({ identityKey: null, identity: null });
      runtime.seedEvent();
      const result = await runtime.invoke(name, {
        ...updateArgs, ...(name === "updateEvent" ? { description: "Personal note" } : {}),
      });
      assert.equal(result.success, true, result.error);
      assert.equal(runtime.calls.modifies.length + runtime.calls.deletes.length, 1);
    }
  });

  for (const [label, options] of [
    ["resolved IMIP identity", { identityKey: null, identity: null, resolvedIdentity: { email: "calendar@example.com" } }],
    ["provider organizerId", { identityKey: null, identity: null, organizerId: "mailto:calendar@example.com" }],
  ]) {
    it(`recognizes self-only attendees through the ${label}`, async () => {
      for (const name of ["updateEvent", "deleteEvent"]) {
        const runtime = loadCalendarRuntime(options);
        runtime.seedEvent({ attendees: [{ id: "MAILTO:CALENDAR@EXAMPLE.COM" }] });
        const result = await runtime.invoke(name, {
          ...updateArgs, ...(name === "updateEvent" ? { title: "Personal appointment" } : {}),
        });
        assert.equal(result.success, true, result.error);
        assert.equal(runtime.calls.modifies.length + runtime.calls.deletes.length, 1);
      }
    });
  }

  it("uses resolved IMIP identity ahead of a stale organizerId for self-attendee exemption", async () => {
    const runtime = loadCalendarRuntime({
      resolvedIdentity: { email: "calendar@example.com" }, organizerId: "mailto:other@example.com",
    });
    runtime.seedEvent({ attendees: [{ id: "mailto:other@example.com" }] });
    const result = await runtime.invoke("deleteEvent", updateArgs);
    assert.match(result.error, /Block skipReview/i);
    assertNoWrites(runtime);
  });

  it("does not treat a non-email organizerId URI as a known self identity", async () => {
    const runtime = loadCalendarRuntime({ identityKey: null, identity: null, organizerId: "urn:uuid:calendar-user" });
    runtime.seedEvent({ attendees: [{ id: "urn:uuid:calendar-user" }] });
    const result = await runtime.invoke("deleteEvent", updateArgs);
    assert.match(result.error, /Block skipReview/i);
    assertNoWrites(runtime);
  });
});

describe("Production organizer ownership and retained attendee records", () => {
  for (const occurrence of [false, true]) {
    it(`preserves the existing organizer on ${occurrence ? "occurrence" : "series"} attendee replacement`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      const organizer = runtime.makeAttendee({
        id: "MAILTO:ORGANIZER@EXAMPLE.COM", commonName: "Custom organizer name",
        isOrganizer: true, role: "CHAIR", participationStatus: "ACCEPTED", sentBy: "mailto:delegate@example.com",
      });
      organizer.setProperty("X-PROVIDER-ORGANIZER", "retained");
      runtime.seedSeries({ organizer });
      const result = await runtime.invoke("updateEvent", {
        ...updateArgs, attendees: [attendee], ...(occurrence ? { recurrenceId } : {}),
      });
      assert.equal(result.success, true, result.error);
      const writtenOrganizer = runtime.calls.modifies[0].item.organizer;
      assert.equal(writtenOrganizer, organizer, "an existing organizer must not be replaced");
      assert.equal(writtenOrganizer.getProperty("X-PROVIDER-ORGANIZER"), "retained");
      assert.equal(writtenOrganizer.commonName, "Custom organizer name");
      assert.equal(writtenOrganizer.sentBy, "mailto:delegate@example.com");
    });

    for (const attendees of [[attendee], []]) {
      it(`rejects ${occurrence ? "occurrence" : "series"} attendee management by a non-organizer (${attendees.length ? "replace" : "clear"})`, async () => {
        const runtime = loadCalendarRuntime({ blockSkipReview: false });
        const organizer = runtime.makeAttendee({ id: "mailto:someone-else@example.com", commonName: "Actual organizer" });
        runtime.seedSeries({ organizer, attendees: [existingAttendee] });
        const result = await runtime.invoke("updateEvent", {
          ...updateArgs, attendees, ...(occurrence ? { recurrenceId } : {}),
        });
        assert.match(result.error, /organizer/i);
        assertNoWrites(runtime);
      });
    }

    it(`retains ${occurrence ? "occurrence" : "series"} attendee status and metadata by normalized email`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      const accepted = runtime.makeAttendee({
        id: "MAILTO:ALICE@EXAMPLE.COM", commonName: "Original Alice", role: "OPT-PARTICIPANT",
        participationStatus: "ACCEPTED", rsvp: false, userType: "RESOURCE", delegatedTo: "mailto:delegate@example.com",
      });
      accepted.setProperty("X-PROVIDER-METADATA", "accepted-data");
      const declined = runtime.makeAttendee({
        id: "mailto:declined@example.com", commonName: "Declined name", role: "REQ-PARTICIPANT",
        participationStatus: "DECLINED", rsvp: true, sentBy: "mailto:assistant@example.com",
      });
      declined.setProperty("X-PROVIDER-METADATA", "declined-data");
      const removed = runtime.makeAttendee({ id: "mailto:removed@example.com", participationStatus: "TENTATIVE" });
      runtime.seedSeries({ attendees: [accepted, declined, removed] });
      const result = await runtime.invoke("updateEvent", {
        ...updateArgs, ...(occurrence ? { recurrenceId } : {}),
        attendees: [
          { email: "alice@example.com" },
          { email: "MAILTO:DECLINED@EXAMPLE.COM", name: "Renamed", role: "optional" },
          { email: "new@example.com" },
        ],
      });
      assert.equal(result.success, true, result.error);
      const [retainedAccepted, retainedDeclined, added] = runtime.calls.modifies[0].item.getAttendees();
      assert.notEqual(retainedAccepted, accepted);
      assert.notEqual(retainedDeclined, declined);
      assert.equal(retainedAccepted.id, accepted.id);
      assert.equal(retainedDeclined.id, declined.id);
      assert.equal(retainedAccepted.participationStatus, "ACCEPTED");
      assert.equal(retainedAccepted.commonName, "Original Alice");
      assert.equal(retainedAccepted.role, "OPT-PARTICIPANT");
      assert.equal(retainedAccepted.rsvp, false);
      assert.equal(retainedAccepted.userType, "RESOURCE");
      assert.equal(retainedAccepted.delegatedTo, "mailto:delegate@example.com");
      assert.equal(retainedAccepted.getProperty("X-PROVIDER-METADATA"), "accepted-data");
      assert.equal(retainedDeclined.participationStatus, "DECLINED");
      assert.equal(retainedDeclined.commonName, "Renamed");
      assert.equal(retainedDeclined.role, "OPT-PARTICIPANT");
      assert.equal(retainedDeclined.rsvp, true);
      assert.equal(retainedDeclined.sentBy, "mailto:assistant@example.com");
      assert.equal(retainedDeclined.getProperty("X-PROVIDER-METADATA"), "declined-data");
      assert.equal(added.id, "mailto:new@example.com");
      assert.equal(added.participationStatus, "NEEDS-ACTION");
      assert.equal(added.role, "REQ-PARTICIPANT");
      assert.equal(declined.commonName, "Declined name", "source records must stay immutable until provider persistence");
      assert.equal(declined.role, "REQ-PARTICIPANT");
      retainedDeclined.setProperty("X-PROVIDER-METADATA", "clone-only");
      assert.equal(declined.getProperty("X-PROVIDER-METADATA"), "declined-data");
    });
  }

  for (const organizerLocation of ["master", "occurrence"]) {
    it(`rejects occurrence attendee changes when the ${organizerLocation} has a foreign organizer`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      const ownOrganizer = runtime.makeAttendee({ id: "mailto:organizer@example.com" });
      const { master, occurrences } = runtime.seedSeries({ organizer: ownOrganizer });
      (organizerLocation === "master" ? master : occurrences[0]).organizer = runtime.makeAttendee({ id: "mailto:foreign@example.com" });
      const result = await runtime.invoke("updateEvent", { ...updateArgs, recurrenceId, attendees: [attendee] });
      assert.match(result.error, /organizer/i);
      assertNoWrites(runtime);
    });
  }

  it("rejects attendee management when organizer ownership cannot be verified", async () => {
    const runtime = loadCalendarRuntime({ blockSkipReview: false, identity: null });
    runtime.seedEvent({ organizer: runtime.makeAttendee({ id: "mailto:organizer@example.com" }) });
    const result = await runtime.invoke("updateEvent", { ...updateArgs, attendees: [] });
    assert.match(result.error, /organizer|identity/i);
    assertNoWrites(runtime);
  });

  it("does not derive organizer authority from a non-email organizerId URI", async () => {
    const runtime = loadCalendarRuntime({
      blockSkipReview: false, identityKey: null, identity: null, organizerId: "urn:uuid:calendar-user",
    });
    runtime.seedEvent({ organizer: runtime.makeAttendee({ id: "urn:uuid:calendar-user" }) });
    const result = await runtime.invoke("updateEvent", { ...updateArgs, attendees: [attendee] });
    assert.match(result.error, /organizer|identity/i);
    assertNoWrites(runtime);
  });

  for (const [label, options] of [
    ["resolved IMIP identity", { identityKey: null, identity: null, resolvedIdentity: { email: "calendar@example.com" } }],
    ["provider organizerId", { identityKey: null, identity: null, organizerId: "MAILTO:CALENDAR@EXAMPLE.COM" }],
  ]) {
    it(`recognizes organizer ownership through the ${label}`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false, ...options });
      const organizer = runtime.makeAttendee({ id: "mailto:calendar@example.com", commonName: "Keep me" });
      runtime.seedEvent({ organizer });
      const result = await runtime.invoke("updateEvent", { ...updateArgs, attendees: [attendee] });
      assert.equal(result.success, true, result.error);
      assert.equal(runtime.calls.modifies[0].item.organizer, organizer);
    });
  }

  it("changes only explicitly supplied editable fields, including clearing a name", async () => {
    const runtime = loadCalendarRuntime({ blockSkipReview: false });
    const existing = runtime.makeAttendee({
      id: "mailto:alice@example.com", commonName: "Alice", role: "OPT-PARTICIPANT", participationStatus: "ACCEPTED",
    });
    runtime.seedEvent({ attendees: [existing] });
    const result = await runtime.invoke("updateEvent", {
      ...updateArgs, attendees: [{ email: "alice@example.com", name: "", role: "required" }],
    });
    assert.equal(result.success, true, result.error);
    const written = runtime.calls.modifies[0].item.getAttendees()[0];
    assert.equal(written.commonName, "");
    assert.equal(written.role, "REQ-PARTICIPANT");
    assert.equal(written.participationStatus, "ACCEPTED");
    assert.equal(existing.commonName, "Alice");
    assert.equal(existing.role, "OPT-PARTICIPANT");
  });

  it("treats null name and role as omitted when retaining an attendee", async () => {
    const runtime = loadCalendarRuntime({ blockSkipReview: false });
    runtime.seedEvent({ attendees: [runtime.makeAttendee({
      id: "mailto:alice@example.com", commonName: "Alice", role: "OPT-PARTICIPANT", participationStatus: "DECLINED",
    })] });
    const result = await runtime.invoke("updateEvent", {
      ...updateArgs, attendees: [{ email: "alice@example.com", name: null, role: null }],
    });
    assert.equal(result.success, true, result.error);
    const written = runtime.calls.modifies[0].item.getAttendees()[0];
    assert.equal(written.commonName, "Alice");
    assert.equal(written.role, "OPT-PARTICIPANT");
    assert.equal(written.participationStatus, "DECLINED");
  });
});

describe("Attendee validation runs against production schemas and handlers", () => {
  const invalidEmails = [
    "", " ", "alice", "alice@", "@example.com", "alice@@example.com",
    "Alice <alice@example.com>", "alice@example.com,bob@example.com",
    "alice@example.com;bob@example.com", "https:alice@example.com",
    "mailto:alice@example.com?subject=hello", "alice @example.com",
    "alice@example.com\r\nATTENDEE:mailto:bob@example.com",
    ...[0, 9, 10, 13, 31, 127, 128, 133, 159].map(code => `alice${String.fromCharCode(code)}@example.com`),
  ];

  for (const name of ["createEvent", "updateEvent"]) {
    it(`${name} rejects malformed email addresses and controls before persistence`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      runtime.seedEvent();
      for (const email of invalidEmails) {
        const result = await runtime.invoke(name, {
          ...(name === "createEvent" ? { ...createArgs, skipReview: true } : updateArgs),
          attendees: [{ email }],
        });
        assert.ok(result.error, `accepted invalid attendee: ${JSON.stringify(email)}`);
        assertNoWrites(runtime);
      }
    });

    it(`${name} rejects controls in attendee names`, async () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      runtime.seedEvent();
      for (const nameValue of ["Alice\r\nInjected", `Alice${String.fromCharCode(0)}`, `Alice${String.fromCharCode(133)}`]) {
        const result = await runtime.invoke(name, {
          ...(name === "createEvent" ? { ...createArgs, skipReview: true } : updateArgs),
          attendees: [{ email: "alice@example.com", name: nameValue }],
        });
        assert.ok(result.error);
        assertNoWrites(runtime);
      }
    });

    it(`${name} rejects malformed attendee arrays with the real nested schema`, () => {
      const runtime = loadCalendarRuntime({ blockSkipReview: false });
      const args = name === "createEvent" ? createArgs : updateArgs;
      for (const attendees of [[null], [{}], [{ email: null }], [{ email: 4 }], [{ email: "alice@example.com", role: "chair" }], "alice@example.com"]) {
        assert.ok(runtime.validateToolArgs(name, { ...args, attendees }).length > 0, JSON.stringify(attendees));
      }
      assert.equal(runtime.validateToolArgs(name, { ...args, attendees: [attendee] }).length, 0);
    });
  }

  it("validates every attendee before a provider write", async () => {
    const runtime = loadCalendarRuntime({ blockSkipReview: false });
    const item = runtime.seedEvent({ attendees: [existingAttendee] });
    const result = await runtime.invoke("updateEvent", {
      ...updateArgs, attendees: [attendee, { email: "not-an-email" }],
    });
    assert.ok(result.error);
    assertNoWrites(runtime);
    assert.deepEqual(item.getAttendees(), [existingAttendee]);
  });
});

describe("Production event attendee serialization", () => {
  it("reduces organizer/attendee fields and includes participation status", () => {
    const runtime = loadCalendarRuntime();
    const item = runtime.makeEvent({
      organizer: { id: "mailto:boss@example.com", commonName: "Boss", role: "CHAIR" },
      attendees: [existingAttendee, null],
      invitedAttendee: { participationStatus: "DECLINED" },
    });
    const result = runtime.formatEvent(item, runtime.calendar);
    assert.deepEqual(plain(result.organizer), { id: "mailto:boss@example.com", commonName: "Boss" });
    assert.deepEqual(plain(result.attendees), [{ ...existingAttendee, commonName: "" }]);
    assert.equal(result.myParticipationStatus, "DECLINED");
    assert.equal(runtime.formatAttendee(null), null);
    assert.deepEqual(plain(runtime.formatAttendee({})), { id: "", commonName: "", participationStatus: "", role: "" });
    assert.equal(runtime.formatOrganizer(null), null);
  });

  it("degrades provider errors independently", () => {
    const runtime = loadCalendarRuntime();
    const item = runtime.makeEvent();
    Object.defineProperty(item, "organizer", { get() { throw new Error("provider failure"); } });
    item.getAttendees = () => { throw new Error("provider failure"); };
    runtime.cal.itip.getInvitedAttendee = () => { throw new Error("provider failure"); };
    const result = runtime.formatEvent(item, runtime.calendar);
    assert.equal(result.organizer, null);
    assert.deepEqual(plain(result.attendees), []);
    assert.equal(result.attendeeCount, 0);
    assert.equal(result.myParticipationStatus, "");
  });

  it("bounds attendees on every expanded occurrence and returns the full count", async () => {
    const runtime = loadCalendarRuntime();
    const attendees = Array.from({ length: 140 }, (_, index) => ({ id: `mailto:user${index}@example.com` }));
    runtime.seedSeries({ attendees });
    const result = await runtime.invoke("listEvents", {
      calendarId: "calendar-1", startDate: "2026-09-28", endDate: "2026-10-06",
    });
    assert.ok(Array.isArray(result), result.error);
    assert.equal(result.length, 2);
    for (const occurrence of result) {
      assert.equal(occurrence.attendees.length, 100);
      assert.equal(occurrence.attendeeCount, 140);
      assert.equal(occurrence.attendees[99].id, "mailto:user99@example.com");
      assert.equal(occurrence.isRecurring, true);
      assert.equal(occurrence.recurrence, "FREQ=WEEKLY");
      assert.ok(occurrence.recurrenceId);
    }
  });
});
