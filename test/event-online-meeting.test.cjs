"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

const PROVIDER_PROPERTY = "X-ONLINE-MEETING-PROVIDER";
const TEAMS_URL_PROPERTY = "X-MICROSOFT-SKYPETEAMSMEETINGURL";

// Mirrors updateEvent's onlineMeeting decision tree. XPCOM is unavailable in
// node:test, so this covers only the property mutation rules.
function applyUpdateOnlineMeeting(newItem, onlineMeeting, changes) {
  if (onlineMeeting !== undefined) {
    if (onlineMeeting) {
      newItem.setProperty(PROVIDER_PROPERTY, "TeamsForBusiness");
    } else {
      newItem.deleteProperty(PROVIDER_PROPERTY);
      newItem.deleteProperty(TEAMS_URL_PROPERTY);
    }
    changes.push("onlineMeeting");
  }
}

function formatOnlineMeetingURL(item) {
  return item.getProperty(TEAMS_URL_PROPERTY) || null;
}

function createMockItem(initialProperties = {}) {
  return {
    properties: new Map(Object.entries(initialProperties)),
    setPropertyCalls: [],
    deletePropertyCalls: [],
    getProperty(name) {
      return this.properties.get(name);
    },
    setProperty(name, value) {
      this.setPropertyCalls.push([name, value]);
      this.properties.set(name, value);
    },
    deleteProperty(name) {
      this.deletePropertyCalls.push(name);
      this.properties.delete(name);
    },
  };
}

describe("updateEvent onlineMeeting decision tree", () => {
  it("true requests Teams meeting generation", () => {
    const item = createMockItem();
    const changes = [];

    applyUpdateOnlineMeeting(item, true, changes);

    assert.deepEqual(item.setPropertyCalls, [[PROVIDER_PROPERTY, "TeamsForBusiness"]]);
    assert.deepEqual(item.deletePropertyCalls, []);
    assert.deepEqual(changes, ["onlineMeeting"]);
  });

  it("false removes Teams meeting properties", () => {
    const item = createMockItem({
      [PROVIDER_PROPERTY]: "TeamsForBusiness",
      [TEAMS_URL_PROPERTY]: "https://teams.example/join",
    });
    const changes = [];

    applyUpdateOnlineMeeting(item, false, changes);

    assert.equal(item.getProperty(PROVIDER_PROPERTY), undefined);
    assert.equal(item.getProperty(TEAMS_URL_PROPERTY), undefined);
    assert.deepEqual(item.deletePropertyCalls, [PROVIDER_PROPERTY, TEAMS_URL_PROPERTY]);
    assert.deepEqual(changes, ["onlineMeeting"]);
  });

  it("undefined is no-op", () => {
    const item = createMockItem({
      [PROVIDER_PROPERTY]: "TeamsForBusiness",
      [TEAMS_URL_PROPERTY]: "https://teams.example/join",
    });
    const changes = [];

    applyUpdateOnlineMeeting(item, undefined, changes);

    assert.equal(item.getProperty(PROVIDER_PROPERTY), "TeamsForBusiness");
    assert.equal(item.getProperty(TEAMS_URL_PROPERTY), "https://teams.example/join");
    assert.deepEqual(item.setPropertyCalls, []);
    assert.deepEqual(item.deletePropertyCalls, []);
    assert.deepEqual(changes, []);
  });

  it("null removes Teams meeting properties", () => {
    const item = createMockItem({
      [PROVIDER_PROPERTY]: "TeamsForBusiness",
      [TEAMS_URL_PROPERTY]: "https://teams.example/join",
    });
    const changes = [];

    applyUpdateOnlineMeeting(item, null, changes);

    assert.equal(item.getProperty(PROVIDER_PROPERTY), undefined);
    assert.equal(item.getProperty(TEAMS_URL_PROPERTY), undefined);
    assert.deepEqual(item.deletePropertyCalls, [PROVIDER_PROPERTY, TEAMS_URL_PROPERTY]);
    assert.deepEqual(changes, ["onlineMeeting"]);
  });
});

describe("formatEvent onlineMeetingURL mapping", () => {
  it("maps the Teams meeting URL property", () => {
    const item = createMockItem({
      [TEAMS_URL_PROPERTY]: "https://teams.example/join",
    });

    assert.equal(formatOnlineMeetingURL(item), "https://teams.example/join");
  });

  it("returns null when the Teams meeting URL property is missing", () => {
    assert.equal(formatOnlineMeetingURL(createMockItem()), null);
  });
});
