"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadContactFieldHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN CONTACT FIELD HELPERS";
  const endMarker = "// END CONTACT FIELD HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker);
  assert.ok(start >= 0, "contact field helper start marker missing");
  assert.ok(end > start, "contact field helper end marker missing");

  const snippet = source.slice(start, end);
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(
    `${snippet}
this.contactFieldHelpers = {
  normalizeContactBirthday,
  contactBirthdayToVCard,
  contactBirthdayToFlatParts,
  validateContactFields,
  readContactFields,
  formatContact,
  applyContactFields,
  shouldSynthesizePhoneDisplayName,
};`,
    sandbox
  );
  return sandbox.contactFieldHelpers;
}

const helpers = loadContactFieldHelpers();

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

class FakeVCardPropertyEntry {
  constructor(name, params, type, value) {
    this.name = name;
    this.params = params;
    this.type = type;
    this.value = value;
  }
}

class FakeVCardProperties {
  constructor(entries = []) {
    this.entries = entries;
  }

  getAllEntries(name) {
    return this.entries.filter(entry => entry.name === name);
  }

  getFirstEntry(name) {
    return this.getAllEntries(name)[0] || null;
  }

  getFirstValue(name) {
    return this.getFirstEntry(name)?.value ?? null;
  }

  addEntry(entry) {
    this.entries.push(entry);
  }

  removeEntry(entry) {
    const index = this.entries.indexOf(entry);
    if (index >= 0) this.entries.splice(index, 1);
  }

  clearValues(name) {
    this.entries = this.entries.filter(entry => entry.name !== name);
  }
}

function makeFlatCard(properties = {}) {
  const stored = new Map(Object.entries(properties));
  return {
    UID: "flat-uid",
    supportsVCard: false,
    displayName: "Flat Contact",
    primaryEmail: "flat@example.com",
    firstName: "Flat",
    lastName: "Contact",
    getProperty(name, fallback = "") {
      return stored.has(name) ? stored.get(name) : fallback;
    },
    setProperty(name, value) {
      stored.set(name, value);
    },
    stored,
  };
}

describe("contact birthday normalization", () => {
  it("accepts API and serialized full-date forms", () => {
    assert.equal(helpers.normalizeContactBirthday("1980-04-15"), "1980-04-15");
    assert.equal(helpers.normalizeContactBirthday("19800415"), "1980-04-15");
    assert.equal(helpers.contactBirthdayToVCard("1980-04-15"), "1980-04-15");
  });

  it("accepts yearless birthdays and emits canonical forms", () => {
    assert.equal(helpers.normalizeContactBirthday("--04-15"), "--04-15");
    assert.equal(helpers.normalizeContactBirthday("--0415"), "--04-15");
    assert.equal(helpers.contactBirthdayToVCard("--04-15"), "--04-15");
    assert.deepEqual(plain(helpers.contactBirthdayToFlatParts("--04-15")), {
      year: "",
      month: "04",
      day: "15",
    });
  });

  it("rejects impossible and malformed dates", () => {
    assert.equal(helpers.normalizeContactBirthday("2023-02-29"), null);
    assert.equal(helpers.normalizeContactBirthday("--13-01"), null);
    assert.equal(helpers.normalizeContactBirthday("04-15"), null);
    assert.equal(helpers.normalizeContactBirthday(" "), null);
    assert.equal(helpers.normalizeContactBirthday("--02-29"), "--02-29");
    assert.equal(helpers.normalizeContactBirthday("2000-02-29"), "2000-02-29");
  });
});

describe("contact handler payload validation", () => {
  it("accepts a phone-only create payload", () => {
    assert.equal(helpers.validateContactFields({
      phones: [{ type: "mobile", number: "+48 123 456 789" }],
    }, true), null);
  });

  it("rejects a wholly empty create payload", () => {
    assert.match(
      helpers.validateContactFields({ email: "", phones: [], addresses: [] }, true),
      /non-empty contact field/
    );
  });

  it("deep-validates phone objects", () => {
    assert.match(
      helpers.validateContactFields({ phones: [{ type: "cell", number: "123" }] }),
      /phones\[0\]\.type/
    );
    assert.match(
      helpers.validateContactFields({ phones: [{ type: "home", number: " " }] }),
      /non-empty string/
    );
    assert.match(
      helpers.validateContactFields({ phones: [{ type: "home", number: "123", label: "x" }] }),
      /Unknown phones\[0\] property/
    );
  });

  it("deep-validates postal address objects", () => {
    assert.equal(helpers.validateContactFields({
      addresses: [{ type: "work", street: "1 Main St", city: "Warsaw" }],
    }), null);
    assert.match(
      helpers.validateContactFields({ addresses: [{ type: "other", city: "Warsaw" }] }),
      /addresses\[0\]\.type/
    );
    assert.match(
      helpers.validateContactFields({ addresses: [{ type: "home" }] }),
      /at least one non-empty address field/
    );
    assert.match(
      helpers.validateContactFields({ addresses: [{ type: "home", city: 42 }] }),
      /addresses\[0\]\.city must be a string/
    );
  });

  it("validates scalar types and birthday calendar values", () => {
    assert.match(helpers.validateContactFields({ email: null }), /email must be a string/);
    assert.match(helpers.validateContactFields({ birthday: "2023-02-29" }), /birthday/);
    assert.equal(helpers.validateContactFields({ birthday: "--02-29" }), null);
    assert.equal(helpers.validateContactFields({ birthday: "" }), null);
  });
});

describe("contact field reads", () => {
  it("formats vCard-backed cards with the full contact shape", () => {
    const vCardProperties = new FakeVCardProperties([
      new FakeVCardPropertyEntry("tel", { type: "WORK" }, "text", "+48 111"),
      new FakeVCardPropertyEntry("tel", { type: ["voice", "cell"] }, "uri", "tel:+48 222"),
      new FakeVCardPropertyEntry(
        "adr",
        { type: "home" },
        "text",
        ["PO 1", "Apt 2", ["Main St", "Floor 3"], "Warsaw", "Mazovia", "00-001", "Poland"]
      ),
      new FakeVCardPropertyEntry("org", {}, "text", ["Example Corp", "Research"]),
      new FakeVCardPropertyEntry("title", {}, "text", "Engineer"),
      new FakeVCardPropertyEntry("note", {}, "text", ["line one", "line two"]),
      new FakeVCardPropertyEntry("bday", {}, "date", "--0415"),
    ]);
    const card = {
      UID: "vcard-uid",
      supportsVCard: true,
      vCardProperties,
      displayName: "Ada Example",
      primaryEmail: "ada@example.com",
      firstName: "Ada",
      lastName: "Example",
    };

    assert.deepEqual(plain(helpers.formatContact(card, {
      dirName: "Personal",
      URI: "jsaddrbook://abook.sqlite",
    })), {
      id: "vcard-uid",
      displayName: "Ada Example",
      email: "ada@example.com",
      firstName: "Ada",
      lastName: "Example",
      phones: [
        { type: "work", number: "+48 111" },
        { type: "mobile", number: "+48 222" },
      ],
      addresses: [{
        type: "home",
        poBox: "PO 1",
        street2: "Apt 2",
        street: "Main St Floor 3",
        city: "Warsaw",
        region: "Mazovia",
        postalCode: "00-001",
        country: "Poland",
      }],
      organization: "Example Corp",
      title: "Engineer",
      note: "line one,line two",
      birthday: "--04-15",
      addressBook: "Personal",
      addressBookId: "jsaddrbook://abook.sqlite",
    });
  });

  it("reads the same fields from legacy flat cards", () => {
    const card = makeFlatCard({
      HomePhone: "555-0100",
      CellularNumber: "555-0101",
      WorkAddress: "1 Main St",
      WorkCity: "Warsaw",
      WorkCountry: "Poland",
      Company: "Legacy Ltd",
      JobTitle: "Director",
      Notes: "legacy note",
      BirthYear: "1980",
      BirthMonth: "4",
      BirthDay: "5",
    });
    const details = plain(helpers.readContactFields(card));
    assert.deepEqual(details.phones, [
      { type: "home", number: "555-0100" },
      { type: "mobile", number: "555-0101" },
    ]);
    assert.deepEqual(details.addresses, [{
      type: "work",
      street: "1 Main St",
      city: "Warsaw",
      country: "Poland",
    }]);
    assert.equal(details.organization, "Legacy Ltd");
    assert.equal(details.title, "Director");
    assert.equal(details.note, "legacy note");
    assert.equal(details.birthday, "1980-04-05");
  });
});

describe("contact field writes", () => {
  it("selectively updates vCard entries and preserves unrelated properties", () => {
    const org = new FakeVCardPropertyEntry("org", {}, "text", ["Old Corp", "Research", "Team A"]);
    const secondOrg = new FakeVCardPropertyEntry("org", { type: "other" }, "text", ["Affiliation"]);
    const url = new FakeVCardPropertyEntry("url", { type: "work" }, "uri", "https://example.com");
    const im = new FakeVCardPropertyEntry("impp", {}, "uri", "xmpp:ada@example.com");
    const custom = new FakeVCardPropertyEntry("x-custom", {}, "text", "keep-me");
    const vCardProperties = new FakeVCardProperties([
      new FakeVCardPropertyEntry("tel", { type: "home" }, "text", "old-phone"),
      new FakeVCardPropertyEntry("adr", { type: "home" }, "text", ["", "", "Old St", "", "", "", ""]),
      org,
      secondOrg,
      new FakeVCardPropertyEntry("title", {}, "text", "Old title"),
      new FakeVCardPropertyEntry("note", {}, "text", "Old note"),
      new FakeVCardPropertyEntry("bday", {}, "date", "19700101"),
      url,
      im,
      custom,
    ]);
    const card = {
      supportsVCard: true,
      vCardProperties,
      primaryEmail: "old@example.com",
      displayName: "Old Name",
      firstName: "Old",
      lastName: "Name",
    };

    helpers.applyContactFields(card, {
      email: "new@example.com",
      phones: [
        { type: "work", number: " 555-1000 " },
        { type: "mobile", number: "555-1001" },
      ],
      addresses: [{ type: "work", street: "New St", city: "Warsaw" }],
      organization: "New Corp",
      title: "",
      note: "first line\nsecond line",
      birthday: "--04-15",
    }, FakeVCardPropertyEntry);

    assert.equal(card.primaryEmail, "new@example.com");
    assert.deepEqual(vCardProperties.getAllEntries("tel").map(entry => ({
      type: entry.params.type,
      value: entry.value,
    })), [
      { type: "work", value: "555-1000" },
      { type: "cell", value: "555-1001" },
    ]);
    assert.deepEqual(plain(vCardProperties.getFirstValue("adr")), ["", "", "New St", "Warsaw", "", "", ""]);
    assert.strictEqual(vCardProperties.getFirstEntry("org"), org);
    assert.deepEqual(plain(org.value), ["New Corp", "Research", "Team A"]);
    assert.strictEqual(vCardProperties.getAllEntries("org")[1], secondOrg);
    assert.equal(vCardProperties.getAllEntries("title").length, 0);
    assert.equal(vCardProperties.getFirstValue("note"), "first line\nsecond line");
    assert.equal(vCardProperties.getFirstValue("bday"), "--04-15");
    assert.ok(vCardProperties.entries.includes(url));
    assert.ok(vCardProperties.entries.includes(im));
    assert.ok(vCardProperties.entries.includes(custom));
  });

  it("keeps later ORG components when organization is cleared", () => {
    const org = new FakeVCardPropertyEntry("org", {}, "text", ["Old Corp", "Department"]);
    const card = {
      supportsVCard: true,
      vCardProperties: new FakeVCardProperties([org]),
    };
    helpers.applyContactFields(card, { organization: "" }, FakeVCardPropertyEntry);
    assert.deepEqual(plain(org.value), ["", "Department"]);
  });

  it("clears the exposed organization while preserving additional ORG entries", () => {
    const primaryOrg = new FakeVCardPropertyEntry("org", {}, "text", "Old Corp");
    const affiliation = new FakeVCardPropertyEntry("org", {}, "text", ["Affiliation"]);
    const card = {
      supportsVCard: true,
      vCardProperties: new FakeVCardProperties([primaryOrg, affiliation]),
    };
    helpers.applyContactFields(card, { organization: "" }, FakeVCardPropertyEntry);
    assert.deepEqual(plain(primaryOrg.value), [""]);
    assert.strictEqual(card.vCardProperties.getAllEntries("org")[1], affiliation);
  });

  it("uses empty values to clear vCard email, phone, and address fields", () => {
    const vCardProperties = new FakeVCardProperties([
      new FakeVCardPropertyEntry("email", { pref: "1" }, "text", "primary@example.com"),
      new FakeVCardPropertyEntry("email", {}, "text", "secondary@example.com"),
      new FakeVCardPropertyEntry("tel", { type: "home" }, "text", "555"),
      new FakeVCardPropertyEntry("adr", { type: "home" }, "text", ["", "", "Street"]),
      new FakeVCardPropertyEntry("url", {}, "uri", "https://example.com"),
    ]);
    helpers.applyContactFields({ supportsVCard: true, vCardProperties }, {
      email: "",
      phones: [],
      addresses: [],
    }, FakeVCardPropertyEntry);
    assert.equal(vCardProperties.getAllEntries("email").length, 0);
    assert.equal(vCardProperties.getAllEntries("tel").length, 0);
    assert.equal(vCardProperties.getAllEntries("adr").length, 0);
    assert.equal(vCardProperties.getAllEntries("url").length, 1);
  });

  it("maps writes to legacy flat properties without touching department", () => {
    const card = makeFlatCard({ Department: "Keep Department", WorkPhone: "old" });
    helpers.applyContactFields(card, {
      phones: [{ type: "fax", number: " 555-0199 " }],
      addresses: [{ type: "home", street: "Home St", postalCode: "00-001" }],
      organization: "Flat Corp",
      title: "Manager",
      note: "flat note",
      birthday: "1980-04-15",
    }, FakeVCardPropertyEntry);

    assert.equal(card.stored.get("WorkPhone"), "");
    assert.equal(card.stored.get("FaxNumber"), "555-0199");
    assert.equal(card.stored.get("HomeAddress"), "Home St");
    assert.equal(card.stored.get("HomeZipCode"), "00-001");
    assert.equal(card.stored.get("Company"), "Flat Corp");
    assert.equal(card.stored.get("Department"), "Keep Department");
    assert.equal(card.stored.get("JobTitle"), "Manager");
    assert.equal(card.stored.get("Notes"), "flat note");
    assert.equal(card.stored.get("BirthYear"), "1980");
    assert.equal(card.stored.get("BirthMonth"), "04");
    assert.equal(card.stored.get("BirthDay"), "15");
  });
});

describe("phone-only display name synthesis", () => {
  it("synthesizes only for a truly phone-only payload", () => {
    assert.equal(helpers.shouldSynthesizePhoneDisplayName({
      phones: [{ type: "home", number: "555-0100" }],
    }), true);
    assert.equal(helpers.shouldSynthesizePhoneDisplayName({
      displayName: "Ada",
      phones: [{ type: "home", number: "555-0100" }],
    }), false);
    assert.equal(helpers.shouldSynthesizePhoneDisplayName({
      organization: "Example Corp",
      phones: [{ type: "home", number: "555-0100" }],
    }), false);
    assert.equal(helpers.shouldSynthesizePhoneDisplayName({
      phones: [{ type: "home", number: "555-0100" }],
      addresses: [{ type: "home", city: "Warsaw" }],
    }), false);
  });
});

describe("contact tool wiring", () => {
  const apiSource = fs.readFileSync(
    path.resolve(__dirname, "../extension/mcp_server/api.js"),
    "utf8"
  );

  it("defines getContact as a UID-based read tool", () => {
    const start = apiSource.indexOf('name: "getContact"');
    const end = apiSource.indexOf('name: "createContact"', start);
    assert.ok(start >= 0 && end > start, "getContact tool definition missing");
    const definition = apiSource.slice(start, end);
    assert.match(definition, /group: "contacts", crud: "read"/);
    assert.match(definition, /contactId: \{ type: "string"/);
    assert.match(definition, /required: \["contactId"\]/);
  });

  it("dispatches getContact through the existing UID lookup and shared formatter", () => {
    assert.match(apiSource, /case "getContact":\s*return getContact\(args\.contactId\);/);
    const start = apiSource.indexOf("function getContact(contactId)");
    const end = apiSource.indexOf("function createContact(", start);
    const handler = apiSource.slice(start, end);
    assert.match(handler, /findContactByUID\(contactId\)/);
    assert.match(handler, /formatContact\(found\.card, found\.book\)/);
  });

  it("uses the shared full-contact formatter in searchContacts", () => {
    assert.match(apiSource, /results\.push\(formatContact\(card, book\)\)/);
    assert.match(apiSource, /if \(card\.isMailList\) continue;\s*if \(card\.UID === contactId\)/);
  });
});
