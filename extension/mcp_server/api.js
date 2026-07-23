// XPCOM globals (ExtensionCommon, ChromeUtils, Services, Cc, Ci) are
// declared in eslint.config.mjs's extension/ file group. Per-file
// /* global */ comment triggered no-redeclare.
"use strict";

/**
 * Thunderbird MCP Server Extension
 * Exposes email, calendar, and contacts via MCP protocol over HTTP.
 *
 * Architecture: MCP Client <-> mcp-bridge.cjs (stdio<->HTTP) <-> This extension (port 8765)
 *
 * Key quirks documented inline:
 * - MIME header decoding (mime2Decoded* properties)
 * - HTML body charset handling (emojis require HTML entity encoding)
 * - Compose window body preservation (must use New type, not Reply)
 * - IMAP folder sync (msgDatabase may be stale)
 */

const resProto = Cc[
  "@mozilla.org/network/protocol;1?name=resource"
].getService(Ci.nsISubstitutingProtocolHandler);

const MCP_DEFAULT_PORT = 8765;
const MCP_MAX_PORT_ATTEMPTS = 10;
const CONNECTION_FILE_REFRESH_MS = 30 * 1000;

// Versions of the MCP protocol this server understands. Behavior never depends
// on the negotiated version inside Thunderbird (the bridge intercepts initialize
// for clients), but for the rare case a client talks directly to the HTTP server
// we still need a spec-compliant negotiated value. Keep in sync with mcp-bridge.cjs.
const MCP_SUPPORTED_PROTOCOL_VERSIONS = new Set([
  "2024-10-07",
  "2024-11-05",
  "2025-03-26",
  "2025-06-18",
  "2025-11-25",
]);
const MCP_LATEST_PROTOCOL_VERSION = "2025-11-25";

// Bridged into serverInfo.version on initialize. Resolved lazily from the
// extension manifest so a single bump in extension/manifest.json propagates here.
let _cachedExtVersion = null;
function isListenAllEnabled() {
  try { return Services.prefs.getBoolPref(PREF_LISTEN_ALL, false); } catch { return false; }
}

function stopConnectionInfoRefreshTimer() {
  if (globalThis.__tbMcpConnectionInfoRefreshTimer) {
    try {
      globalThis.__tbMcpConnectionInfoRefreshTimer.cancel();
    } catch (e) {
      console.warn("thunderbird-mcp: failed to stop connection info refresh timer:", e);
    }
    globalThis.__tbMcpConnectionInfoRefreshTimer = null;
  }
}

// BEGIN CONNECTION INFO REFRESH HELPERS
function ensureFreshConnectionInfo({
  port,
  token,
  expectedPid,
  readConnectionInfo,
  writeConnectionInfo,
  onCheckError,
}) {
  try {
    const current = readConnectionInfo();
    const data = current && current.data;
    if (
      data &&
      data.port === port &&
      data.token === token &&
      data.pid === expectedPid
    ) {
      return current.path;
    }
  } catch (e) {
    if (typeof onCheckError === "function") {
      onCheckError(e);
    }
  }
  return writeConnectionInfo(port, token);
}
// END CONNECTION INFO REFRESH HELPERS

// BEGIN CONTACT FIELD HELPERS
// BEGIN CONTACT FIELD CONSTANTS
const CONTACT_PHONE_TYPES = ["work", "home", "mobile", "fax", "pager"];
const CONTACT_ADDRESS_TYPES = ["home", "work"];
const CONTACT_ADDRESS_FIELDS = [
  "poBox",
  "street2",
  "street",
  "city",
  "region",
  "postalCode",
  "country",
];
const CONTACT_SCALAR_FIELDS = [
  "email",
  "displayName",
  "firstName",
  "lastName",
  "organization",
  "title",
  "note",
  "birthday",
];
const CONTACT_PHONE_FLAT_PROPERTIES = {
  work: "WorkPhone",
  home: "HomePhone",
  mobile: "CellularNumber",
  fax: "FaxNumber",
  pager: "PagerNumber",
};
const CONTACT_ADDRESS_FLAT_PROPERTIES = {
  home: [
    "HomePOBox",
    "HomeAddress2",
    "HomeAddress",
    "HomeCity",
    "HomeState",
    "HomeZipCode",
    "HomeCountry",
  ],
  work: [
    "WorkPOBox",
    "WorkAddress2",
    "WorkAddress",
    "WorkCity",
    "WorkState",
    "WorkZipCode",
    "WorkCountry",
  ],
};
// END CONTACT FIELD CONSTANTS

function contactValueToString(value, separator = ",") {
  if (Array.isArray(value)) {
    return value.map(part => {
      if (Array.isArray(part)) return part.join(" ");
      return part === null || part === undefined ? "" : String(part);
    }).join(separator);
  }
  return value === null || value === undefined ? "" : String(value);
}

function contactStructuredValuePart(value) {
  if (Array.isArray(value)) return value.join(" ");
  return value === null || value === undefined ? "" : String(value);
}

function getContactVCardTypes(entry) {
  const type = entry?.params?.type;
  if (!type) return [];
  const types = Array.isArray(type) ? type : [type];
  return types
    .filter(value => typeof value === "string")
    .map(value => value.toLowerCase());
}

function getContactPhoneType(entry) {
  const vCardType = getContactVCardTypes(entry)
    .find(type => ["home", "work", "cell", "fax", "pager"].includes(type));
  if (vCardType === "cell") return "mobile";
  return vCardType || "work";
}

function getContactAddressType(entry) {
  return getContactVCardTypes(entry).includes("home") ? "home" : "work";
}

function isContactLeapYear(year) {
  return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
}

function isValidContactBirthdayParts(year, month, day) {
  const numericMonth = Number(month);
  const numericDay = Number(day);
  if (!Number.isInteger(numericMonth) || numericMonth < 1 || numericMonth > 12) {
    return false;
  }
  const numericYear = year ? Number(year) : 2000;
  if (year && (!/^\d{4}$/.test(year) || numericYear < 1)) return false;
  const daysInMonth = [
    31,
    isContactLeapYear(numericYear) ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31,
  ];
  return Number.isInteger(numericDay) && numericDay >= 1 && numericDay <= daysInMonth[numericMonth - 1];
}

/**
 * Normalize API and serialized vCard birthday forms to YYYY-MM-DD/--MM-DD.
 * Returns null for invalid input and an empty string for an explicit clear.
 */
function normalizeContactBirthday(value) {
  if (typeof value !== "string") return null;
  if (value === "") return "";
  const trimmed = value.trim();
  if (!trimmed) return null;

  let match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed);
  if (!match) match = /^(\d{4})(\d{2})(\d{2})$/.exec(trimmed);
  if (match) {
    const [, year, month, day] = match;
    return isValidContactBirthdayParts(year, month, day)
      ? `${year}-${month}-${day}`
      : null;
  }

  match = /^--(\d{2})-(\d{2})$/.exec(trimmed);
  if (!match) match = /^--(\d{2})(\d{2})$/.exec(trimmed);
  if (match) {
    const [, month, day] = match;
    return isValidContactBirthdayParts("", month, day)
      ? `--${month}-${day}`
      : null;
  }
  return null;
}

function contactBirthdayToVCard(value) {
  // VCardPropertyEntry stores ICAL's normalized jCard value. ICAL removes the
  // separators as needed when the card is serialized.
  return normalizeContactBirthday(value) || "";
}

function contactBirthdayToFlatParts(value) {
  const normalized = normalizeContactBirthday(value);
  if (!normalized) return { year: "", month: "", day: "" };
  if (normalized.startsWith("--")) {
    return {
      year: "",
      month: normalized.slice(2, 4),
      day: normalized.slice(5, 7),
    };
  }
  return {
    year: normalized.slice(0, 4),
    month: normalized.slice(5, 7),
    day: normalized.slice(8, 10),
  };
}

function getContactCardProperty(card, name) {
  try {
    const value = card.getProperty(name, "");
    return value === null || value === undefined ? "" : String(value);
  } catch {
    return "";
  }
}

function readVCardContactFields(card) {
  const vCardProperties = card.vCardProperties;
  const phones = vCardProperties.getAllEntries("tel").map(entry => ({
    type: getContactPhoneType(entry),
    number: normalizeContactPhoneValue(entry.value),
  })).filter(phone => phone.number);

  const addresses = [];
  for (const entry of vCardProperties.getAllEntries("adr")) {
    const rawParts = Array.isArray(entry.value) ? entry.value : [entry.value];
    const parts = CONTACT_ADDRESS_FIELDS.map((field, index) =>
      contactStructuredValuePart(rawParts[index])
    );
    if (!parts.some(Boolean)) continue;
    const address = { type: getContactAddressType(entry) };
    for (let i = 0; i < CONTACT_ADDRESS_FIELDS.length; i++) {
      if (parts[i]) address[CONTACT_ADDRESS_FIELDS[i]] = parts[i];
    }
    addresses.push(address);
  }

  const organizationValue = vCardProperties.getFirstValue("org");
  const organization = Array.isArray(organizationValue)
    ? contactStructuredValuePart(organizationValue[0])
    : contactValueToString(organizationValue);
  const birthdayValue = contactValueToString(vCardProperties.getFirstValue("bday"));

  return {
    phones,
    addresses,
    organization,
    title: contactValueToString(vCardProperties.getFirstValue("title")),
    note: contactValueToString(vCardProperties.getFirstValue("note")),
    birthday: normalizeContactBirthday(birthdayValue) || "",
  };
}

function readFlatContactFields(card) {
  const phones = [];
  for (const type of CONTACT_PHONE_TYPES) {
    const number = getContactCardProperty(card, CONTACT_PHONE_FLAT_PROPERTIES[type]);
    if (number) phones.push({ type, number });
  }

  const addresses = [];
  for (const type of CONTACT_ADDRESS_TYPES) {
    const properties = CONTACT_ADDRESS_FLAT_PROPERTIES[type];
    const values = properties.map(name => getContactCardProperty(card, name));
    if (!values.some(Boolean)) continue;
    const address = { type };
    for (let i = 0; i < CONTACT_ADDRESS_FIELDS.length; i++) {
      if (values[i]) address[CONTACT_ADDRESS_FIELDS[i]] = values[i];
    }
    addresses.push(address);
  }

  const year = getContactCardProperty(card, "BirthYear").trim();
  const rawMonth = getContactCardProperty(card, "BirthMonth").trim();
  const rawDay = getContactCardProperty(card, "BirthDay").trim();
  let birthday = "";
  if (rawMonth && rawDay) {
    const month = rawMonth.padStart(2, "0");
    const day = rawDay.padStart(2, "0");
    birthday = normalizeContactBirthday(year ? `${year}-${month}-${day}` : `--${month}-${day}`) || "";
  }

  return {
    phones,
    addresses,
    organization: getContactCardProperty(card, "Company"),
    title: getContactCardProperty(card, "JobTitle"),
    note: getContactCardProperty(card, "Notes"),
    birthday,
  };
}

function readContactFields(card) {
  if (card.supportsVCard) {
    try {
      return readVCardContactFields(card);
    } catch {
      // A malformed vCard should not prevent the rest of an address book from
      // being searched. Legacy properties are the best available fallback.
    }
  }
  return readFlatContactFields(card);
}

function formatContact(card, book) {
  const details = readContactFields(card);
  return {
    id: card.UID,
    displayName: card.displayName || "",
    email: card.primaryEmail || "",
    firstName: card.firstName || "",
    lastName: card.lastName || "",
    phones: details.phones,
    addresses: details.addresses,
    organization: details.organization,
    title: details.title,
    note: details.note,
    birthday: details.birthday,
    addressBook: book.dirName,
    addressBookId: book.URI,
  };
}

function contactFieldsHaveContent(fields) {
  if (CONTACT_SCALAR_FIELDS.some(name =>
    typeof fields[name] === "string" && fields[name].trim().length > 0
  )) {
    return true;
  }
  return (Array.isArray(fields.phones) && fields.phones.length > 0) ||
    (Array.isArray(fields.addresses) && fields.addresses.length > 0);
}

/**
 * Deep validation used inside contact handlers before any card is mutated.
 * Returns an error string, or null for a valid payload.
 */
function validateContactFields(fields, requireContent = false) {
  for (const name of CONTACT_SCALAR_FIELDS) {
    if (fields[name] !== undefined && typeof fields[name] !== "string") {
      return `${name} must be a string`;
    }
  }

  if (fields.phones !== undefined) {
    if (!Array.isArray(fields.phones)) return "phones must be an array";
    for (let i = 0; i < fields.phones.length; i++) {
      const phone = fields.phones[i];
      if (!phone || typeof phone !== "object" || Array.isArray(phone)) {
        return `phones[${i}] must be an object`;
      }
      const unknown = Object.keys(phone).find(key => !["type", "number"].includes(key));
      if (unknown) return `Unknown phones[${i}] property: ${unknown}`;
      if (!CONTACT_PHONE_TYPES.includes(phone.type)) {
        return `phones[${i}].type must be one of: ${CONTACT_PHONE_TYPES.join(", ")}`;
      }
      if (typeof phone.number !== "string" || !phone.number.trim()) {
        return `phones[${i}].number must be a non-empty string`;
      }
    }
  }

  if (fields.addresses !== undefined) {
    if (!Array.isArray(fields.addresses)) return "addresses must be an array";
    for (let i = 0; i < fields.addresses.length; i++) {
      const address = fields.addresses[i];
      if (!address || typeof address !== "object" || Array.isArray(address)) {
        return `addresses[${i}] must be an object`;
      }
      const unknown = Object.keys(address)
        .find(key => key !== "type" && !CONTACT_ADDRESS_FIELDS.includes(key));
      if (unknown) return `Unknown addresses[${i}] property: ${unknown}`;
      if (!CONTACT_ADDRESS_TYPES.includes(address.type)) {
        return `addresses[${i}].type must be one of: ${CONTACT_ADDRESS_TYPES.join(", ")}`;
      }
      for (const field of CONTACT_ADDRESS_FIELDS) {
        if (address[field] !== undefined && typeof address[field] !== "string") {
          return `addresses[${i}].${field} must be a string`;
        }
      }
      if (!CONTACT_ADDRESS_FIELDS.some(field =>
        typeof address[field] === "string" && address[field].trim().length > 0
      )) {
        return `addresses[${i}] must contain at least one non-empty address field`;
      }
    }
  }

  if (fields.birthday !== undefined && normalizeContactBirthday(fields.birthday) === null) {
    return "birthday must be YYYY-MM-DD or --MM-DD with a valid calendar date";
  }
  if (requireContent && !contactFieldsHaveContent(fields)) {
    return "At least one non-empty contact field is required";
  }
  return null;
}

function updateVCardOrganization(vCardProperties, organization, VCardPropertyEntry) {
  const entries = vCardProperties.getAllEntries("org");
  const entry = entries[0];
  if (!entry) {
    if (organization) {
      vCardProperties.addEntry(new VCardPropertyEntry("org", {}, "text", [organization]));
    }
    return;
  }

  if (!Array.isArray(entry.value)) {
    if (organization) entry.value = organization;
    else if (entries.length > 1) entry.value = [""];
    else vCardProperties.removeEntry(entry);
    return;
  }

  const remainingComponents = entry.value.slice(1);
  if (organization || remainingComponents.some(value => contactStructuredValuePart(value))) {
    entry.value = [organization, ...remainingComponents];
  } else if (entries.length > 1) {
    entry.value = [""];
  } else {
    vCardProperties.removeEntry(entry);
  }
}

function reconcileVCardContactEntries(vCardProperties, name, items, config) {
  const existingEntries = vCardProperties.getAllEntries(name);
  const matchedEntries = new Array(items.length).fill(null);
  const retainedEntries = new Set();

  // Prefer an exact normalized type/value match. Besides avoiding unnecessary
  // writes, this keeps URI-backed TEL values and structured ADR values exactly
  // as Thunderbird parsed them.
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const entry = existingEntries.find(candidate =>
      !retainedEntries.has(candidate) &&
      config.getEntryType(candidate) === config.getItemType(item) &&
      config.getEntryValue(candidate) === config.getItemValue(item)
    );
    if (entry) {
      matchedEntries[i] = entry;
      retainedEntries.add(entry);
    }
  }

  // A changed value still reuses the entry in the same type bucket. Mutating
  // only its value preserves PREF, extra TYPE values, groups/labels, and the
  // parsed vCard value type.
  for (let i = 0; i < items.length; i++) {
    if (matchedEntries[i]) continue;
    const item = items[i];
    const entry = existingEntries.find(candidate =>
      !retainedEntries.has(candidate) &&
      config.getEntryType(candidate) === config.getItemType(item)
    );
    if (entry) {
      matchedEntries[i] = entry;
      retainedEntries.add(entry);
      config.updateEntry(entry, item);
    }
  }

  for (const entry of existingEntries) {
    if (!retainedEntries.has(entry)) vCardProperties.removeEntry(entry);
  }
  for (let i = 0; i < items.length; i++) {
    if (!matchedEntries[i]) vCardProperties.addEntry(config.createEntry(items[i]));
  }
}

function normalizeContactPhoneValue(value) {
  return contactValueToString(value).trim().replace(/^tel:/i, "").trim();
}

function contactAddressValueParts(value) {
  const rawParts = Array.isArray(value) ? value : [value];
  return CONTACT_ADDRESS_FIELDS.map((_, index) =>
    contactStructuredValuePart(rawParts[index])
  );
}

function contactAddressItemParts(address) {
  return CONTACT_ADDRESS_FIELDS.map(field => address[field] || "");
}

function normalizeContactAddressValue(value) {
  return JSON.stringify(contactAddressValueParts(value).map(part => part.trim()));
}

function updateVCardPhoneEntry(entry, phone) {
  const number = phone.number.trim();
  const hasUriValueType = typeof entry.type === "string" &&
    entry.type.toLowerCase() === "uri";
  const hadTelUri = typeof entry.value === "string" && /^tel:/i.test(entry.value.trim());
  const hasUriScheme = /^[a-z][a-z\d+.-]*:/i.test(number);
  entry.value = (hasUriValueType || hadTelUri) && !hasUriScheme
    ? `tel:${number}`
    : number;
}

function getDuplicateContactType(items) {
  if (!Array.isArray(items)) return null;
  const seen = new Set();
  for (const item of items) {
    if (seen.has(item.type)) return item.type;
    seen.add(item.type);
  }
  return null;
}

function getFlatContactCollectionError(card, fields) {
  if (card.supportsVCard) return null;

  const duplicatePhoneType = getDuplicateContactType(fields.phones);
  if (duplicatePhoneType) {
    return `Non-vCard contact cards support only one phone number per type; duplicate phone type "${duplicatePhoneType}" is not supported`;
  }
  const duplicateAddressType = getDuplicateContactType(fields.addresses);
  if (duplicateAddressType) {
    return `Non-vCard contact cards support only one address per type; duplicate address type "${duplicateAddressType}" is not supported`;
  }
  return null;
}

function applyVCardContactFields(card, fields, VCardPropertyEntry) {
  const vCardProperties = card.vCardProperties;

  if (fields.phones !== undefined) {
    reconcileVCardContactEntries(vCardProperties, "tel", fields.phones, {
      getEntryType: getContactPhoneType,
      getItemType: phone => phone.type,
      getEntryValue: entry => normalizeContactPhoneValue(entry.value),
      getItemValue: phone => normalizeContactPhoneValue(phone.number),
      updateEntry: updateVCardPhoneEntry,
      createEntry: phone => new VCardPropertyEntry(
        "tel",
        { type: phone.type === "mobile" ? "cell" : phone.type },
        "text",
        phone.number.trim()
      ),
    });
  }

  if (fields.addresses !== undefined) {
    reconcileVCardContactEntries(vCardProperties, "adr", fields.addresses, {
      getEntryType: getContactAddressType,
      getItemType: address => address.type,
      getEntryValue: entry => normalizeContactAddressValue(entry.value),
      getItemValue: address => normalizeContactAddressValue(contactAddressItemParts(address)),
      updateEntry: (entry, address) => {
        entry.value = contactAddressItemParts(address);
      },
      createEntry: address => new VCardPropertyEntry(
        "adr",
        { type: address.type },
        "text",
        contactAddressItemParts(address)
      ),
    });
  }

  if (fields.organization !== undefined) {
    updateVCardOrganization(vCardProperties, fields.organization, VCardPropertyEntry);
  }
  for (const [field, vCardName] of [["title", "title"], ["note", "note"]]) {
    if (fields[field] === undefined) continue;
    vCardProperties.clearValues(vCardName);
    if (fields[field]) {
      vCardProperties.addEntry(new VCardPropertyEntry(vCardName, {}, "text", fields[field]));
    }
  }
  if (fields.birthday !== undefined) {
    vCardProperties.clearValues("bday");
    const birthday = contactBirthdayToVCard(fields.birthday);
    if (birthday) {
      vCardProperties.addEntry(new VCardPropertyEntry("bday", {}, "date", birthday));
    }
  }
}

function applyFlatContactFields(card, fields) {
  if (fields.phones !== undefined) {
    for (const property of Object.values(CONTACT_PHONE_FLAT_PROPERTIES)) {
      card.setProperty(property, "");
    }
    for (const phone of fields.phones) {
      card.setProperty(CONTACT_PHONE_FLAT_PROPERTIES[phone.type], phone.number.trim());
    }
  }

  if (fields.addresses !== undefined) {
    for (const properties of Object.values(CONTACT_ADDRESS_FLAT_PROPERTIES)) {
      for (const property of properties) card.setProperty(property, "");
    }
    for (const address of fields.addresses) {
      const properties = CONTACT_ADDRESS_FLAT_PROPERTIES[address.type];
      for (let i = 0; i < CONTACT_ADDRESS_FIELDS.length; i++) {
        card.setProperty(properties[i], address[CONTACT_ADDRESS_FIELDS[i]] || "");
      }
    }
  }

  if (fields.organization !== undefined) card.setProperty("Company", fields.organization);
  if (fields.title !== undefined) card.setProperty("JobTitle", fields.title);
  if (fields.note !== undefined) card.setProperty("Notes", fields.note);
  if (fields.birthday !== undefined) {
    const birthday = contactBirthdayToFlatParts(fields.birthday);
    card.setProperty("BirthYear", birthday.year);
    card.setProperty("BirthMonth", birthday.month);
    card.setProperty("BirthDay", birthday.day);
  }
}

function applyContactFields(card, fields, VCardPropertyEntry) {
  const supportsVCard = !!card.supportsVCard;
  const flatCollectionError = getFlatContactCollectionError(card, fields);
  if (flatCollectionError) return { error: flatCollectionError };

  if (fields.email !== undefined) {
    if (supportsVCard && fields.email === "") {
      card.vCardProperties.clearValues("email");
    } else {
      card.primaryEmail = fields.email;
    }
  }
  if (fields.displayName !== undefined) card.displayName = fields.displayName;
  if (fields.firstName !== undefined) card.firstName = fields.firstName;
  if (fields.lastName !== undefined) card.lastName = fields.lastName;

  if (supportsVCard) {
    applyVCardContactFields(card, fields, VCardPropertyEntry);
  } else {
    applyFlatContactFields(card, fields);
  }
  return null;
}

function shouldSynthesizePhoneDisplayName(fields) {
  if (!Array.isArray(fields.phones) || fields.phones.length === 0) return false;
  if (CONTACT_SCALAR_FIELDS.some(name =>
    typeof fields[name] === "string" && fields[name].trim()
  )) {
    return false;
  }
  return !Array.isArray(fields.addresses) || fields.addresses.length === 0;
}
// END CONTACT FIELD HELPERS

function getExtVersion() {
  if (_cachedExtVersion) return _cachedExtVersion;
  try {
    const uri = Services.io.newURI("resource://thunderbird-mcp/manifest.json");
    const channel = Services.io.newChannelFromURI(uri, null,
      Services.scriptSecurityManager.getSystemPrincipal(), null,
      Ci.nsILoadInfo.SEC_ALLOW_CROSS_ORIGIN_SEC_CONTEXT_IS_NULL,
      Ci.nsIContentPolicy.TYPE_OTHER);
    const sis = Cc["@mozilla.org/scriptableinputstream;1"]
      .createInstance(Ci.nsIScriptableInputStream);
    sis.init(channel.open());
    const text = sis.read(sis.available());
    sis.close();
    _cachedExtVersion = JSON.parse(text).version || "0.0.0";
  } catch (e) {
    console.warn("thunderbird-mcp: could not read extension manifest version:", e);
    _cachedExtVersion = "0.0.0";
  }
  return _cachedExtVersion;
}
// Track temp files created for inline base64 attachments (cleaned up on shutdown).
const _tempAttachFiles = new Set();
// Track compose windows already claimed by an in-flight replyToMessage or
// forwardMessage call, so concurrent reply/forward operations on the same
// original message never bind two observers to the same compose window
// (which would double-inject the body/attachments).
// WeakSet so entries are collected automatically when the window is destroyed.
const _claimedComposeWindows = new WeakSet();
// BEGIN INLINE ATTACHMENT BASE64 HELPERS
// Require canonical RFC 4648 base64: complete quartets with padding only in
// the final quartet. In particular, do not silently discard invalid bytes.
const STRICT_BASE64_PATTERN = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;
function isValidBase64(value) {
  return typeof value === "string" && value.length > 0 && STRICT_BASE64_PATTERN.test(value);
}
// END INLINE ATTACHMENT BASE64 HELPERS
// BEGIN OUTBOUND ATTACHMENT LIMITS
const MAX_BASE64_SIZE = 25 * 1024 * 1024; // 25 MB limit for inline base64 data (encoded)
// Cap file-path attachments to the same magnitude as saved-message attachments.
// Prevents an MCP caller from attaching multi-GB files to a single outgoing message.
const MAX_FILE_PATH_ATTACHMENT_BYTES = 50 * 1024 * 1024;
const MAX_TOTAL_ATTACHMENT_BYTES = 50 * 1024 * 1024;
const MAX_ATTACHMENTS_PER_MESSAGE = 20;
// END OUTBOUND ATTACHMENT LIMITS
// Must be large enough to carry MAX_BASE64_SIZE plus JSON-RPC framing overhead.
// The httpd.sys.mjs pre-buffer cap uses the same value.
const MAX_REQUEST_BODY = 32 * 1024 * 1024; // 32 MB limit for incoming HTTP request bodies

// BEGIN INLINE IMAGE CONTENT HELPERS
// MCP image payloads are base64 text, so budget the encoded representation that
// actually enters the client's context rather than only the decoded MIME bytes.
const MAX_INLINE_IMAGE_BASE64_BYTES = 1 * 1024 * 1024;
const MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES = 4 * 1024 * 1024;
const SUPPORTED_INLINE_IMAGE_MIME_TYPES = new Set([
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/webp",
]);
const MCP_EXTRA_CONTENT_BLOCKS = Symbol("thunderbird-mcp.extra-content-blocks");

function normalizeInlineImageMimeType(contentType) {
  return ((String(contentType || "").split(";")[0] || "").trim().toLowerCase());
}

function normalizeInlineImageContentId(contentId) {
  let normalized = String(contentId || "").trim();
  if (/^cid:/i.test(normalized)) normalized = normalized.slice(4).trim();
  try {
    normalized = decodeURIComponent(normalized);
  } catch {
    // Keep malformed-but-usable identifiers in their original form.
  }
  return normalized.replace(/^<+|>+$/g, "").trim().toLowerCase();
}

function findInlineImageRecordIndex(records, target) {
  const targetContentId = normalizeInlineImageContentId(target?.contentId);
  if (targetContentId) {
    const contentIdIndex = records.findIndex(record =>
      normalizeInlineImageContentId(record?.contentId) === targetContentId
    );
    if (contentIdIndex >= 0) return contentIdIndex;
  }

  const targetPartName = String(target?.partName || "").trim();
  if (!targetPartName) return -1;
  return records.findIndex(record =>
    String(record?.partName || "").trim() === targetPartName
  );
}

/**
 * Correlates Gloda's allUserAttachments records with inline MIME-tree parts.
 * Content-ID is authoritative when available; MIME part name is the fallback
 * for Gloda representations where disposition and Content-ID were stripped.
 * The returned arrays are new, and input records are not mutated.
 */
function correlateInlineImageRecords(knownAttachments, inlineImages) {
  const metadataEntries = Array.isArray(knownAttachments)
    ? knownAttachments.slice()
    : [];
  const inlineImageEntries = [];

  for (const inlineImage of Array.isArray(inlineImages) ? inlineImages : []) {
    if (findInlineImageRecordIndex(
      inlineImageEntries.map(entry => entry.inlineImage),
      inlineImage
    ) >= 0) {
      continue;
    }

    const knownAttachmentIndex = findInlineImageRecordIndex(metadataEntries, inlineImage);
    const matchedKnownAttachment = knownAttachmentIndex >= 0;
    let metadataRecord;
    let metadataIndex = knownAttachmentIndex;
    if (matchedKnownAttachment) {
      metadataRecord = metadataEntries[knownAttachmentIndex];
    } else {
      metadataRecord = inlineImage;
      metadataIndex = metadataEntries.length;
      metadataEntries.push(metadataRecord);
    }

    inlineImageEntries.push({
      metadataRecord,
      metadataIndex,
      inlineImage,
      matchedKnownAttachment,
    });
  }

  return { metadataEntries, inlineImageEntries };
}

function getInlineImageContentIdReferences(body) {
  const references = [];
  const seen = new Set();
  const cidPattern = /\bcid\s*:\s*(?:<([^>]+)>|([^"'<>\s)\]]+))/gi;
  let match;
  while ((match = cidPattern.exec(String(body || ""))) !== null) {
    const contentId = normalizeInlineImageContentId(match[1] || match[2]);
    if (!contentId || seen.has(contentId)) continue;
    seen.add(contentId);
    references.push(contentId);
  }
  return references;
}

function orderInlineImageRecordsForBody(inlineImages, body) {
  const remaining = Array.isArray(inlineImages) ? inlineImages.slice() : [];
  const ordered = [];

  // Attempt rendered CID images first, in first-reference document order.
  // Any inline MIME parts not referenced by the rendered body retain MIME order.
  for (const referencedContentId of getInlineImageContentIdReferences(body)) {
    for (let i = 0; i < remaining.length;) {
      const recordContentId = normalizeInlineImageContentId(
        remaining[i]?.contentId ?? remaining[i]?.info?.contentId
      );
      if (recordContentId === referencedContentId) {
        ordered.push(remaining.splice(i, 1)[0]);
      } else {
        i++;
      }
    }
  }

  return ordered.concat(remaining);
}

function getBase64EncodedSize(byteLength) {
  if (!Number.isFinite(byteLength) || byteLength <= 0) return 0;
  return 4 * Math.ceil(byteLength / 3);
}

function getInlineImageSkipReason(mimeType, encodedSize, totalEncodedSize) {
  const normalizedMimeType = normalizeInlineImageMimeType(mimeType);
  if (!SUPPORTED_INLINE_IMAGE_MIME_TYPES.has(normalizedMimeType)) {
    return `Unsupported MIME type "${normalizedMimeType || "(missing)"}"`;
  }
  if (!Number.isFinite(encodedSize) || encodedSize <= 0) {
    return "Image data is empty";
  }
  if (encodedSize > MAX_INLINE_IMAGE_BASE64_BYTES) {
    return `Image exceeds per-image base64 limit (${encodedSize} bytes > ${MAX_INLINE_IMAGE_BASE64_BYTES} bytes)`;
  }
  if (totalEncodedSize + encodedSize > MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES) {
    return `Image would exceed total base64 limit (${totalEncodedSize + encodedSize} bytes > ${MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES} bytes)`;
  }
  return "";
}

function encodeByteStringToBase64(byteString) {
  return btoa(String(byteString || ""));
}

function setExtraMcpContentBlocks(toolResult, blocks) {
  if (!toolResult || typeof toolResult !== "object" || !Array.isArray(blocks) || blocks.length === 0) {
    return toolResult;
  }
  Object.defineProperty(toolResult, MCP_EXTRA_CONTENT_BLOCKS, {
    value: blocks,
    enumerable: false,
    configurable: true,
  });
  return toolResult;
}

function buildToolResultContent(toolResult) {
  const content = [{
    type: "text",
    text: JSON.stringify(toolResult, null, 2),
  }];
  const extraBlocks = toolResult && toolResult[MCP_EXTRA_CONTENT_BLOCKS];
  if (Array.isArray(extraBlocks)) content.push(...extraBlocks);
  return content;
}
// END INLINE IMAGE CONTENT HELPERS

// File paths that an MCP caller must never be allowed to attach to outbound
// mail. Protects against the LLM-confused-deputy chain where attacker-controlled
// email content prompt-injects an assistant into running
// sendMail({attachments: ["/home/user/.ssh/id_rsa"], skipReview: true}).
//
// Patterns match the path AFTER backslashes are normalized to forward slashes
// and the whole string is lower-cased, so a single set covers POSIX and Windows.
// This is a deny-list, not an allow-list -- it intentionally errs toward
// blocking known-sensitive locations rather than restricting users to a
// downloads-only sandbox. Extend it as new high-value targets surface.
// BEGIN SENSITIVE ATTACHMENT PATH HELPERS
// Keep in sync with mcp-bridge.cjs isSensitiveFilePath.
const SENSITIVE_ATTACHMENT_PATTERNS = [
  // SSH / PGP / cloud / kube / docker credentials
  /\/\.ssh(\/|$)/,
  /\/\.gnupg(\/|$)/,
  /\/\.aws(\/|$)/,
  /\/\.azure(\/|$)/,
  /\/\.config\/gcloud(\/|$)/,
  /\/\.kube(\/|$)/,
  /\/\.docker(\/|$)/,
  /\/\.netrc$/,
  /\/\.npmrc$/,
  /\/\.pypirc$/,
  // Common key / secret file extensions anywhere on disk
  /\/id_(rsa|dsa|ecdsa|ed25519)(\.pub)?$/,
  /\.pem$/,
  /\.pfx$/,
  /\.p12$/,
  /\.kdbx$/,
  /\.key$/,
  /\.asc$/,
  /\.gpg$/,
  // Linux / macOS system directories
  /^\/etc\//,
  /^\/proc\//,
  /^\/sys\//,
  /^\/root\//,
  /^\/var\/log\//,
  /^\/var\/lib\/sudo\//,
  // macOS keychain locations
  /\/library\/keychains\//,
  // Windows system directories
  /^[a-z]:\/windows\//,
  /^[a-z]:\/programdata\/microsoft\/(crypto|protect)\//,
  /\/appdata\/(local|roaming)\/microsoft\/(credentials|crypto|protect|vault)(\/|$)/,
  // Browser credential stores (Firefox / Chrome / Edge)
  /\/(logins\.json|key3\.db|key4\.db|cookies(\.sqlite)?|login data)$/,
  // Thunderbird's own profile (contains the user's entire mail store + prefs).
  // Linux profile directories and profiles.ini live directly under
  // ~/.thunderbird (or ~/.icedove), while macOS and Windows use the platform
  // application-data directories below. Block each profile root in full.
  /\/\.(?:thunderbird|icedove)(\/|$)/,
  /\/library\/thunderbird(\/|$)/,
  /\/appdata\/roaming\/thunderbird(\/|$)/,
];

/**
 * Return true if `attachmentPath` looks like a credential, secret, or system
 * file that an MCP caller should not be able to attach to outgoing mail.
 * Path is normalized (backslashes → forward slashes, lower-cased) before
 * matching so the same pattern set works on POSIX and Windows.
 */
function isSensitiveFilePath(attachmentPath) {
  if (typeof attachmentPath !== "string" || !attachmentPath) return false;
  const normalized = attachmentPath.replace(/\\/g, "/").toLowerCase();
  return SENSITIVE_ATTACHMENT_PATTERNS.some(re => re.test(normalized));
}
// END SENSITIVE ATTACHMENT PATH HELPERS
let _tempFileCounter = 0;
const DEFAULT_MAX_RESULTS = 50;
const PREF_ALLOWED_ACCOUNTS = "extensions.thunderbird-mcp.allowedAccounts";
const PREF_DISABLED_TOOLS = "extensions.thunderbird-mcp.disabledTools";
const PREF_BLOCK_SKIPREVIEW = "extensions.thunderbird-mcp.blockSkipReview";
const PREF_STABLE_AUTH_TOKEN = "extensions.thunderbird-mcp.stableAuthToken";
const PREF_GET_MESSAGES_LIMIT = "extensions.thunderbird-mcp.getMessagesLimit";
const PREF_LISTEN_ALL = "extensions.thunderbird-mcp.listenAll";
const AUTH_TOKEN_PATTERN = /^[0-9a-f]{64}$/;
// Valid group and CRUD values for tool metadata validation
const VALID_GROUPS = ["messages", "folders", "contacts", "calendar", "filters", "system"];
const VALID_CRUD = ["create", "read", "update", "delete"];
// CRUD sort order: read first, then create, update, delete (safe → destructive)
const CRUD_ORDER = { read: 0, create: 1, update: 2, delete: 3 };
// Tools that cannot be disabled via the settings page (infrastructure tools)
const UNDISABLEABLE_TOOLS = new Set(["listAccounts", "listFolders", "getAccountAccess"]);
const MAX_SEARCH_RESULTS_CAP = 200;
const SEARCH_COLLECTION_CAP = 10000;
const DEFAULT_GET_MESSAGES_LIMIT = 10;
// 20 is a reasonable upper bound for now; adjust later if usage supports it.
const MAX_GET_MESSAGES_LIMIT = 20;
// Internal IMAP/Thunderbird keywords that should not appear as user-visible tags
const INTERNAL_KEYWORDS = new Set([
  "junk", "notjunk", "$forwarded", "$replied",
  "\\seen", "\\answered", "\\flagged", "\\deleted", "\\draft", "\\recent",
  // Some IMAP servers store flags without the backslash prefix
  "seen", "answered", "flagged", "deleted", "draft", "recent",
]);

var mcpServer = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const extensionRoot = context.extension.rootURI;
    const resourceName = "thunderbird-mcp";

    resProto.setSubstitutionWithFlags(
      resourceName,
      extensionRoot,
      resProto.ALLOW_CONTENT_ACCESS
    );

    function normalizeGetMessagesLimit(value) {
      const limit = Number(value);
      if (!Number.isInteger(limit)) return DEFAULT_GET_MESSAGES_LIMIT;
      if (limit < 1) return 1;
      if (limit > MAX_GET_MESSAGES_LIMIT) return MAX_GET_MESSAGES_LIMIT;
      return limit;
    }

    function getConfiguredGetMessagesLimit() {
      try {
        return normalizeGetMessagesLimit(
          Services.prefs.getIntPref(PREF_GET_MESSAGES_LIMIT, DEFAULT_GET_MESSAGES_LIMIT)
        );
      } catch {
        return DEFAULT_GET_MESSAGES_LIMIT;
      }
    }

    // BEGIN TOOL SCHEMA BUILDER
    function buildTools() {
      const getMessagesLimit = getConfiguredGetMessagesLimit();
      const contactFieldProperties = {
        email: { type: "string", description: "Primary email address. May be omitted for phone-only contacts." },
        displayName: { type: "string", description: "Display name" },
        firstName: { type: "string", description: "First name" },
        lastName: { type: "string", description: "Last name" },
        phones: {
          type: "array",
          description: "Phone numbers. On update, replaces the phone collection; use [] to clear it.",
          items: {
            type: "object",
            properties: {
              type: { type: "string", enum: CONTACT_PHONE_TYPES, description: "Phone type" },
              number: { type: "string", description: "Phone number" },
            },
            required: ["type", "number"],
            additionalProperties: false,
          },
        },
        addresses: {
          type: "array",
          description: "Postal addresses. On update, replaces the address collection; use [] to clear it.",
          items: {
            type: "object",
            properties: {
              type: { type: "string", enum: CONTACT_ADDRESS_TYPES, description: "Address type" },
              poBox: { type: "string", description: "Post office box" },
              street: { type: "string", description: "Street address" },
              street2: { type: "string", description: "Additional street/address line" },
              city: { type: "string", description: "City or locality" },
              region: { type: "string", description: "State, province, or region" },
              postalCode: { type: "string", description: "Postal or ZIP code" },
              country: { type: "string", description: "Country" },
            },
            required: ["type"],
            additionalProperties: false,
          },
        },
        organization: { type: "string", description: "Organization or company name" },
        title: { type: "string", description: "Job title" },
        note: { type: "string", description: "Contact note; may contain multiple lines" },
        birthday: { type: "string", description: "Birthday as YYYY-MM-DD or --MM-DD when the year is unknown" },
      };
      return [
      {
        name: "listAccounts",
        group: "system", crud: "read",
        title: "List Accounts",
        description: "List all email accounts and their identities",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "listFolders",
        group: "system", crud: "read",
        title: "List Folders",
        description: "List all mail folders with URIs and message counts",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Optional account ID (from listAccounts) to limit results to a single account" },
            folderPath: { type: "string", description: "Optional folder URI (from listFolders) to list only that folder and its subfolders" },
            format: { type: "string", enum: ["objects", "table"], description: "Response format: 'objects' (default, existing array of folder objects) or 'table' ({ columns, rows } compact form)" },
          },
          required: [],
        },
      },
      {
        name: "searchMessages",
        group: "messages", crud: "read",
        title: "Search Mail",
        description: "Search message headers and return IDs/folder paths you can use with getMessage to read full email content",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Text to search. Multi-word queries are AND-of-tokens: every word must appear somewhere across subject/author/recipients/ccList/preview (or inside the selected field when an operator is used). Prefix with 'from:', 'subject:', 'to:', or 'cc:' to restrict matching to one field (e.g. 'from:Alice Smith' requires both tokens in the author field). Use empty string to match all." },
            folderPath: { type: "string", description: "Optional folder URI (from listFolders) to limit search to that folder and its subfolders" },
            startDate: { type: "string", description: "Filter messages on or after this ISO 8601 date" },
            endDate: { type: "string", description: "Filter messages on or before this ISO 8601 date. Date-only strings (e.g. '2024-01-15') include the full day." },
            maxResults: { type: "number", description: "Maximum number of results to return (default 50, max 200)" },
            offset: { type: "number", description: "Number of results to skip for pagination (default 0). When provided, returns {messages, totalMatches, offset, limit, hasMore} instead of a plain array. Note: totalMatches is capped at 10000." },
            sortOrder: { type: "string", description: "Date sort order: asc (oldest first) or desc (newest first, default)" },
            unreadOnly: { type: "boolean", description: "Only return unread messages (default: false)" },
            flaggedOnly: { type: "boolean", description: "Only return flagged/starred messages (default: false)" },
            tag: { type: "string", description: "Filter by tag keyword (e.g. '$label1' for Important, or a custom tag). Only messages with this tag are returned." },
            includeSubfolders: { type: "boolean", description: "If false, only search the specified folder — not its subfolders. Default: true." },
            countOnly: { type: "boolean", description: "If true, return only the match count instead of full results. Much faster for 'how many unread?' queries." },
            searchBody: { type: "boolean", description: "If true, search full message bodies using Thunderbird's Gloda index (slower but finds text beyond the ~200 char preview). Requires query. IMAP accounts need offline sync enabled for body indexing." },
            dedupByMessageId: { type: "boolean", description: "If false, return every folder/label location for messages found in multiple folders. Default: true, which collapses the same RFC Message-ID into one row and lists the other folder paths in dupLocations." },
          },
          required: ["query"],
        },
      },
      {
        name: "getMessage",
        group: "messages", crud: "read",
        title: "Get Message",
        description: "Read the full content of an email message by its ID",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            saveAttachments: { type: "boolean", description: "If true, save attachments to <OS temp dir>/thunderbird-mcp/<messageId>/ and include filePath in response (default: false)" },
            includeInlineImages: { type: "boolean", description: "If true, append supported inline email images as MCP image content blocks after the text result (default: false; max 1 MiB base64 per image and 4 MiB total). Images referenced by the rendered body are attempted first in document order, followed by remaining inline images in MIME order. Ignored when rawSource is true." },
            bodyFormat: { type: "string", enum: ["markdown", "text", "html"], description: "Body output format: 'markdown' (default, preserves structure), 'text' (plain text), 'html' (raw HTML)" },
            rawSource: { type: "boolean", description: "If true, return the full raw RFC 2822 message source (all headers + MIME parts). Useful for extracting calendar invites, S/MIME data, or debugging. Other fields (body, attachments) are omitted when this is set. Note: requires local/offline message copy; IMAP messages not cached offline may fail." },
          },
          required: ["messageId", "folderPath"],
        },
      },
      {
        name: "getMessages",
        group: "messages", crud: "read",
        title: "Get Messages",
        description: `Read full email content for up to ${getMessagesLimit} messages in one call. Each item needs messageId and folderPath from searchMessages/getRecentMessages results.`,
        inputSchema: {
          type: "object",
          properties: {
            messages: {
              type: "array",
              minItems: 1,
              maxItems: getMessagesLimit,
              description: `Messages to read, max ${getMessagesLimit}. Each item is { messageId, folderPath }.`,
              items: {
                type: "object",
                properties: {
                  messageId: { type: "string", description: "The message ID" },
                  folderPath: { type: "string", description: "The folder URI path containing the message" },
                },
                required: ["messageId", "folderPath"],
                additionalProperties: false,
              },
            },
            saveAttachments: { type: "boolean", description: "If true, save attachments for each message and include filePath in attachment metadata (default: false)" },
            bodyFormat: { type: "string", enum: ["markdown", "text", "html"], description: "Body output format shared by all messages: 'markdown' (default), 'text', or 'html'" },
            rawSource: { type: "boolean", description: "If true, return raw RFC 2822 source for each message instead of parsed body fields" },
          },
          required: ["messages"],
        },
      },
      {
        name: "sendMail",
        group: "messages", crud: "create",
        title: "Compose Mail",
        description: "Compose a new email in a review window. The skipReview safety block is on by default; direct sending is honored only when the user explicitly disables that preference.",
        inputSchema: {
          type: "object",
          properties: {
            to: { type: "string", description: "Recipient email address" },
            subject: { type: "string", description: "Email subject line" },
            body: { type: "string", description: "Email body text" },
            cc: { type: "string", description: "CC recipients (comma-separated)" },
            bcc: { type: "string", description: "BCC recipients (comma-separated)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            from: { type: "string", description: "Sender identity (email address or identity ID from listAccounts)" },
            skipReview: { type: "boolean", description: "Request direct sending without a compose window. Honored only when the user explicitly disables the default-on skipReview safety block (default: false)." },
            attachments: {
              type: "array",
              maxItems: MAX_ATTACHMENTS_PER_MESSAGE,
              description: "Attachments: file paths (strings) or inline objects ({name, contentType, base64})",
              items: {
                oneOf: [
                  { type: "string", description: "Absolute file path to attach" },
                  {
                    type: "object",
                    properties: {
                      name: { type: "string", minLength: 1, description: "Attachment filename" },
                      contentType: { type: "string", description: "MIME type, e.g. application/pdf" },
                      base64: { type: "string", minLength: 1, contentEncoding: "base64", description: "Base64-encoded file content" },
                      content: { type: "string", minLength: 1, contentEncoding: "base64", description: "Alias for base64 (accepted for backwards compatibility); base64 takes precedence when both are set" },
                    },
                    required: ["name"],
                    anyOf: [
                      { type: "object", required: ["base64"] },
                      { type: "object", required: ["content"] },
                    ],
                    additionalProperties: false,
                  },
                ],
              },
            },
          },
          required: ["to", "subject", "body"],
        },
      },
      {
        name: "saveDraft",
        group: "messages", crud: "create",
        title: "Save Draft",
        description: "Save a composed message to the identity's Drafts folder without sending or opening a compose window. Useful when a human will review and send the message later from Thunderbird.",
        inputSchema: {
          type: "object",
          properties: {
            to: { type: "string", description: "Recipient email address(es), comma-separated. Optional -- a draft can have no recipient." },
            subject: { type: "string", description: "Email subject line (optional)" },
            body: { type: "string", description: "Email body (optional)" },
            cc: { type: "string", description: "CC recipients (comma-separated)" },
            bcc: { type: "string", description: "BCC recipients (comma-separated)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            from: { type: "string", description: "Sender identity (email address or identity ID from listAccounts)" },
            attachments: {
              type: "array",
              maxItems: MAX_ATTACHMENTS_PER_MESSAGE,
              description: "Attachments: file paths (strings) or inline objects ({name, contentType, base64})",
              items: {
                oneOf: [
                  { type: "string", description: "Absolute file path to attach" },
                  {
                    type: "object",
                    properties: {
                      name: { type: "string", minLength: 1, description: "Attachment filename" },
                      contentType: { type: "string", description: "MIME type, e.g. application/pdf" },
                      base64: { type: "string", minLength: 1, contentEncoding: "base64", description: "Base64-encoded file content" },
                      content: { type: "string", minLength: 1, contentEncoding: "base64", description: "Alias for base64 (accepted for backwards compatibility); base64 takes precedence when both are set" },
                    },
                    required: ["name"],
                    anyOf: [
                      { type: "object", required: ["base64"] },
                      { type: "object", required: ["content"] },
                    ],
                    additionalProperties: false,
                  },
                ],
              },
            },
          },
          required: [],
        },
      },
      {
        name: "listCalendars",
        group: "calendar", crud: "read",
        title: "List Calendars",
        description: "Return the user's calendars",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "createEvent",
        group: "calendar", crud: "create",
        title: "Create Event",
        description: "Create a calendar event through a review dialog. The skipReview safety block is on by default; direct creation is honored only when the user explicitly disables that preference.",
        inputSchema: {
          type: "object",
          properties: {
            title: { type: "string", description: "Event title" },
            startDate: { type: "string", description: "Start date/time in ISO 8601 format" },
            endDate: { type: "string", description: "End date/time in ISO 8601 (defaults to startDate + 1h for timed, +1 day for all-day)" },
            location: { type: "string", description: "Event location" },
            description: { type: "string", description: "Event description" },
            calendarId: { type: "string", description: "Target calendar ID (from listCalendars, defaults to first writable calendar)" },
            allDay: { type: "boolean", description: "Create an all-day event (default: false)" },
            status: { type: "string", description: "VEVENT STATUS: 'tentative', 'confirmed', or 'cancelled'. Defaults to confirmed if omitted." },
            showAs: { type: "string", enum: ["busy", "free"], description: "How the event appears in the calendar: 'busy' (solid block, TRANSP:OPAQUE + STATUS:CONFIRMED) or 'free' (hatched, TRANSP:TRANSPARENT + STATUS:TENTATIVE). Defaults to 'busy'. Overridden per-property by explicit status parameter." },
            categories: { type: "array", items: { type: "string" }, description: "Category labels (optional). Category names are case-sensitive; use listCategories to get exact existing names before setting." },
            onlineMeeting: { type: "boolean", description: "If true, generates a Microsoft Teams meeting link via Exchange (OWL/Office 365 accounts only). After creation, OWL embeds the join URL in the event description and exposes it via listEvents (onlineMeetingURL). No-op on non-OWL backends." },
            skipReview: { type: "boolean", description: "Request direct creation without a review dialog. Honored only when the user explicitly disables the default-on skipReview safety block (default: false)." },
          },
          required: ["title", "startDate"],
        },
      },
      {
        name: "listEvents",
        group: "calendar", crud: "read",
        title: "List Events",
        description: "List calendar events within a date range",
        inputSchema: {
          type: "object",
          properties: {
            calendarId: { type: "string", description: "Calendar ID to query (from listCalendars). If omitted, queries all calendars." },
            startDate: { type: "string", description: "Start of date range in ISO 8601 format (default: now)" },
            endDate: { type: "string", description: "End of date range in ISO 8601 format (default: 30 days from startDate)" },
            maxResults: { type: "number", description: "Maximum number of events to return (default: 100, max: 500)" },
          },
          required: [],
        },
      },
      {
        name: "updateEvent",
        group: "calendar", crud: "update",
        title: "Update Event",
        description: "Update an existing calendar event's title, dates, location, or description",
        inputSchema: {
          type: "object",
          properties: {
            eventId: { type: "string", description: "The event ID (from listEvents results)" },
            calendarId: { type: "string", description: "The calendar ID containing the event (from listEvents results)" },
            title: { type: "string", description: "New event title (optional)" },
            startDate: { type: "string", description: "New start date/time in ISO 8601 format (optional)" },
            endDate: { type: "string", description: "New end date/time in ISO 8601 format (optional)" },
            location: { type: "string", description: "New event location (optional)" },
            description: { type: "string", description: "New event description (optional)" },
            status: { type: "string", description: "New VEVENT STATUS: 'tentative', 'confirmed', or 'cancelled' (optional)" },
            showAs: { type: "string", enum: ["busy", "free"], description: "How the event appears in the calendar: 'busy' (solid, TRANSP:OPAQUE + STATUS:CONFIRMED) or 'free' (hatched, TRANSP:TRANSPARENT + STATUS:TENTATIVE). Pass null to clear TRANSP only. Explicit status parameter overrides the STATUS coupling." },
            categories: { type: "array", items: { type: "string" }, description: "Category labels (optional). Category names are case-sensitive; pass an empty array to clear all categories. Use listCategories to get exact existing names before setting." },
            onlineMeeting: { type: "boolean", description: "If true, generates a Microsoft Teams meeting link via Exchange (OWL/Office 365 accounts only). Pass false to remove an existing Teams link." },
          },
          required: ["eventId", "calendarId"],
        },
      },
      {
        name: "deleteEvent",
        group: "calendar", crud: "delete",
        title: "Delete Event",
        description: "Delete a calendar event",
        inputSchema: {
          type: "object",
          properties: {
            eventId: { type: "string", description: "The event ID (from listEvents results)" },
            calendarId: { type: "string", description: "The calendar ID containing the event (from listEvents results)" },
          },
          required: ["eventId", "calendarId"],
        },
      },
      {
        name: "createTask",
        group: "calendar", crud: "create",
        title: "Create Task",
        description: "Open a pre-filled task dialog for review. The skipReview safety block is on by default; direct saving is honored only when the user explicitly disables that preference.",
        inputSchema: {
          type: "object",
          properties: {
            title: { type: "string", description: "Task title" },
            dueDate: { type: "string", description: "Due date in ISO 8601 format (optional)" },
            calendarId: { type: "string", description: "Target calendar ID (from listCalendars, must have supportsTasks=true)" },
            description: { type: "string", description: "Task description/body (optional)" },
            priority: { type: "integer", description: "Priority: 1=high, 5=normal, 9=low (optional)" },
            categories: { type: "array", items: { type: "string" }, description: "Category labels (optional). Use listCategories to get exact existing names before setting." },
            skipReview: { type: "boolean", description: "Request direct saving without a review dialog. Honored only when the user explicitly disables the default-on skipReview safety block (default: false)." },
          },
          required: ["title"],
        },
      },
      {
        name: "listCategories",
        group: "calendar", crud: "read",
        title: "List Categories",
        description: "Return all calendar category names defined in Thunderbird preferences. Use this before creating tasks or events to get exact category names (case-sensitive).",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "listTasks",
        group: "calendar", crud: "read",
        title: "List Tasks",
        description: "List tasks/to-dos from Thunderbird calendars, optionally filtered by completion status or due date",
        inputSchema: {
          type: "object",
          properties: {
            calendarId: { type: "string", description: "Calendar ID to query (from listCalendars). If omitted, queries all task-capable calendars." },
            completed: { type: "boolean", description: "Filter by completion status. true = completed only, false = outstanding only. Omit for all tasks." },
            dueBefore: { type: "string", description: "Return tasks due before this ISO 8601 date" },
            maxResults: { type: "integer", description: "Maximum number of tasks to return (default: 100, max: 500)" },
          },
          required: [],
        },
      },
      {
        name: "updateTask",
        group: "calendar", crud: "update",
        title: "Update Task",
        description: "Update an existing task/to-do: change title, due date, description, priority, completion status, or percent complete",
        inputSchema: {
          type: "object",
          properties: {
            taskId: { type: "string", description: "Task ID (from listTasks results)" },
            calendarId: { type: "string", description: "Calendar ID containing the task (from listTasks results)" },
            title: { type: "string", description: "New task title (optional)" },
            dueDate: { type: "string", description: "New due date in ISO 8601 format (optional)" },
            description: { type: "string", description: "New task description/body (optional)" },
            completed: { type: "boolean", description: "Set to true to mark the task done (sets percentComplete=100 and records completedDate), false to reopen it (optional)" },
            percentComplete: { type: "integer", description: "Completion percentage 0–100 (optional)" },
            priority: { type: "integer", description: "Priority: 1=high, 5=normal, 9=low (optional)" },
          },
          required: ["taskId", "calendarId"],
        },
      },
      {
        name: "searchContacts",
        group: "contacts", crud: "read",
        title: "Search Contacts",
        description: "Search contacts across all address books by email address or name",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Email address or name to search for" },
            maxResults: { type: "number", description: "Maximum number of results to return (default 50, max 200). If truncated, response includes hasMore: true." },
          },
          required: ["query"],
        },
      },
      {
        name: "getContact",
        group: "contacts", crud: "read",
        title: "Get Contact",
        description: "Read a contact by UID",
        inputSchema: {
          type: "object",
          properties: {
            contactId: { type: "string", description: "Contact UID (from searchContacts results)" },
          },
          required: ["contactId"],
        },
      },
      {
        name: "createContact",
        group: "contacts", crud: "create",
        title: "Create Contact",
        description: "Create a new contact in an address book",
        inputSchema: {
          type: "object",
          properties: {
            ...contactFieldProperties,
            addressBookId: { type: "string", description: "Address book directory ID (from searchContacts results). Defaults to the first writable address book." },
          },
          required: [],
        },
      },
      {
        name: "updateContact",
        group: "contacts", crud: "update",
        title: "Update Contact",
        description: "Update an existing contact's properties",
        inputSchema: {
          type: "object",
          properties: {
            contactId: { type: "string", description: "Contact UID (from searchContacts results)" },
            ...contactFieldProperties,
          },
          required: ["contactId"],
        },
      },
      {
        name: "deleteContact",
        group: "contacts", crud: "delete",
        title: "Delete Contact",
        description: "Delete a contact from its address book",
        inputSchema: {
          type: "object",
          properties: {
            contactId: { type: "string", description: "Contact UID (from searchContacts results)" },
          },
          required: ["contactId"],
        },
      },
      {
        name: "replyToMessage",
        group: "messages", crud: "create",
        title: "Reply to Message",
        description: "Reply in a compose window with quoted original text for review. The skipReview safety block is on by default; direct sending is honored only when the user explicitly disables that preference.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID to reply to (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            body: { type: "string", description: "Reply body text" },
            replyAll: { type: "boolean", description: "Reply to all recipients (default: false)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            to: { type: "string", description: "Override recipient email (default: original sender)" },
            cc: { type: "string", description: "CC recipients (comma-separated)" },
            bcc: { type: "string", description: "BCC recipients (comma-separated)" },
            from: { type: "string", description: "Sender identity (email address or identity ID from listAccounts)" },
            skipReview: { type: "boolean", description: "Request direct sending without a compose window. Honored only when the user explicitly disables the default-on skipReview safety block (default: false)." },
            attachments: {
              type: "array",
              maxItems: MAX_ATTACHMENTS_PER_MESSAGE,
              description: "Attachments: file paths (strings) or inline objects ({name, contentType, base64})",
              items: {
                oneOf: [
                  { type: "string", description: "Absolute file path to attach" },
                  {
                    type: "object",
                    properties: {
                      name: { type: "string", minLength: 1, description: "Attachment filename" },
                      contentType: { type: "string", description: "MIME type, e.g. application/pdf" },
                      base64: { type: "string", minLength: 1, contentEncoding: "base64", description: "Base64-encoded file content" },
                      content: { type: "string", minLength: 1, contentEncoding: "base64", description: "Alias for base64 (accepted for backwards compatibility); base64 takes precedence when both are set" },
                    },
                    required: ["name"],
                    anyOf: [
                      { type: "object", required: ["base64"] },
                      { type: "object", required: ["content"] },
                    ],
                    additionalProperties: false,
                  },
                ],
              },
            },
          },
          required: ["messageId", "folderPath", "body"],
        },
      },
      {
        name: "forwardMessage",
        group: "messages", crud: "create",
        title: "Forward Message",
        description: "Forward in a compose window with original content for review. The skipReview safety block is on by default; direct sending is honored only when the user explicitly disables that preference.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID to forward (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            to: { type: "string", description: "Recipient email address" },
            body: { type: "string", description: "Additional text to prepend (optional)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            cc: { type: "string", description: "CC recipients (comma-separated)" },
            bcc: { type: "string", description: "BCC recipients (comma-separated)" },
            from: { type: "string", description: "Sender identity (email address or identity ID from listAccounts)" },
            skipReview: { type: "boolean", description: "Request direct sending without a compose window. Honored only when the user explicitly disables the default-on skipReview safety block (default: false)." },
            attachments: {
              type: "array",
              maxItems: MAX_ATTACHMENTS_PER_MESSAGE,
              description: "Additional attachments: file paths (strings) or inline objects ({name, contentType, base64})",
              items: {
                oneOf: [
                  { type: "string", description: "Absolute file path to attach" },
                  {
                    type: "object",
                    properties: {
                      name: { type: "string", minLength: 1, description: "Attachment filename" },
                      contentType: { type: "string", description: "MIME type, e.g. application/pdf" },
                      base64: { type: "string", minLength: 1, contentEncoding: "base64", description: "Base64-encoded file content" },
                      content: { type: "string", minLength: 1, contentEncoding: "base64", description: "Alias for base64 (accepted for backwards compatibility); base64 takes precedence when both are set" },
                    },
                    required: ["name"],
                    anyOf: [
                      { type: "object", required: ["base64"] },
                      { type: "object", required: ["content"] },
                    ],
                    additionalProperties: false,
                  },
                ],
              },
            },
          },
          required: ["messageId", "folderPath", "to"],
        },
      },
      {
        name: "getRecentMessages",
        group: "messages", crud: "read",
        title: "Get Recent Messages",
        description: "Get recent messages sorted newest-first from a specific folder or all Inboxes, with date and unread filtering",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Folder URI (from listFolders) to list messages from. If omitted, returns messages from all Inboxes." },
            daysBack: { type: "number", description: "Only return messages from the last N days (default: 7). Use a larger value like 365 for older messages." },
            maxResults: { type: "number", description: "Maximum number of results (default: 50, max: 200)" },
            offset: { type: "number", description: "Number of results to skip for pagination (default 0). When provided, returns {messages, totalMatches, offset, limit, hasMore} instead of a plain array. Note: totalMatches is capped at 10000." },
            unreadOnly: { type: "boolean", description: "Only return unread messages (default: false)" },
            flaggedOnly: { type: "boolean", description: "Only return flagged/starred messages (default: false)" },
            includeSubfolders: { type: "boolean", description: "If false, only return messages from the specified folder — not its subfolders. Default: true." },
          },
          required: [],
        },
      },
      {
        name: "displayMessage",
        group: "messages", crud: "read",
        title: "Display Message",
        description: "Open or navigate to a message in the Thunderbird GUI. Use '3pane' (default) to select the message in the mail view, 'tab' to open in a new tab, or 'window' to open in a standalone window.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            displayMode: { type: "string", enum: ["3pane", "tab", "window"], description: "How to display: '3pane' (navigate in mail view, default), 'tab' (new tab), or 'window' (new window)" },
          },
          required: ["messageId", "folderPath"],
        },
      },
      {
        name: "deleteMessages",
        group: "messages", crud: "delete",
        title: "Delete Messages",
        description: "Delete messages from a folder. Drafts are moved to Trash instead of permanently deleted.",
        inputSchema: {
          type: "object",
          properties: {
            messageIds: { type: "array", items: { type: "string" }, description: "Array of message IDs to delete" },
            folderPath: { type: "string", description: "The folder URI containing the messages (from listFolders or searchMessages results)" },
          },
          required: ["messageIds", "folderPath"],
        },
      },
      {
        name: "updateMessage",
        group: "messages", crud: "update",
        title: "Update Message",
        description: "Update one or more messages' read/flagged/tagged state and optionally move them. Supply messageId for a single message or messageIds for bulk operations. Tags are Thunderbird keywords (e.g. '$label1' for Important, '$label2' for Work, or any custom string). Note: combining tags with moveTo/trash on IMAP may not preserve tags on the moved copy — use separate calls if needed.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "A single message ID (from searchMessages results). Required unless messageIds is provided." },
            messageIds: { type: "array", items: { type: "string" }, description: "Array of message IDs for bulk operations. Required unless messageId is provided." },
            folderPath: { type: "string", description: "The folder URI containing the message(s) (from searchMessages results)" },
            read: { type: "boolean", description: "Set to true/false to mark read/unread (optional)" },
            flagged: { type: "boolean", description: "Set to true/false to flag/unflag (optional)" },
            addTags: { type: "array", items: { type: "string" }, description: "Tag keywords to add (e.g. ['$label1', 'project-x']). Thunderbird built-in tags: $label1 (Important), $label2 (Work), $label3 (Personal), $label4 (To Do), $label5 (Later)" },
            removeTags: { type: "array", items: { type: "string" }, description: "Tag keywords to remove from the message(s)" },
            moveTo: { type: "string", description: "Destination folder URI (optional). Cannot be used with trash." },
            trash: { type: "boolean", description: "Set to true to move message to Trash (optional). Cannot be used with moveTo." },
          },
          required: ["folderPath"],
        },
      },
      {
        name: "createFolder",
        group: "folders", crud: "create",
        title: "Create Folder",
        description: "Create a new mail subfolder under an existing folder. Note: on IMAP accounts, server-side completion is asynchronous; verify with listFolders.",
        inputSchema: {
          type: "object",
          properties: {
            parentFolderPath: { type: "string", description: "URI of the parent folder (from listFolders)" },
            name: { type: "string", description: "Name for the new subfolder" },
          },
          required: ["parentFolderPath", "name"],
        },
      },
      {
        name: "renameFolder",
        group: "folders", crud: "update",
        title: "Rename Folder",
        description: "Rename an existing mail folder. Note: on IMAP accounts, server-side completion is asynchronous; verify with listFolders.",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "URI of the folder to rename (from listFolders)" },
            newName: { type: "string", description: "New name for the folder" },
          },
          required: ["folderPath", "newName"],
        },
      },
      {
        name: "deleteFolder",
        group: "folders", crud: "delete",
        title: "Delete Folder",
        description: "Delete a mail folder and all its contents. Moves to Trash, or permanently deletes if already in Trash. Note: permanent deletion may prompt the user for confirmation. On IMAP accounts, server-side completion is asynchronous; verify with listFolders.",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "URI of the folder to delete (from listFolders)" },
          },
          required: ["folderPath"],
        },
      },
      {
        name: "emptyTrash",
        group: "folders", crud: "delete",
        title: "Empty Trash",
        description: "Permanently delete all messages in the Trash folder for an account.",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID (from listAccounts). If omitted, empties Trash for all accessible accounts." },
          },
          required: [],
        },
      },
      {
        name: "emptyJunk",
        group: "folders", crud: "delete",
        title: "Empty Junk",
        description: "Permanently delete all messages in the Junk/Spam folder for an account.",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID (from listAccounts). If omitted, empties Junk for all accessible accounts." },
          },
          required: [],
        },
      },
      {
        name: "moveFolder",
        group: "folders", crud: "update",
        title: "Move Folder",
        description: "Move a mail folder to a new parent folder within the same account. Note: on IMAP accounts, server-side completion is asynchronous; verify with listFolders.",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "URI of the folder to move (from listFolders)" },
            newParentPath: { type: "string", description: "URI of the destination parent folder (from listFolders)" },
          },
          required: ["folderPath", "newParentPath"],
        },
      },
      {
        name: "listFilters",
        group: "filters", crud: "read",
        title: "List Filters",
        description: "List all mail filters/rules for an account with their conditions and actions",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID from listAccounts (omit for all accounts)" },
          },
          required: [],
        },
      },
      {
        name: "createFilter",
        group: "filters", crud: "create",
        title: "Create Filter",
        description: "Create a new mail filter rule on an account",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID" },
            name: { type: "string", description: "Filter name" },
            enabled: { type: "boolean", description: "Whether filter is active (default: true)" },
            type: { type: "number", description: "Filter type bitmask (default: 17 = inbox + manual). 1=inbox, 16=manual, 32=post-plugin, 64=post-outgoing" },
            conditions: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  attrib: { type: "string", description: "Attribute: subject, from, to, cc, toOrCc, body, date, priority, status, size, ageInDays, hasAttachment, junkStatus, tag, otherHeader" },
                  op: { type: "string", description: "Operator: contains, doesntContain, is, isnt, isEmpty, beginsWith, endsWith, isGreaterThan, isLessThan, isBefore, isAfter, matches, doesntMatch" },
                  value: { type: "string", description: "Value to match against" },
                  booleanAnd: { type: "boolean", description: "true=AND with previous, false=OR (default: true)" },
                  header: { type: "string", description: "Custom header name (only when attrib is otherHeader)" },
                },
              },
              description: "Array of filter conditions",
            },
            actions: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  type: { type: "string", description: "Action: moveToFolder, copyToFolder, markRead, markUnread, markFlagged, addTag, changePriority, delete, stopExecution, forward, reply" },
                  value: { type: "string", description: "Action parameter (folder URI for move/copy, tag name for addTag, priority for changePriority, email for forward)" },
                },
              },
              description: "Array of actions to perform",
            },
            insertAtIndex: { type: "number", description: "Position to insert (0 = top priority, default: end of list)" },
          },
          required: ["accountId", "name", "conditions", "actions"],
        },
      },
      {
        name: "updateFilter",
        group: "filters", crud: "update",
        title: "Update Filter",
        description: "Modify an existing filter's properties, conditions, or actions",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID" },
            filterIndex: { type: "number", description: "Filter index (from listFilters)" },
            name: { type: "string", description: "New filter name (optional)" },
            enabled: { type: "boolean", description: "Enable/disable (optional)" },
            type: { type: "number", description: "New filter type bitmask (optional)" },
            conditions: {
              type: "array",
              description: "Replace all conditions (optional, same format as createFilter)",
              items: {
                type: "object",
                properties: {
                  attrib: { type: "string", description: "Attribute: subject, from, to, cc, toOrCc, body, date, priority, status, size, ageInDays, hasAttachment, junkStatus, tag, otherHeader" },
                  op: { type: "string", description: "Operator: contains, doesntContain, is, isnt, isEmpty, beginsWith, endsWith, isGreaterThan, isLessThan, isBefore, isAfter, matches, doesntMatch" },
                  value: { type: "string", description: "Value to match against" },
                  booleanAnd: { type: "boolean", description: "true=AND with previous, false=OR (default: true)" },
                  header: { type: "string", description: "Custom header name (only when attrib is otherHeader)" },
                },
              },
            },
            actions: {
              type: "array",
              description: "Replace all actions (optional, same format as createFilter)",
              items: {
                type: "object",
                properties: {
                  type: { type: "string", description: "Action: moveToFolder, copyToFolder, markRead, markUnread, markFlagged, addTag, changePriority, delete, stopExecution, forward, reply" },
                  value: { type: "string", description: "Action parameter (folder URI for move/copy, tag name for addTag, priority for changePriority, email for forward)" },
                },
              },
            },
          },
          required: ["accountId", "filterIndex"],
        },
      },
      {
        name: "deleteFilter",
        group: "filters", crud: "delete",
        title: "Delete Filter",
        description: "Delete a mail filter by index",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID" },
            filterIndex: { type: "number", description: "Filter index to delete (from listFilters)" },
          },
          required: ["accountId", "filterIndex"],
        },
      },
      {
        name: "reorderFilters",
        group: "filters", crud: "update",
        title: "Reorder Filters",
        description: "Move a filter to a different position in the execution order",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID" },
            fromIndex: { type: "number", description: "Current filter index" },
            toIndex: { type: "number", description: "Target index (0 = highest priority)" },
          },
          required: ["accountId", "fromIndex", "toIndex"],
        },
      },
      {
        name: "applyFilters",
        group: "filters", crud: "update",
        title: "Apply Filters",
        description: "Manually run all enabled filters on a folder to organize existing messages",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID (uses its filters)" },
            folderPath: { type: "string", description: "Folder URI to apply filters to (from listFolders)" },
          },
          required: ["accountId", "folderPath"],
        },
      },
      {
        name: "getAccountAccess",
        group: "system", crud: "read",
        title: "Get Account Access",
        description: "Get the current account access control list. Shows which accounts the MCP server can access. Account access is configured by the user in the extension settings page (Tools > Add-ons > Thunderbird MCP > Options) and cannot be changed via MCP tools.",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      ];
    }
    // END TOOL SCHEMA BUILDER

    const tools = buildTools();

    // Validate tool metadata: every tool must have valid group and crud fields.
    // This prevents tools from being silently hidden in the settings UI.
    const toolErrors = [];
    for (const tool of tools) {
      if (!tool.group || !VALID_GROUPS.includes(tool.group)) {
        toolErrors.push(`Tool "${tool.name}" has invalid or missing group: "${tool.group}" (valid: ${VALID_GROUPS.join(", ")})`);
      }
      if (!tool.crud || !VALID_CRUD.includes(tool.crud)) {
        toolErrors.push(`Tool "${tool.name}" has invalid or missing crud: "${tool.crud}" (valid: ${VALID_CRUD.join(", ")})`);
      }
    }
    if (toolErrors.length > 0) {
      console.error("thunderbird-mcp: Tool metadata validation failed:\n  " + toolErrors.join("\n  "));
    }

    // Derive ALL_TOOL_NAMES from the tools array (single source of truth)
    const ALL_TOOL_NAMES = tools.map(t => t.name);

    // Group display order for settings UI
    const GROUP_ORDER = { system: 0, messages: 1, folders: 2, contacts: 3, calendar: 4, filters: 5 };
    // Group display labels
    const GROUP_LABELS = { system: "System", messages: "Messages", folders: "Folders", contacts: "Contacts", calendar: "Calendar", filters: "Filters" };

    /**
     * Generate a cryptographically random auth token (hex string).
     * Used to authenticate bridge requests to the HTTP server.
     */
    function generateAuthToken() {
      const bytes = new Uint8Array(32);
      // crypto.getRandomValues is not available in Thunderbird experiment API scope;
      // use the XPCOM random generator instead.
      const rng = Cc["@mozilla.org/security/random-generator;1"]
        .createInstance(Ci.nsIRandomGenerator);
      const randomBytes = rng.generateRandomBytes(32);
      for (let i = 0; i < 32; i++) bytes[i] = randomBytes[i];
      return Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
    }

    function getStableAuthTokenPref() {
      try {
        const pref = Services.prefs.getStringPref(PREF_STABLE_AUTH_TOKEN, "");
        const token = pref.trim();
        if (!token) {
          if (pref) {
            console.warn("thunderbird-mcp: stableAuthToken preference is malformed; expected 64 lowercase hex characters, ignoring stored value");
          }
          return "";
        }
        if (!AUTH_TOKEN_PATTERN.test(token)) {
          console.warn("thunderbird-mcp: stableAuthToken preference is malformed; expected 64 lowercase hex characters, ignoring stored value");
          return "";
        }
        return token;
      } catch {
        return "";
      }
    }

    function readConnectionInfo() {
      const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
      tmpDir.append("thunderbird-mcp");
      const connFile = tmpDir.clone();
      connFile.append("connection.json");
      if (!connFile.exists()) {
        return { path: connFile.path, data: null };
      }
      const fis = Cc["@mozilla.org/network/file-input-stream;1"]
        .createInstance(Ci.nsIFileInputStream);
      fis.init(connFile, 0x01, 0, 0);
      const sis = Cc["@mozilla.org/scriptableinputstream;1"]
        .createInstance(Ci.nsIScriptableInputStream);
      sis.init(fis);
      const text = sis.read(sis.available());
      sis.close();
      return { path: connFile.path, data: JSON.parse(text) };
    }

    /**
     * Remove the connection info file during startup or shutdown cleanup.
     */
    function removeConnectionInfo() {
      try {
        const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
        tmpDir.append("thunderbird-mcp");
        const connFile = tmpDir.clone();
        connFile.append("connection.json");
        if (connFile.exists()) {
          connFile.remove(false);
        }
      } catch {
        // Best-effort cleanup
      }
    }

    return {
      mcpServer: {
        start: async function() {
          // Guard against double-start on extension reload (port conflict)
          if (globalThis.__tbMcpStartPromise) {
            return await globalThis.__tbMcpStartPromise;
          }
          const startPromise = (async () => {
          try {
            // Stop any previously running server (e.g. extension reload)
            if (globalThis.__tbMcpServer) {
              try { globalThis.__tbMcpServer.stop(() => {}); } catch { /* ignore */ }
              globalThis.__tbMcpServer = null;
              stopConnectionInfoRefreshTimer();
            }
            const { HttpServer } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/httpd.sys.mjs?" + Date.now()
            );
            const { NetUtil } = ChromeUtils.importESModule(
              "resource://gre/modules/NetUtil.sys.mjs"
            );
            const { MailServices } = ChromeUtils.importESModule(
              "resource:///modules/MailServices.sys.mjs"
            );
            let VCardPropertyEntry;
            try {
              ({ VCardPropertyEntry } = ChromeUtils.importESModule(
                "resource:///modules/VCardUtils.sys.mjs"
              ));
            } catch {
              ({ VCardPropertyEntry } = ChromeUtils.import(
                "resource:///modules/VCardUtils.jsm"
              ));
            }

            let cal = null;
            let CalEvent = null;
            let CalTodo = null;
            try {
              const calModule = ChromeUtils.importESModule(
                "resource:///modules/calendar/calUtils.sys.mjs"
              );
              cal = calModule.cal;
              const { CalEvent: CE } = ChromeUtils.importESModule(
                "resource:///modules/CalEvent.sys.mjs"
              );
              CalEvent = CE;
              const { CalTodo: CT } = ChromeUtils.importESModule(
                "resource:///modules/CalTodo.sys.mjs"
              );
              CalTodo = CT;
            } catch {
              // Calendar not available
            }

            let GlodaMsgSearcher = null;
            try {
              const glodaModule = ChromeUtils.importESModule(
                "resource:///modules/gloda/GlodaMsgSearcher.sys.mjs"
              );
              GlodaMsgSearcher = glodaModule.GlodaMsgSearcher;
            } catch {
              // Gloda not available
            }

            /**
             * CRITICAL: Must specify { charset: "UTF-8" } or emojis/special chars
             * will be corrupted. NetUtil defaults to Latin-1.
             */
            function readRequestBody(request) {
              const stream = request.bodyInputStream;
              return NetUtil.readInputStreamToString(stream, stream.available(), { charset: "UTF-8" });
            }

            /**
             * Apply offset-based pagination to a sorted results array.
             * Removes the internal _dateTs property from each result.
             *
             * Backward-compatible: when offset is undefined/null (not provided),
             * returns a plain array. When offset is explicitly provided (even 0),
             * returns structured { messages, totalMatches, offset, limit, hasMore }.
             * Note: totalMatches is capped at SEARCH_COLLECTION_CAP and may underreport.
             */
            function paginate(results, offset, effectiveLimit) {
              const offsetProvided = offset !== undefined && offset !== null;
              const effectiveOffset = (offset > 0) ? Math.floor(offset) : 0;
              const page = results.slice(effectiveOffset, effectiveOffset + effectiveLimit).map(r => {
                delete r._dateTs;
                return r;
              });
              if (!offsetProvided) {
                return page;
              }
              return {
                messages: page,
                totalMatches: results.length,
                offset: effectiveOffset,
                limit: effectiveLimit,
                hasMore: effectiveOffset + effectiveLimit < results.length
              };
            }

            function normalizeMessageIdForDedup(value) {
              // Compare RFC Message-IDs without surrounding angle brackets / whitespace.
              // Case is preserved on purpose: the local part of a Message-ID is
              // case-sensitive per RFC 5322, so lowercasing could collapse two
              // genuinely-distinct messages and hide one. Showing a duplicate is the
              // safe failure direction; hiding a message is not.
              if (value === undefined || value === null) return "";
              let normalized = String(value).trim();
              if (!normalized) return "";
              if (normalized.startsWith("<") && normalized.endsWith(">")) {
                normalized = normalized.slice(1, -1).trim();
              }
              return normalized;
            }

            function dedupeSearchMessageResults(results) {
              const seen = new Map();
              const deduped = [];

              function addDupLocation(survivor, folderPath) {
                if (!folderPath || folderPath === survivor.folderPath) return;
                if (!Array.isArray(survivor.dupLocations)) survivor.dupLocations = [];
                if (!survivor.dupLocations.includes(folderPath)) {
                  survivor.dupLocations.push(folderPath);
                }
              }

              function mergeDupLocations(survivor, row) {
                addDupLocation(survivor, row.folderPath);
                if (Array.isArray(row.dupLocations)) {
                  for (const folderPath of row.dupLocations) {
                    addDupLocation(survivor, folderPath);
                  }
                }
              }

              for (const row of results) {
                const normalizedId = normalizeMessageIdForDedup(row?.id);
                if (!normalizedId) {
                  deduped.push(row);
                  continue;
                }

                const survivor = seen.get(normalizedId);
                if (survivor) {
                  mergeDupLocations(survivor, row);
                  continue;
                }

                if (Array.isArray(row.dupLocations)) {
                  const existingDupLocations = row.dupLocations;
                  delete row.dupLocations;
                  for (const folderPath of existingDupLocations) {
                    addDupLocation(row, folderPath);
                  }
                }
                seen.set(normalizedId, row);
                deduped.push(row);
              }

              return deduped;
            }

            /**
             * Write connection info (port + auth token) to a well-known file
             * so the bridge can discover how to connect.
             * File: <TmpD>/thunderbird-mcp/connection.json
             */
            function writeConnectionInfo(port, token) {
              const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
              tmpDir.append("thunderbird-mcp");
              if (!tmpDir.exists()) {
                tmpDir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
              } else if (tmpDir.isSymlink()) {
                throw new Error("thunderbird-mcp tmp directory is a symlink — refusing to write connection info");
              } else {
                // POSIX hardening: on a shared /tmp another local user could
                // pre-create the directory with group/world bits set, then race
                // the connection file. The O_EXCL on the file itself blocks a
                // straight overwrite, but a permissive directory still lets the
                // attacker read or rename our file. Force perms back to 0o700.
                // permissions is 0 on platforms that don't expose POSIX modes
                // (Windows ACLs), so the chmod is a no-op there.
                try {
                  const mode = tmpDir.permissions;
                  if (mode && (mode & 0o077) !== 0) {
                    try { tmpDir.permissions = 0o700; } catch { /* best-effort */ }
                    if ((tmpDir.permissions & 0o077) !== 0) {
                      throw new Error("thunderbird-mcp tmp directory has group/world permissions — refusing to write connection info");
                    }
                  }
                } catch (e) {
                  if (e && e.message && e.message.startsWith("thunderbird-mcp tmp directory")) throw e;
                  // ignore: permissions accessor unsupported on this platform
                }
              }
              const connFile = tmpDir.clone();
              connFile.append("connection.json");
              // Symlink defense: remove any existing file first, then create
              // with O_CREAT|O_EXCL (0x08|0x80) to fail if a symlink appeared
              // between remove and create.
              if (connFile.exists()) {
                connFile.remove(false);
              }
              const data = JSON.stringify({ port, token, pid: Services.appinfo.processID });
              const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                .createInstance(Ci.nsIFileOutputStream);
              // 0x02 = O_WRONLY, 0x08 = O_CREAT, 0x80 = O_EXCL
              ostream.init(connFile, 0x02 | 0x08 | 0x80, 0o600, 0);
              const converter = Cc["@mozilla.org/intl/converter-output-stream;1"]
                .createInstance(Ci.nsIConverterOutputStream);
              converter.init(ostream, "UTF-8");
              converter.writeString(data);
              converter.close();
              return connFile.path;
            }

            function ensureConnectionInfo(port, token) {
              return ensureFreshConnectionInfo({
                port,
                token,
                expectedPid: Services.appinfo.processID,
                readConnectionInfo,
                writeConnectionInfo,
                onCheckError: (e) => {
                  console.warn("thunderbird-mcp: connection info check failed; rewriting:", e);
                },
              });
            }

            function startConnectionInfoRefresh(port, token) {
              stopConnectionInfoRefreshTimer();
              const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
              timer.initWithCallback(() => {
                try {
                  ensureConnectionInfo(port, token);
                } catch (e) {
                  console.warn("thunderbird-mcp: failed to refresh connection info:", e);
                }
              }, CONNECTION_FILE_REFRESH_MS, Ci.nsITimer.TYPE_REPEATING_SLACK);
              globalThis.__tbMcpConnectionInfoRefreshTimer = timer;
            }

            const authToken = getStableAuthTokenPref() || generateAuthToken();

            /**
             * Constant-time string comparison to prevent timing side-channel attacks.
             */
            function timingSafeEqual(a, b) {
              const aStr = String(a);
              const bStr = String(b);
              const len = Math.max(aStr.length, bStr.length);
              let result = aStr.length ^ bStr.length;
              for (let i = 0; i < len; i++) {
                result |= (aStr.charCodeAt(i) || 0) ^ (bStr.charCodeAt(i) || 0);
              }
              return result === 0;
            }

            /**
             * Get the list of allowed account IDs from preferences.
             * Returns an empty array if no restriction is set (all accounts allowed).
             */
            function getAllowedAccountIds() {
              try {
                const pref = Services.prefs.getStringPref(PREF_ALLOWED_ACCOUNTS, "");
                if (!pref) return [];
                const parsed = JSON.parse(pref);
                if (!Array.isArray(parsed)) {
                  console.error("thunderbird-mcp: allowed accounts pref is not an array, blocking all accounts");
                  return ["__invalid__"];
                }
                return parsed;
              } catch (e) {
                // Fail closed: corrupt pref means block all accounts, not allow all
                console.error("thunderbird-mcp: failed to parse allowed accounts pref, blocking all accounts:", e);
                return ["__invalid__"];
              }
            }

            /**
             * Check if an account is accessible based on the allowed accounts list.
             * When the list is empty, all accounts are accessible (default).
             */
            function isAccountAllowed(accountKey) {
              const allowed = getAllowedAccountIds();
              if (allowed.length === 0) return true;
              return allowed.includes(accountKey);
            }

            /**
             * Check if the user has disabled the skipReview shortcut.
             * When true, send/reply/forward/createEvent/createTask tools must open
             * the review window/dialog even if the caller passed skipReview: true.
             *
             * Default is true: an LLM that reads attacker-controlled email content
             * can be prompt-injected into invoking sendMail with skipReview, so the
             * safe default is to require human review. Users can explicitly opt
             * into silent sends from the options page.
             */
            function isSkipReviewBlocked() {
              try {
                return Services.prefs.getBoolPref(PREF_BLOCK_SKIPREVIEW, true);
              } catch {
                // Fail closed: if we can't read the pref, assume blocked so the
                // user retains ability to review before send.
                return true;
              }
            }

            /**
             * Get the list of disabled tool names from preferences.
             * Returns an empty array if no tools are disabled (all enabled).
             * Fails closed: corrupt pref disables all tools.
             */
            function getDisabledTools() {
              try {
                const pref = Services.prefs.getStringPref(PREF_DISABLED_TOOLS, "");
                if (!pref) return [];
                const parsed = JSON.parse(pref);
                if (!Array.isArray(parsed) || !parsed.every(v => typeof v === "string")) {
                  console.error("thunderbird-mcp: disabled tools pref is invalid, disabling all tools");
                  return ["__all__"];
                }
                return parsed;
              } catch (e) {
                console.error("thunderbird-mcp: failed to parse disabled tools pref, disabling all tools:", e);
                return ["__all__"];
              }
            }

            /**
             * Check if a tool is enabled.
             * Undisableable tools (listAccounts, listFolders, getAccountAccess) always return true.
             */
            function isToolEnabled(toolName) {
              if (UNDISABLEABLE_TOOLS.has(toolName)) return true;
              const disabled = getDisabledTools();
              if (disabled.includes("__all__")) return false;
              return !disabled.includes(toolName);
            }

            /**
             * Check if a resolved folder belongs to an allowed account.
             * Returns true if the folder's account is accessible, false otherwise.
             */
            function isFolderAccessible(folder) {
              if (!folder || !folder.server) return false;
              const account = MailServices.accounts.findAccountForServer(folder.server);
              return account ? isAccountAllowed(account.key) : false;
            }

            /**
             * Lookup a folder by URI and verify it exists and is accessible.
             * Returns { folder } on success, or { error } if not found or restricted.
             */
            function getAccessibleFolder(folderPath) {
              const folder = MailServices.folderLookup.getFolderForURL(folderPath);
              if (!folder) return { error: `Folder not found: ${folderPath}` };
              if (!isFolderAccessible(folder)) return { error: `Account not accessible for folder: ${folderPath}` };
              return { folder };
            }

            /**
             * Get all accessible Thunderbird accounts, filtered by allowed list.
             */
            function getAccessibleAccounts() {
              const result = [];
              for (const account of MailServices.accounts.accounts) {
                if (isAccountAllowed(account.key)) {
                  result.push(account);
                }
              }
              return result;
            }

            /**
             * Best-effort refresh for IMAP folders. updateFolder starts async work
             * and may not complete before callers read from Thunderbird's cache.
             */
            function refreshImapFolderSync(folder) {
              if (folder.server && folder.server.type === "imap") {
                try {
                  folder.updateFolder(null);
                } catch {
                  // updateFolder may fail, continue anyway
                }
              }
            }

            function toColumnarTable(items, keys) {
              const columns = Array.from(keys).sort();
              return {
                columns,
                rows: items.map(item => columns.map(column => item[column])),
              };
            }

            function listAccounts() {
              const accounts = [];
              for (const account of getAccessibleAccounts()) {
                const server = account.incomingServer;
                const identities = [];
                for (const identity of account.identities) {
                  identities.push({
                    id: identity.key,
                    email: identity.email,
                    name: identity.fullName,
                    isDefault: identity === account.defaultIdentity
                  });
                }
                accounts.push({
                  id: account.key,
                  name: server.prettyName,
                  type: server.type,
                  identities
                });
              }
              return accounts;
            }

            /**
             * Get the current account access control list.
             */
            function getAccountAccess() {
              const allowed = getAllowedAccountIds();
              // Only return accessible accounts — restricted accounts are hidden
              const accessibleAccounts = [];
              for (const account of MailServices.accounts.accounts) {
                if (!isAccountAllowed(account.key)) continue;
                const server = account.incomingServer;
                accessibleAccounts.push({
                  id: account.key,
                  name: server.prettyName,
                  type: server.type,
                });
              }
              return {
                mode: allowed.length === 0 ? "all" : "restricted",
                accounts: accessibleAccounts,
              };
            }

            /**
             * Lists all folders (optionally limited to a single account).
             * Depth is 0 for root children, increasing for subfolders.
             */
            function listFolders(accountId, folderPath, format) {
              const results = [];
              const outputFormat = format == null ? "objects" : format;
              const folderKeys = ["name", "path", "type", "accountId", "totalMessages", "unreadMessages", "depth"];

              if (outputFormat !== "objects" && outputFormat !== "table") {
                return { error: `Invalid format: "${outputFormat}". Must be one of: objects, table` };
              }

              function formatFolderResults() {
                return outputFormat === "table" ? toColumnarTable(results, folderKeys) : results;
              }

              function folderType(flags) {
                if (flags & 0x00001000) return "inbox";
                if (flags & 0x00000200) return "sent";
                if (flags & 0x00000400) return "drafts";
                if (flags & 0x00000100) return "trash";
                if (flags & 0x00400000) return "templates";
                if (flags & 0x00000800) return "queue";
                if (flags & 0x40000000) return "junk";
                if (flags & 0x00004000) return "archive";
                return "folder";
              }

              function walkFolder(folder, accountKey, depth) {
                try {
                  // Skip virtual/search folders to avoid duplicates
                  if (folder.flags & 0x00000020) return;

                  const prettyName = folder.prettyName;
                  results.push({
                    name: prettyName || folder.name || "(unnamed)",
                    path: folder.URI,
                    type: folderType(folder.flags),
                    accountId: accountKey,
                    totalMessages: folder.getTotalMessages(false),
                    unreadMessages: folder.getNumUnread(false),
                    depth
                  });
                } catch (e) {
                  console.warn("thunderbird-mcp: listFolders skipped inaccessible folder", folder?.URI || folder?.name, e);
                }

                try {
                  if (folder.hasSubFolders) {
                    for (const subfolder of folder.subFolders) {
                      walkFolder(subfolder, accountKey, depth + 1);
                    }
                  }
                } catch (e) {
                  console.warn("thunderbird-mcp: listFolders subfolder traversal failed for folder", folder?.URI || folder?.name, e);
                }
              }

              // folderPath filter: list that folder and its subtree
              if (folderPath) {
                const result = getAccessibleFolder(folderPath);
                if (result.error) return result;
                const folder = result.folder;
                const accountKey = folder.server
                  ? (MailServices.accounts.findAccountForServer(folder.server)?.key || "unknown")
                  : "unknown";
                walkFolder(folder, accountKey, 0);
                return formatFolderResults();
              }

              if (accountId) {
                if (!isAccountAllowed(accountId)) {
                  return { error: `Account not accessible: "${accountId}". Call listAccounts to see which account IDs are available.` };
                }
                let target = null;
                for (const account of MailServices.accounts.accounts) {
                  if (account.key === accountId) {
                    target = account;
                    break;
                  }
                }
                if (!target) {
                  return { error: `Account not found: "${accountId}". Account IDs come from listAccounts (internal keys like "account1"), not email addresses. Omit accountId to list folders across all accounts.` };
                }
                try {
                  const root = target.incomingServer.rootFolder;
                  if (root && root.hasSubFolders) {
                    for (const subfolder of root.subFolders) {
                      walkFolder(subfolder, target.key, 0);
                    }
                  }
                } catch (e) {
                  console.warn("thunderbird-mcp: listFolders failed to enumerate account root", target.key || accountId, e);
                }
                return formatFolderResults();
              }

              for (const account of getAccessibleAccounts()) {
                try {
                  const root = account.incomingServer.rootFolder;
                  if (!root) continue;
                  if (root.hasSubFolders) {
                    for (const subfolder of root.subFolders) {
                      walkFolder(subfolder, account.key, 0);
                    }
                  }
                } catch (e) {
                  console.warn("thunderbird-mcp: listFolders failed to enumerate account", account?.key, e);
                }
              }

              return formatFolderResults();
            }

            /**
             * Searches the given accounts for an identity matching emailOrId
             * (by key or case-insensitive email).
             * Returns the identity object, or null if not found.
             */
            function findIdentityIn(accounts, emailOrId) {
              if (!emailOrId) return null;
              const lowerInput = emailOrId.toLowerCase();
              for (const account of accounts) {
                for (const identity of account.identities) {
                  if (identity.key === emailOrId || (identity.email || "").toLowerCase() === lowerInput) {
                    return identity;
                  }
                }
              }
              return null;
            }

            /**
             * Finds an identity by email address or identity ID
             * among accessible accounts only.  Returns null if not found.
             */
            function findIdentity(emailOrId) {
              return findIdentityIn(getAccessibleAccounts(), emailOrId);
            }

            /** Creates an nsIFile instance for the given path. */
            function createLocalFile(path) {
              const file = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
              file.initWithPath(path);
              return file;
            }

            /** Returns user-visible tag keywords from a message header, filtering out internal IMAP flags. */
            function getUserTags(msgHdr) {
              return (msgHdr.getStringProperty("keywords") || "").split(/\s+/).filter(k => k && !INTERNAL_KEYWORDS.has(k.toLowerCase()));
            }

            /**
             * Converts attachment entries to attachment descriptors.
             * Each entry can be:
             *   - A string (file path) — resolved from disk
             *   - An object { name, contentType, base64 } — decoded and written
             *     to a temp file under <TmpD>/thunderbird-mcp/attachments/
             * Returns { descs: [{url, name, size, contentType?}], failed: string[] }
             */
            // BEGIN OUTBOUND ATTACHMENT CONVERSION
            function filePathsToAttachDescs(filePaths) {
              const descs = [];
              const failed = [];
              if (!filePaths || !Array.isArray(filePaths)) return { descs, failed };
              let attachmentEntries = filePaths;
              if (filePaths.length > MAX_ATTACHMENTS_PER_MESSAGE) {
                failed.push(`Attachment count ${filePaths.length} exceeds the ${MAX_ATTACHMENTS_PER_MESSAGE} attachment limit; skipped ${filePaths.length - MAX_ATTACHMENTS_PER_MESSAGE} attachment(s)`);
                attachmentEntries = filePaths.slice(0, MAX_ATTACHMENTS_PER_MESSAGE);
              }
              let totalAttachmentBytes = 0;
              for (const entry of attachmentEntries) {
                try {
                  if (typeof entry === "string") {
                    // File path attachment.
                    //
                    // SECURITY: reject paths that point at credentials, system
                    // files, or browser/mail profile data BEFORE touching the
                    // filesystem. This is the LLM-confused-deputy defense:
                    // attacker-controlled email content can prompt-inject an
                    // assistant into calling sendMail with attachments=["/path/to/id_rsa"]
                    // and we never want that to succeed regardless of skipReview.
                    if (isSensitiveFilePath(entry)) {
                      failed.push(`${entry} (sensitive path blocked)`);
                      continue;
                    }
                    const file = createLocalFile(entry);
                    if (!file.exists()) {
                      failed.push(entry);
                      continue;
                    }
                    let isSymlink;
                    try {
                      isSymlink = file.isSymlink();
                    } catch {
                      failed.push(`${entry} (symlink check failed)`);
                      continue;
                    }
                    if (isSymlink) {
                      failed.push(`${entry} (symlinked path blocked)`);
                      continue;
                    }
                    // SECURITY: normalize for lexical cleanup and re-run the
                    // deny-list against the cleaned path. Do not rely on
                    // nsIFile.normalize() to resolve symlinks or junctions: that
                    // is not its cross-platform contract (notably on Windows).
                    try {
                      file.normalize();
                    } catch {
                      failed.push(`${entry} (path normalization failed)`);
                      continue;
                    }
                    if (isSensitiveFilePath(file.path)) {
                      failed.push(`${entry} (sensitive path blocked)`);
                      continue;
                    }
                    let isRegularFile;
                    try {
                      isRegularFile = file.isFile();
                    } catch {
                      failed.push(`${entry} (file type check failed)`);
                      continue;
                    }
                    if (!isRegularFile) {
                      failed.push(`${entry} (not a regular file)`);
                      continue;
                    }
                    // Size cap mirrors the saved-attachment ceiling and avoids
                    // ballooning outgoing messages when a caller points at a huge file.
                    let fileSize;
                    try {
                      fileSize = file.fileSize;
                    } catch {
                      failed.push(`${entry} (file size check failed)`);
                      continue;
                    }
                    if (!Number.isSafeInteger(fileSize) || fileSize < 0) {
                      failed.push(`${entry} (invalid file size)`);
                      continue;
                    }
                    if (fileSize > MAX_FILE_PATH_ATTACHMENT_BYTES) {
                      failed.push(`${entry} (exceeds ${MAX_FILE_PATH_ATTACHMENT_BYTES / 1024 / 1024}MB size limit)`);
                      continue;
                    }
                    if (fileSize > MAX_TOTAL_ATTACHMENT_BYTES - totalAttachmentBytes) {
                      failed.push(`${entry} (exceeds ${MAX_TOTAL_ATTACHMENT_BYTES / 1024 / 1024}MB aggregate attachment limit)`);
                      continue;
                    }
                    // SECURITY: Parent-component symlinks/junctions and a TOCTOU
                    // window remain between these checks and Thunderbird's later
                    // MIME read. Fully closing those residual risks would require
                    // copying each file to a private temp directory before sending.
                    const desc = { url: Services.io.newFileURI(file).spec, name: file.leafName, size: fileSize };
                    descs.push(desc);
                    totalAttachmentBytes += fileSize;
                  } else if (entry && typeof entry === "object" && (entry.base64 || entry.content) && entry.name) {
                    // Inline base64 attachment — decode and write to temp file
                    const b64Data = entry.base64 || entry.content;
                    if (!isValidBase64(b64Data)) {
                      failed.push(`${entry.name} (invalid base64 data)`);
                      continue;
                    }
                    if (b64Data.length > MAX_BASE64_SIZE) {
                      failed.push(`${entry.name} (exceeds ${MAX_BASE64_SIZE / 1024 / 1024}MB size limit)`);
                      continue;
                    }
                    // Decode base64 to binary bytes
                    let bytes;
                    try {
                      // Use the global atob when available, otherwise fall back
                      const raw = typeof atob === "function" ? atob(b64Data) : ChromeUtils.base64Decode(b64Data);
                      bytes = new Uint8Array(raw.length);
                      for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
                    } catch {
                      // Fallback: manual base64 decode (atob may not be available in XPCOM context)
                      try {
                        const lookup = new Uint8Array(256);
                        const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
                        for (let i = 0; i < chars.length; i++) lookup[chars.charCodeAt(i)] = i;
                        // Shape was validated above; remove only legal trailing
                        // padding rather than stripping arbitrary invalid bytes.
                        const clean = b64Data.replace(/=+$/, "");
                        const len = clean.length;
                        const outLen = (len * 3) >> 2;
                        bytes = new Uint8Array(outLen);
                        let p = 0;
                        for (let i = 0; i < len; i += 4) {
                          const a = lookup[clean.charCodeAt(i)];
                          const b = lookup[clean.charCodeAt(i + 1)];
                          const c = lookup[clean.charCodeAt(i + 2)];
                          const d = lookup[clean.charCodeAt(i + 3)];
                          bytes[p++] = (a << 2) | (b >> 4);
                          if (i + 2 < len) bytes[p++] = ((b & 15) << 4) | (c >> 2);
                          if (i + 3 < len) bytes[p++] = ((c & 3) << 6) | d;
                        }
                        bytes = bytes.subarray(0, p);
                      } catch {
                        failed.push(`${entry.name} (invalid base64 data)`);
                        continue;
                      }
                    }
                    if (bytes.length > MAX_TOTAL_ATTACHMENT_BYTES - totalAttachmentBytes) {
                      failed.push(`${entry.name} (exceeds ${MAX_TOTAL_ATTACHMENT_BYTES / 1024 / 1024}MB aggregate attachment limit)`);
                      continue;
                    }
                    const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
                    tmpDir.append("thunderbird-mcp");
                    tmpDir.append("attachments");
                    if (!tmpDir.exists()) {
                      tmpDir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
                    }
                    const tmpFile = tmpDir.clone();
                    let safeName = (entry.name || entry.filename || "attachment").replace(/[^a-zA-Z0-9._-]/g, "_");
                    if (!safeName || safeName === "." || safeName === "..") safeName = "attachment";
                    tmpFile.append(`${Date.now()}_${++_tempFileCounter}_${safeName}`);
                    // Write via XPCOM binary stream
                    const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                      .createInstance(Ci.nsIFileOutputStream);
                    ostream.init(tmpFile, 0x02 | 0x08 | 0x20, 0o600, 0);
                    const bstream = Cc["@mozilla.org/binaryoutputstream;1"]
                      .createInstance(Ci.nsIBinaryOutputStream);
                    bstream.setOutputStream(ostream);
                    bstream.writeByteArray(bytes, bytes.length);
                    bstream.close();
                    ostream.close();
                    _tempAttachFiles.add(tmpFile.path);
                    const desc = { url: Services.io.newFileURI(tmpFile).spec, name: entry.name || entry.filename, size: tmpFile.fileSize };
                    if (entry.contentType) desc.contentType = entry.contentType;
                    descs.push(desc);
                    totalAttachmentBytes += bytes.length;
                  } else {
                    failed.push(typeof entry === "object" ? JSON.stringify(entry) : String(entry));
                  }
                } catch (e) {
                  failed.push(typeof entry === "object" ? (entry.name || JSON.stringify(entry)) : String(entry));
                }
              }
              return { descs, failed };
            }
            // END OUTBOUND ATTACHMENT CONVERSION

            /**
             * Converts attachment descriptors to nsIMsgAttachment objects.
             * Shared by the new-compose path (composeFields.addAttachment),
             * the reply/forward observer path (addAttachmentsToComposeWindow),
             * and sendMessageDirectly (headless send).
             */
            function descsToMsgAttachments(attachDescs) {
              const result = [];
              for (const desc of attachDescs) {
                try {
                  const att = Cc["@mozilla.org/messengercompose/attachment;1"]
                    .createInstance(Ci.nsIMsgAttachment);
                  att.url = desc.url;
                  att.name = desc.name;
                  if (desc.size != null) att.size = desc.size;
                  if (desc.contentType) att.contentType = desc.contentType;
                  result.push(att);
                } catch (e) {
                  console.warn("thunderbird-mcp: failed to convert attachment descriptor:", desc?.name || desc?.url || desc, e);
                }
              }
              return result;
            }

            function addAttachmentsToComposeWindow(composeWin, attachDescs) {
              if (!composeWin) {
                console.warn("thunderbird-mcp: skipping attachment add — no compose window");
                return;
              }
              if (typeof composeWin.AddAttachments !== "function") {
                console.warn("thunderbird-mcp: skipping attachment add — composeWin.AddAttachments not a function");
                return;
              }
              const attachList = descsToMsgAttachments(attachDescs);
              if (attachList.length > 0) {
                // Caller's try/catch (or fire-and-forget caller) is responsible for
                // surfacing failures — do not swallow here.
                composeWin.AddAttachments(attachList);
              }
            }

            function splitAddressHeader(header) {
              return (header || "").match(/(?:[^,"]|"[^"]*")+/g) || [];
            }

            function extractAddressEmail(address) {
              return (address.match(/<([^>]+)>/)?.[1] || address.trim()).toLowerCase();
            }

            function mergeAddressHeaders(...headers) {
              const seen = new Set();
              const merged = [];
              for (const header of headers) {
                for (const raw of splitAddressHeader(header)) {
                  const address = raw.trim();
                  if (!address) continue;
                  const email = extractAddressEmail(address);
                  if (seen.has(email)) continue;
                  seen.add(email);
                  merged.push(address);
                }
              }
              return merged.join(", ");
            }

            function getReplyAllCcRecipients(msgHdr, folder) {
              const ownAccount = MailServices.accounts.findAccountForServer(folder.server);
              const ownEmails = new Set();
              if (ownAccount) {
                for (const identity of ownAccount.identities) {
                  if (identity.email) ownEmails.add(identity.email.toLowerCase());
                }
              }

              const allRecipients = [
                ...splitAddressHeader(msgHdr.recipients),
                ...splitAddressHeader(msgHdr.ccList)
              ]
                .map(r => r.trim())
                .filter(r => r && (ownEmails.size === 0 || !ownEmails.has(extractAddressEmail(r))));

              const seen = new Set();
              const uniqueRecipients = allRecipients.filter(r => {
                const email = extractAddressEmail(r);
                if (seen.has(email)) return false;
                seen.add(email);
                return true;
              });

              return uniqueRecipients.join(", ");
            }

            function getIdentityAutoRecipientHeader(identity, kind) {
              if (!identity) return "";
              try {
                if (kind === "cc") {
                  return identity.doCc ? (identity.doCcList || "") : "";
                }
                if (kind === "bcc") {
                  return identity.doBcc ? (identity.doBccList || "") : "";
                }
              } catch {}
              return "";
            }

            function applyComposeRecipientOverrides(composeWin, identity, to, cc, bcc) {
              if (!composeWin) return;
              const overrides = { identityKey: null };
              if (to) overrides.to = to;
              if (cc) overrides.cc = mergeAddressHeaders(getIdentityAutoRecipientHeader(identity, "cc"), cc);
              if (bcc) overrides.bcc = mergeAddressHeaders(getIdentityAutoRecipientHeader(identity, "bcc"), bcc);
              if (Object.keys(overrides).length === 1) return;

              if (typeof composeWin.SetComposeDetails === "function") {
                composeWin.SetComposeDetails(overrides);
                return;
              }

              const fields = composeWin.gMsgCompose?.compFields;
              if (!fields) return;
              if (Object.prototype.hasOwnProperty.call(overrides, "to")) fields.to = overrides.to;
              if (Object.prototype.hasOwnProperty.call(overrides, "cc")) fields.cc = overrides.cc;
              if (Object.prototype.hasOwnProperty.call(overrides, "bcc")) fields.bcc = overrides.bcc;
              if (typeof composeWin.CompFields2Recipients === "function") {
                composeWin.CompFields2Recipients(fields);
              }
            }

            function formatBodyFragmentHtml(body, isHtml) {
              const formatted = formatBodyHtml(body, isHtml);
              if (!isHtml) return formatted;
              if (!formatted) return "";

              const needsParsing = /<(?:html|body|head)\b/i.test(formatted) || /\bmoz-signature\b/i.test(formatted);
              if (!needsParsing) return formatted;

              try {
                const doc = new DOMParser().parseFromString(formatted, "text/html");
                for (const node of doc.querySelectorAll("div.moz-signature, pre.moz-signature")) {
                  node.remove();
                }
                return doc.body ? doc.body.innerHTML : formatted;
              } catch {
                return formatted;
              }
            }

            function moveComposeSelectionToBodyStartIfRange(composeWin) {
              const browser = typeof composeWin?.getBrowser === "function" ? composeWin.getBrowser() : null;
              const editorDoc = browser?.contentDocument;
              const root = editorDoc?.body;
              const selection = typeof editorDoc?.getSelection === "function" ? editorDoc.getSelection() : null;
              let shouldMove = false;
              let moveError = null;

              if (selection) {
                if (selection.isCollapsed) return true;
                shouldMove = true;
              }

              if (shouldMove && root && typeof editorDoc.createRange === "function") {
                try {
                  const range = editorDoc.createRange();
                  range.setStart(root, 0);
                  range.collapse(true);
                  selection.removeAllRanges();
                  selection.addRange(range);
                  return true;
                } catch (e) {
                  moveError = e;
                }
              }

              const editor = typeof composeWin?.GetCurrentEditor === "function" ? composeWin.GetCurrentEditor() : null;
              let editorSelection = null;
              if (!selection && editor) {
                try {
                  editorSelection = editor.selection || null;
                } catch (e) {
                  moveError = e;
                }
              }

              if (!selection && editorSelection) {
                try {
                  if (editorSelection.isCollapsed) return true;
                  shouldMove = true;
                } catch (e) {
                  moveError = e;
                }
              }

              if (!selection && !editorSelection) {
                // Legacy editor-only path does not expose a reliable DOM
                // selection, so anchor defensively before insertHTML.
                shouldMove = true;
              }

              if (shouldMove && editor && typeof editor.beginningOfDocument === "function") {
                try {
                  editor.beginningOfDocument();
                  return true;
                } catch (e) {
                  moveError = e;
                }
              }

              if (shouldMove) {
                console.warn("thunderbird-mcp: could not anchor compose body insertion; insertHTML may replace selected quote", moveError);
              }
              return !shouldMove;
            }

            function insertReplyBodyIntoComposeWindow(composeWin, body, isHtml) {
              if (!composeWin || !body) return;
              const fragment = formatBodyFragmentHtml(body, isHtml);
              if (!fragment) return;

              const browser = typeof composeWin.getBrowser === "function" ? composeWin.getBrowser() : null;
              const editorDoc = browser?.contentDocument;
              if (editorDoc && typeof editorDoc.execCommand === "function") {
                // Body-ready can leave the original quote selected. insertHTML
                // replaces active ranges, so anchor only when TB selected text.
                moveComposeSelectionToBodyStartIfRange(composeWin);
                editorDoc.execCommand("insertHTML", false, fragment);
              } else {
                const editor = typeof composeWin.GetCurrentEditor === "function" ? composeWin.GetCurrentEditor() : null;
                if (editor && typeof editor.insertHTML === "function") {
                  moveComposeSelectionToBodyStartIfRange(composeWin);
                  editor.insertHTML(fragment);
                }
              }

              if (composeWin.gMsgCompose) {
                composeWin.gMsgCompose.bodyModified = true;
              }
              if ("gContentChanged" in composeWin) {
                composeWin.gContentChanged = true;
              }
            }

            function shouldUseDirectComposeOpen(compType) {
              return compType === Ci.nsIMsgCompType.ForwardInline;
            }

            function openComposeWindowWithCustomizations(msgComposeParams, originalMsgURI, compType, identity, body, isHtml, to, cc, bcc, attachDescs) {
              return new Promise((resolve) => {
                const OPEN_TIMEOUT_MS = shouldUseDirectComposeOpen(compType) ? 60000 : 15000;
                let settled = false;
                let matchedWindow = null;
                let pendingStateListener = null;
                let pendingStateCompose = null;

                const finish = (result) => {
                  if (settled) return;
                  settled = true;
                  try { Services.ww.unregisterNotification(windowObserver); } catch {}
                  try { timeout.cancel(); } catch {}
                  // Unregister any dangling state listener so a late
                  // NotifyComposeBodyReady cannot mutate the compose window
                  // after we have already resolved (e.g. after a timeout).
                  if (pendingStateListener && pendingStateCompose) {
                    try { pendingStateCompose.UnregisterStateListener(pendingStateListener); } catch {}
                  }
                  pendingStateListener = null;
                  pendingStateCompose = null;
                  resolve(result);
                };

                const timeout = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                timeout.initWithCallback({
                  notify() {
                    finish({ error: "Timed out waiting for compose window" });
                  }
                }, OPEN_TIMEOUT_MS, Ci.nsITimer.TYPE_ONE_SHOT);

                const maybeCustomizeWindow = (composeWin) => {
                  try {
                    if (!composeWin || composeWin === matchedWindow) return;
                    if (composeWin.document?.documentElement?.getAttribute("windowtype") !== "msgcompose") return;
                    if (!composeWin.gMsgCompose) return;
                    if (composeWin.gMsgCompose.originalMsgURI !== originalMsgURI) return;
                    if (composeWin.gComposeType !== compType) return;
                    // When two callers reply to the same message concurrently,
                    // both observers see both compose windows. Skip any window
                    // that has already been claimed by a prior observer so each
                    // call binds to exactly one compose window.
                    if (_claimedComposeWindows.has(composeWin)) return;
                    _claimedComposeWindows.add(composeWin);

                    matchedWindow = composeWin;
                    try { Services.ww.unregisterNotification(windowObserver); } catch {}

                    const stateListener = {
                      QueryInterface: ChromeUtils.generateQI(["nsIMsgComposeStateListener"]),
                      NotifyComposeFieldsReady() {},
                      ComposeProcessDone() {},
                      SaveInFolderDone() {},
                      NotifyComposeBodyReady() {
                        // Guard against a late body-ready firing after the
                        // caller already timed out -- don't mutate the compose
                        // window once the promise is settled.
                        if (settled) {
                          try { composeWin.gMsgCompose.UnregisterStateListener(stateListener); } catch {}
                          return;
                        }

                        try {
                          composeWin.gMsgCompose.UnregisterStateListener(stateListener);
                        } catch {}
                        pendingStateListener = null;
                        pendingStateCompose = null;

                        try {
                          applyComposeRecipientOverrides(composeWin, identity, to, cc, bcc);
                          insertReplyBodyIntoComposeWindow(composeWin, body, isHtml);
                          addAttachmentsToComposeWindow(composeWin, attachDescs);
                          finish({ success: true });
                        } catch (e) {
                          finish({ error: e.toString() });
                        }
                      },
                    };

                    pendingStateListener = stateListener;
                    pendingStateCompose = composeWin.gMsgCompose;
                    composeWin.gMsgCompose.RegisterStateListener(stateListener);
                  } catch (e) {
                    finish({ error: e.toString() });
                  }
                };

                const windowObserver = {
                  observe(subject, topic) {
                    if (topic !== "domwindowopened") return;
                    const composeWin = subject;
                    if (!composeWin || typeof composeWin.addEventListener !== "function") return;

                    // Thunderbird dispatches a non-bubbling compose-window-init event
                    // from MsgComposeCommands.js after gMsgCompose is initialized and
                    // the built-in state listener is registered, but before editor
                    // creation begins. Capturing it on the window lets us register our
                    // own ComposeBodyReady listener for the specific reply window
                    // without relying on getMostRecentWindow("msgcompose").
                    composeWin.addEventListener("compose-window-init", () => {
                      maybeCustomizeWindow(composeWin);
                    }, { once: true, capture: true });
                  },
                };

                try {
                  Services.ww.registerNotification(windowObserver);
                  const msgComposeService = MailServices.compose || Cc["@mozilla.org/messengercompose;1"]
                    .getService(Ci.nsIMsgComposeService);
                  if (shouldUseDirectComposeOpen(compType)) {
                    // Native inline-forward body/attachment population only runs through OpenComposeWindow.
                    msgComposeService.OpenComposeWindow(
                      null,
                      msgComposeParams.origMsgHdr,
                      originalMsgURI,
                      compType,
                      msgComposeParams.format,
                      identity,
                      identity?.email || "",
                      null
                    );
                  } else {
                    msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);
                  }
                } catch (e) {
                  finish({ error: e.toString() });
                }
              });
            }

            function markMessageDispositionState(msgHdr, dispositionState) {
              try {
                const folder = msgHdr?.folder;
                if (!folder || dispositionState == null) return false;
                if (typeof folder.addMessageDispositionState === "function") {
                  folder.addMessageDispositionState(msgHdr, dispositionState);
                  return true;
                }
                if (typeof folder.AddMessageDispositionState === "function") {
                  folder.AddMessageDispositionState(msgHdr, dispositionState);
                  return true;
                }
              } catch {}
              return false;
            }

            /**
             * Sends a message directly via nsIMsgSend without opening a compose window.
             * Used by composeMail, replyToMessage, forwardMessage when skipReview=true.
             *
             * Handles two createAndSendMessage signatures:
             * - TB 102-127 (C++): 18 args, includes aAttachments + aPreloadedAttachments
             * - TB 128+   (JS):  16 args, attachments via composeFields only
             * Attachments are always added to composeFields (works in both).
             * We try the modern 16-arg call first; if TB throws
             * NS_ERROR_XPC_NOT_ENOUGH_ARGS, fall back to the legacy 18-arg call.
             */
            function sendMessageDirectly(composeFields, identity, attachDescs, originalMsgURI, compType, deliverMode, bodyType) {
              if (!identity) {
                return Promise.resolve({ error: "No identity available for direct send" });
              }

              const mode = deliverMode ?? Ci.nsIMsgCompDeliverMode.Now;
              const bodyMimeType = bodyType || "text/html";
              const SEND_TIMEOUT_MS = 120000; // 2 min safety timeout

              return new Promise((resolve) => {
                let settled = false;
                const settle = (result) => {
                  if (!settled) {
                    settled = true;
                    resolve(result);
                  }
                };

                // Safety timeout -- if neither listener callback nor error fires
                const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                timer.initWithCallback({
                  notify() { settle({ error: "Send timed out after " + (SEND_TIMEOUT_MS / 1000) + "s" }); }
                }, SEND_TIMEOUT_MS, Ci.nsITimer.TYPE_ONE_SHOT);

                try {
                  const msgSend = Cc["@mozilla.org/messengercompose/send;1"]
                    .createInstance(Ci.nsIMsgSend);

                  // Populate sender fields from identity (normally done by compose window)
                  if (identity.email) {
                    const name = identity.fullName || "";
                    composeFields.from = name
                      ? `"${name}" <${identity.email}>`
                      : identity.email;
                  }
                  if (identity.organization) {
                    composeFields.organization = identity.organization;
                  }

                  // Add attachments to composeFields (works in all TB versions)
                  for (const att of descsToMsgAttachments(attachDescs)) {
                    composeFields.addAttachment(att);
                  }

                  // Extract body -- createAndSendMessage takes it as a separate param
                  const body = composeFields.body || "";

                  // Resolve account key from identity
                  let accountKey = "";
                  try {
                    for (const account of MailServices.accounts.accounts) {
                      for (let i = 0; i < account.identities.length; i++) {
                        if (account.identities[i].key === identity.key) {
                          accountKey = account.key;
                          break;
                        }
                      }
                      if (accountKey) break;
                    }
                  } catch {}

                  // SaveAsDraft mode routes through _mimeDoFcc() which does not
                  // call onStopSending. Completion is signaled via the copy
                  // service's onStopCopy after the message lands in the Drafts
                  // folder, so the listener also QIs nsIMsgCopyServiceListener.
                  const listener = {
                    QueryInterface: ChromeUtils.generateQI(["nsIMsgSendListener", "nsIMsgCopyServiceListener"]),
                    // nsIMsgSendListener -- fires for SMTP send paths
                    onStartSending() {},
                    onProgress() {},
                    onSendProgress() {},
                    onStatus() {},
                    onStopSending(msgID, status) {
                      timer.cancel();
                      if (Components.isSuccessCode(status)) {
                        settle({ success: true, message: "Message sent" });
                      } else {
                        settle({ error: `Send failed (status: 0x${status.toString(16)})` });
                      }
                    },
                    onGetDraftFolderURI() {},
                    onSendNotPerformed(msgID, status) {
                      timer.cancel();
                      settle({ error: "Send was not performed" });
                    },
                    onTransportSecurityError(msgID, status, secInfo, location) {
                      timer.cancel();
                      settle({ error: `Transport security error${location ? ": " + location : ""}` });
                    },
                    // nsIMsgCopyServiceListener -- fires for SaveAsDraft / Sent-folder copy
                    onStartCopy() {},
                    setMessageKey() {},
                    onStopCopy(status) {
                      timer.cancel();
                      if (Components.isSuccessCode(status)) {
                        settle({ success: true, message: "Saved" });
                      } else {
                        settle({ error: `Save failed (status: 0x${status.toString(16)})` });
                      }
                    },
                  };

                  // Common args shared by both signatures (positions 1-10)
                  const commonArgs = [
                    null,                           // editor
                    identity,                       // identity
                    accountKey,                     // account key
                    composeFields,                  // fields
                    false,                          // isDigest
                    false,                          // dontDeliver
                    mode,                           // deliver mode
                    null,                           // msgToReplace
                    bodyMimeType,                   // body type
                    body,                           // body
                  ];

                  // Tail args shared by both (parentWindow..compType)
                  const tailArgs = [
                    null,                           // parent window
                    null,                           // progress
                    listener,                       // listener
                    "",                             // password
                    originalMsgURI || "",           // original msg URI
                    compType,                       // compose type
                  ];

                  // Try modern 16-arg signature first (TB 128+).
                  // On TB 102-127, XPCOM throws NS_ERROR_XPC_NOT_ENOUGH_ARGS
                  // (0x80570001), so we fall back to legacy 18-arg with null
                  // attachment params (attachments already on composeFields).
                  // Modern TB may return a Promise -- catch async rejections.
                  let sendResult;
                  try {
                    sendResult = msgSend.createAndSendMessage(...commonArgs, ...tailArgs);
                  } catch (e) {
                    const isArgError = (e && e.result === 0x80570001) ||
                      String(e).includes("Not enough arguments");
                    if (isArgError) {
                      sendResult = msgSend.createAndSendMessage(...commonArgs, null, null, ...tailArgs);
                    } else {
                      throw e;
                    }
                  }
                  // Modern TB (128+) returns a Promise from createAndSendMessage.
                  // Handle both fulfillment and rejection -- belt-and-suspenders
                  // with the listener (settle is idempotent). For SaveAsDraft on
                  // older TB without the copy listener, the Promise fulfillment
                  // can be the only completion signal we get.
                  if (sendResult && typeof sendResult.then === "function") {
                    sendResult.then(
                      () => {
                        timer.cancel();
                        settle({ success: true });
                      },
                      e => {
                        timer.cancel();
                        settle({ error: e.toString() });
                      }
                    );
                  }
                } catch (e) {
                  timer.cancel();
                  settle({ error: e.toString() });
                }
              });
            }

            function escapeHtml(s) {
              return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
            }

            function stripHtml(html) {
              if (!html) return "";
              let text = String(html);

              // Remove style/script blocks
              text = text.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, " ");
              text = text.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, " ");

              // Convert block-level tags to newlines before stripping
              text = text.replace(/<br\s*\/?>/gi, "\n");
              text = text.replace(/<\/(p|div|li|tr|h[1-6]|blockquote|pre)>/gi, "\n");
              text = text.replace(/<(p|div|li|tr|h[1-6]|blockquote|pre)\b[^>]*>/gi, "\n");

              // Strip remaining tags
              text = text.replace(/<[^>]+>/g, " ");

              // Decode entities in a single pass
              const NAMED_ENTITIES = {
                nbsp: " ", amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'",
                "#39": "'",
                mdash: "\u2014", ndash: "\u2013", hellip: "\u2026",
                lsquo: "\u2018", rsquo: "\u2019", ldquo: "\u201C", rdquo: "\u201D",
                bull: "\u2022", middot: "\u00B7", ensp: "\u2002", emsp: "\u2003",
                thinsp: "\u2009", zwnj: "\u200C", zwj: "\u200D",
                laquo: "\u00AB", raquo: "\u00BB",
                copy: "\u00A9", reg: "\u00AE", trade: "\u2122", deg: "\u00B0",
                plusmn: "\u00B1", times: "\u00D7", divide: "\u00F7",
                micro: "\u00B5", para: "\u00B6", sect: "\u00A7",
                euro: "\u20AC", pound: "\u00A3", yen: "\u00A5", cent: "\u00A2",
              };
              text = text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/gi, (match, entity) => {
                if (entity.startsWith("#x") || entity.startsWith("#X")) {
                  const cp = parseInt(entity.slice(2), 16);
                  if (!cp || cp > 0x10FFFF) return match;
                  try { return String.fromCodePoint(cp); } catch { return match; }
                }
                if (entity.startsWith("#")) {
                  const cp = parseInt(entity.slice(1), 10);
                  if (!cp || cp > 0x10FFFF) return match;
                  try { return String.fromCodePoint(cp); } catch { return match; }
                }
                return NAMED_ENTITIES[entity.toLowerCase()] || match;
              });

              // Normalize newlines/spaces
              text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
              text = text.replace(/\n{3,}/g, "\n\n");
              text = text.replace(/[ \t\f\v]+/g, " ");
              text = text.replace(/ *\n */g, "\n");
              text = text.trim();
              return text;
            }

            /**
             * Converts HTML to markdown using DOMParser for structure-preserving
             * body extraction. Handles headings, links, bold/italic, lists,
             * blockquotes, code blocks, images, and horizontal rules. Email
             * tables (usually layout, not data) are flattened to text.
             * Falls back to stripHtml if DOMParser is unavailable.
             */
            function htmlToMarkdown(html) {
              if (!html) return "";
              try {
                const doc = new DOMParser().parseFromString(html, "text/html");

                function walkChildren(node) {
                  return Array.from(node.childNodes).map(walk).join("");
                }

                function walk(node) {
                  if (node.nodeType === 3) { // Text
                    return node.textContent.replace(/[ \t]+/g, " ");
                  }
                  if (node.nodeType !== 1) return "";
                  const tag = node.tagName.toLowerCase();
                  const inner = () => walkChildren(node);

                  switch (tag) {
                    case "script": case "style": case "head": return "";
                    case "br": return "\n";
                    case "hr": return "\n\n---\n\n";
                    case "p": case "div": case "section": case "article":
                      return "\n\n" + inner().trim() + "\n\n";
                    case "h1": return "\n\n# " + inner().trim() + "\n\n";
                    case "h2": return "\n\n## " + inner().trim() + "\n\n";
                    case "h3": return "\n\n### " + inner().trim() + "\n\n";
                    case "h4": return "\n\n#### " + inner().trim() + "\n\n";
                    case "h5": return "\n\n##### " + inner().trim() + "\n\n";
                    case "h6": return "\n\n###### " + inner().trim() + "\n\n";
                    case "strong": case "b": {
                      const t = inner().trim();
                      return t ? "**" + t + "**" : "";
                    }
                    case "em": case "i": {
                      const t = inner().trim();
                      return t ? "*" + t + "*" : "";
                    }
                    case "a": {
                      const href = node.getAttribute("href") || "";
                      const text = inner().trim();
                      // Skip empty/anchor-only links and mailto: without text
                      if (!text && !href) return "";
                      if (href && text && text !== href) return `[${text}](${href})`;
                      return text || href;
                    }
                    case "img": {
                      const alt = node.getAttribute("alt") || "";
                      const src = node.getAttribute("src") || "";
                      // Skip tracking pixels (1x1, tiny, or data: without alt)
                      const w = parseInt(node.getAttribute("width")) || 0;
                      const h = parseInt(node.getAttribute("height")) || 0;
                      if ((w > 0 && w <= 3) || (h > 0 && h <= 3)) return "";
                      if (src.startsWith("data:") && !alt) return "";
                      if (src) return `![${alt}](${src})`;
                      return alt;
                    }
                    case "code": return "`" + node.textContent + "`";
                    case "pre": return "\n\n```\n" + node.textContent.trim() + "\n```\n\n";
                    case "blockquote": {
                      const text = inner().trim();
                      return "\n\n" + text.split("\n").map(l => "> " + l).join("\n") + "\n\n";
                    }
                    case "ul": case "ol": return "\n" + inner() + "\n";
                    case "li": {
                      const parent = node.parentElement;
                      const isOl = parent && parent.tagName.toLowerCase() === "ol";
                      return (isOl ? "1. " : "- ") + inner().trim() + "\n";
                    }
                    // Tables: extract text with spacing (email tables are usually layout)
                    case "table": return "\n\n" + inner().trim() + "\n\n";
                    case "tr": return inner().trim() + "\n";
                    case "td": case "th": return inner().trim() + " ";
                    case "thead": case "tbody": case "tfoot": return inner();
                    default: return inner();
                  }
                }

                const body = doc.body || doc.documentElement;
                let result = walk(body);
                // Collapse excessive newlines, trim
                result = result.replace(/\n{3,}/g, "\n\n").trim();
                return result;
              } catch {
                // DOMParser unavailable or parse failure -- fall back to stripHtml
                return stripHtml(html);
              }
            }

            /**
             * Walks the MIME tree to find the raw body content.
             * Returns { text, isHtml } without any format conversion.
             * Does NOT use coerceBodyToPlaintext -- callers that want
             * the raw HTML (for markdown/html output) need this.
             */
            function extractBodyContent(aMimeMsg) {
              if (!aMimeMsg) return { text: "", isHtml: false };
              try {
                function findBody(part, isRoot = false) {
                  const ct = ((part.contentType || "").split(";")[0] || "").trim().toLowerCase();
                  if (ct === "message/rfc822" && !isRoot) return null;
                  if (ct !== "message/rfc822") {
                    if (ct === "text/plain" && part.body) return { text: part.body, isHtml: false };
                    if (ct === "text/html" && part.body) return { text: part.body, isHtml: true };
                  }
                  if (part.parts) {
                    let htmlFallback = null;
                    for (const sub of part.parts) {
                      const r = findBody(sub);
                      if (r && !r.isHtml) return r;
                      if (r && r.isHtml && !htmlFallback) htmlFallback = r;
                    }
                    if (htmlFallback) return htmlFallback;
                  }
                  return null;
                }
                const found = findBody(aMimeMsg, true);
                if (found) return found;
              } catch { /* give up */ }
              return { text: "", isHtml: false };
            }

            /**
             * Extracts plain text body from a MIME message.
             * Uses coerceBodyToPlaintext as fast path, then MIME tree fallback.
             * Used by reply/forward quoting where plain text is appropriate.
             */
            function extractPlainTextBody(aMimeMsg) {
              if (!aMimeMsg) return "";
              try {
                const text = aMimeMsg.coerceBodyToPlaintext();
                if (text) return text;
              } catch { /* fall through */ }
              const { text, isHtml } = extractBodyContent(aMimeMsg);
              return isHtml ? stripHtml(text) : text;
            }

            /**
             * Extracts body from a MIME message in the requested format.
             * For "text": uses coerceBodyToPlaintext fast path (original behavior).
             * For "markdown"/"html": walks MIME tree to find raw HTML content.
             */
            function extractFormattedBody(aMimeMsg, bodyFormat) {
              if (bodyFormat === "text") {
                return { body: extractPlainTextBody(aMimeMsg), bodyIsHtml: false };
              }
              // For markdown/html: need raw MIME content, not coerced text
              const { text, isHtml } = extractBodyContent(aMimeMsg);
              if (!text) {
                // MIME tree empty -- try coerce as last resort
                const fallback = extractPlainTextBody(aMimeMsg);
                return { body: fallback, bodyIsHtml: false };
              }
              if (!isHtml) return { body: text, bodyIsHtml: false };
              if (bodyFormat === "html") return { body: text, bodyIsHtml: true };
              // Default: markdown
              return { body: htmlToMarkdown(text), bodyIsHtml: false };
            }

            /**
             * Converts body text to HTML for compose fields.
             * Handles both HTML input (entity-encodes non-ASCII) and plain text.
             */
            function formatBodyHtml(body, isHtml) {
              if (isHtml) {
                let text = (body || "").replace(/\n/g, '');
                text = [...text].map(c => c.codePointAt(0) > 127 ? `&#${c.codePointAt(0)};` : c).join('');
                return text;
              }
              return escapeHtml(body || "").replace(/\n/g, '<br>');
            }

            /**
             * Decides whether a compose operation will (or should) run in HTML
             * mode, and returns the matching msgComposeParams.format value.
             *
             * The caller's explicit isHtml wins (true/false). When isHtml is
             * omitted, the identity's compose-format preference is consulted.
             *
             * ForwardInline is a special case: Thunderbird's compose service
             * only passes format through when it's Default or OppositeOfDefault
             * for that compType, so when the caller's explicit isHtml conflicts
             * with the identity pref on a forward, we have to ask for
             * OppositeOfDefault instead of HTML/PlainText.
             *
             * Identity must be resolved (setComposeIdentity) before calling this.
             */
            function resolveComposeFormat(identity, isHtml, compType) {
              const identityUsesHtml = identity?.composeHtml !== false;
              const useHtml = isHtml === true || (isHtml !== false && identityUsesHtml);
              const isForward = compType === Ci.nsIMsgCompType.ForwardInline
                             || compType === Ci.nsIMsgCompType.ForwardAsAttachment;

              let format;
              if (isHtml === undefined) {
                format = Ci.nsIMsgCompFormat.Default;
              } else if (isForward) {
                const explicitMatchesPref = (isHtml === true) === identityUsesHtml;
                format = explicitMatchesPref
                  ? Ci.nsIMsgCompFormat.Default
                  : Ci.nsIMsgCompFormat.OppositeOfDefault;
              } else {
                format = isHtml === true ? Ci.nsIMsgCompFormat.HTML : Ci.nsIMsgCompFormat.PlainText;
              }
              return { useHtml, format };
            }

            /**
             * Sets compose identity from `from` param or falls back to default.
             * Returns "" on success, or { error } if `from` was explicitly
             * provided but not found / restricted.  Fallback to default only
             * applies when `from` is omitted.
             */
            function setComposeIdentity(msgComposeParams, from, fallbackServer) {
              if (from) {
                // Explicit `from` -- must resolve or fail, never silently substitute
                const identity = findIdentity(from);
                if (identity) {
                  // findIdentity searches accessible accounts, so this is safe
                  msgComposeParams.identity = identity;
                  return "";
                }
                // Not found in accessible accounts -- check ALL accounts to
                // distinguish "restricted" from "genuinely unknown"
                if (findIdentityIn(MailServices.accounts.accounts, from)) {
                  return { error: `identity ${from} belongs to a restricted account` };
                }
                return { error: `unknown identity: ${from} -- no matching account configured in Thunderbird` };
              }
              // No explicit `from` -- fall back to contextual default
              if (fallbackServer) {
                const account = MailServices.accounts.findAccountForServer(fallbackServer);
                if (account && isAccountAllowed(account.key)) {
                  msgComposeParams.identity = account.defaultIdentity;
                }
              } else {
                const defaultAccount = MailServices.accounts.defaultAccount;
                if (defaultAccount && isAccountAllowed(defaultAccount.key)) {
                  msgComposeParams.identity = defaultAccount.defaultIdentity;
                }
              }
              // If no identity was set (all fallbacks restricted), explicitly set
              // the first accessible identity. Without this, Thunderbird's
              // OpenComposeWindowWithParams fills identity from defaultAccount
              // internally, bypassing account restrictions.
              if (!msgComposeParams.identity) {
                for (const account of getAccessibleAccounts()) {
                  if (account.defaultIdentity) {
                    msgComposeParams.identity = account.defaultIdentity;
                    break;
                  }
                }
                if (!msgComposeParams.identity) {
                  return { error: "No accessible identity found -- all accounts are restricted" };
                }
              }
              return "";
            }

	            /**
	             * Opens a folder and its message database.
	             * Best-effort refresh for IMAP folders (db may be stale).
	             * Returns { folder, db } or { error }.
	             */
	            function openFolder(folderPath) {
	              try {
	                const result = getAccessibleFolder(folderPath);
	                if (result.error) return result;
	                const folder = result.folder;

	                refreshImapFolderSync(folder);

	                const db = folder.msgDatabase;
	                if (!db) {
	                  return { error: "Could not access folder database" };
	                }

	                return { folder, db };
	              } catch (e) {
	                return { error: e.toString() };
	              }
	            }

	            /**
	             * Finds a single message header by messageId within a folderPath.
	             * Returns { msgHdr, folder, db } or { error }.
	             */
            function findTrashFolder(folder) {
              let account;
              try {
                account = MailServices.accounts.findAccountForServer(folder.server);
              } catch {
                return null;
              }
              const root = account?.incomingServer?.rootFolder;
              if (!root) return null;

              let fallback = null;
              const TRASH_NAMES = ["trash", "deleted items"];
              const stack = [root];
              while (stack.length > 0) {
                const current = stack.pop();
                try {
                  if (current && typeof current.getFlag === "function" && current.getFlag(Ci.nsMsgFolderFlags.Trash)) {
                    return current;
                  }
                } catch {}
                if (!fallback && current?.prettyName && TRASH_NAMES.includes(current.prettyName.toLowerCase())) {
                  fallback = current;
                }
                try {
                  if (current?.hasSubFolders) {
                    for (const sf of current.subFolders) stack.push(sf);
                  }
                } catch {}
              }
              return fallback;
            }

	            function findMessage(messageId, folderPath) {
	              const opened = openFolder(folderPath);
	              if (opened.error) return opened;

	              const { folder, db } = opened;
	              let msgHdr = null;

	              const hasDirectLookup = typeof db.getMsgHdrForMessageID === "function";
	              if (hasDirectLookup) {
	                try {
	                  msgHdr = db.getMsgHdrForMessageID(messageId);
	                } catch {
	                  msgHdr = null;
	                }
	              }

	              if (!msgHdr) {
	                for (const hdr of db.enumerateMessages()) {
	                  if (hdr.messageId === messageId) {
	                    msgHdr = hdr;
	                    break;
	                  }
	                }
	              }

	              if (!msgHdr) {
	                return { error: `Message not found: ${messageId}` };
	              }

	              return { msgHdr, folder, db };
	            }

            /**
             * Full-text body search using Thunderbird's Gloda index via
             * GlodaMsgSearcher. Searches subject, body, and attachment
             * names. Returns a Promise resolving to the same format as
             * searchMessages. IMAP accounts need offline sync for body
             * indexing; without it only headers are searched.
             */
            function glodaBodySearch(query, folderPath, startDate, endDate, maxResults, offset, sortOrder, unreadOnly, flaggedOnly, tag, countOnly, dedupByMessageId) {
              const requestedLimit = Number(maxResults);
              const effectiveLimit = Math.min(
                Number.isFinite(requestedLimit) && requestedLimit > 0 ? Math.floor(requestedLimit) : DEFAULT_MAX_RESULTS,
                MAX_SEARCH_RESULTS_CAP
              );
              const normalizedSortOrder = sortOrder === "asc" ? "asc" : "desc";
              const parsedStartDate = startDate ? new Date(startDate).getTime() : null;
              const parsedEndDate = endDate ? new Date(endDate).getTime() : null;
              if (parsedStartDate !== null && isNaN(parsedStartDate)) return { error: `Invalid startDate: ${startDate}` };
              if (parsedEndDate !== null && isNaN(parsedEndDate)) return { error: `Invalid endDate: ${endDate}` };
              // Match the regular search path: expand date-only endDate to end of day
              const isDateOnly = endDate && /^\d{4}-\d{2}-\d{2}$/.test(endDate.trim());
              const endDateOffset = isDateOnly ? 86400000 : 0;
              const endDateTs = parsedEndDate !== null ? (parsedEndDate + endDateOffset) * 1000 : null;
              const startDateTs = parsedStartDate !== null ? parsedStartDate * 1000 : null;

              // Resolve folder filter upfront -- match by URI prefix for subfolder inclusion
              let folderFilterURI = null;
              if (folderPath) {
                const result = getAccessibleFolder(folderPath);
                if (result.error) return result;
                folderFilterURI = result.folder.URI;
              }

              return new Promise((resolve) => {
                try {
                  const listener = {
                    onItemsAdded() {},
                    onItemsModified() {},
                    onItemsRemoved() {},
                    onQueryCompleted(collection) {
                      try {
                        const results = [];
                        for (const glodaMsg of collection.items) {
                          if (results.length >= SEARCH_COLLECTION_CAP) break;
                          // Get the underlying msgHdr
                          let msgHdr;
                          try {
                            msgHdr = glodaMsg.folderMessage;
                          } catch { continue; }
                          if (!msgHdr) continue;

                          // Account access control
                          const folder = msgHdr.folder;
                          if (!folder) continue;
                          if (!isFolderAccessible(folder)) continue;

                          // Folder filter (URI prefix match includes subfolders)
                          if (folderFilterURI && !folder.URI.startsWith(folderFilterURI)) continue;

                          // Date filters (timestamps in microseconds)
                          const msgDateTs = msgHdr.date || 0;
                          if (startDateTs !== null && msgDateTs < startDateTs) continue;
                          if (endDateTs !== null && msgDateTs > endDateTs) continue;

                          // Boolean filters
                          if (unreadOnly && msgHdr.isRead) continue;
                          if (flaggedOnly && !msgHdr.isFlagged) continue;
                          if (tag) {
                            const keywords = (msgHdr.getStringProperty("keywords") || "").split(/\s+/);
                            if (!keywords.includes(tag)) continue;
                          }

                          const msgTags = getUserTags(msgHdr);
                          const preview = msgHdr.getStringProperty("preview") || "";
                          const result = {
                            id: msgHdr.messageId,
                            threadId: msgHdr.threadId,
                            subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                            author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                            recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                            ccList: msgHdr.ccList,
                            date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                            folder: folder.prettyName,
                            folderPath: folder.URI,
                            read: msgHdr.isRead,
                            flagged: msgHdr.isFlagged,
                            tags: msgTags,
                            _dateTs: msgDateTs
                          };
                          if (preview) result.preview = preview;
                          results.push(result);
                        }

                        const finalResults = dedupByMessageId !== false ? dedupeSearchMessageResults(results) : results;

                        if (countOnly) {
                          resolve({ count: finalResults.length });
                          return;
                        }
                        finalResults.sort((a, b) => normalizedSortOrder === "asc" ? a._dateTs - b._dateTs : b._dateTs - a._dateTs);
                        resolve(paginate(finalResults, offset, effectiveLimit));
                      } catch (e) {
                        resolve({ error: e.toString() });
                      }
                    }
                  };
                  const searcher = new GlodaMsgSearcher(listener, query);
                  searcher.getCollection();
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

	            function searchMessages(query, folderPath, startDate, endDate, maxResults, offset, sortOrder, unreadOnly, flaggedOnly, tag, includeSubfolders, countOnly, searchBody, dedupByMessageId) {
	              // Gloda full-body search path (async)
	              if (searchBody) {
	                if (!GlodaMsgSearcher) return { error: "Gloda full-text index is not available" };
	                if (!query) return { error: "searchBody requires a non-empty query" };
	                return glodaBodySearch(query, folderPath, startDate, endDate, maxResults, offset, sortOrder, unreadOnly, flaggedOnly, tag, countOnly, dedupByMessageId);
	              }
	              const results = [];
	              const lowerQuery = (query || "").toLowerCase();
	              const hasQuery = !!lowerQuery;
	              // Parse optional field-operator prefix and split into AND tokens.
	              // Supports: from:Name, subject:Text, to:Email, cc:Email
	              // Without an operator, every token must appear somewhere across all fields.
	              const OPERATOR_RE = /^(from|subject|to|cc):\s*/;
	              let fieldTarget = null;
	              let queryTokens = [];
	              if (hasQuery) {
	                const opMatch = lowerQuery.match(OPERATOR_RE);
	                if (opMatch) {
	                  const opMap = { from: 'author', subject: 'subject', to: 'recipients', cc: 'ccList' };
	                  fieldTarget = opMap[opMatch[1]];
	                  queryTokens = lowerQuery.slice(opMatch[0].length).trim().split(/\s+/).filter(Boolean);
	                } else {
	                  queryTokens = lowerQuery.split(/\s+/).filter(Boolean);
	                }
	              }
	              // Treat whitespace-only queries and bare field operators (e.g. "from:"
	              // with nothing after) as failed queries that match nothing, rather
	              // than silently matching every message. The documented way to match
	              // all messages is to pass an empty string, which keeps hasQuery=false.
	              const failedQuery = hasQuery && queryTokens.length === 0;
	              const parsedStartDate = startDate ? new Date(startDate).getTime() : NaN;
              const parsedEndDate = endDate ? new Date(endDate).getTime() : NaN;
              const startDateTs = Number.isFinite(parsedStartDate) ? parsedStartDate * 1000 : null;
              // Add 24h only for date-only strings (e.g. "2024-01-15") to include the full day.
              // Use regex to detect ISO date-only format rather than checking for "T" which
              // would match arbitrary strings like "Totally invalid".
              const isDateOnly = endDate && /^\d{4}-\d{2}-\d{2}$/.test(endDate.trim());
              const endDateOffset = isDateOnly ? 86400000 : 0;
              const endDateTs = Number.isFinite(parsedEndDate) ? (parsedEndDate + endDateOffset) * 1000 : null;
              const requestedLimit = Number(maxResults);
              const effectiveLimit = Math.min(
                Number.isFinite(requestedLimit) && requestedLimit > 0 ? Math.floor(requestedLimit) : DEFAULT_MAX_RESULTS,
                MAX_SEARCH_RESULTS_CAP
              );
              const normalizedSortOrder = sortOrder === "asc" ? "asc" : "desc";

              function searchFolder(folder) {
                if (results.length >= SEARCH_COLLECTION_CAP) return;

                try {
                  refreshImapFolderSync(folder);

                  const db = folder.msgDatabase;
                  if (!db) return;

                  for (const msgHdr of db.enumerateMessages()) {
                    if (results.length >= SEARCH_COLLECTION_CAP) break;

                    // Check cheap numeric/boolean filters before string work
                    const msgDateTs = msgHdr.date || 0;
                    if (startDateTs !== null && msgDateTs < startDateTs) continue;
                    if (endDateTs !== null && msgDateTs > endDateTs) continue;
                    if (unreadOnly && msgHdr.isRead) continue;
                    if (flaggedOnly && !msgHdr.isFlagged) continue;
                    if (tag) {
                      const keywords = (msgHdr.getStringProperty("keywords") || "").split(/\s+/);
                      if (!keywords.includes(tag)) continue;
                    }

                    // IMPORTANT: Use mime2Decoded* properties for searching.
                    // Raw headers contain MIME encoding like "=?UTF-8?Q?...?="
                    // which won't match plain text searches.
                    const preview = msgHdr.getStringProperty("preview") || "";
                    if (failedQuery) continue;
                    if (hasQuery) {
                      const subject = (msgHdr.mime2DecodedSubject || msgHdr.subject || "").toLowerCase();
                      const author = (msgHdr.mime2DecodedAuthor || msgHdr.author || "").toLowerCase();
                      const recipients = (msgHdr.mime2DecodedRecipients || msgHdr.recipients || "").toLowerCase();
                      const ccList = (msgHdr.ccList || "").toLowerCase();
                      // AND-of-tokens: every token must appear somewhere across the fields.
                      // If a field operator (from:, subject:, to:, cc:) was given,
                      // restrict matching to that specific field only.
                      const fieldValues = { subject, author, recipients, ccList };
                      const matches = fieldTarget
                        ? queryTokens.every(t => (fieldValues[fieldTarget] || "").includes(t))
                        : queryTokens.every(t =>
                            subject.includes(t) ||
                            author.includes(t) ||
                            recipients.includes(t) ||
                            ccList.includes(t) ||
                            preview.toLowerCase().includes(t)
                          );
                      if (!matches) continue;
                    }

                    const msgTags = getUserTags(msgHdr);
                    const result = {
                      id: msgHdr.messageId,
                      threadId: msgHdr.threadId, // folder-local, use with folderPath for grouping
                      subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                      author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                      recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                      ccList: msgHdr.ccList,
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                      folder: folder.prettyName,
                      folderPath: folder.URI,
                      read: msgHdr.isRead,
                      flagged: msgHdr.isFlagged,
                      tags: msgTags,
                      _dateTs: msgDateTs
                    };
                    if (preview) result.preview = preview;
                    results.push(result);
                  }
                } catch {
                  // Skip inaccessible folders
                }

                const recurse = includeSubfolders !== false; // default true
                if (recurse && folder.hasSubFolders) {
                  for (const subfolder of folder.subFolders) {
                    if (results.length >= SEARCH_COLLECTION_CAP) break;
                    searchFolder(subfolder);
                  }
                }
              }

              if (folderPath) {
                const result = getAccessibleFolder(folderPath);
                if (result.error) return result;
                searchFolder(result.folder);
              } else {
                for (const account of getAccessibleAccounts()) {
                  if (results.length >= SEARCH_COLLECTION_CAP) break;
                  searchFolder(account.incomingServer.rootFolder);
                }
              }

              const finalResults = dedupByMessageId !== false ? dedupeSearchMessageResults(results) : results;

              if (countOnly) {
                return { count: finalResults.length };
              }

              finalResults.sort((a, b) => normalizedSortOrder === "asc" ? a._dateTs - b._dateTs : b._dateTs - a._dateTs);

              return paginate(finalResults, offset, effectiveLimit);
            }

            function searchContacts(query, maxResults) {
              const results = [];
              const lowerQuery = (query || "").toLowerCase();
              const hasQuery = !!lowerQuery;
              let queryTokens = [];
              // Tokenize so CardDAV "Lastname, Firstname" names match natural-order searches.
              if (hasQuery) {
                queryTokens = lowerQuery.split(/[,\s]+/).filter(Boolean);
              }
              const failedQuery = hasQuery && queryTokens.length === 0;
              const requestedLimit = Number(maxResults);
              const limit = Number.isFinite(requestedLimit) && requestedLimit > 0
                ? Math.min(Math.floor(requestedLimit), MAX_SEARCH_RESULTS_CAP)
                : DEFAULT_MAX_RESULTS;
              let truncated = false;

              for (const book of MailServices.ab.directories) {
                for (const card of book.childCards) {
                  if (card.isMailList) continue;

                  const email = (card.primaryEmail || "").toLowerCase();
                  const displayName = (card.displayName || "").toLowerCase();
                  const firstName = (card.firstName || "").toLowerCase();
                  const lastName = (card.lastName || "").toLowerCase();
                  const fields = [email, displayName, firstName, lastName];

                  if (failedQuery) continue;
                  const matches = !hasQuery || queryTokens.every(token =>
                    fields.some(field => field.includes(token))
                  );
                  if (matches) {
                    results.push(formatContact(card, book));
                  }

                  if (results.length >= limit) { truncated = true; break; }
                }
                if (truncated) break;
              }

              if (truncated) {
                return { contacts: results, hasMore: true, message: `Results limited to ${limit}. Refine your query to see more.` };
              }
              return results;
            }

            /**
             * Find a contact card by UID across all address books.
             * Returns { card, book } or { error }.
             */
            function findContactByUID(contactId) {
              for (const book of MailServices.ab.directories) {
                for (const card of book.childCards) {
                  if (card.isMailList) continue;
                  if (card.UID === contactId) {
                    return { card, book };
                  }
                }
              }
              return { error: `Contact not found: ${contactId}` };
            }

            function getContact(contactId) {
              try {
                if (typeof contactId !== "string" || !contactId) {
                  return { error: "contactId must be a non-empty string" };
                }
                const found = findContactByUID(contactId);
                if (found.error) return found;
                return formatContact(found.card, found.book);
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function createContact(
              email,
              displayName,
              firstName,
              lastName,
              phones,
              addresses,
              organization,
              title,
              note,
              birthday,
              addressBookId
            ) {
              try {
                const fields = {
                  email,
                  displayName,
                  firstName,
                  lastName,
                  phones,
                  addresses,
                  organization,
                  title,
                  note,
                  birthday,
                };
                const validationError = validateContactFields(fields, true);
                if (validationError) return { error: validationError };

                // Find the target address book
                let targetBook = null;
                if (addressBookId) {
                  for (const book of MailServices.ab.directories) {
                    if (book.dirPrefId === addressBookId || book.UID === addressBookId || book.URI === addressBookId) {
                      targetBook = book;
                      break;
                    }
                  }
                  if (!targetBook) {
                    return { error: `Address book not found: ${addressBookId}` };
                  }
                } else {
                  // Use the first writable address book
                  for (const book of MailServices.ab.directories) {
                    if (!book.readOnly) {
                      targetBook = book;
                      break;
                    }
                  }
                  if (!targetBook) {
                    return { error: "No writable address book found" };
                  }
                }

                const card = Cc["@mozilla.org/addressbook/cardproperty;1"]
                  .createInstance(Ci.nsIAbCard);
                const applyError = applyContactFields(card, fields, VCardPropertyEntry);
                if (applyError) return applyError;
                if (shouldSynthesizePhoneDisplayName(fields) && !card.displayName) {
                  card.displayName = phones[0].number.trim();
                }

                const newCard = targetBook.addCard(card);
                return {
                  success: true,
                  id: newCard.UID,
                  email: newCard.primaryEmail,
                  displayName: newCard.displayName,
                  addressBook: targetBook.dirName,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function updateContact(
              contactId,
              email,
              displayName,
              firstName,
              lastName,
              phones,
              addresses,
              organization,
              title,
              note,
              birthday
            ) {
              try {
                if (typeof contactId !== "string" || !contactId) {
                  return { error: "contactId must be a non-empty string" };
                }
                const fields = {
                  email,
                  displayName,
                  firstName,
                  lastName,
                  phones,
                  addresses,
                  organization,
                  title,
                  note,
                  birthday,
                };
                const validationError = validateContactFields(fields);
                if (validationError) return { error: validationError };

                const found = findContactByUID(contactId);
                if (found.error) return found;
                const { card, book } = found;

                const applyError = applyContactFields(card, fields, VCardPropertyEntry);
                if (applyError) return applyError;

                book.modifyCard(card);
                return {
                  success: true,
                  id: card.UID,
                  email: card.primaryEmail,
                  displayName: card.displayName,
                  firstName: card.firstName,
                  lastName: card.lastName,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function deleteContact(contactId) {
              try {
                if (typeof contactId !== "string" || !contactId) {
                  return { error: "contactId must be a non-empty string" };
                }

                const found = findContactByUID(contactId);
                if (found.error) return found;
                const { card, book } = found;

                book.deleteCards([card]);
                return {
                  success: true,
                  message: `Contact "${card.displayName || card.primaryEmail}" deleted`,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function listCalendars() {
              if (!cal) {
                return { error: "Calendar not available" };
              }
              try {
                return cal.manager.getCalendars().map(c => ({
                  id: c.id,
                  name: c.name,
                  type: c.type,
                  readOnly: c.readOnly,
                  supportsEvents: c.getProperty("capabilities.events.supported") !== false,
                  supportsTasks: c.getProperty("capabilities.tasks.supported") !== false,
                }));
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function createEvent(title, startDate, endDate, location, description, calendarId, allDay, skipReview, status, showAs, categories, onlineMeeting) {
              if (!cal || !CalEvent) {
                return { error: "Calendar module not available" };
              }
              if (skipReview && isSkipReviewBlocked()) {
                return { error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review dialog instead." };
              }
              try {
                const win = Services.wm.getMostRecentWindow("mail:3pane");
                if (!win && !skipReview) {
                  return { error: "No Thunderbird window found" };
                }

                const startJs = new Date(startDate);
                if (isNaN(startJs.getTime())) {
                  return { error: `Invalid startDate: ${startDate}` };
                }

                let endJs = endDate ? new Date(endDate) : null;
                if (endDate && (!endJs || isNaN(endJs.getTime()))) {
                  return { error: `Invalid endDate: ${endDate}` };
                }

                if (endJs) {
                  if (allDay) {
                    const startDay = new Date(startJs.getFullYear(), startJs.getMonth(), startJs.getDate());
                    const endDay = new Date(endJs.getFullYear(), endJs.getMonth(), endJs.getDate());
                    if (endDay.getTime() < startDay.getTime()) {
                      return { error: "endDate must not be before startDate" };
                    }
                  } else if (endJs.getTime() <= startJs.getTime()) {
                    return { error: "endDate must be after startDate" };
                  }
                }

                const event = new CalEvent();
                event.title = title;

                if (allDay) {
                  const startDt = cal.createDateTime();
                  startDt.resetTo(startJs.getFullYear(), startJs.getMonth(), startJs.getDate(), 0, 0, 0, cal.dtz.floating);
                  startDt.isDate = true;
                  event.startDate = startDt;

                  const endDt = cal.createDateTime();
                  if (endJs) {
                    endDt.resetTo(endJs.getFullYear(), endJs.getMonth(), endJs.getDate(), 0, 0, 0, cal.dtz.floating);
                    endDt.isDate = true;
                    // iCal DTEND is exclusive — bump if same as start
                    if (endDt.compare(startDt) <= 0) {
                      const bumpedEnd = new Date(endJs.getFullYear(), endJs.getMonth(), endJs.getDate());
                      bumpedEnd.setDate(bumpedEnd.getDate() + 1);
                      endDt.resetTo(
                        bumpedEnd.getFullYear(),
                        bumpedEnd.getMonth(),
                        bumpedEnd.getDate(),
                        0,
                        0,
                        0,
                        cal.dtz.floating
                      );
                      endDt.isDate = true;
                    }
                  } else {
                    const defaultEnd = new Date(startJs.getTime());
                    defaultEnd.setDate(defaultEnd.getDate() + 1);
                    endDt.resetTo(
                      defaultEnd.getFullYear(),
                      defaultEnd.getMonth(),
                      defaultEnd.getDate(),
                      0,
                      0,
                      0,
                      cal.dtz.floating
                    );
                    endDt.isDate = true;
                  }
                  event.endDate = endDt;
                } else {
                  event.startDate = cal.dtz.jsDateToDateTime(startJs, cal.dtz.defaultTimezone);
                  if (endJs) {
                    event.endDate = cal.dtz.jsDateToDateTime(endJs, cal.dtz.defaultTimezone);
                  } else {
                    const defaultEnd = new Date(startJs.getTime() + 3600000);
                    event.endDate = cal.dtz.jsDateToDateTime(defaultEnd, cal.dtz.defaultTimezone);
                  }
                }

                if (location) event.setProperty("LOCATION", location);
                if (description) event.setProperty("DESCRIPTION", description);
                if (showAs !== undefined && showAs !== null && showAs !== "busy" && showAs !== "free") {
                  return { error: `Invalid showAs: "${showAs}". Expected "busy" or "free".` };
                }
                // STATUS: explicit param wins; otherwise derive from showAs so Thunderbird renders busy=solid, free=hatched
                const effectiveStatus = (status !== undefined && status !== null && status !== "")
                  ? status
                  : (showAs === "free" ? "tentative" : "confirmed");
                const normalizedStatus = normalizeEventStatus(effectiveStatus);
                if (!normalizedStatus) {
                  return { error: `Invalid status: "${effectiveStatus}". Expected tentative, confirmed, or cancelled.` };
                }
                event.setProperty("STATUS", normalizedStatus);
                event.setProperty("TRANSP", showAs === "free" ? "TRANSPARENT" : "OPAQUE");
                if (categories && categories.length > 0) event.setCategories(categories);
                if (onlineMeeting) event.setProperty("X-ONLINE-MEETING-PROVIDER", "TeamsForBusiness");

                // Find target calendar
                const calendars = cal.manager.getCalendars();
                let targetCalendar = null;
                if (calendarId) {
                  targetCalendar = calendars.find(c => c.id === calendarId);
                  if (!targetCalendar) {
                    return { error: `Calendar not found: ${calendarId}` };
                  }
                  if (targetCalendar.readOnly) {
                    return { error: `Calendar is read-only: ${targetCalendar.name}` };
                  }
                } else {
                  targetCalendar = calendars.find(c => !c.readOnly);
                  if (!targetCalendar) {
                    return { error: "No writable calendar found" };
                  }
                }

                event.calendar = targetCalendar;

                if (skipReview) {
                  await targetCalendar.addItem(event);
                  return { success: true, message: `Event "${title}" added to calendar "${targetCalendar.name}"` };
                }

                const args = {
                  calendarEvent: event,
                  calendar: targetCalendar,
                  mode: "new",
                  inTab: false,
                  onOk(item, calendar) {
                    calendar.addItem(item);
                  },
                };

                win.openDialog(
                  "chrome://calendar/content/calendar-event-dialog.xhtml",
                  "_blank",
                  "centerscreen,chrome,titlebar,toolbar,resizable",
                  args
                );

                return { success: true, message: `Event dialog opened for "${title}" on calendar "${targetCalendar.name}"` };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function getCalendarItems(calendar, rangeStart, rangeEnd) {
              const FILTER_EVENT = 1 << 3;
              if (typeof calendar.getItemsAsArray === "function") {
                return await calendar.getItemsAsArray(FILTER_EVENT, 0, rangeStart, rangeEnd);
              }
              // Fallback for older Thunderbird versions using ReadableStream
              const items = [];
              const stream = cal.iterate.streamValues(calendar.getItems(FILTER_EVENT, 0, rangeStart, rangeEnd));
              for await (const chunk of stream) {
                for (const i of chunk) items.push(i);
              }
              return items;
            }

            function calDateToISO(dt) {
              if (!dt) return null;
              try { return new Date(dt.nativeTime / 1000).toISOString(); }
              catch { return dt.icalString || null; }
            }

            // VEVENT STATUS values per iCal RFC 5545 § 3.8.1.11.
            const VEVENT_STATUS_MAP = {
              tentative: "TENTATIVE",
              confirmed: "CONFIRMED",
              cancelled: "CANCELLED",
              canceled: "CANCELLED",
            };
            function normalizeEventStatus(status) {
              if (status === undefined || status === null) return null;
              return VEVENT_STATUS_MAP[String(status).trim().toLowerCase()] || null;
            }

            function formatEvent(item, calendar) {
              const allDay = item.startDate ? item.startDate.isDate : false;
              // For all-day events, iCal DTEND is exclusive. Convert to inclusive
              // (last day of event) so the API is intuitive and round-trips correctly.
              let endDateISO = calDateToISO(item.endDate);
              if (allDay && item.endDate) {
                try {
                  const raw = new Date(item.endDate.nativeTime / 1000);
                  raw.setDate(raw.getDate() - 1);
                  endDateISO = raw.toISOString();
                } catch { /* keep raw value */ }
              }
              const result = {
                id: item.id,
                calendarId: calendar.id,
                calendarName: calendar.name,
                title: item.title || "",
                startDate: calDateToISO(item.startDate),
                endDate: endDateISO,
                location: item.getProperty("LOCATION") || "",
                description: item.getProperty("DESCRIPTION") || "",
                // VEVENT STATUS (tentative/confirmed/cancelled). Empty string
                // when the event has no explicit status (iCal spec treats this
                // as implicit -- Thunderbird renders it like confirmed).
                status: (item.getProperty("STATUS") || "").toLowerCase(),
                categories: item.getCategories(),
                onlineMeetingURL: item.getProperty("X-MICROSOFT-SKYPETEAMSMEETINGURL") || null,
                allDay,
                isRecurring: !!item.recurrenceInfo,
              };
              // Occurrences of recurring events share the parent's id.
              // Include recurrenceId so callers can distinguish them.
              if (item.recurrenceId) {
                result.recurrenceId = calDateToISO(item.recurrenceId);
              }
              return result;
            }

            function formatTask(item, calendar) {
              const completed = item.isCompleted || (item.percentComplete === 100);
              const priority = item.priority || 0; // 0=undefined, 1=high, 5=normal, 9=low per iCal
              return {
                id: item.id,
                calendarId: calendar.id,
                calendarName: calendar.name,
                title: item.title || "",
                dueDate: calDateToISO(item.dueDate),
                startDate: calDateToISO(item.entryDate),
                completedDate: calDateToISO(item.completedDate),
                completed,
                percentComplete: item.percentComplete || 0,
                priority,
                description: item.getProperty("DESCRIPTION") || "",
              };
            }

            async function updateTask(taskId, calendarId, title, dueDate, description, completed, percentComplete, priority) {
              if (!cal) return { error: "Calendar not available" };
              try {
                if (!taskId) return { error: "taskId is required" };
                if (!calendarId) return { error: "calendarId is required" };

                const calendar = cal.manager.getCalendars().find(c => c.id === calendarId);
                if (!calendar) return { error: `Calendar not found: ${calendarId}` };
                if (calendar.readOnly) return { error: `Calendar is read-only: ${calendar.name}` };
                if (calendar.getProperty("capabilities.tasks.supported") === false) {
                  return { error: `Calendar "${calendar.name}" does not support tasks. Use listCalendars to find one with supportsTasks=true.` };
                }

                // Try direct lookup first, then fall back to scanning all tasks
                let oldItem = null;
                if (typeof calendar.getItem === "function") {
                  try { oldItem = await calendar.getItem(taskId); } catch {}
                }
                if (!oldItem) {
                  const FILTER_TODO = 1 << 2;
                  const COMPLETED_YES = 1 << 0;
                  const COMPLETED_NO = 1 << 1;
                  let items;
                  if (typeof calendar.getItemsAsArray === "function") {
                    items = await calendar.getItemsAsArray(FILTER_TODO | COMPLETED_YES | COMPLETED_NO, 0, null, null);
                  } else {
                    items = [];
                    const stream = cal.iterate.streamValues(calendar.getItems(FILTER_TODO | COMPLETED_YES | COMPLETED_NO, 0, null, null));
                    for await (const chunk of stream) {
                      for (const i of chunk) items.push(i);
                    }
                  }
                  oldItem = items.find(i => i.id === taskId) || null;
                }
                if (!oldItem) return { error: `Task not found: ${taskId}` };

                const newItem = oldItem.clone();
                const changes = [];

                if (title !== undefined) { newItem.title = title; changes.push("title"); }
                if (description !== undefined) { newItem.descriptionHTML = descriptionToHTML(description); changes.push("description"); }
                if (priority !== undefined) {
                  if (priority !== null && (!Number.isInteger(priority) || priority < 0 || priority > 9)) {
                    return { error: "priority must be an integer between 0 and 9 (0=unset, 1=high, 5=normal, 9=low)" };
                  }
                  newItem.priority = priority ?? 0;
                  changes.push("priority");
                }

                if (dueDate !== undefined) {
                  // Explicit null or empty string clears the due date.
                  // Without this, `new Date(null).getTime() === 0` would
                  // silently write Unix epoch (1970-01-01) instead.
                  if (dueDate === null || dueDate === "") {
                    newItem.dueDate = null;
                  } else {
                    const js = new Date(dueDate);
                    if (isNaN(js.getTime())) return { error: `Invalid dueDate: ${dueDate}` };
                    if (/^\d{4}-\d{2}-\d{2}$/.test(dueDate.trim())) {
                      const dt = cal.createDateTime();
                      dt.resetTo(js.getFullYear(), js.getMonth(), js.getDate(), 0, 0, 0, cal.dtz.floating);
                      dt.isDate = true;
                      newItem.dueDate = dt;
                    } else {
                      newItem.dueDate = cal.dtz.jsDateToDateTime(js, cal.dtz.defaultTimezone);
                    }
                  }
                  changes.push("dueDate");
                }

                // 'completed' and 'percentComplete' both control completion state.
                // Reject ambiguous input rather than guessing precedence.
                if (completed !== undefined && percentComplete !== undefined) {
                  return { error: "Specify either 'completed' or 'percentComplete', not both" };
                }

                // Apply completion state keeping STATUS, PERCENT-COMPLETE, and
                // COMPLETED consistent per iCal RFC 5545 VTODO rules -- so
                // Thunderbird's UI and other consumers see a valid task state.
                function applyCompletionState(pct) {
                  const clamped = Math.min(100, Math.max(0, pct));
                  newItem.percentComplete = clamped;
                  if (clamped === 100) {
                    newItem.setProperty("STATUS", "COMPLETED");
                    newItem.completedDate = cal.dtz.jsDateToDateTime(new Date(), cal.dtz.defaultTimezone);
                  } else if (clamped === 0) {
                    newItem.setProperty("STATUS", "NEEDS-ACTION");
                    newItem.completedDate = null;
                  } else {
                    newItem.setProperty("STATUS", "IN-PROCESS");
                    newItem.completedDate = null;
                  }
                }

                if (percentComplete !== undefined) {
                  applyCompletionState(percentComplete);
                  changes.push("percentComplete");
                }

                if (completed !== undefined) {
                  applyCompletionState(completed ? 100 : 0);
                  changes.push("completed");
                }

                if (changes.length === 0) return { error: "No changes specified" };

                await calendar.modifyItem(newItem, oldItem);
                const result = { success: true, updated: changes, task: formatTask(newItem, calendar) };
                if (newItem.recurrenceInfo) {
                  result.warning = "This is a recurring task -- changes apply to the entire series.";
                }
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function listEvents(calendarId, startDate, endDate, maxResults) {
              if (!cal) {
                return { error: "Calendar not available" };
              }
              try {
                const calendars = cal.manager.getCalendars();
                let targets = calendars;
                if (calendarId) {
                  const found = calendars.find(c => c.id === calendarId);
                  if (!found) return { error: `Calendar not found: ${calendarId}` };
                  targets = [found];
                }

                const startJs = startDate ? new Date(startDate) : new Date();
                if (isNaN(startJs.getTime())) return { error: `Invalid startDate: ${startDate}` };
                const endJs = endDate ? new Date(endDate) : new Date(startJs.getTime() + 30 * 86400000);
                if (isNaN(endJs.getTime())) return { error: `Invalid endDate: ${endDate}` };

                const rangeStart = cal.dtz.jsDateToDateTime(startJs, cal.dtz.defaultTimezone);
                const rangeEnd = cal.dtz.jsDateToDateTime(endJs, cal.dtz.defaultTimezone);
                const limit = Math.min(Math.max(maxResults || 100, 1), 500);

                // Two-phase query to correctly handle recurring events.
                //
                // The FILTER_OCCURRENCES flag is intended to make the storage layer
                // expand recurring masters and return individual occurrences within the
                // date range. In practice this does not work for offline-backed calendars
                // (e.g. OWL/Exchange): recurring masters are stored with event_start set
                // to their original first occurrence date, which is typically outside the
                // queried range, so a date-range query never returns them and the manual
                // expansion at the call site never runs.
                //
                // Fix — Phase 1: fetch non-recurring events and modified-occurrence
                // exceptions within the date range (their event_start is inside the
                // range). Phase 2: fetch ALL recurring masters without a date filter,
                // then expand each with recurrenceInfo.getOccurrences() and keep only
                // occurrences that fall within the queried range.
                const FILTER_EVENT = 1 << 3;
                const results = [];
                for (const calendar of targets) {
                  // Phase 1: non-recurring events + modified-occurrence exceptions.
                  // Bounded by the date range so no expansion cap is needed here --
                  // applying one would risk starving Phase 2 on wide ranges.
                  const rangeItems = await getCalendarItems(calendar, rangeStart, rangeEnd);
                  for (const item of rangeItems) {
                    if (!item.recurrenceInfo) results.push(formatEvent(item, calendar));
                  }

                  // Phase 2: all recurring masters (no date filter) → manual expansion.
                  let allItems = [];
                  try {
                    if (typeof calendar.getItemsAsArray === "function") {
                      allItems = await calendar.getItemsAsArray(FILTER_EVENT, 0, null, null);
                    } else {
                      const stream = cal.iterate.streamValues(calendar.getItems(FILTER_EVENT, 0, null, null));
                      for await (const chunk of stream) {
                        for (const i of chunk) allItems.push(i);
                      }
                    }
                  } catch {
                    allItems = rangeItems;
                  }

                  // Cap recurring expansion to prevent a malformed daily-over-decades
                  // master from blowing up the result. Tracked independently of Phase 1
                  // so a busy date range cannot starve recurring expansion.
                  const OCCURRENCE_EXPANSION_CAP = limit * 10;
                  let expanded = 0;
                  for (const item of allItems) {
                    if (expanded >= OCCURRENCE_EXPANSION_CAP) break;
                    if (!item.recurrenceInfo) continue;
                    try {
                      const occurrences = item.recurrenceInfo.getOccurrences(rangeStart, rangeEnd, 0);
                      for (const occ of occurrences) {
                        if (expanded >= OCCURRENCE_EXPANSION_CAP) break;
                        results.push(formatEvent(occ, calendar));
                        expanded++;
                      }
                    } catch (e) {
                      // Don't push the master on failure -- its event_start is the original
                      // first-occurrence date and is almost certainly outside the queried
                      // range, which would pollute results with stale events.
                      console.warn("thunderbird-mcp: recurrence expansion failed for", item.id || item.title, e);
                    }
                  }
                }

                results.sort((a, b) => new Date(a.startDate) - new Date(b.startDate));
                return results.slice(0, limit);
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function listCategories() {
              try {
                return cal.category.fromPrefs().sort((a, b) => a.localeCompare(b));
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function listTasks(calendarId, completed, dueBefore, maxResults) {
              if (!cal) return { error: "Calendar not available" };
              try {
                const calendars = cal.manager.getCalendars();
                let targets = calendars.filter(c =>
                  c.getProperty("capabilities.tasks.supported") !== false
                );
                if (calendarId) {
                  const found = calendars.find(c => c.id === calendarId);
                  if (!found) return { error: `Calendar not found: ${calendarId}` };
                  if (found.getProperty("capabilities.tasks.supported") === false) {
                    return { error: `Calendar "${found.name}" does not support tasks` };
                  }
                  targets = [found];
                }

                let dueBeforeDt = null;
                if (dueBefore) {
                  const js = new Date(dueBefore);
                  if (isNaN(js.getTime())) return { error: `Invalid dueBefore: ${dueBefore}` };
                  dueBeforeDt = js;
                }

                const limit = Math.min(Math.max(maxResults || 100, 1), 500);
                // Thunderbird calICalendar filter bits:
                // TYPE_TODO = 1<<2, COMPLETED_YES = 1<<0, COMPLETED_NO = 1<<1
                const FILTER_TODO = 1 << 2;
                const COMPLETED_YES = 1 << 0;
                const COMPLETED_NO = 1 << 1;
                let itemFilter = FILTER_TODO;
                if (completed === true) {
                  itemFilter |= COMPLETED_YES;
                } else if (completed === false) {
                  itemFilter |= COMPLETED_NO;
                } else {
                  itemFilter |= COMPLETED_YES | COMPLETED_NO;
                }
                const TASK_COLLECTION_CAP = limit * 10;
                const results = [];

                for (const calendar of targets) {
                  let items;
                  try {
                    if (typeof calendar.getItemsAsArray === "function") {
                      items = await calendar.getItemsAsArray(itemFilter, 0, null, null);
                    } else {
                      items = [];
                      const stream = cal.iterate.streamValues(calendar.getItems(itemFilter, 0, null, null));
                      for await (const chunk of stream) {
                        for (const i of chunk) items.push(i);
                      }
                    }
                  } catch {
                    continue; // Skip calendars that fail to query
                  }

                  for (const item of items) {
                    if (results.length >= TASK_COLLECTION_CAP) break;
                    // Filter by due date -- exclude undated tasks when dueBefore is set
                    if (dueBeforeDt) {
                      if (!item.dueDate) continue;
                      try {
                        const due = new Date(item.dueDate.nativeTime / 1000);
                        if (due >= dueBeforeDt) continue;
                      } catch { /* include if we can't parse */ }
                    }
                    results.push(formatTask(item, calendar));
                  }
                }

                // Sort by dueDate (nulls last), then title
                results.sort((a, b) => {
                  if (a.dueDate && b.dueDate) return new Date(a.dueDate) - new Date(b.dueDate);
                  if (a.dueDate) return -1;
                  if (b.dueDate) return 1;
                  return (a.title || "").localeCompare(b.title || "");
                });
                return results.slice(0, limit);
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function updateEvent(eventId, calendarId, title, startDate, endDate, location, description, status, showAs, categories, onlineMeeting) {
              if (!cal) return { error: "Calendar not available" };
              try {
                if (!eventId) return { error: "eventId is required" };
                if (!calendarId) return { error: "calendarId is required" };

                const calendar = cal.manager.getCalendars().find(c => c.id === calendarId);
                if (!calendar) return { error: `Calendar not found: ${calendarId}` };
                if (calendar.readOnly) return { error: `Calendar is read-only: ${calendar.name}` };

                // Use getItem API if available, else scan
                let oldItem = null;
                if (typeof calendar.getItem === "function") {
                  try { oldItem = await calendar.getItem(eventId); } catch {}
                }
                if (!oldItem) {
                  // Fallback: scan all events
                  const all = await getCalendarItems(calendar, null, null);
                  oldItem = all.find(i => i.id === eventId) || null;
                }
                if (!oldItem) return { error: `Event not found: ${eventId}` };

                const newItem = oldItem.clone();
                const changes = [];

                if (title !== undefined) { newItem.title = title; changes.push("title"); }

                if (startDate !== undefined) {
                  const js = new Date(startDate);
                  if (isNaN(js.getTime())) return { error: `Invalid startDate: ${startDate}` };
                  if (newItem.startDate && newItem.startDate.isDate) {
                    const dt = cal.createDateTime();
                    dt.resetTo(js.getFullYear(), js.getMonth(), js.getDate(), 0, 0, 0, cal.dtz.floating);
                    dt.isDate = true;
                    newItem.startDate = dt;
                  } else {
                    newItem.startDate = cal.dtz.jsDateToDateTime(js, cal.dtz.defaultTimezone);
                  }
                  changes.push("startDate");
                }

                if (endDate !== undefined) {
                  const js = new Date(endDate);
                  if (isNaN(js.getTime())) return { error: `Invalid endDate: ${endDate}` };
                  if (newItem.endDate && newItem.endDate.isDate) {
                    const dt = cal.createDateTime();
                    // iCal DTEND is exclusive for all-day -- bump by 1 day
                    const next = new Date(js.getFullYear(), js.getMonth(), js.getDate());
                    next.setDate(next.getDate() + 1);
                    dt.resetTo(next.getFullYear(), next.getMonth(), next.getDate(), 0, 0, 0, cal.dtz.floating);
                    dt.isDate = true;
                    newItem.endDate = dt;
                  } else {
                    newItem.endDate = cal.dtz.jsDateToDateTime(js, cal.dtz.defaultTimezone);
                  }
                  changes.push("endDate");
                }

                if (location !== undefined) { newItem.setProperty("LOCATION", location); changes.push("location"); }
                if (description !== undefined) { newItem.setProperty("DESCRIPTION", description); changes.push("description"); }
                if (status !== undefined) {
                  if (status === null || status === "") {
                    newItem.deleteProperty("STATUS");
                  } else {
                    const normalized = normalizeEventStatus(status);
                    if (!normalized) {
                      return { error: `Invalid status: "${status}". Expected tentative, confirmed, or cancelled.` };
                    }
                    newItem.setProperty("STATUS", normalized);
                  }
                  changes.push("status");
                }
                if (showAs !== undefined) {
                  if (showAs === null || showAs === "") {
                    newItem.deleteProperty("TRANSP");
                  } else if (showAs === "free") {
                    newItem.setProperty("TRANSP", "TRANSPARENT");
                    // Also set STATUS:TENTATIVE for Thunderbird visual display unless caller overrides
                    if (status === undefined) { newItem.setProperty("STATUS", "TENTATIVE"); changes.push("status"); }
                  } else if (showAs === "busy") {
                    newItem.setProperty("TRANSP", "OPAQUE");
                    // Also set STATUS:CONFIRMED for Thunderbird visual display unless caller overrides
                    if (status === undefined) { newItem.setProperty("STATUS", "CONFIRMED"); changes.push("status"); }
                  } else {
                    return { error: `Invalid showAs: "${showAs}". Expected "busy" or "free".` };
                  }
                  changes.push("showAs");
                }
                // null/undefined preserves existing categories; empty array clears them.
                if (Array.isArray(categories)) {
                  newItem.setCategories(categories);
                  changes.push("categories");
                }
                if (onlineMeeting !== undefined) {
                  if (onlineMeeting) {
                    newItem.setProperty("X-ONLINE-MEETING-PROVIDER", "TeamsForBusiness");
                  } else {
                    newItem.deleteProperty("X-ONLINE-MEETING-PROVIDER");
                    newItem.deleteProperty("X-MICROSOFT-SKYPETEAMSMEETINGURL");
                  }
                  changes.push("onlineMeeting");
                }

                if (changes.length === 0) return { error: "No changes specified" };

                // Validate end > start after all changes
                if (newItem.startDate && newItem.endDate && newItem.endDate.compare(newItem.startDate) <= 0) {
                  return { error: "endDate must be after startDate" };
                }

                await calendar.modifyItem(newItem, oldItem);
                const result = { success: true, updated: changes };
                if (oldItem.recurrenceInfo) {
                  result.warning = "This is a recurring event -- changes apply to the entire series.";
                }
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function deleteEvent(eventId, calendarId) {
              if (!cal) return { error: "Calendar not available" };
              try {
                if (!eventId) return { error: "eventId is required" };
                if (!calendarId) return { error: "calendarId is required" };

                const calendar = cal.manager.getCalendars().find(c => c.id === calendarId);
                if (!calendar) return { error: `Calendar not found: ${calendarId}` };
                if (calendar.readOnly) return { error: `Calendar is read-only: ${calendar.name}` };

                let item = null;
                if (typeof calendar.getItem === "function") {
                  try { item = await calendar.getItem(eventId); } catch {}
                }
                if (!item) {
                  const all = await getCalendarItems(calendar, null, null);
                  item = all.find(i => i.id === eventId) || null;
                }
                if (!item) return { error: `Event not found: ${eventId}` };

                const isRecurring = !!item.recurrenceInfo;
                await calendar.deleteItem(item);
                const result = { success: true, deleted: eventId };
                if (isRecurring) {
                  result.warning = "This was a recurring event -- the entire series was deleted.";
                }
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function descriptionToHTML(text) {
              if (text == null || text === "") return "";
              // Strip null bytes so the sentinel below can't collide with user input.
              const input = String(text).replace(/\x00/g, "");
              const escapeText = s => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
              const escapeAttr = s => escapeText(s).replace(/"/g, "&quot;");
              const SAFE_HREF = /^(?:https?:|mailto:)/i;

              // Stash sanitized anchors so the global HTML-escape doesn't double-escape
              // them. Anchors with unsafe (e.g. javascript:, data:) or missing href are
              // dropped to plain text -- TB's task UI sanitizes on render but the raw
              // markup is persisted in ALTREP and may be consumed by other clients.
              const anchors = [];
              let processed = input.replace(/<a\b[^>]*>([\s\S]*?)<\/a>/gi, (whole, inner) => {
                const hrefMatch = whole.match(/href\s*=\s*(['"])([^'"]*)\1/i);
                const innerText = escapeText(inner.replace(/<[^>]+>/g, ""));
                if (!hrefMatch || !SAFE_HREF.test(hrefMatch[2].trim())) {
                  return innerText;
                }
                anchors.push(`<a href="${escapeAttr(hrefMatch[2].trim())}">${innerText}</a>`);
                return `\x00ANCHOR${anchors.length - 1}\x00`;
              });

              // Escape remaining plain text, then auto-link bare URLs.
              processed = escapeText(processed).replace(
                /https?:\/\/[^\s<>"]+/g,
                url => `<a href="${escapeAttr(url)}">${url}</a>`
              );

              // Restore sanitized anchors and convert newlines.
              processed = processed.replace(/\x00ANCHOR(\d+)\x00/g, (_, i) => anchors[i]);
              return `<html><body><div>${processed.replace(/\n/g, "<br>")}</div></body></html>`;
            }

            async function createTask(title, dueDate, calendarId, description, priority, categories, skipReview) {
              if (!cal || !CalTodo) return { error: "Calendar module not available" };
              if (skipReview && isSkipReviewBlocked()) {
                return { error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review dialog instead." };
              }
              try {
                let dueDt = null;
                if (dueDate) {
                  const js = new Date(dueDate);
                  if (isNaN(js.getTime())) return { error: `Invalid dueDate: ${dueDate}` };
                  // Date-only string (YYYY-MM-DD) means all-day
                  if (/^\d{4}-\d{2}-\d{2}$/.test(dueDate.trim())) {
                    dueDt = cal.createDateTime();
                    dueDt.resetTo(js.getFullYear(), js.getMonth(), js.getDate(), 0, 0, 0, cal.dtz.floating);
                    dueDt.isDate = true;
                  } else {
                    dueDt = cal.dtz.jsDateToDateTime(js, cal.dtz.defaultTimezone);
                  }
                }

                // Find target calendar (must support tasks)
                let targetCalendar = null;
                if (calendarId) {
                  targetCalendar = cal.manager.getCalendars().find(c => c.id === calendarId);
                  if (!targetCalendar) return { error: `Calendar not found: ${calendarId}` };
                  if (targetCalendar.readOnly) return { error: `Calendar is read-only: ${targetCalendar.name}` };
                  if (targetCalendar.getProperty("capabilities.tasks.supported") === false) {
                    return { error: `Calendar "${targetCalendar.name}" does not support tasks. Use listCalendars to find one with supportsTasks=true.` };
                  }
                }

                // Build a fully-populated CalTodo. The extension runs in addon_parent
                // (main-process privileged context) and imports CalTodo from the same
                // resource:/// ESModule singleton as the chrome, so the object is fully
                // interoperable with createTodoWithDialog and the task edit dialog.
                if (priority !== undefined && priority !== null) {
                  if (!Number.isInteger(priority) || priority < 0 || priority > 9) {
                    return { error: "priority must be an integer between 0 and 9 (0=unset, 1=high, 5=normal, 9=low)" };
                  }
                }

                const todo = new CalTodo();
                todo.title = title;
                if (dueDt) todo.dueDate = dueDt;
                if (description) todo.descriptionHTML = descriptionToHTML(description);
                if (priority !== undefined && priority !== null) todo.priority = priority;
                if (categories && categories.length > 0) todo.setCategories(categories);
                if (targetCalendar) todo.calendar = targetCalendar;

                if (skipReview) {
                  if (!targetCalendar) {
                    targetCalendar = cal.manager.getCalendars().find(
                      c => !c.readOnly && c.getProperty("capabilities.tasks.supported") !== false
                    );
                    if (!targetCalendar) return { error: "No writable task-capable calendar found" };
                    todo.calendar = targetCalendar;
                  }
                  await targetCalendar.addItem(todo);
                  return { success: true, message: `Task "${title}" created in calendar "${targetCalendar.name}"` };
                }

                const win = Services.wm.getMostRecentWindow("mail:3pane");
                if (!win) return { error: "No Thunderbird window found" };

                // Pass the pre-populated CalTodo to the dialog. createTodoWithDialog
                // clones it (clearing the id) before opening, so all fields are
                // pre-filled and the user can review or cancel without side effects.
                win.createTodoWithDialog(targetCalendar, dueDt, null, todo);

                return { success: true, message: `Task dialog opened for "${title}"` };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            // BEGIN RAW MIME PARSING HELPERS
            function rawMimeToByteString(input) {
              if (typeof input === "string") return input;
              if (input instanceof Uint8Array) {
                let out = "";
                const chunkSize = 0x8000;
                for (let i = 0; i < input.length; i += chunkSize) {
                  out += String.fromCharCode(...input.subarray(i, i + chunkSize));
                }
                return out;
              }
              return "";
            }

            function rawMimeBytesFromByteString(s) {
              const bytes = new Uint8Array(s.length);
              for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i) & 0xFF;
              return bytes;
            }

            function findRawMimeHeaderBodySplit(s) {
              const matches = [
                { idx: s.indexOf("\r\n\r\n"), len: 4 },
                { idx: s.indexOf("\n\n"), len: 2 },
                { idx: s.indexOf("\r\r"), len: 2 },
              ].filter(m => m.idx >= 0).sort((a, b) => a.idx - b.idx);
              if (matches.length === 0) return null;
              return {
                header: s.slice(0, matches[0].idx),
                body: s.slice(matches[0].idx + matches[0].len),
              };
            }

            function parseRawMimeHeaders(headerBlock) {
              const headers = Object.create(null);
              const unfolded = String(headerBlock || "").replace(/(?:\r\n|\r|\n)[ \t]+/g, " ");
              for (const line of unfolded.split(/\r\n|\r|\n/)) {
                const colonIdx = line.indexOf(":");
                if (colonIdx < 0) continue;
                const name = line.slice(0, colonIdx).trim().toLowerCase();
                const value = line.slice(colonIdx + 1).trim();
                if (!name) continue;
                if (!headers[name]) headers[name] = [];
                headers[name].push(value);
              }
              return headers;
            }

            function splitRawMimeHeaderParameters(value) {
              const parts = [];
              let current = "";
              let quoted = false;
              let escaped = false;
              for (let i = 0; i < value.length; i++) {
                const ch = value[i];
                if (escaped) {
                  current += ch;
                  escaped = false;
                  continue;
                }
                if (quoted && ch === "\\") {
                  current += ch;
                  escaped = true;
                  continue;
                }
                if (quoted) {
                  current += ch;
                  if (ch === "\"") quoted = false;
                  continue;
                }
                if (ch === "\"") {
                  current += ch;
                  quoted = true;
                  continue;
                }
                if (ch === ";") {
                  parts.push(current.trim());
                  current = "";
                  continue;
                }
                current += ch;
              }
              parts.push(current.trim());
              return parts;
            }

            function unquoteRawMimeParameter(value) {
              let v = String(value || "").trim();
              if (v.startsWith("\"") && v.endsWith("\"")) {
                v = v.slice(1, -1).replace(/\\(["'\\])/g, "$1");
              }
              return v;
            }

            function decodeRawMimePercentBytes(s) {
              const bytes = [];
              for (let i = 0; i < s.length; i++) {
                if (s[i] === "%" && i + 2 < s.length && /^[0-9A-Fa-f]{2}$/.test(s.slice(i + 1, i + 3))) {
                  bytes.push(parseInt(s.slice(i + 1, i + 3), 16));
                  i += 2;
                } else {
                  bytes.push(s.charCodeAt(i) & 0xFF);
                }
              }
              return new Uint8Array(bytes);
            }

            function decodeRawMimeExtendedParameter(value) {
              const raw = unquoteRawMimeParameter(value);
              const match = raw.match(/^([^']*)'[^']*'(.*)$/);
              if (!match) return raw;
              const charset = (match[1] || "utf-8").trim() || "utf-8";
              const encoded = match[2] || "";
              try {
                return new TextDecoder(charset, { fatal: false }).decode(decodeRawMimePercentBytes(encoded));
              } catch {
                try {
                  return new TextDecoder("utf-8", { fatal: false }).decode(decodeRawMimePercentBytes(encoded));
                } catch {
                  return raw;
                }
              }
            }

            function parseRawMimeHeaderValue(value) {
              const pieces = splitRawMimeHeaderParameters(String(value || ""));
              const main = (pieces.shift() || "").trim().toLowerCase();
              const params = Object.create(null);
              for (const piece of pieces) {
                const eqIdx = piece.indexOf("=");
                if (eqIdx < 0) continue;
                const key = piece.slice(0, eqIdx).trim().toLowerCase();
                const val = piece.slice(eqIdx + 1).trim();
                if (!key) continue;
                params[key] = key.endsWith("*")
                  ? decodeRawMimeExtendedParameter(val)
                  : unquoteRawMimeParameter(val);
              }
              return { value: main, params };
            }

            function getRawMimeHeader(headers, name) {
              return headers[name]?.[0] || "";
            }

            function getRawMimeFilename(contentDisposition, contentType) {
              return contentDisposition.params["filename*"] ||
                contentDisposition.params.filename ||
                contentType.params["name*"] ||
                contentType.params.name ||
                "";
            }

            function normalizeRawMimeContentId(value) {
              return String(value || "").trim().replace(/^<|>$/g, "");
            }

            function normalizeRawMimeContentIdForMatch(value) {
              return String(value || "")
                .trim()
                .replace(/^<+|>+$/g, "")
                .trim()
                .toLowerCase();
            }

            function decodeRawMimeBase64ToBytes(body, strict = false) {
              const raw = String(body || "");
              const clean = strict
                ? raw.replace(/\s/g, "")
                : raw.replace(/[^A-Za-z0-9+/=]/g, "");
              if (!clean) return new Uint8Array(0);
              if (typeof atob === "function") {
                const binary = atob(clean);
                return rawMimeBytesFromByteString(binary);
              }
              if (strict && /[^A-Za-z0-9+/=]/.test(clean)) {
                throw new Error("invalid base64 body");
              }
              const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
              const lookup = new Uint8Array(256);
              for (let i = 0; i < chars.length; i++) lookup[chars.charCodeAt(i)] = i;
              const out = [];
              for (let i = 0; i < clean.length; i += 4) {
                const a = clean[i];
                const b = clean[i + 1];
                const c = clean[i + 2];
                const d = clean[i + 3];
                if (!a || !b || a === "=" || b === "=") break;
                const av = lookup[a.charCodeAt(0)];
                const bv = lookup[b.charCodeAt(0)];
                out.push((av << 2) | (bv >> 4));
                if (c && c !== "=") {
                  const cv = lookup[c.charCodeAt(0)];
                  out.push(((bv & 15) << 4) | (cv >> 2));
                  if (d && d !== "=") {
                    const dv = lookup[d.charCodeAt(0)];
                    out.push(((cv & 3) << 6) | dv);
                  }
                }
              }
              return new Uint8Array(out);
            }

            function decodeRawMimeQuotedPrintableToBytes(body) {
              const qpBody = String(body || "").replace(/=(?:\r\n|\r|\n)/g, "");
              const decodedBytes = [];
              for (let i = 0; i < qpBody.length; i++) {
                if (qpBody[i] === "=" && i + 2 < qpBody.length) {
                  const hex = qpBody.slice(i + 1, i + 3);
                  if (/^[0-9A-Fa-f]{2}$/.test(hex)) {
                    decodedBytes.push(parseInt(hex, 16));
                    i += 2;
                    continue;
                  }
                }
                decodedBytes.push(qpBody.charCodeAt(i) & 0xFF);
              }
              return new Uint8Array(decodedBytes);
            }

            function decodeRawMimeTransferBody(body, transferEncoding, options = {}) {
              const cte = (transferEncoding || "7bit").split(";")[0].trim().toLowerCase() || "7bit";
              if (cte === "base64") {
                return decodeRawMimeBase64ToBytes(body, options.strictBase64 === true);
              }
              if (cte === "quoted-printable") return decodeRawMimeQuotedPrintableToBytes(body);
              if (cte === "7bit" || cte === "8bit" || cte === "binary") {
                return rawMimeBytesFromByteString(body || "");
              }
              return null;
            }

            function escapeRawMimeRegExp(s) {
              return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            }

            function splitRawMimeMultipartBody(body, boundary) {
              if (!boundary) return [];
              const markerRe = new RegExp(
                "(^|\\r\\n|\\n|\\r)--" + escapeRawMimeRegExp(boundary) +
                  "(--)?[ \\t]*(?:\\r\\n|\\n|\\r|$)",
                "g"
              );
              const parts = [];
              let partStart = null;
              let match;
              while ((match = markerRe.exec(body)) !== null) {
                if (partStart !== null) parts.push(body.slice(partStart, match.index));
                if (match[2]) {
                  partStart = null;
                  break;
                }
                partStart = markerRe.lastIndex;
                if (match[0].length === 0) markerRe.lastIndex++;
              }
              // A missing terminal boundary is common in damaged/imported mbox
              // messages. Preserve the final open part instead of rejecting it.
              if (partStart !== null && partStart < body.length) {
                parts.push(body.slice(partStart));
              }
              return parts;
            }

            function parseRawMimeEntity(rawBytes, options = {}) {
              const maxDepth = Number.isInteger(options.maxDepth) && options.maxDepth >= 0
                ? options.maxDepth
                : 32;
              const raw = rawMimeToByteString(rawBytes);

              function parseEntity(partRaw, depth, partName) {
                const split = findRawMimeHeaderBodySplit(partRaw);
                if (!split) return null;
                const headers = parseRawMimeHeaders(split.header);
                const contentType = parseRawMimeHeaderValue(
                  getRawMimeHeader(headers, "content-type") || "text/plain"
                );
                const contentDisposition = parseRawMimeHeaderValue(
                  getRawMimeHeader(headers, "content-disposition") || ""
                );
                const entity = {
                  headers,
                  contentType,
                  contentDisposition,
                  body: split.body,
                  parts: [],
                  partName,
                  depthLimitReached: false,
                };

                if (contentType.value.startsWith("multipart/")) {
                  entity.body = "";
                  if (depth >= maxDepth) {
                    entity.depthLimitReached = true;
                    return entity;
                  }
                  const children = splitRawMimeMultipartBody(
                    split.body,
                    contentType.params.boundary || ""
                  );
                  for (let index = 0; index < children.length; index++) {
                    const child = parseEntity(children[index], depth + 1, `${partName}.${index + 1}`);
                    if (child) entity.parts.push(child);
                  }
                }
                return entity;
              }

              return parseEntity(raw, 0, "1");
            }

            function decodeRawMimeTextPart(entity) {
              const contentType = entity?.contentType?.value || "text/plain";
              if (contentType !== "text/plain" && contentType !== "text/html") return null;
              const transferEncoding = getRawMimeHeader(entity.headers, "content-transfer-encoding");
              const bodyBytes = decodeRawMimeTransferBody(entity.body, transferEncoding, {
                // Preserve the original singlepart fallback: whitespace is ignored,
                // but other non-base64 input makes atob reject the part.
                strictBase64: true,
              });
              if (!bodyBytes) return null;

              const contentTypeValue = getRawMimeHeader(entity.headers, "content-type") || "text/plain";
              const charsetMatch = contentTypeValue.match(
                /(?:^|;)\s*charset\s*=\s*(?:"([^"]+)"|'([^']+)'|([^;\s]+))/i
              );
              const charset = (
                charsetMatch?.[1] || charsetMatch?.[2] || charsetMatch?.[3] || "utf-8"
              ).trim();
              try {
                return {
                  text: new TextDecoder(charset, { fatal: false }).decode(bodyBytes),
                  isHtml: contentType === "text/html",
                  charset,
                  charsetFallback: false,
                };
              } catch (e) {
                if (!(e instanceof RangeError) && e?.name !== "RangeError") throw e;
                return {
                  text: new TextDecoder("utf-8", { fatal: false }).decode(bodyBytes),
                  isHtml: contentType === "text/html",
                  charset,
                  charsetFallback: true,
                };
              }
            }

            function extractBodyPartFromRawMime(rawBytes, bodyFormat, diagnostic = null) {
              // Body extraction runs synchronously on Thunderbird's main thread.
              // Ten MIME levels covers normal nesting while bounding adversarial input.
              const root = parseRawMimeEntity(rawBytes, { maxDepth: 10 });
              if (!root) {
                if (diagnostic) {
                  diagnostic.bodyNote = "raw MIME body extraction could not parse message";
                }
                return null;
              }
              const preferHtml = bodyFormat === "html" || bodyFormat === "markdown";
              let depthLimitReached = false;

              function findBody(entity, isRoot = false) {
                const contentType = entity.contentType.value || "text/plain";
                // Attached messages are intentionally out of scope for this fallback.
                if (contentType === "message/rfc822") return null;
                if (!isRoot && entity.contentDisposition.value === "attachment") return null;

                if (contentType.startsWith("multipart/")) {
                  if (entity.depthLimitReached) {
                    depthLimitReached = true;
                    return null;
                  }
                  if (contentType === "multipart/alternative") {
                    let fallback = null;
                    for (const child of entity.parts) {
                      const candidate = findBody(child);
                      if (!candidate) continue;
                      if (candidate.isHtml === preferHtml) return candidate;
                      if (!fallback) fallback = candidate;
                    }
                    return fallback;
                  }

                  if (contentType === "multipart/related") {
                    const start = normalizeRawMimeContentIdForMatch(
                      entity.contentType.params.start
                    );
                    const matchingRoot = start
                      ? entity.parts.find(child => normalizeRawMimeContentIdForMatch(
                        getRawMimeHeader(child.headers, "content-id")
                      ) === start)
                      : null;
                    // RFC 2387 defines one related root. An absent or unmatched
                    // start parameter falls back to the first child.
                    const relatedRoot = matchingRoot || entity.parts[0];
                    return relatedRoot ? findBody(relatedRoot) : null;
                  }

                  // mixed (and uncommon multipart subtypes) use the first suitable
                  // text part in depth-first message order.
                  for (const child of entity.parts) {
                    const candidate = findBody(child);
                    if (candidate) return candidate;
                  }
                  return null;
                }

                if (contentType !== "text/plain" && contentType !== "text/html") return null;
                try {
                  const decoded = decodeRawMimeTextPart(entity);
                  // Preserve the singlepart fallback's empty-body result shape;
                  // empty multipart candidates are not suitable alternatives.
                  return decoded && (isRoot || decoded.text) ? decoded : null;
                } catch {
                  return null;
                }
              }

              const bodyPart = findBody(root, true);
              if (!bodyPart && diagnostic) {
                diagnostic.bodyNote = depthLimitReached
                  ? "multipart body extraction hit depth cap"
                  : root.contentType.value.startsWith("multipart/")
                    ? "multipart body extraction found no suitable text part"
                    : "raw MIME body extraction found no suitable text part";
              }
              return bodyPart;
            }
            // END RAW MIME PARSING HELPERS

            // BEGIN RAW MIME ATTACHMENT HELPERS
            function attachmentSaveLooksWrong(declaredSize, actualSize) {
              if (typeof actualSize !== "number" || actualSize < 0) return true;
              if (actualSize === 0 && declaredSize !== 0) return true;
              if (typeof declaredSize !== "number" || declaredSize <= 0) return false;
              return Math.abs(actualSize - declaredSize) > Math.max(8, declaredSize * 0.05);
            }

            // Reads a message stream fully, looping on stream.available() to handle
            // mbox-stored messages where msgHdr.offlineMessageSize/messageSize can
            // underreport (observed at ~56% of true size for locally-injected mbox
            // entries). Returns the raw Latin-1 bytestring or throws. Caller must
            // close the stream.
            function readMessageStreamFully(stream, maxBytes) {
              let raw = "";
              for (let safety = 0; safety < 1024; safety++) {
                const available = stream.available();
                if (available <= 0) break;
                if (typeof maxBytes === "number" && raw.length + available > maxBytes) {
                  throw new Error(`message too large (> ${maxBytes} bytes)`);
                }
                const chunk = NetUtil.readInputStreamToString(stream, available);
                if (!chunk || chunk.length === 0) break;
                raw += chunk;
              }
              return raw;
            }

            function parseAttachmentPartsFromRawMime(rawBytes, options = {}) {
              const includeInlineImages = options.includeInlineImages === true;
              // Keep the attachment walk's historical depth allowance. Body
              // extraction uses the stricter main-thread cap in its own consumer.
              const top = parseRawMimeEntity(rawBytes, { maxDepth: 33 });
              if (!top || !top.contentType.value.startsWith("multipart/")) return [];

              const results = [];
              function walkPart(entity, insideRelated) {
                const contentType = entity.contentType;
                const contentDisposition = entity.contentDisposition;
                const ct = contentType.value || "text/plain";
                const disposition = contentDisposition.value || "";
                if (ct === "message/rfc822") return;
                if (ct.startsWith("multipart/")) {
                  const childInsideRelated = insideRelated || ct === "multipart/related";
                  for (const child of entity.parts) walkPart(child, childInsideRelated);
                  return;
                }

                const filename = getRawMimeFilename(contentDisposition, contentType);
                const contentId = normalizeRawMimeContentId(
                  getRawMimeHeader(entity.headers, "content-id")
                );
                const hasAttachmentDisposition = disposition === "attachment";
                const hasInlineFilename = disposition === "inline" && !!filename;
                const hasNonTextFilename = !!filename && !ct.startsWith("text/");
                const isInlineImage = ct.startsWith("image/") && disposition !== "attachment" &&
                  (insideRelated || disposition === "inline" || !!contentId);
                if (!hasAttachmentDisposition && !hasInlineFilename && !hasNonTextFilename &&
                    !(includeInlineImages && isInlineImage)) return;

                let bytes;
                try {
                  bytes = decodeRawMimeTransferBody(
                    entity.body,
                    getRawMimeHeader(entity.headers, "content-transfer-encoding")
                  );
                } catch {
                  bytes = null;
                }
                if (!bytes) return;
                results.push({
                  filename,
                  contentType: ct,
                  contentId,
                  disposition,
                  partName: entity.partName,
                  isInline: isInlineImage,
                  bytes,
                });
              }

              const insideRelated = top.contentType.value === "multipart/related";
              for (const child of top.parts) walkPart(child, insideRelated);
              return results;
            }
            // END RAW MIME ATTACHMENT HELPERS

	            function getMessage(messageId, folderPath, saveAttachments, bodyFormat, rawSource, includeInlineImages) {
	              return new Promise((resolve) => {
	                try {
	                  const found = findMessage(messageId, folderPath);
	                  if (found.error) {
	                    resolve({ error: found.error });
	                    return;
	                  }
	                  const { msgHdr } = found;

	                  // Raw source mode: return full RFC 2822 message
	                  if (rawSource) {
	                    let stream = null;
	                    try {
	                      const folder = msgHdr.folder;
	                      stream = folder.getMsgInputStream(msgHdr, {});
	                      // Latin-1 default preserves raw bytes; UTF-8 corrupts 8-bit content.
	                      const raw = readMessageStreamFully(stream);
	                      if (!raw || raw.length === 0) {
	                        resolve({ error: "Message has zero size - cannot read raw source" });
	                        return;
	                      }
	                      resolve({
	                        id: msgHdr.messageId,
	                        subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
	                        rawSource: raw,
	                      });
	                    } catch (e) {
	                      resolve({ error: `Failed to read raw source: ${e}` });
	                    } finally {
	                      if (stream) try { stream.close(); } catch { /* ignore */ }
	                    }
	                    return;
	                  }

	                  const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
	                    "resource:///modules/gloda/MimeMessage.sys.mjs"
	                  );

                  MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
                    if (!aMimeMsg) {
                      resolve({ error: "Could not parse message" });
                      return;
                    }

                    const requestedBodyFormat = bodyFormat || "markdown";
                    const fmt = extractFormattedBody(aMimeMsg, requestedBodyFormat);
                    let body = fmt.body;
                    let bodyIsHtml = fmt.bodyIsHtml;
                    let bodyNote = "";

                    // Bound all synchronous raw-MIME work with the existing
                    // attachment-recovery ceiling.
                    const MAX_ATTACHMENT_BYTES = 50 * 1024 * 1024;
                    let rawMimeContent = null;
                    let rawMimeAttachmentParts = null;
                    let rawMimePartsWithInlineImages = null;
                    let rawMimeAttachmentError = null;

                    // If structured MIME extraction failed, try the raw stream for
                    // local mbox folders where MsgHdrToMimeMessage returns empty parts.
                    if (!body) {
                      const fallbackContext = `thunderbird-mcp: raw MIME body fallback (${msgHdr.messageId})`;
                      let rawStream = null;
                      try {
                        const rawFolder = msgHdr.folder;
                        rawStream = rawFolder.getMsgInputStream(msgHdr, {});
                        // Latin-1 default preserves raw bytes for transfer decoding.
                        rawMimeContent = readMessageStreamFully(rawStream, MAX_ATTACHMENT_BYTES);
                        if (!rawMimeContent || rawMimeContent.length === 0) {
                          bodyNote = "raw MIME body extraction could not read message stream";
                          console.error(`${fallbackContext}: message stream has zero size`);
                        } else {
                          const bodyDiagnostic = {};
                          const extracted = extractBodyPartFromRawMime(
                            rawMimeContent,
                            requestedBodyFormat,
                            bodyDiagnostic
                          );
                          if (!extracted) {
                            bodyNote = bodyDiagnostic.bodyNote ||
                              "raw MIME body extraction found no suitable text part";
                            console.error(`${fallbackContext}: ${bodyNote}`);
                          } else {
                            if (extracted.charsetFallback) {
                              console.error(
                                `${fallbackContext}: unknown charset "${extracted.charset}", retrying with utf-8`
                              );
                            }
                            if (extracted.isHtml) {
                              if (requestedBodyFormat === "html") {
                                body = extracted.text;
                                bodyIsHtml = true;
                              } else if (requestedBodyFormat === "markdown") {
                                body = htmlToMarkdown(extracted.text);
                                bodyIsHtml = false;
                              } else {
                                body = stripHtml(extracted.text);
                                bodyIsHtml = false;
                              }
                            } else {
                              body = extracted.text;
                              bodyIsHtml = false;
                            }
                          }
                        }
                      } catch (e) {
                        bodyNote = String(e?.message || e).startsWith("message too large")
                          ? "raw MIME body extraction hit 50 MiB size cap"
                          : "raw MIME body extraction failed";
                        console.error(`${fallbackContext}: failed`, e);
                      } finally {
                        if (rawStream) try { rawStream.close(); } catch (e) {
                          console.error(`${fallbackContext}: failed to close stream`, e);
                        }
                      }
                    }

                    function getRawMimeAttachmentParts(includeInlineParts = false) {
                      const cached = includeInlineParts ? rawMimePartsWithInlineImages : rawMimeAttachmentParts;
                      if (cached) return { parts: cached };
                      if (rawMimeAttachmentError) return { error: rawMimeAttachmentError };
                      let rawStream = null;
                      try {
                        if (rawMimeContent === null) {
                          const rawFolder = msgHdr.folder;
                          rawStream = rawFolder.getMsgInputStream(msgHdr, {});
                          rawMimeContent = readMessageStreamFully(rawStream, MAX_ATTACHMENT_BYTES);
                          if (!rawMimeContent || rawMimeContent.length === 0) {
                            throw new Error("message stream has zero size");
                          }
                        }
                        const parts = parseAttachmentPartsFromRawMime(rawMimeContent, {
                          includeInlineImages: includeInlineParts,
                        });
                        if (includeInlineParts) rawMimePartsWithInlineImages = parts;
                        else rawMimeAttachmentParts = parts;
                        return { parts };
                      } catch (e) {
                        rawMimeAttachmentError = e;
                        return { error: e };
                      } finally {
                        if (rawStream) try { rawStream.close(); } catch {
                          // ignore close failure during best-effort fallback
                        }
                      }
                    }

                    // Always collect attachment metadata
                    const attachments = [];
                    const attachmentSources = [];
                    const inlineImageSources = [];
                    const knownAttachmentRecords = [];

                    function getGlodaInlineContentId(part) {
                      const rawContentId = part?.contentId || part?.contentID || part?.cid ||
                        part?.headers?.["content-id"]?.[0] || "";
                      return String(rawContentId).trim().replace(/^<+|>+$/g, "").trim();
                    }

                    if (aMimeMsg && aMimeMsg.allUserAttachments) {
                      for (const att of aMimeMsg.allUserAttachments) {
                        const info = {
                          name: att?.name || "",
                          contentType: att?.contentType || "",
                          size: typeof att?.size === "number" ? att.size : null,
                          isInline: false,
                        };
                        const source = {
                          info,
                          url: att?.url || "",
                          size: typeof att?.size === "number" ? att.size : null
                        };
                        attachments.push(info);
                        attachmentSources.push(source);
                        if (includeInlineImages) {
                          knownAttachmentRecords.push({
                            info,
                            source,
                            contentId: getGlodaInlineContentId(att),
                            partName: att?.partName || "",
                          });
                        }
                      }
                    }

                    // Find inline CID images not included in allUserAttachments.
                    // Gloda's MimeMessage commonly strips content-id headers, so the
                    // primary signal is an image/* part inside multipart/related. An
                    // explicit inline disposition or surviving Content-ID also counts;
                    // an explicit attachment disposition never does. URLs are resolved
                    // through the message service because imap-message:// is not
                    // directly fetchable by NetUtil.
                    if (aMimeMsg) {
                      // For the opt-in path this set only deduplicates the MIME-tree
                      // walk. allUserAttachments is reconciled after collection so a
                      // named inline image remains eligible for an image block.
                      const existingPartNames = includeInlineImages
                        ? new Set()
                        : new Set(attachments.map(a => a.partName).filter(Boolean));
                      function collectInlineImages(part, insideRelated, results) {
                        const ct = ((part.contentType || "").split(";")[0] || "").trim().toLowerCase();
                        // Skip nested messages -- their inline images are not ours. The
                        // Gloda root is also message/rfc822 with an empty partName; walk
                        // that wrapper for the opt-in path while preserving legacy output
                        // when includeInlineImages is omitted.
                        if (ct === "message/rfc822" && (part.partName || !includeInlineImages)) return;
                        if (ct === "multipart/related") insideRelated = true;
                        const dispositionHeader = includeInlineImages
                          ? part.headers?.["content-disposition"]?.[0] || ""
                          : "";
                        const disposition = includeInlineImages
                          ? (dispositionHeader.split(";")[0] || "").trim().toLowerCase()
                          : "";
                        const rawContentId = includeInlineImages ? getGlodaInlineContentId(part) : "";
                        const contentId = includeInlineImages
                          ? String(rawContentId).trim().replace(/^<+|>+$/g, "").trim()
                          : "";
                        const isInline = includeInlineImages
                          ? disposition !== "attachment" &&
                            (insideRelated || disposition === "inline" || !!contentId)
                          : insideRelated;
                        if (isInline && ct.startsWith("image/") && part.partName) {
                          // Deduplicate by partName (stable ID), not filename (can collide)
                          if (existingPartNames.has(part.partName)) return;
                          existingPartNames.add(part.partName);
                          // Extract filename from headers (contentType field lacks params)
                          const ctHeader = part.headers?.["content-type"]?.[0] || "";
                          const nameMatch = includeInlineImages
                            ? `${dispositionHeader};${ctHeader}`.match(/(?:filename|name)\s*=\s*"?([^";]+)"?/i)
                            : ctHeader.match(/name\s*=\s*"?([^";]+)"?/i);
                          const name = nameMatch ? nameMatch[1] : `inline_${part.partName}`;
                          results.push({
                            part,
                            name,
                            ct,
                            contentId,
                            partName: part.partName,
                          });
                        }
                        if (part.parts) {
                          for (const sub of part.parts) collectInlineImages(sub, insideRelated, results);
                        }
                      }
                      const inlineImages = [];
                      collectInlineImages(aMimeMsg, false, inlineImages);
                      if (inlineImages.length > 0) {
                        const msgUri = msgHdr.folder.getUriForMsg(msgHdr);
                        const correlations = includeInlineImages
                          ? correlateInlineImageRecords(knownAttachmentRecords, inlineImages)
                          : {
                              inlineImageEntries: inlineImages.map(inlineImage => ({
                                inlineImage,
                                matchedKnownAttachment: false,
                              })),
                            };
                        for (const correlation of correlations.inlineImageEntries) {
                          const { part, name, ct, contentId } = correlation.inlineImage;
                          // Resolve to a fetchable URL via the message service
                          let partUrl;
                          try {
                            const svc = MailServices.messageServiceFromURI(msgUri);
                            const baseUri = svc.getUrlForUri(msgUri);
                            // Append part parameter to the resolved fetchable URL
                            const sep = baseUri.spec.includes("?") ? "&" : "?";
                            partUrl = `${baseUri.spec}${sep}part=${part.partName}`;
                          } catch {
                            partUrl = "";
                          }
                          const partSize = typeof part.size === "number" && part.size > 0
                            ? part.size
                            : null;
                          let info;
                          let fallbackUrl = "";
                          if (includeInlineImages && correlation.matchedKnownAttachment) {
                            const knownRecord = correlation.metadataRecord;
                            info = knownRecord.info;
                            info.name = info.name || name;
                            info.contentType = ct || info.contentType;
                            if (!(typeof info.size === "number" && info.size > 0)) {
                              info.size = partSize;
                            }
                            info.partName = part.partName;
                            info.isInline = true;
                            info.contentId = contentId || knownRecord.contentId || null;
                            fallbackUrl = knownRecord.source?.url || "";
                          } else {
                            info = {
                              name,
                              contentType: ct,
                              size: partSize,
                              partName: part.partName,
                              isInline: true,
                            };
                            if (includeInlineImages) info.contentId = contentId || null;
                            attachments.push(info);
                            if (partUrl) {
                              attachmentSources.push({ info, url: partUrl, size: info.size });
                            }
                          }
                          inlineImageSources.push({
                            info,
                            url: partUrl || fallbackUrl,
                            size: info.size,
                            partName: part.partName,
                          });
                        }
                      }
                    }

                    // Recover Content-ID from the raw MIME tree for attribution. This
                    // only runs for the opt-in path; default getMessage output and I/O
                    // remain unchanged.
                    if (includeInlineImages && inlineImageSources.length > 0) {
                      const parsed = getRawMimeAttachmentParts(true);
                      if (!parsed.error) {
                        const rawInlineParts = (parsed.parts || []).filter(part => part.isInline);
                        const sourceRecords = inlineImageSources.map(source => ({
                          contentId: source.info.contentId,
                          partName: source.partName,
                          source,
                        }));
                        const rawPartRecords = rawInlineParts.map(rawPart => ({
                          contentId: rawPart.contentId,
                          partName: rawPart.partName,
                          rawPart,
                        }));
                        const rawCorrelations = correlateInlineImageRecords(
                          sourceRecords,
                          rawPartRecords
                        );
                        for (const correlation of rawCorrelations.inlineImageEntries) {
                          if (!correlation.matchedKnownAttachment) continue;
                          const source = correlation.metadataRecord.source;
                          const rawPart = correlation.inlineImage.rawPart;
                          if (rawPart.contentId) source.info.contentId = rawPart.contentId;
                          if (rawPart.filename && source.info.name.startsWith("inline_")) {
                            source.info.name = rawPart.filename;
                          }
                        }
                      }
                    }

                    const msgTags = getUserTags(msgHdr);
                    const baseResponse = {
                      id: msgHdr.messageId,
                      subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                      author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                      recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                      ccList: msgHdr.ccList,
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                      tags: msgTags,
                      body,
                      bodyIsHtml,
                      attachments
                    };
                    if (bodyNote) baseResponse.bodyNote = bodyNote;

                    function fetchInlineImageBase64(source) {
                      return new Promise((resolve) => {
                        if (!source.url) {
                          resolve({ error: "Inline image has no fetchable message-part URL" });
                          return;
                        }
                        try {
                          const channel = NetUtil.newChannel({
                            uri: source.url,
                            loadUsingSystemPrincipal: true,
                          });
                          NetUtil.asyncFetch(channel, (inputStream, status, request) => {
                            try {
                              if (status && status !== 0) {
                                resolve({ error: `Inline image fetch failed: ${status}` });
                                return;
                              }
                              if (!inputStream) {
                                resolve({ error: "Inline image fetch returned no data" });
                                return;
                              }
                              const requestLength = request && typeof request.contentLength === "number"
                                ? request.contentLength
                                : -1;
                              if (requestLength > 0 &&
                                  getBase64EncodedSize(requestLength) > MAX_INLINE_IMAGE_BASE64_BYTES) {
                                resolve({
                                  error: `Image exceeds per-image base64 limit (${getBase64EncodedSize(requestLength)} bytes > ${MAX_INLINE_IMAGE_BASE64_BYTES} bytes)`,
                                });
                                return;
                              }
                              // Largest decoded payload whose base64 representation fits
                              // exactly inside the per-image encoded budget.
                              const maxRawBytes = Math.floor(MAX_INLINE_IMAGE_BASE64_BYTES / 4) * 3;
                              let byteString;
                              try {
                                byteString = readMessageStreamFully(inputStream, maxRawBytes);
                              } catch {
                                resolve({
                                  error: `Image exceeds per-image base64 limit (${MAX_INLINE_IMAGE_BASE64_BYTES} bytes)`,
                                });
                                return;
                              }
                              resolve({ data: encodeByteStringToBase64(byteString) });
                            } catch (e) {
                              resolve({ error: `Inline image fetch failed: ${e}` });
                            } finally {
                              try { inputStream?.close(); } catch {}
                            }
                          });
                        } catch (e) {
                          resolve({ error: `Inline image fetch failed: ${e}` });
                        }
                      });
                    }

                    async function appendInlineImageContent() {
                      const blocks = [];
                      let totalBase64Bytes = 0;
                      const orderedInlineImageSources = orderInlineImageRecordsForBody(
                        inlineImageSources,
                        body
                      );

                      for (const source of orderedInlineImageSources) {
                        const mimeType = normalizeInlineImageMimeType(source.info.contentType);
                        let skipReason = "";

                        if (!SUPPORTED_INLINE_IMAGE_MIME_TYPES.has(mimeType)) {
                          skipReason = getInlineImageSkipReason(mimeType, 1, totalBase64Bytes);
                        } else if (totalBase64Bytes >= MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES) {
                          skipReason = `Total base64 limit reached (${MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES} bytes)`;
                        } else if (typeof source.size === "number" && source.size >= 0) {
                          skipReason = getInlineImageSkipReason(
                            mimeType,
                            getBase64EncodedSize(source.size),
                            totalBase64Bytes
                          );
                        }

                        if (skipReason) {
                          source.info.mcpImage = { status: "skipped", reason: skipReason };
                          continue;
                        }

                        const fetched = await fetchInlineImageBase64(source);
                        if (fetched.error) {
                          source.info.mcpImage = { status: "skipped", reason: fetched.error };
                          continue;
                        }

                        skipReason = getInlineImageSkipReason(
                          mimeType,
                          fetched.data.length,
                          totalBase64Bytes
                        );
                        if (skipReason) {
                          source.info.mcpImage = { status: "skipped", reason: skipReason };
                          continue;
                        }

                        const contentBlockIndex = blocks.length + 1; // text block is index 0
                        blocks.push({ type: "image", data: fetched.data, mimeType });
                        totalBase64Bytes += fetched.data.length;
                        source.info.mcpImage = {
                          status: "included",
                          contentBlockIndex,
                          base64Bytes: fetched.data.length,
                        };
                      }

                      baseResponse.inlineImageContent = {
                        included: blocks.length,
                        skipped: inlineImageSources.length - blocks.length,
                        totalBase64Bytes,
                        limits: {
                          perImageBase64Bytes: MAX_INLINE_IMAGE_BASE64_BYTES,
                          totalBase64Bytes: MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES,
                        },
                      };
                      setExtraMcpContentBlocks(baseResponse, blocks);
                    }

                    function resolveBaseResponse() {
                      if (!includeInlineImages) {
                        resolve(baseResponse);
                        return;
                      }
                      appendInlineImageContent()
                        .then(() => resolve(baseResponse))
                        .catch((e) => {
                          baseResponse.inlineImageContent = {
                            included: 0,
                            skipped: inlineImageSources.length,
                            error: `Failed to include inline images: ${e}`,
                            limits: {
                              perImageBase64Bytes: MAX_INLINE_IMAGE_BASE64_BYTES,
                              totalBase64Bytes: MAX_INLINE_IMAGES_TOTAL_BASE64_BYTES,
                            },
                          };
                          resolve(baseResponse);
                        });
                    }

                    if (!saveAttachments || attachmentSources.length === 0) {
                      resolveBaseResponse();
                      return;
                    }

                    function sanitizePathSegment(s) {
                      const sanitized = String(s || "").replace(/[^a-zA-Z0-9]/g, "_");
                      return sanitized || "message";
                    }

                    function sanitizeFilename(s) {
                      let name = String(s || "").trim();
                      if (!name) name = "attachment";
                      name = name.replace(/[^a-zA-Z0-9._-]/g, "_");
                      name = name.replace(/^_+/, "").replace(/_+$/, "");
                      return name || "attachment";
                    }

                    function ensureAttachmentDir(sanitizedId) {
                      const root = Services.dirsvc.get("TmpD", Ci.nsIFile);
                      root.append("thunderbird-mcp");
                      try {
                        root.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
                      } catch (e) {
                        if (!root.exists() || !root.isDirectory()) throw e;
                        // already exists, fine
                      }
                      const dir = root.clone();
                      dir.append(sanitizedId);
                      try {
                        dir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
                      } catch (e) {
                        if (!dir.exists() || !dir.isDirectory()) throw e;
                        // already exists, fine
                      }
                      return dir;
                    }

                    const sanitizedId = sanitizePathSegment(messageId);
                    let dir;
                    try {
                      dir = ensureAttachmentDir(sanitizedId);
                    } catch (e) {
                      for (const { info } of attachmentSources) {
                        info.error = `Failed to create attachment directory: ${e}`;
                      }
                      resolveBaseResponse();
                      return;
                    }

                    const _consumedRawMimeParts = new Set();

                                        function normalizeAttachmentFilename(name) {
                                          return String(name || "").trim().toLowerCase();
                                        }

                                        function normalizeAttachmentContentType(contentType) {
                                          return ((String(contentType || "").split(";")[0] || "").trim().toLowerCase());
                                        }

                                        function normalizeAttachmentContentId(contentId) {
                                          return String(contentId || "").trim().replace(/^<|>$/g, "").toLowerCase();
                                        }

                                        function findRawMimeAttachmentPart(info, expectedSize) {
                                          const parsed = getRawMimeAttachmentParts();
                      if (parsed.error) return { error: parsed.error };
                      const parts = parsed.parts || [];
                      const availableParts = parts.filter(part => !_consumedRawMimeParts.has(part));
                      const expectedName = normalizeAttachmentFilename(info?.name);
                      const expectedType = normalizeAttachmentContentType(info?.contentType);
                      const expectedCid = normalizeAttachmentContentId(info?.contentId || info?.contentID || info?.cid);

                      let candidates = expectedName
                        ? availableParts.filter(part => normalizeAttachmentFilename(part.filename) === expectedName)
                        : [];
                      if (candidates.length === 0 && expectedCid) {
                        candidates = availableParts.filter(part => normalizeAttachmentContentId(part.contentId) === expectedCid);
                      }
                      if (candidates.length === 0 && expectedType) {
                        candidates = availableParts.filter(part => normalizeAttachmentContentType(part.contentType) === expectedType);
                      }
                                          if (candidates.length === 0) return { part: null };

                                          candidates.sort((a, b) => {
                                            const aName = normalizeAttachmentFilename(a.filename) === expectedName ? 1 : 0;
                                            const bName = normalizeAttachmentFilename(b.filename) === expectedName ? 1 : 0;
                                            if (aName !== bName) return bName - aName;
                                            const aType = normalizeAttachmentContentType(a.contentType) === expectedType ? 1 : 0;
                                            const bType = normalizeAttachmentContentType(b.contentType) === expectedType ? 1 : 0;
                                            if (aType !== bType) return bType - aType;
                                            const aCid = normalizeAttachmentContentId(a.contentId) === expectedCid ? 1 : 0;
                                            const bCid = normalizeAttachmentContentId(b.contentId) === expectedCid ? 1 : 0;
                                            if (aCid !== bCid) return bCid - aCid;
                                            if (typeof expectedSize === "number" && expectedSize > 0) {
                                              const aDelta = Math.abs((a.bytes?.length || 0) - expectedSize);
                                              const bDelta = Math.abs((b.bytes?.length || 0) - expectedSize);
                                              if (aDelta !== bDelta) return aDelta - bDelta;
                                            }
                                            const aAttachment = a.disposition === "attachment" ? 1 : 0;
                                            const bAttachment = b.disposition === "attachment" ? 1 : 0;
                                            return bAttachment - aAttachment;
                                          });
                                          return { part: candidates[0] };
                                        }

                                        function writeBytesToFile(file, bytes) {
                                          const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                                            .createInstance(Ci.nsIFileOutputStream);
                                          ostream.init(file, 0x02 | 0x08 | 0x20, 0o600, 0);
                                          const bstream = Cc["@mozilla.org/binaryoutputstream;1"]
                                            .createInstance(Ci.nsIBinaryOutputStream);
                                          try {
                                            bstream.setOutputStream(ostream);
                                            bstream.writeByteArray(bytes, bytes.length);
                                          } finally {
                                            try { bstream.close(); } catch {}
                                            try { ostream.close(); } catch {}
                                          }
                                        }

                                        function recoverAttachmentFromRawMime(info, expectedSize, file) {
                                          const found = findRawMimeAttachmentPart(info, expectedSize);
                                          if (found.error) {
                                            return { error: `attachment body not recoverable from raw MIME: ${found.error}` };
                                          }
                                          if (!found.part) {
                                            return { error: "attachment body not recoverable from raw MIME" };
                                          }
                                          const bytes = found.part.bytes;
                                          if (!bytes || bytes.length === 0) {
                                            return { error: "attachment body not recoverable from raw MIME" };
                                          }
                                          if (bytes.length > MAX_ATTACHMENT_BYTES) {
                                            return { error: `Attachment too large (${bytes.length} bytes, limit ${MAX_ATTACHMENT_BYTES})` };
                                          }
                                          try {
                                            writeBytesToFile(file, bytes);
                                          } catch (e) {
                                            return { error: `attachment raw MIME recovery write failed: ${e}` };
                                          }
                                          let recoveredSize;
                                          try {
                                            recoveredSize = file.fileSize;
                                            if (recoveredSize > MAX_ATTACHMENT_BYTES) {
                                              return { error: `Attachment too large (${recoveredSize} bytes, limit ${MAX_ATTACHMENT_BYTES})` };
                                            }
                                          } catch {
                                            return { error: "attachment raw MIME recovery size check failed" };
                                          }
                      if (attachmentSaveLooksWrong(expectedSize, recoveredSize)) {
                        const expectedText = typeof expectedSize === "number" ? `, expected ${expectedSize}` : "";
                        return { error: `attachment body recovered from raw MIME but size mismatch (${recoveredSize} bytes${expectedText})` };
                      }
                      _consumedRawMimeParts.add(found.part);
                      return { size: recoveredSize };
                    }

                                        const saveOne = ({ info, url, size }, index) =>
                                          new Promise((done) => {
                        try {
                          if (!url) {
                            info.error = "Missing attachment URL";
                            done();
                            return;
                          }

                          const knownSize = typeof size === "number" ? size : null;
                          if (knownSize !== null && knownSize > MAX_ATTACHMENT_BYTES) {
                            info.error = `Attachment too large (${knownSize} bytes, limit ${MAX_ATTACHMENT_BYTES})`;
                            done();
                            return;
                          }

                          const idx = typeof index === "number" && Number.isFinite(index) ? index : 0;
                          let safeName = sanitizeFilename(info.name);
                          if (!safeName || safeName === "." || safeName === "..") {
                            safeName = `attachment_${idx}`;
                          }
                          const file = dir.clone();
                          file.append(safeName);

                          try {
                            file.createUnique(Ci.nsIFile.NORMAL_FILE_TYPE, 0o600);
                          } catch (e) {
                            info.error = `Failed to create file: ${e}`;
                            done();
                            return;
                          }

                          const channel = NetUtil.newChannel({
                            uri: url,
                            loadUsingSystemPrincipal: true
                          });

                          NetUtil.asyncFetch(channel, (inputStream, status, request) => {
                            try {
                              if (status && status !== 0) {
                                try { inputStream?.close(); } catch {}
                                info.error = `Fetch failed: ${status}`;
                                try { file.remove(false); } catch {}
                                done();
                                return;
                              }
                              if (!inputStream) {
                                info.error = "Fetch returned no data";
                                try { file.remove(false); } catch {}
                                done();
                                return;
                              }

                              try {
                                const reqLen = request && typeof request.contentLength === "number" ? request.contentLength : -1;
                                if (reqLen >= 0 && reqLen > MAX_ATTACHMENT_BYTES) {
                                  try { inputStream.close(); } catch {}
                                  info.error = `Attachment too large (${reqLen} bytes, limit ${MAX_ATTACHMENT_BYTES})`;
                                  try { file.remove(false); } catch {}
                                  done();
                                  return;
                                }
                              } catch {
                                // ignore contentLength failures
                              }

                              const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                                .createInstance(Ci.nsIFileOutputStream);
                              ostream.init(file, -1, -1, 0);

                              NetUtil.asyncCopy(inputStream, ostream, (copyStatus) => {
                                try {
                                  if (copyStatus && copyStatus !== 0) {
                                    info.error = `Write failed: ${copyStatus}`;
                                    try { file.remove(false); } catch {}
                                    done();
                                    return;
                                  }

                                  let actualSize = null;
                                  try {
                                    actualSize = file.fileSize;
                                    if (actualSize > MAX_ATTACHMENT_BYTES) {
                                      info.error = `Attachment too large (${actualSize} bytes, limit ${MAX_ATTACHMENT_BYTES})`;
                                      try { file.remove(false); } catch {}
                                      done();
                                      return;
                                    }
                                  } catch {
                                    actualSize = null;
                                  }

                                  if (attachmentSaveLooksWrong(knownSize, actualSize)) {
                                    const recovered = recoverAttachmentFromRawMime(info, knownSize, file);
                                    if (recovered.error) {
                                      info.error = recovered.error;
                                      delete info.filePath;
                                      try { file.remove(false); } catch {}
                                      done();
                                      return;
                                    }
                                    actualSize = recovered.size;
                                  }

                                  info.filePath = file.path;
                                  done();
                                                                } catch (e) {
                                  info.error = `Write failed: ${e}`;
                                  try { file.remove(false); } catch {}
                                  done();
                                }
                              });
                            } catch (e) {
                              info.error = `Fetch failed: ${e}`;
                              try { file.remove(false); } catch {}
                              done();
                            }
                          });
                        } catch (e) {
                          info.error = String(e);
                          done();
                        }
                      });

                    (async () => {
                      try {
                        await Promise.all(attachmentSources.map((src, i) => saveOne(src, i)));
                      } catch (e) {
                        // Per-attachment errors are handled; this is just a safeguard.
                        for (const { info } of attachmentSources) {
                          if (!info.error) info.error = `Unexpected save error: ${e}`;
                        }
                      }
                      resolveBaseResponse();
                    })();
                  }, true, { examineEncryptedParts: true });

	                } catch (e) {
	                  resolve({ error: e.toString() });
	                }
	              });
	            }

            async function getMessages(messages, saveAttachments, bodyFormat, rawSource) {
              if (typeof messages === "string") {
                try { messages = JSON.parse(messages); } catch { /* leave as-is */ }
              }
              if (!Array.isArray(messages) || messages.length === 0) {
                return { error: "messages must be a non-empty array of { messageId, folderPath } objects" };
              }
              const getMessagesLimit = getConfiguredGetMessagesLimit();
              if (messages.length > getMessagesLimit) {
                return { error: `getMessages accepts at most ${getMessagesLimit} messages per call` };
              }

              const results = [];
              for (let i = 0; i < messages.length; i++) {
                const ref = messages[i];
                if (!ref || typeof ref !== "object" || Array.isArray(ref)) {
                  results.push({
                    index: i,
                    error: "Message reference must be an object with messageId and folderPath",
                  });
                  continue;
                }

                const { messageId, folderPath } = ref;
                if (typeof messageId !== "string" || !messageId) {
                  results.push({
                    index: i,
                    folderPath,
                    error: "messageId must be a non-empty string",
                  });
                  continue;
                }
                if (typeof folderPath !== "string" || !folderPath) {
                  results.push({
                    index: i,
                    messageId,
                    error: "folderPath must be a non-empty string",
                  });
                  continue;
                }

                const result = await getMessage(
                  messageId,
                  folderPath,
                  saveAttachments,
                  bodyFormat,
                  rawSource
                );
                results.push({ index: i, messageId, folderPath, ...result });
              }

              const failed = results.filter(result => result.error).length;
              return {
                messages: results,
                requested: messages.length,
                succeeded: results.length - failed,
                failed,
                max: getMessagesLimit,
              };
            }

            /**
             * Composes a new email. Opens a compose window for review, or sends
             * directly when skipReview is true.
             *
             * HTML body handling quirks:
             * 1. Strip newlines from HTML - Thunderbird adds <br> for each \n
             * 2. Encode non-ASCII as HTML entities - compose window has charset issues
             *    with emojis/unicode even with <meta charset="UTF-8">
             */
            function composeMail(to, subject, body, cc, bcc, isHtml, from, attachments, skipReview) {
              try {
                if (skipReview && isSkipReviewBlocked()) {
                  return { error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review window instead." };
                }
                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                composeFields.to = to || "";
                composeFields.cc = cc || "";
                composeFields.bcc = bcc || "";
                composeFields.subject = subject || "";

                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.composeFields = composeFields;

                const identityResult = setComposeIdentity(msgComposeParams, from, null);
                if (identityResult && identityResult.error) return identityResult;

                // Match body shape and format to caller intent / identity pref.
                // When the resolved mode is plain, ship a plain body -- the HTML
                // envelope would otherwise render as literal text in plain-mode
                // editors and recipients.
                const { useHtml, format } = resolveComposeFormat(msgComposeParams.identity, isHtml, Ci.nsIMsgCompType.New);
                msgComposeParams.format = format;
                if (useHtml) {
                  const formatted = formatBodyHtml(body, isHtml);
                  composeFields.body = isHtml && formatted.includes('<html')
                    ? formatted
                    : `<html><head><meta charset="UTF-8"></head><body>${formatted}</body></html>`;
                } else {
                  composeFields.body = body || "";
                }

                const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

                if (skipReview) {
                  return sendMessageDirectly(composeFields, msgComposeParams.identity, fileDescs, null, Ci.nsIMsgCompType.New, Ci.nsIMsgCompDeliverMode.Now, useHtml ? "text/html" : "text/plain").then(result => {
                    if (result.success) {
                      let msg = "Message sent";
                      if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                      result.message = msg;
                    }
                    return result;
                  });
                }

                // Attach via composeFields BEFORE opening the window so it
                // renders with the attachments already present. This avoids the
                // old getMostRecentWindow("msgcompose") race: that helper returns
                // the most recently *focused* compose window, so a second
                // sendMail call -- or the user simply clicking an older compose
                // window while this one opens -- would steal or drop the
                // attachments. composeFields binds them to this exact message,
                // exactly like the direct-send path (sendMessageDirectly).
                for (const att of descsToMsgAttachments(fileDescs)) {
                  composeFields.addAttachment(att);
                }

                const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                  .getService(Ci.nsIMsgComposeService);
                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                let msg = "Compose window opened";
                if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                return { success: true, message: msg };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Saves a composed message to the identity's Drafts folder without
             * sending or opening a compose window. The destination folder is
             * resolved by Thunderbird from the identity's draft-folder pref.
             */
            function saveDraft(to, subject, body, cc, bcc, isHtml, from, attachments) {
              try {
                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                composeFields.to = to || "";
                composeFields.cc = cc || "";
                composeFields.bcc = bcc || "";
                composeFields.subject = subject || "";

                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.composeFields = composeFields;

                const identityResult = setComposeIdentity(msgComposeParams, from, null);
                if (identityResult && identityResult.error) return identityResult;

                const { useHtml, format } = resolveComposeFormat(msgComposeParams.identity, isHtml, Ci.nsIMsgCompType.New);
                msgComposeParams.format = format;
                if (useHtml) {
                  const formatted = formatBodyHtml(body, isHtml);
                  composeFields.body = isHtml && formatted.includes('<html')
                    ? formatted
                    : `<html><head><meta charset="UTF-8"></head><body>${formatted}</body></html>`;
                } else {
                  composeFields.body = body || "";
                }

                const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

                return sendMessageDirectly(
                  composeFields,
                  msgComposeParams.identity,
                  fileDescs,
                  null,
                  Ci.nsIMsgCompType.New,
                  Ci.nsIMsgCompDeliverMode.SaveAsDraft,
                  useHtml ? "text/html" : "text/plain"
                ).then(result => {
                  if (result.success) {
                    let msg = "Draft saved";
                    if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                    result.message = msg;
                  }
                  return result;
                });
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Replies to a message with quoted original. Opens a compose window
             * for review, or sends directly when skipReview is true.
             *
             * Review path uses Thunderbird's native reply compose flow so it can
             * build the quoted original, place the identity signature according
             * to user preferences, and set threading headers/disposition flags.
             * skipReview still uses direct send, so it keeps a manual quoted body
             * and manually marks the original as replied after a successful send.
             */
	            function replyToMessage(messageId, folderPath, body, replyAll, isHtml, to, cc, bcc, from, attachments, skipReview) {
	              return new Promise((resolve) => {
	                try {
	                  if (skipReview && isSkipReviewBlocked()) {
	                    resolve({ error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review window instead." });
	                    return;
	                  }
	                  const found = findMessage(messageId, folderPath);
	                  if (found.error) {
	                    resolve({ error: found.error });
	                    return;
	                  }
	                  const { msgHdr, folder } = found;
	                  const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);
	                  const msgURI = folder.getUriForMsg(msgHdr);
	                  const compType = replyAll ? Ci.nsIMsgCompType.ReplyAll : Ci.nsIMsgCompType.Reply;

	                  const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
	                    .createInstance(Ci.nsIMsgComposeParams);

	                  const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
	                    .createInstance(Ci.nsIMsgCompFields);

	                  msgComposeParams.type = compType;
	                  msgComposeParams.originalMsgURI = msgURI;
	                  msgComposeParams.composeFields = composeFields;

	                  try {
	                    msgComposeParams.origMsgHdr = msgHdr;
	                  } catch {}

	                  const identityResult = setComposeIdentity(msgComposeParams, from, folder.server);
	                  if (identityResult && identityResult.error) {
	                    resolve(identityResult);
	                    return;
	                  }

	                  // Resolve compose mode against caller intent + identity pref.
	                  // The skipReview branch reads useHtml below to shape the body.
	                  const { useHtml: replyUseHtml, format: replyFormat } =
	                    resolveComposeFormat(msgComposeParams.identity, isHtml, compType);
	                  msgComposeParams.format = replyFormat;

	                  // Pass through only the fields the caller explicitly provided.
	                  // Any field left undefined is filled in by Thunderbird's native
	                  // reply/reply-all machinery (including proper Reply-To,
	                  // Mail-Followup-To, mailing-list handling, and self-filtering
	                  // against the selected identity). Our old custom
	                  // getReplyAllCcRecipients path bypassed all of that.
	                  const reviewTo = to;
	                  const reviewCc = cc;

	                  if (skipReview) {
	                    const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
	                      "resource:///modules/gloda/MimeMessage.sys.mjs"
                      );

	                    MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
	                      try {
	                        const originalBody = extractPlainTextBody(aMimeMsg);

	                        if (replyAll) {
	                          composeFields.to = to || msgHdr.author;
	                          if (cc) {
	                            composeFields.cc = cc;
	                          } else {
	                            const replyAllCc = getReplyAllCcRecipients(msgHdr, folder);
	                            if (replyAllCc) composeFields.cc = replyAllCc;
	                          }
	                        } else {
	                          composeFields.to = to || msgHdr.author;
	                          if (cc) composeFields.cc = cc;
	                        }

	                        composeFields.bcc = bcc || "";

	                        const origSubject = msgHdr.mime2DecodedSubject || msgHdr.subject || "";
	                        composeFields.subject = /^re:/i.test(origSubject) ? origSubject : `Re: ${origSubject}`;
	                        composeFields.references = `<${messageId}>`;
	                        composeFields.setHeader("In-Reply-To", `<${messageId}>`);

	                        const dateStr = msgHdr.date ? new Date(msgHdr.date / 1000).toLocaleString() : "";
	                        const author = msgHdr.mime2DecodedAuthor || msgHdr.author || "";

	                        // Direct send goes through nsIMsgSend, not nsIMsgCompose, so
	                        // it still uses a hand-built quoted body and cannot place the
	                        // identity signature according to reply preferences. The shape
	                        // matches the resolved compose mode -- shipping an HTML envelope
	                        // for a plain-format send would otherwise render as literal
	                        // markup in the recipient's mail client.
	                        if (replyUseHtml) {
	                          const quotedLines = originalBody.split('\n').map(line =>
	                            `&gt; ${escapeHtml(line)}`
	                          ).join('<br>');
	                          const quotedHtml = escapeHtml(originalBody).replace(/\n/g, '<br>');
	                          const quoteBlock = isHtml
	                            ? `<br><br>On ${dateStr}, ${escapeHtml(author)} wrote:<blockquote type="cite">${quotedHtml}</blockquote>`
	                            : `<br><br>On ${dateStr}, ${escapeHtml(author)} wrote:<br>${quotedLines}`;
	                          composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${formatBodyHtml(body, isHtml)}${quoteBlock}</body></html>`;
	                        } else {
	                          const quotedLines = originalBody.split('\n').map(line => `> ${line}`).join('\n');
	                          composeFields.body = `${body || ""}\n\nOn ${dateStr}, ${author} wrote:\n${quotedLines}`;
	                        }

	                        sendMessageDirectly(composeFields, msgComposeParams.identity, fileDescs, msgURI, compType, Ci.nsIMsgCompDeliverMode.Now, replyUseHtml ? "text/html" : "text/plain").then(result => {
	                          if (result.success) {
	                            let repliedDisposition = null;
	                            try {
	                              repliedDisposition = Ci.nsIMsgFolder.nsMsgDispositionState_Replied;
	                            } catch {}
	                            markMessageDispositionState(msgHdr, repliedDisposition);

	                            let msg = "Reply sent";
	                            if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
	                            result.message = msg;
	                          }
	                          resolve(result);
	                        });
	                      } catch (e) {
	                        resolve({ error: e.toString() });
	                      }
	                    }, true, { examineEncryptedParts: true });
	                    return;
	                  }

	                  openComposeWindowWithCustomizations(
	                    msgComposeParams,
	                    msgURI,
	                    compType,
	                    msgComposeParams.identity,
	                    body,
	                    isHtml,
	                    reviewTo,
	                    reviewCc,
	                    bcc,
	                    fileDescs
	                  ).then(result => {
	                    if (result.success) {
	                      let msg = "Reply window opened";
	                      if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
	                      result.message = msg;
	                    }
	                    resolve(result);
	                  });

	                } catch (e) {
	                  resolve({ error: e.toString() });
	                }
	              });
            }

            /**
             * Forwards a message with original content and attachments.
             *
             * Review path uses Thunderbird's native ForwardInline compose flow so
             * TB builds the forward body with a proper <blockquote type="cite">
             * quote, auto-attaches the original message's attachments, places the
             * identity signature per user preferences, and sets the $Forwarded
             * disposition on the original after a successful send. The caller's
             * intro body is injected via NotifyComposeBodyReady, mirroring how
             * replyToMessage handles intro injection.
             *
             * skipReview still uses direct send, so it keeps a manual forward
             * block + auto-attaches originals from MsgHdrToMimeMessage + manually
             * marks the original as forwarded after a successful send.
             */
            function forwardMessage(messageId, folderPath, to, body, isHtml, cc, bcc, from, attachments, skipReview) {
              return new Promise((resolve) => {
                try {
                  if (skipReview && isSkipReviewBlocked()) {
                    resolve({ error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review window instead." });
                    return;
                  }
                  const found = findMessage(messageId, folderPath);
                  if (found.error) {
                    resolve({ error: found.error });
                    return;
                  }
                  const { msgHdr, folder } = found;
                  const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);
                  const msgURI = folder.getUriForMsg(msgHdr);
                  const compType = Ci.nsIMsgCompType.ForwardInline;

                  const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                    .createInstance(Ci.nsIMsgComposeParams);

                  const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                    .createInstance(Ci.nsIMsgCompFields);

                  msgComposeParams.type = compType;
                  msgComposeParams.originalMsgURI = msgURI;
                  msgComposeParams.composeFields = composeFields;

                  try {
                    msgComposeParams.origMsgHdr = msgHdr;
                  } catch {}

                  const identityResult = setComposeIdentity(msgComposeParams, from, folder.server);
                  if (identityResult && identityResult.error) {
                    resolve(identityResult);
                    return;
                  }

                  // ForwardInline only passes the format flag through when it is
                  // Default or OppositeOfDefault -- HTML/PlainText are ignored and
                  // the identity's compose pref always wins. resolveComposeFormat
                  // returns OppositeOfDefault when the caller's explicit isHtml
                  // conflicts with the identity pref so we can still force the
                  // intended editor mode.
                  const { useHtml: fwdUseHtml, format: fwdFormat } =
                    resolveComposeFormat(msgComposeParams.identity, isHtml, compType);
                  msgComposeParams.format = fwdFormat;

                  if (skipReview) {
                    const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
                      "resource:///modules/gloda/MimeMessage.sys.mjs"
                    );

                    MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
                      try {
                        const originalBody = extractPlainTextBody(aMimeMsg);

                        composeFields.to = to;
                        composeFields.cc = cc || "";
                        composeFields.bcc = bcc || "";

                        const origSubject = msgHdr.mime2DecodedSubject || msgHdr.subject || "";
                        composeFields.subject = /^fwd:/i.test(origSubject) ? origSubject : `Fwd: ${origSubject}`;

                        const dateStr = msgHdr.date ? new Date(msgHdr.date / 1000).toLocaleString() : "";
                        const fwdAuthor = msgHdr.mime2DecodedAuthor || msgHdr.author || "";
                        const fwdRecipients = msgHdr.mime2DecodedRecipients || msgHdr.recipients || "";

                        // Direct send goes through nsIMsgSend, not nsIMsgCompose,
                        // so we hand-build the forward block. The shape matches the
                        // resolved compose mode -- shipping an HTML envelope for a
                        // plain-format send would render as literal markup in the
                        // recipient's mail client.
                        if (fwdUseHtml) {
                          const fwdHeaderHtml =
                            `-------- Forwarded Message --------<br>` +
                            `Subject: ${escapeHtml(origSubject)}<br>` +
                            `Date: ${dateStr}<br>` +
                            `From: ${escapeHtml(fwdAuthor)}<br>` +
                            `To: ${escapeHtml(fwdRecipients)}<br><br>`;
                          const quotedHtml = escapeHtml(originalBody).replace(/\n/g, '<br>');
                          const quotedLinesHtml = originalBody.split('\n').map(line =>
                            `&gt; ${escapeHtml(line)}`
                          ).join('<br>');
                          const forwardBlock = isHtml
                            ? `<blockquote type="cite">${fwdHeaderHtml}${quotedHtml}</blockquote>`
                            : `${fwdHeaderHtml}${quotedLinesHtml}`;
                          const introHtml = body ? formatBodyHtml(body, isHtml) + '<br><br>' : "";
                          composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${introHtml}${forwardBlock}</body></html>`;
                        } else {
                          const fwdHeader =
                            `-------- Forwarded Message --------\n` +
                            `Subject: ${origSubject}\n` +
                            `Date: ${dateStr}\n` +
                            `From: ${fwdAuthor}\n` +
                            `To: ${fwdRecipients}\n\n`;
                          composeFields.body = `${body ? body + '\n\n' : ''}${fwdHeader}${originalBody}`;
                        }

                        const origDescs = [];
                        if (aMimeMsg && aMimeMsg.allUserAttachments) {
                          for (const att of aMimeMsg.allUserAttachments) {
                            try {
                              origDescs.push({ url: att.url, name: att.name, contentType: att.contentType });
                            } catch {
                              // Skip unreadable original attachments
                            }
                          }
                        }
                        const allDescs = [...origDescs, ...fileDescs];

                        sendMessageDirectly(composeFields, msgComposeParams.identity, allDescs, msgURI, compType, Ci.nsIMsgCompDeliverMode.Now, fwdUseHtml ? "text/html" : "text/plain").then(result => {
                          if (result.success) {
                            let forwardedDisposition = null;
                            try {
                              forwardedDisposition = Ci.nsIMsgFolder.nsMsgDispositionState_Forwarded;
                            } catch {}
                            markMessageDispositionState(msgHdr, forwardedDisposition);

                            let msg = `Forward sent with ${allDescs.length} attachment(s)`;
                            if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                            result.message = msg;
                          }
                          resolve(result);
                        });
                      } catch (e) {
                        resolve({ error: e.toString() });
                      }
                    }, true, { examineEncryptedParts: true });
                    return;
                  }

                  // Review path: TB builds the forward body and auto-attaches the
                  // original's attachments via ForwardInline. The intro body and
                  // user-specified extra attachments are injected once the compose
                  // window's editor signals NotifyComposeBodyReady. Subject is set
                  // by TB from origMsgHdr.
                  openComposeWindowWithCustomizations(
                    msgComposeParams,
                    msgURI,
                    compType,
                    msgComposeParams.identity,
                    body,
                    isHtml,
                    to,
                    cc,
                    bcc,
                    fileDescs
                  ).then(result => {
                    if (result.success) {
                      let msg = "Forward window opened";
                      if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                      result.message = msg;
                    }
                    resolve(result);
                  });
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

            function displayMessage(messageId, folderPath, displayMode) {
              const found = findMessage(messageId, folderPath);
              if (found.error) return found;
              const { msgHdr } = found;

              const VALID_DISPLAY_MODES = ["3pane", "tab", "window"];
              const mode = displayMode || "3pane";
              if (!VALID_DISPLAY_MODES.includes(mode)) {
                return { error: `Invalid displayMode: "${mode}". Must be one of: ${VALID_DISPLAY_MODES.join(", ")}` };
              }

              try {
                const { MailUtils } = ChromeUtils.importESModule(
                  "resource:///modules/MailUtils.sys.mjs"
                );

                switch (mode) {
                  case "tab": {
                    const win = Services.wm.getMostRecentWindow("mail:3pane");
                    if (!win) return { error: "No Thunderbird mail window found (required for tab mode)" };
                    const msgUri = msgHdr.folder.getUriForMsg(msgHdr);
                    const tabmail = win.document.getElementById("tabmail");
                    if (!tabmail) return { error: "Could not access tabmail interface" };
                    tabmail.openTab("mailMessageTab", { messageURI: msgUri, msgHdr });
                    break;
                  }
                  case "window":
                    // openMessageInNewWindow doesn't need an existing 3pane window
                    MailUtils.openMessageInNewWindow(msgHdr);
                    break;
                  case "3pane":
                    MailUtils.displayMessageInFolderTab(msgHdr);
                    break;
                }
              } catch (e) {
                return { error: `Failed to display message: ${e.message || e}` };
              }

              return { success: true, displayMode: mode, subject: msgHdr.mime2DecodedSubject || msgHdr.subject || "" };
            }

            function getRecentMessages(folderPath, daysBack, maxResults, offset, unreadOnly, flaggedOnly, includeSubfolders) {
              const results = [];
              const days = Number.isFinite(Number(daysBack)) && Number(daysBack) > 0 ? Math.floor(Number(daysBack)) : 7;
              const cutoffTs = (Date.now() - days * 86400000) * 1000; // Thunderbird uses microseconds
              const requestedLimit = Number(maxResults);
              const effectiveLimit = Math.min(
                Number.isFinite(requestedLimit) && requestedLimit > 0 ? Math.floor(requestedLimit) : DEFAULT_MAX_RESULTS,
                MAX_SEARCH_RESULTS_CAP
              );

              function collectFromFolder(folder) {
                if (results.length >= SEARCH_COLLECTION_CAP) return;

                try {
                  const db = folder.msgDatabase;
                  if (!db) return;

                  for (const msgHdr of db.enumerateMessages()) {
                    if (results.length >= SEARCH_COLLECTION_CAP) break;

                    const msgDateTs = msgHdr.date || 0;
                    if (msgDateTs < cutoffTs) continue;
                    if (unreadOnly && msgHdr.isRead) continue;
                    if (flaggedOnly && !msgHdr.isFlagged) continue;

                    const msgTags = getUserTags(msgHdr);
                    const preview = msgHdr.getStringProperty("preview") || "";
                    const result = {
                      id: msgHdr.messageId,
                      threadId: msgHdr.threadId, // folder-local, use with folderPath for grouping
                      subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                      author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                      recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                      folder: folder.prettyName,
                      folderPath: folder.URI,
                      read: msgHdr.isRead,
                      flagged: msgHdr.isFlagged,
                      tags: msgTags,
                      _dateTs: msgDateTs
                    };
                    if (preview) result.preview = preview;
                    results.push(result);
                  }
                } catch {
                  // Skip inaccessible folders
                }

                const recurse = includeSubfolders !== false; // default true
                if (recurse && folder.hasSubFolders) {
                  for (const subfolder of folder.subFolders) {
                    if (results.length >= SEARCH_COLLECTION_CAP) break;
                    collectFromFolder(subfolder);
                  }
                }
              }

              if (folderPath) {
                // Specific folder
                const opened = openFolder(folderPath);
                if (opened.error) return { error: opened.error };
                collectFromFolder(opened.folder);
              } else {
                // All folders across accessible accounts
                for (const account of getAccessibleAccounts()) {
                  if (results.length >= SEARCH_COLLECTION_CAP) break;
                  try {
                    const root = account.incomingServer.rootFolder;
                    collectFromFolder(root);
                  } catch {
                    // Skip inaccessible accounts
                  }
                }
              }

              results.sort((a, b) => b._dateTs - a._dateTs);

              return paginate(results, offset, effectiveLimit);
            }

            function isTrashOrDescendant(folder) {
              try {
                return folder.isSpecialFolder(Ci.nsMsgFolderFlags.Trash, true);
              } catch {
                return false;
              }
            }

            function deleteMessages(messageIds, folderPath) {
              try {
                // MCP clients may send arrays as JSON strings
                if (typeof messageIds === "string") {
                  try { messageIds = JSON.parse(messageIds); } catch { /* leave as-is */ }
                }
                if (!Array.isArray(messageIds) || messageIds.length === 0) {
                  return { error: "messageIds must be a non-empty array of strings" };
                }
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }

                const opened = openFolder(folderPath);
                if (opened.error) return { error: opened.error };
                const { folder, db } = opened;

                // Find all requested message headers
                const found = [];
                const notFound = [];
                for (const msgId of messageIds) {
                  if (typeof msgId !== "string" || !msgId) {
                    notFound.push(msgId);
                    continue;
                  }
                  let hdr = null;
                  const hasDirectLookup = typeof db.getMsgHdrForMessageID === "function";
                  if (hasDirectLookup) {
                    try { hdr = db.getMsgHdrForMessageID(msgId); } catch { hdr = null; }
                  }
                  if (!hdr) {
                    for (const h of db.enumerateMessages()) {
                      if (h.messageId === msgId) { hdr = h; break; }
                    }
                  }
                  if (hdr) {
                    found.push(hdr);
                  } else {
                    notFound.push(msgId);
                  }
                }

                if (found.length === 0) {
                  return { error: "No matching messages found" };
                }

                // Drafts get moved to Trash instead of hard-deleted
                const isDrafts = typeof folder.getFlag === "function" && folder.getFlag(Ci.nsMsgFolderFlags.Drafts);
                let trashFolder = null;

                if (isDrafts) {
                  trashFolder = findTrashFolder(folder);

                  if (trashFolder) {
                    MailServices.copy.copyMessages(folder, found, trashFolder, true, null, null, false);
                  } else {
                    // No trash found, fall back to regular delete
                    folder.deleteMessages(found, null, false, true, null, false);
                  }
                } else {
                  // Thunderbird permanently deletes from Trash (including its
                  // descendants) and for non-move IMAP delete models.
                  const server = folder.server;
                  const deletionMovesToTrash = !isTrashOrDescendant(folder) &&
                    (server?.type !== "imap" ||
                      server.QueryInterface(Ci.nsIImapIncomingServer).deleteModel ===
                        Ci.nsMsgImapDeleteModels.MoveToTrash);
                  if (deletionMovesToTrash) {
                    trashFolder = findTrashFolder(folder);
                    if (!trashFolder) {
                      return { error: "Trash folder not found" };
                    }
                  }
                  folder.deleteMessages(found, null, false, true, null, false);
                }

                let result = { success: true, deleted: found.length };
                if (isDrafts && trashFolder) result.movedToTrash = true;
                if (notFound.length > 0) result.notFound = notFound;
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function updateMessage(messageId, messageIds, folderPath, read, flagged, addTags, removeTags, moveTo, trash) {
              try {
                // Normalize to an array of IDs
                if (typeof messageIds === "string") {
                  try { messageIds = JSON.parse(messageIds); } catch { /* leave as-is */ }
                }
                if (messageId && messageIds) {
                  return { error: "Specify messageId or messageIds, not both" };
                }
                if (messageId) {
                  messageIds = [messageId];
                }
                if (!Array.isArray(messageIds) || messageIds.length === 0) {
                  return { error: "messageId or messageIds is required" };
                }
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }

                // Coerce boolean params (MCP clients may send strings)
                if (read !== undefined) read = read === true || read === "true";
                if (flagged !== undefined) flagged = flagged === true || flagged === "true";
                if (trash !== undefined) trash = trash === true || trash === "true";
                if (moveTo !== undefined && (typeof moveTo !== "string" || !moveTo)) {
                  return { error: "moveTo must be a non-empty string" };
                }
                // Coerce tag arrays (MCP clients may send JSON strings)
                if (typeof addTags === "string") {
                  try { addTags = JSON.parse(addTags); } catch { /* leave as-is */ }
                }
                if (typeof removeTags === "string") {
                  try { removeTags = JSON.parse(removeTags); } catch { /* leave as-is */ }
                }
                if (addTags !== undefined && !Array.isArray(addTags)) {
                  return { error: "addTags must be an array of tag keyword strings" };
                }
                if (removeTags !== undefined && !Array.isArray(removeTags)) {
                  return { error: "removeTags must be an array of tag keyword strings" };
                }

                if (moveTo && trash === true) {
                  return { error: "Cannot specify both moveTo and trash" };
                }

                // Find all requested message headers
                const opened = openFolder(folderPath);
                if (opened.error) return { error: opened.error };
                const { folder, db } = opened;

                const foundHdrs = [];
                const notFound = [];
                for (const msgId of messageIds) {
                  if (typeof msgId !== "string" || !msgId) {
                    notFound.push(msgId);
                    continue;
                  }
                  let hdr = null;
                  const hasDirectLookup = typeof db.getMsgHdrForMessageID === "function";
                  if (hasDirectLookup) {
                    try { hdr = db.getMsgHdrForMessageID(msgId); } catch { hdr = null; }
                  }
                  if (!hdr) {
                    for (const h of db.enumerateMessages()) {
                      if (h.messageId === msgId) { hdr = h; break; }
                    }
                  }
                  if (hdr) {
                    foundHdrs.push(hdr);
                  } else {
                    notFound.push(msgId);
                  }
                }

                if (foundHdrs.length === 0) {
                  return { error: "No matching messages found" };
                }

                const actions = [];

                if (read !== undefined) {
                  // Use folder-level API for proper IMAP sync (hdr.markRead
                  // only updates the local DB, doesn't queue IMAP commands)
                  folder.markMessagesRead(foundHdrs, read);
                  actions.push({ type: "read", value: read });
                }

                if (flagged !== undefined) {
                  folder.markMessagesFlagged(foundHdrs, flagged);
                  actions.push({ type: "flagged", value: flagged });
                }

                if (addTags || removeTags) {
                  // Validate: allow IMAP atom chars per RFC 3501 plus & for modified UTF-7
                  // tag keys that Thunderbird generates for non-ASCII labels.
                  // Blocks whitespace, null bytes, parens, braces, wildcards, quotes, backslash.
                  const VALID_TAG = /^[a-zA-Z0-9_$.\-&+!']+$/;
                  const tagsToAdd = (addTags || []).filter(t => typeof t === "string" && VALID_TAG.test(t));
                  const tagsToRemove = (removeTags || []).filter(t => typeof t === "string" && VALID_TAG.test(t));
                  // Use folder-level keyword APIs for proper IMAP sync
                  if (tagsToAdd.length > 0) {
                    folder.addKeywordsToMessages(foundHdrs, tagsToAdd.join(" "));
                    actions.push({ type: "addTags", value: tagsToAdd });
                  }
                  if (tagsToRemove.length > 0) {
                    folder.removeKeywordsFromMessages(foundHdrs, tagsToRemove.join(" "));
                    actions.push({ type: "removeTags", value: tagsToRemove });
                  }
                }

                let targetFolder = null;

                if (trash === true) {
                  targetFolder = findTrashFolder(folder);
                  if (!targetFolder) {
                    return { error: "Trash folder not found" };
                  }
                } else if (moveTo) {
                  const moveResult = getAccessibleFolder(moveTo);
                  if (moveResult.error) return moveResult;
                  targetFolder = moveResult.folder;
                }

                if (targetFolder) {
                  // Note: on IMAP, tags/flags set above may not transfer to the
                  // moved copy. If both tags and move are needed, consider making
                  // two separate updateMessage calls (tags first, then move).
                  MailServices.copy.copyMessages(folder, foundHdrs, targetFolder, true, null, null, false);
                  actions.push({ type: "move", to: targetFolder.URI });
                }

                const result = { success: true, updated: foundHdrs.length, actions };
                if (targetFolder && (addTags || removeTags)) {
                  result.warning = "Tags were applied before move; on IMAP accounts, tags may not transfer to the moved copy. Consider separate calls if tags are missing.";
                }
                if (notFound.length > 0) result.notFound = notFound;
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function createFolder(parentFolderPath, name) {
              try {
                if (typeof parentFolderPath !== "string" || !parentFolderPath) {
                  return { error: "parentFolderPath must be a non-empty string" };
                }
                if (typeof name !== "string" || !name) {
                  return { error: "name must be a non-empty string" };
                }

                const parentResult = getAccessibleFolder(parentFolderPath);
                if (parentResult.error) return parentResult;
                const parent = parentResult.folder;

                parent.createSubfolder(name, null);

                // Try to return the new folder's URI
                let newPath = null;
                try {
                  if (parent.hasSubFolders) {
                    for (const sub of parent.subFolders) {
                      if (sub.prettyName === name || sub.name === name) {
                        newPath = sub.URI;
                        break;
                      }
                    }
                  }
                } catch {
                  // Folder may not be immediately visible (IMAP)
                }

                return {
                  success: true,
                  message: `Folder "${name}" created`,
                  path: newPath
                };
              } catch (e) {
                const msg = e.toString();
                if (msg.includes("NS_MSG_FOLDER_EXISTS")) {
                  return { error: `Folder "${name}" already exists under this parent` };
                }
                return { error: msg };
              }
            }

            function renameFolder(folderPath, newName) {
              try {
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }
                if (typeof newName !== "string" || !newName) {
                  return { error: "newName must be a non-empty string" };
                }

                const renameResult = getAccessibleFolder(folderPath);
                if (renameResult.error) return renameResult;
                const folder = renameResult.folder;

                folder.rename(newName, null);
                return {
                  success: true,
                  message: `Folder renamed to "${newName}"`,
                  oldPath: folderPath,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function deleteFolder(folderPath) {
              try {
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }

                const delResult = getAccessibleFolder(folderPath);
                if (delResult.error) return delResult;
                const folder = delResult.folder;
                const folderName = folder.prettyName || folder.name || folderPath;

                const parent = folder.parent;
                if (!parent) {
                  return { error: "Cannot delete a root folder" };
                }

                // Check if folder is already in Trash — if so, permanently delete
                if (isTrashOrDescendant(folder)) {
                  // Permanently delete — deleteSelf requires a msgWindow
                  const win = Services.wm.getMostRecentWindow("mail:3pane");
                  folder.deleteSelf(win?.msgWindow ?? null);
                  return { success: true, message: `Folder "${folderName}" permanently deleted` };
                } else {
                  // Move to trash
                  const trashFolder = findTrashFolder(folder);
                  if (!trashFolder) {
                    return { error: "Trash folder not found" };
                  }
                  MailServices.copy.copyFolder(folder, trashFolder, true, null, null);
                  return { success: true, message: `Folder "${folderName}" moved to Trash` };
                }
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Find a special folder by flag bit, searching the account's folder tree.
             */
            function findSpecialFolder(root, flagBit) {
              const search = (folder) => {
                try {
                  if (folder.getFlag && folder.getFlag(flagBit)) return folder;
                } catch {}
                if (folder.hasSubFolders) {
                  for (const sub of folder.subFolders) {
                    const found = search(sub);
                    if (found) return found;
                  }
                }
                return null;
              };
              return search(root);
            }

            /**
             * Recursively delete all messages in a folder and its subfolders.
             * Returns total count of messages deleted.
             */
            function deleteAllMessagesRecursive(folder) {
              let count = 0;
              try {
                const db = folder.msgDatabase;
                if (db) {
                  const hdrs = [];
                  for (const hdr of db.enumerateMessages()) hdrs.push(hdr);
                  if (hdrs.length > 0) {
                    folder.deleteMessages(hdrs, null, true, false, null, false);
                    count += hdrs.length;
                  }
                }
              } catch (e) {
                // Continue traversal; log per-folder failures so partial empties are visible.
                console.error("thunderbird-mcp: deleteMessages failed for folder", folder?.URI || folder?.name, ":", e);
              }
              if (folder.hasSubFolders) {
                for (const sub of folder.subFolders) {
                  count += deleteAllMessagesRecursive(sub);
                }
              }
              return count;
            }

            function emptyTrash(accountId) {
              try {
                const accounts = accountId
                  ? [MailServices.accounts.getAccount(accountId)].filter(Boolean)
                  : Array.from(getAccessibleAccounts());
                if (accountId && accounts.length === 0) {
                  return { error: `Account not found: ${accountId}` };
                }
                if (accountId && !isAccountAllowed(accountId)) {
                  return { error: `Account not accessible: ${accountId}` };
                }

                const results = [];
                for (const account of accounts) {
                  const root = account.incomingServer?.rootFolder;
                  if (!root) continue;
                  const trash = findSpecialFolder(root, Ci.nsMsgFolderFlags.Trash);
                  if (!trash) {
                    results.push({ account: account.key, status: "no Trash folder found" });
                    continue;
                  }
                  // Use Thunderbird's native emptyTrash when available (handles
                  // IMAP expunge, subfolders, and compaction correctly)
                  if (typeof trash.emptyTrash === "function") {
                    try {
                      // TB 128+ dropped msgWindow arg
                      trash.emptyTrash(null);
                    } catch (e) {
                      const isArgError = (e && (e.result === 0x80570001 || e.result === 0x80570009)) ||
                        String(e).includes("Not enough arguments") ||
                        String(e).includes("Could not convert JavaScript argument");
                      if (isArgError) {
                        const win = Services.wm.getMostRecentWindow("mail:3pane");
                        trash.emptyTrash(win?.msgWindow ?? null, null);
                      } else {
                        throw e;
                      }
                    }
                    results.push({ account: account.key, folder: trash.URI, status: "emptied" });
                  } else {
                    const deleted = deleteAllMessagesRecursive(trash);
                    results.push({ account: account.key, folder: trash.URI, deleted });
                  }
                }
                return { success: true, results };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function emptyJunk(accountId) {
              try {
                const JUNK_FLAG = 0x40000000;
                const accounts = accountId
                  ? [MailServices.accounts.getAccount(accountId)].filter(Boolean)
                  : Array.from(getAccessibleAccounts());
                if (accountId && accounts.length === 0) {
                  return { error: `Account not found: ${accountId}` };
                }
                if (accountId && !isAccountAllowed(accountId)) {
                  return { error: `Account not accessible: ${accountId}` };
                }

                const results = [];
                for (const account of accounts) {
                  const root = account.incomingServer?.rootFolder;
                  if (!root) continue;
                  const junk = findSpecialFolder(root, JUNK_FLAG);
                  if (!junk) {
                    results.push({ account: account.key, status: "no Junk folder found" });
                    continue;
                  }
                  // No native emptyJunk in Thunderbird, delete recursively
                  const deleted = deleteAllMessagesRecursive(junk);
                  results.push({ account: account.key, folder: junk.URI, deleted });
                }
                return { success: true, results };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function moveFolder(folderPath, newParentPath) {
              try {
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }
                if (typeof newParentPath !== "string" || !newParentPath) {
                  return { error: "newParentPath must be a non-empty string" };
                }

                const srcResult = getAccessibleFolder(folderPath);
                if (srcResult.error) return srcResult;
                const folder = srcResult.folder;
                const folderName = folder.prettyName || folder.name || folderPath;

                const destResult = getAccessibleFolder(newParentPath);
                if (destResult.error) return destResult;
                const newParent = destResult.folder;
                const parentName = newParent.prettyName || newParent.name || newParentPath;

                if (folder.parent && folder.parent.URI === newParentPath) {
                  return { error: "Folder is already under this parent" };
                }

                MailServices.copy.copyFolder(folder, newParent, true, null, null);
                return {
                  success: true,
                  message: `Folder "${folderName}" moved to "${parentName}"`,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            // ── Filter constant maps ──

            // Values are the real nsMsgSearchAttrib enum (nsMsgSearchCore.idl).
            // They are NOT sequential past allAddresses -- do not renumber.
            const ATTRIB_MAP = {
              subject: 0, from: 1, body: 2, date: 3, priority: 4,
              status: 5, to: 6, cc: 7, toOrCc: 8, allAddresses: 9,
              ageInDays: 12, size: 14, tag: 16, hasAttachment: 44,
              junkStatus: 45, junkPercent: 46, otherHeader: 52,
            };
            const ATTRIB_NAMES = Object.fromEntries(Object.entries(ATTRIB_MAP).map(([k, v]) => [v, k]));

            const OP_MAP = {
              contains: 0, doesntContain: 1, is: 2, isnt: 3, isEmpty: 4,
              isBefore: 5, isAfter: 6, isHigherThan: 7, isLowerThan: 8,
              beginsWith: 9, endsWith: 10,
              soundsLike: 11, ldapDwim: 12,
              isGreaterThan: 13, isLessThan: 14,
              nameCompletion: 15, isInAB: 16, isntInAB: 17, isntEmpty: 18,
              matches: 19, doesntMatch: 20,
            };
            const OP_NAMES = Object.fromEntries(Object.entries(OP_MAP).map(([k, v]) => [v, k]));

            // Values are the real nsMsgRuleActionType enum (nsMsgFilterCore.idl).
            // They are NOT sequential -- copyToFolder jumps to 16, value 8 (Label)
            // no longer exists, and 18 is killSubthread. Do not renumber.
            const ACTION_MAP = {
              moveToFolder: 1, changePriority: 2, delete: 3,
              markRead: 4, killThread: 5, watchThread: 6,
              markFlagged: 7, reply: 9, forward: 10,
              stopExecution: 11, deleteFromServer: 12, leaveOnServer: 13,
              junkScore: 14, fetchBody: 15, copyToFolder: 16,
              addTag: 17, killSubthread: 18, markUnread: 19,
              custom: -1,
            };
            const ACTION_NAMES = Object.fromEntries(Object.entries(ACTION_MAP).map(([k, v]) => [v, k]));

            // nsIMsgSearchValue is a tagged union: the value lives in a different
            // member depending on the attribute. Attributes not listed here are
            // string-valued and use `.str`. Setting the wrong member (e.g. `.str`
            // on a Date attrib) makes XPCOM throw NS_ERROR_ILLEGAL_VALUE, which
            // otherwise fails the whole filter operation. Every mapping below
            // matches upstream nsMsgSearchCore.idl (IS_STRING_ATTRIBUTE),
            // nsMsgSearchValue.cpp and nsMsgSearchTerm.cpp @ 72b8ba0.
            const VALUE_KIND = {
              3: "date",           // Date -> PRTime (microseconds since epoch)
              4: "priority",       // Priority
              5: "status",         // MsgStatus bitmask -> .status (u.msgStatus)
              12: "age",           // AgeInDays (whole days, int32)
              14: "size",          // Size (bytes)
              44: "hasAttachment", // HasAttachmentStatus -> .status = Attachment flag
              45: "junkStatus",    // JunkStatus
              46: "junkPercent",   // JunkPercent
            };

            // nsMsgMessageFlags.Attachment: the fixed flag a HasAttachmentStatus
            // term stores in .status. Has-vs-hasn't is decided by the operator
            // (is/isnt), not the value; Thunderbird ignores the file token too.
            const MSG_FLAG_ATTACHMENT = 0x10000000;

            function parseIntStrict(raw, label) {
              const n = parseInt(raw, 10);
              if (Number.isNaN(n)) throw new Error(`Invalid numeric value for ${label}: ${raw}`);
              return n;
            }

            // Write a condition's value into an nsIMsgSearchValue using the union
            // member that matches its attribute. Throws on malformed input so the
            // caller surfaces a clear error rather than persisting a broken filter.
            function setSearchValue(value, attribNum, raw) {
              value.attrib = attribNum;
              switch (VALUE_KIND[attribNum]) {
                case "date": {
                  const ms = Date.parse(raw);
                  if (Number.isNaN(ms)) throw new Error(`Invalid date value: ${raw}`);
                  value.date = ms * 1000; // PRTime is microseconds since epoch
                  break;
                }
                case "age": value.age = parseIntStrict(raw, "ageInDays"); break;
                case "size": value.size = parseIntStrict(raw, "size"); break;
                case "priority": value.priority = parseIntStrict(raw, "priority"); break;
                case "status": value.status = parseIntStrict(raw, "status"); break;
                case "hasAttachment": value.status = MSG_FLAG_ATTACHMENT; break; // raw ignored, like TB
                case "junkStatus": value.junkStatus = parseIntStrict(raw, "junkStatus"); break;
                case "junkPercent": value.junkPercent = parseIntStrict(raw, "junkPercent"); break;
                default: value.str = raw || "";
              }
            }

            // Read an nsIMsgSearchValue back into a display string, reading the
            // union member that matches its attribute. Falls back to `.str`.
            function readSearchValue(term) {
              try {
                switch (VALUE_KIND[term.attrib]) {
                  case "date": {
                    const d = term.value.date;
                    return d ? new Date(d / 1000).toISOString() : "";
                  }
                  case "age": return String(term.value.age);
                  case "size": return String(term.value.size);
                  case "priority": return String(term.value.priority);
                  case "status": return String(term.value.status);
                  case "hasAttachment": return (term.value.status & MSG_FLAG_ATTACHMENT) ? "true" : "false";
                  case "junkStatus": return String(term.value.junkStatus);
                  case "junkPercent": return String(term.value.junkPercent);
                  default: return term.value.str || "";
                }
              } catch {
                try { return term.value.str || ""; } catch { return ""; }
              }
            }

            function getFilterListForAccount(accountId) {
              if (!isAccountAllowed(accountId)) {
                return { error: `Account not accessible: ${accountId}` };
              }
              const account = MailServices.accounts.getAccount(accountId);
              if (!account) return { error: `Account not found: ${accountId}` };
              const server = account.incomingServer;
              if (!server) return { error: "Account has no server" };
              if (server.canHaveFilters === false) return { error: "Account does not support filters" };
              const filterList = server.getFilterList(null);
              if (!filterList) return { error: "Could not access filter list" };
              return { account, server, filterList };
            }

            function serializeFilter(filter, index) {
              const terms = [];
              try {
                for (const term of filter.searchTerms) {
                  const t = {
                    attrib: ATTRIB_NAMES[term.attrib] || String(term.attrib),
                    op: OP_NAMES[term.op] || String(term.op),
                    booleanAnd: term.booleanAnd,
                  };
                  t.value = readSearchValue(term);
                  if (term.arbitraryHeader) t.header = term.arbitraryHeader;
                  terms.push(t);
                }
              } catch {
                // searchTerms iteration may fail on some TB versions
                // Try indexed access via termAsString as fallback
              }

              const actions = [];
              for (let a = 0; a < filter.actionCount; a++) {
                try {
                  const action = filter.getActionAt(a);
                  const act = { type: ACTION_NAMES[action.type] || String(action.type) };
                  if (action.type === 1 || action.type === 16) {
                    // moveToFolder or copyToFolder
                    act.value = action.targetFolderUri || "";
                  } else if (action.type === 2) {
                    act.value = String(action.priority);
                  } else if (action.type === 14) {
                    act.value = String(action.junkScore);
                  } else {
                    try { if (action.strValue) act.value = action.strValue; } catch {}
                  }
                  actions.push(act);
                } catch {
                  // Skip unreadable actions
                }
              }

              return {
                index,
                name: filter.filterName,
                enabled: filter.enabled,
                type: filter.filterType,
                temporary: filter.temporary,
                terms,
                actions,
              };
            }

            function buildTerms(filter, conditions) {
              for (const cond of conditions) {
                const term = filter.createTerm();
                // SECURITY: strict allow-list. The previous `?? parseInt(...)`
                // fallback let callers pass raw nsMsgSearchAttrib enum values that
                // aren't in ATTRIB_MAP, bypassing the intended named-action set.
                if (!Object.prototype.hasOwnProperty.call(ATTRIB_MAP, cond.attrib)) {
                  throw new Error(`Unknown attribute: ${cond.attrib}`);
                }
                term.attrib = ATTRIB_MAP[cond.attrib];

                if (!Object.prototype.hasOwnProperty.call(OP_MAP, cond.op)) {
                  throw new Error(`Unknown operator: ${cond.op}`);
                }
                term.op = OP_MAP[cond.op];

                const value = term.value;
                setSearchValue(value, term.attrib, cond.value);
                term.value = value;

                term.booleanAnd = cond.booleanAnd !== false;
                if (cond.header) term.arbitraryHeader = cond.header;
                filter.appendTerm(term);
              }
            }

            function buildActions(filter, actions) {
              for (const act of actions) {
                const action = filter.createAction();
                // SECURITY: strict allow-list. The previous `?? parseInt(...)`
                // fallback accepted any numeric nsMsgFilterAction value, which
                // would auto-expose new (or legacy) action types we never
                // intended to surface -- including historic "run program" flavors.
                if (!Object.prototype.hasOwnProperty.call(ACTION_MAP, act.type)) {
                  throw new Error(`Unknown action type: ${act.type}`);
                }
                const typeNum = ACTION_MAP[act.type];
                action.type = typeNum;

                if (act.value) {
                  if (typeNum === 1 || typeNum === 16) {
                    // Move/Copy to folder -- verify target is accessible
                    const targetCheck = getAccessibleFolder(act.value);
                    if (targetCheck.error) throw new Error(`Filter target folder not accessible: ${act.value}`);
                    action.targetFolderUri = act.value;
                  } else if (typeNum === 2) {
                    action.priority = parseInt(act.value);
                  } else if (typeNum === 14) {
                    action.junkScore = parseInt(act.value);
                  } else {
                    action.strValue = act.value;
                  }
                }
                filter.appendAction(action);
              }
            }

            // ── Filter tool handlers ──

            function listFilters(accountId) {
              try {
                const results = [];
                let accounts;
                if (accountId) {
                  if (!isAccountAllowed(accountId)) {
                    return { error: `Account not accessible: ${accountId}` };
                  }
                  const account = MailServices.accounts.getAccount(accountId);
                  if (!account) return { error: `Account not found: ${accountId}` };
                  accounts = [account];
                } else {
                  accounts = Array.from(getAccessibleAccounts());
                }

                for (const account of accounts) {
                  if (!account) continue;
                  try {
                    const server = account.incomingServer;
                    if (!server || server.canHaveFilters === false) continue;

                    const filterList = server.getFilterList(null);
                    if (!filterList) continue;

                    const filters = [];
                    for (let i = 0; i < filterList.filterCount; i++) {
                      try {
                        filters.push(serializeFilter(filterList.getFilterAt(i), i));
                      } catch {
                        // Skip unreadable filters
                      }
                    }

                    results.push({
                      accountId: account.key,
                      accountName: server.prettyName,
                      filterCount: filterList.filterCount,
                      loggingEnabled: filterList.loggingEnabled,
                      filters,
                    });
                  } catch {
                    // Skip inaccessible accounts
                  }
                }

                return results;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function createFilter(accountId, name, enabled, type, conditions, actions, insertAtIndex) {
              try {
                // Coerce arrays from MCP client string serialization
                if (typeof conditions === "string") {
                  try { conditions = JSON.parse(conditions); } catch { /* leave as-is */ }
                }
                if (typeof actions === "string") {
                  try { actions = JSON.parse(actions); } catch { /* leave as-is */ }
                }
                if (typeof enabled === "string") enabled = enabled === "true";
                if (typeof type === "string") type = parseInt(type);
                if (typeof insertAtIndex === "string") insertAtIndex = parseInt(insertAtIndex);

                if (!Array.isArray(conditions) || conditions.length === 0) {
                  return { error: "conditions must be a non-empty array" };
                }
                if (!Array.isArray(actions) || actions.length === 0) {
                  return { error: "actions must be a non-empty array" };
                }

                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                const filter = filterList.createFilter(name);
                filter.enabled = enabled !== false;
                filter.filterType = (Number.isFinite(type) && type > 0) ? type : 17; // inbox + manual

                buildTerms(filter, conditions);
                buildActions(filter, actions);

                const idx = (insertAtIndex != null && insertAtIndex >= 0)
                  ? Math.min(insertAtIndex, filterList.filterCount)
                  : filterList.filterCount;
                filterList.insertFilterAt(idx, filter);
                filterList.saveToDefaultFile();

                return {
                  success: true,
                  name: filter.filterName,
                  index: idx,
                  filterCount: filterList.filterCount,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function updateFilter(accountId, filterIndex, name, enabled, type, conditions, actions) {
              try {
                // Coerce from MCP client
                if (typeof filterIndex === "string") filterIndex = parseInt(filterIndex);
                if (!Number.isInteger(filterIndex)) return { error: "filterIndex must be an integer" };
                if (typeof enabled === "string") enabled = enabled === "true";
                if (typeof type === "string") type = parseInt(type);
                if (typeof conditions === "string") {
                  try { conditions = JSON.parse(conditions); } catch {
                    return { error: "conditions must be a valid JSON array" };
                  }
                }
                if (typeof actions === "string") {
                  try { actions = JSON.parse(actions); } catch {
                    return { error: "actions must be a valid JSON array" };
                  }
                }

                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                if (filterIndex < 0 || filterIndex >= filterList.filterCount) {
                  return { error: `Invalid filter index: ${filterIndex}` };
                }

                const filter = filterList.getFilterAt(filterIndex);
                const changes = [];

                if (name !== undefined) {
                  filter.filterName = name;
                  changes.push("name");
                }
                if (enabled !== undefined) {
                  filter.enabled = enabled;
                  changes.push("enabled");
                }
                if (type !== undefined) {
                  filter.filterType = type;
                  changes.push("type");
                }

                const replaceConditions = Array.isArray(conditions) && conditions.length > 0;
                const replaceActions = Array.isArray(actions) && actions.length > 0;

                if (replaceConditions || replaceActions) {
                  // No clearTerms/clearActions API -- rebuild filter via remove+insert
                  const newFilter = filterList.createFilter(filter.filterName);
                  newFilter.enabled = filter.enabled;
                  newFilter.filterType = filter.filterType;

                  // Build or copy conditions
                  if (replaceConditions) {
                    buildTerms(newFilter, conditions);
                    changes.push("conditions");
                  } else {
                    // Copy existing terms -- abort on failure to prevent data loss
                    let termsCopied = 0;
                    try {
                      for (const term of filter.searchTerms) {
                        const newTerm = newFilter.createTerm();
                        newTerm.attrib = term.attrib;
                        newTerm.op = term.op;
                        const val = newTerm.value;
                        try { setSearchValue(val, term.attrib, readSearchValue(term)); } catch {}
                        newTerm.value = val;
                        newTerm.booleanAnd = term.booleanAnd;
                        try { newTerm.beginsGrouping = term.beginsGrouping; } catch {}
                        try { newTerm.endsGrouping = term.endsGrouping; } catch {}
                        try { if (term.arbitraryHeader) newTerm.arbitraryHeader = term.arbitraryHeader; } catch {}
                        newFilter.appendTerm(newTerm);
                        termsCopied++;
                      }
                    } catch (e) {
                      return { error: `Failed to copy existing conditions: ${e.toString()}` };
                    }
                    if (termsCopied === 0) {
                      return { error: "Cannot update: failed to read existing filter conditions" };
                    }
                  }

                  // Build or copy actions
                  if (replaceActions) {
                    buildActions(newFilter, actions);
                    changes.push("actions");
                  } else {
                    for (let a = 0; a < filter.actionCount; a++) {
                      try {
                        const origAction = filter.getActionAt(a);
                        const newAction = newFilter.createAction();
                        newAction.type = origAction.type;
                        try { newAction.targetFolderUri = origAction.targetFolderUri; } catch {}
                        try { newAction.priority = origAction.priority; } catch {}
                        try { newAction.strValue = origAction.strValue; } catch {}
                        try { newAction.junkScore = origAction.junkScore; } catch {}
                        newFilter.appendAction(newAction);
                      } catch {}
                    }
                  }

                  filterList.removeFilterAt(filterIndex);
                  filterList.insertFilterAt(filterIndex, newFilter);
                }

                filterList.saveToDefaultFile();

                return {
                  success: true,
                  changes,
                  filter: serializeFilter(filterList.getFilterAt(filterIndex), filterIndex),
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function deleteFilter(accountId, filterIndex) {
              try {
                if (typeof filterIndex === "string") filterIndex = parseInt(filterIndex);
                if (!Number.isInteger(filterIndex)) return { error: "filterIndex must be an integer" };

                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                if (filterIndex < 0 || filterIndex >= filterList.filterCount) {
                  return { error: `Invalid filter index: ${filterIndex}` };
                }

                const filter = filterList.getFilterAt(filterIndex);
                const filterName = filter.filterName;
                filterList.removeFilterAt(filterIndex);
                filterList.saveToDefaultFile();

                return { success: true, deleted: filterName, remainingCount: filterList.filterCount };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function reorderFilters(accountId, fromIndex, toIndex) {
              try {
                if (typeof fromIndex === "string") fromIndex = parseInt(fromIndex);
                if (typeof toIndex === "string") toIndex = parseInt(toIndex);
                if (!Number.isInteger(fromIndex)) return { error: "fromIndex must be an integer" };
                if (!Number.isInteger(toIndex)) return { error: "toIndex must be an integer" };

                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                if (fromIndex < 0 || fromIndex >= filterList.filterCount) {
                  return { error: `Invalid source index: ${fromIndex}` };
                }
                if (toIndex < 0 || toIndex >= filterList.filterCount) {
                  return { error: `Invalid target index: ${toIndex}` };
                }

                // moveFilterAt is unreliable — use remove + insert instead
                // Adjust toIndex after removal: if moving down, indices shift
                const filter = filterList.getFilterAt(fromIndex);
                filterList.removeFilterAt(fromIndex);
                const adjustedTo = (fromIndex < toIndex) ? toIndex - 1 : toIndex;
                filterList.insertFilterAt(adjustedTo, filter);
                filterList.saveToDefaultFile();

                return { success: true, name: filter.filterName, fromIndex, toIndex };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function applyFilters(accountId, folderPath) {
              try {
                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                const afResult = getAccessibleFolder(folderPath);
                if (afResult.error) return afResult;
                const folder = afResult.folder;

                // Try MailServices.filters first, fall back to XPCOM contract ID
                let filterService;
                try {
                  filterService = MailServices.filters;
                } catch {}
                if (!filterService) {
                  try {
                    filterService = Cc["@mozilla.org/messenger/filter-service;1"]
                      .getService(Ci.nsIMsgFilterService);
                  } catch {}
                }
                if (!filterService) {
                  return { error: "Filter service not available in this Thunderbird version" };
                }
                filterService.applyFiltersToFolders(filterList, [folder], null);

                // applyFiltersToFolders is async — returns immediately
                return {
                  success: true,
                  message: "Filters applied (processing may take a moment)",
                  folder: folderPath,
                  enabledFilters: (() => {
                    let count = 0;
                    for (let i = 0; i < filterList.filterCount; i++) {
                      if (filterList.getFilterAt(i).enabled) count++;
                    }
                    return count;
                  })(),
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Validate tool arguments against the tool's inputSchema.
             * Checks required fields, types (string, number, boolean, array, object),
             * and rejects unknown properties.
             * Returns an array of error strings (empty = valid).
             */
            /**
             * Walk a JSON-Schema subtree and report any errors against `value`.
             * Not a full JSON Schema implementation -- intentionally minimal --
             * but covers the keywords actually used by toolSchemas:
             *   - type (string/number/integer/boolean/array/object)
             *   - enum, minLength, and base64 contentEncoding
             *   - properties + required + additionalProperties (on objects)
             *   - items (on arrays), oneOf, and anyOf
             * `path` is the dotted property path used in error messages.
             */
            // BEGIN TOOL SCHEMA VALIDATOR
            function validateAgainstSchema(value, schema, path, errors) {
              if (!schema || value === undefined || value === null) return;

              const expectedType = schema.type;
              if (expectedType === "array") {
                if (!Array.isArray(value)) {
                  errors.push(`Parameter '${path}' must be an array, got ${typeof value}`);
                  return;
                }
                if (schema.items) {
                  for (let i = 0; i < value.length; i++) {
                    // Array items are never nullable: validateAgainstSchema
                    // returns early on null/undefined, so a null item would
                    // otherwise skip the item schema entirely. Reject explicitly.
                    if (value[i] === null || value[i] === undefined) {
                      errors.push(`Parameter '${path}[${i}]' must not be null`);
                      continue;
                    }
                    validateAgainstSchema(value[i], schema.items, `${path}[${i}]`, errors);
                  }
                }
              } else if (expectedType === "object") {
                if (typeof value !== "object" || Array.isArray(value)) {
                  errors.push(`Parameter '${path}' must be an object, got ${Array.isArray(value) ? "array" : typeof value}`);
                  return;
                }
                const nestedProps = schema.properties || {};
                const nestedRequired = schema.required || [];
                // Inline attachment objects accept `content` as a legacy alias,
                // but runtime uses `base64` whenever it is present. Do not reject
                // an ignored `content` value after the preferred payload validates.
                const hasPreferredBase64 = value.base64 !== undefined
                  && value.base64 !== null
                  && nestedProps.base64?.contentEncoding === "base64"
                  && nestedProps.content?.contentEncoding === "base64";
                for (const r of nestedRequired) {
                  if (value[r] === undefined || value[r] === null) {
                    errors.push(`Missing required parameter: ${path}.${r}`);
                  }
                }
                for (const [k, v] of Object.entries(value)) {
                  if (k === "content" && hasPreferredBase64) continue;
                  const has = Object.prototype.hasOwnProperty.call(nestedProps, k);
                  if (!has) {
                    if (schema.additionalProperties === false) {
                      errors.push(`Unknown parameter: ${path}.${k}`);
                    }
                    continue;
                  }
                  validateAgainstSchema(v, nestedProps[k], `${path}.${k}`, errors);
                }
              } else if (expectedType === "integer") {
                if (typeof value !== "number" || !Number.isInteger(value)) {
                  errors.push(`Parameter '${path}' must be an integer, got ${typeof value === "number" ? "non-integer number" : typeof value}`);
                  return;
                }
              } else if (expectedType && typeof value !== expectedType) {
                errors.push(`Parameter '${path}' must be ${expectedType}, got ${typeof value}`);
                return;
              }

              if (expectedType === "string") {
                if (schema.minLength !== undefined && value.length < schema.minLength) {
                  errors.push(`Parameter '${path}' must contain at least ${schema.minLength} character(s)`);
                }
                if (schema.contentEncoding === "base64" && !isValidBase64(value)) {
                  errors.push(`Parameter '${path}' must contain valid base64 data`);
                }
              }

              // anyOf: accept the value when one or more branches validate.
              if (Array.isArray(schema.anyOf) && schema.anyOf.length > 0) {
                let matched = 0;
                const branchFailures = [];
                for (const branch of schema.anyOf) {
                  const branchErrors = [];
                  validateAgainstSchema(value, branch, path, branchErrors);
                  if (branchErrors.length === 0) matched++;
                  else branchFailures.push(branchErrors);
                }
                if (matched === 0) {
                  const details = [...new Set(branchFailures.flat())].join("; ");
                  errors.push(`Parameter '${path}' did not match any required schema alternative${details ? `: ${details}` : ""}`);
                }
              }

              // oneOf: accept the value if exactly one branch validates clean.
              // Used by the attachments array items (string | object).
              if (Array.isArray(schema.oneOf) && schema.oneOf.length > 0) {
                let matched = 0;
                const branchFailures = [];
                for (const branch of schema.oneOf) {
                  const branchErrors = [];
                  validateAgainstSchema(value, branch, path, branchErrors);
                  if (branchErrors.length === 0) matched++;
                  else branchFailures.push(branchErrors);
                }
                if (matched === 0) {
                  const details = [...new Set(branchFailures.flat())].join("; ");
                  errors.push(`Parameter '${path}' did not match any allowed schema variant${details ? `: ${details}` : ""}`);
                } else if (matched > 1) {
                  errors.push(`Parameter '${path}' matched more than one schema variant`);
                }
              }

              // enum: explicit value allow-list (e.g. bodyFormat).
              if (Array.isArray(schema.enum) && !schema.enum.includes(value)) {
                errors.push(`Parameter '${path}' must be one of ${JSON.stringify(schema.enum)}, got ${JSON.stringify(value)}`);
              }
            }
            // END TOOL SCHEMA VALIDATOR

            function validateToolArgs(name, args) {
              const tool = buildTools().find(t => t.name === name);
              const schema = tool?.inputSchema;
              if (!schema) return [`Unknown tool: ${name}`];

              const errors = [];
              const props = schema.properties || {};
              const required = schema.required || [];

              // Check required fields
              for (const key of required) {
                if (args[key] === undefined || args[key] === null) {
                  errors.push(`Missing required parameter: ${key}`);
                }
              }

              // Check types and reject unknown properties
              for (const [key, value] of Object.entries(args)) {
                // Use hasOwnProperty to prevent inherited properties like
                // 'constructor' or 'toString' from bypassing unknown-param checks.
                const propSchema = Object.prototype.hasOwnProperty.call(props, key) ? props[key] : undefined;
                if (!propSchema) {
                  errors.push(`Unknown parameter: ${key}`);
                  continue;
                }
                if (value === undefined || value === null) continue;

                validateAgainstSchema(value, propSchema, key, errors);
                // minItems/maxItems sit outside validateAgainstSchema (which covers
                // type/items/object/integer/oneOf/enum); keep the array-length bounds
                // so the v0.6.0 getMessages-batch caps stay enforced.
                if (propSchema.type === "array" && Array.isArray(value)) {
                  if (propSchema.minItems !== undefined && value.length < propSchema.minItems) {
                    errors.push(`Parameter '${key}' must contain at least ${propSchema.minItems} item(s)`);
                  }
                  if (propSchema.maxItems !== undefined && value.length > propSchema.maxItems) {
                    errors.push(`Parameter '${key}' must contain at most ${propSchema.maxItems} item(s)`);
                  }
                }
              }

              return errors;
            }

            /**
             * Coerce tool arguments to match expected schema types.
             * MCP clients may send "true"/"false" as strings for booleans,
             * "50" as strings for numbers, or JSON-encoded arrays as strings.
             * Mutates and returns the args object.
             */
            function coerceToolArgs(name, args) {
              const tool = buildTools().find(t => t.name === name);
              const schema = tool?.inputSchema;
              if (!schema) return args;
              const props = schema.properties || {};
              for (const [key, value] of Object.entries(args)) {
                if (value === undefined || value === null) continue;
                const propSchema = Object.prototype.hasOwnProperty.call(props, key) ? props[key] : undefined;
                if (!propSchema) continue;
                const expected = propSchema.type;
                if (expected === "boolean" && typeof value === "string") {
                  if (value === "true") args[key] = true;
                  else if (value === "false") args[key] = false;
                } else if (expected === "number" && typeof value === "string") {
                  // Reject blank/whitespace strings -- Number("") is 0 which
                  // would silently coerce empty input into a valid number.
                  if (value.trim() === "") continue;
                  const n = Number(value);
                  if (Number.isFinite(n)) args[key] = n;
                } else if (expected === "integer" && typeof value === "string") {
                  if (value.trim() === "") continue;
                  const n = Number(value);
                  if (Number.isFinite(n) && Number.isInteger(n)) args[key] = n;
                } else if (expected === "array" && typeof value === "string") {
                  try {
                    const parsed = JSON.parse(value);
                    if (Array.isArray(parsed)) args[key] = parsed;
                  } catch (e) {
                    // Leave value as-is so validator surfaces a typed error to the client.
                    console.warn(`thunderbird-mcp: coerceToolArgs JSON.parse failed for key=${key}:`, e.message);
                  }
                }
              }
              return args;
            }

            async function callTool(name, args) {
              switch (name) {
                case "listAccounts":
                  return listAccounts();
                case "listFolders":
                  return listFolders(args.accountId, args.folderPath, args.format);
                case "searchMessages":
                  return await searchMessages(args.query || "", args.folderPath, args.startDate, args.endDate, args.maxResults, args.offset, args.sortOrder, args.unreadOnly, args.flaggedOnly, args.tag, args.includeSubfolders, args.countOnly, args.searchBody, args.dedupByMessageId);
                case "getMessage":
                  return await getMessage(args.messageId, args.folderPath, args.saveAttachments, args.bodyFormat, args.rawSource, args.includeInlineImages);
                case "getMessages":
                  return await getMessages(args.messages, args.saveAttachments, args.bodyFormat, args.rawSource);
                case "searchContacts":
                  return searchContacts(args.query || "", args.maxResults);
                case "getContact":
                  return getContact(args.contactId);
                case "createContact":
                  return createContact(args.email, args.displayName, args.firstName, args.lastName, args.phones, args.addresses, args.organization, args.title, args.note, args.birthday, args.addressBookId);
                case "updateContact":
                  return updateContact(args.contactId, args.email, args.displayName, args.firstName, args.lastName, args.phones, args.addresses, args.organization, args.title, args.note, args.birthday);
                case "deleteContact":
                  return deleteContact(args.contactId);
                case "listCalendars":
                  return listCalendars();
                case "createEvent":
                  return await createEvent(args.title, args.startDate, args.endDate, args.location, args.description, args.calendarId, args.allDay, args.skipReview, args.status, args.showAs, args.categories, args.onlineMeeting);
                case "listEvents":
                  return await listEvents(args.calendarId, args.startDate, args.endDate, args.maxResults);
                case "updateEvent":
                  return await updateEvent(args.eventId, args.calendarId, args.title, args.startDate, args.endDate, args.location, args.description, args.status, args.showAs, args.categories, args.onlineMeeting);
                case "deleteEvent":
                  return await deleteEvent(args.eventId, args.calendarId);
                case "listCategories":
                  return listCategories();
                case "createTask":
                  return await createTask(args.title, args.dueDate, args.calendarId, args.description, args.priority, args.categories, args.skipReview);
                case "listTasks":
                  return await listTasks(args.calendarId, args.completed, args.dueBefore, args.maxResults);
                case "updateTask":
                  return await updateTask(args.taskId, args.calendarId, args.title, args.dueDate, args.description, args.completed, args.percentComplete, args.priority);
                case "sendMail":
                  return await composeMail(args.to, args.subject, args.body, args.cc, args.bcc, args.isHtml, args.from, args.attachments, args.skipReview);
                case "saveDraft":
                  return await saveDraft(args.to, args.subject, args.body, args.cc, args.bcc, args.isHtml, args.from, args.attachments);
                case "replyToMessage":
                  return await replyToMessage(args.messageId, args.folderPath, args.body, args.replyAll, args.isHtml, args.to, args.cc, args.bcc, args.from, args.attachments, args.skipReview);
                case "forwardMessage":
                  return await forwardMessage(args.messageId, args.folderPath, args.to, args.body, args.isHtml, args.cc, args.bcc, args.from, args.attachments, args.skipReview);
                case "getRecentMessages":
                  return getRecentMessages(args.folderPath, args.daysBack, args.maxResults, args.offset, args.unreadOnly, args.flaggedOnly, args.includeSubfolders);
                case "displayMessage":
                  return displayMessage(args.messageId, args.folderPath, args.displayMode);
                case "deleteMessages":
                  return deleteMessages(args.messageIds, args.folderPath);
                case "updateMessage":
                  return updateMessage(args.messageId, args.messageIds, args.folderPath, args.read, args.flagged, args.addTags, args.removeTags, args.moveTo, args.trash);
                case "createFolder":
                  return createFolder(args.parentFolderPath, args.name);
                case "renameFolder":
                  return renameFolder(args.folderPath, args.newName);
                case "deleteFolder":
                  return deleteFolder(args.folderPath);
                case "emptyTrash":
                  return emptyTrash(args.accountId);
                case "emptyJunk":
                  return emptyJunk(args.accountId);
                case "moveFolder":
                  return moveFolder(args.folderPath, args.newParentPath);
                case "listFilters":
                  return listFilters(args.accountId);
                case "createFilter":
                  return createFilter(args.accountId, args.name, args.enabled, args.type, args.conditions, args.actions, args.insertAtIndex);
                case "updateFilter":
                  return updateFilter(args.accountId, args.filterIndex, args.name, args.enabled, args.type, args.conditions, args.actions);
                case "deleteFilter":
                  return deleteFilter(args.accountId, args.filterIndex);
                case "reorderFilters":
                  return reorderFilters(args.accountId, args.fromIndex, args.toIndex);
                case "applyFilters":
                  return applyFilters(args.accountId, args.folderPath);
                case "getAccountAccess":
                  return getAccountAccess();
                default:
                  throw new Error(`Unknown tool: ${name}`);
              }
            }

            const server = new HttpServer();

            server.registerPathHandler("/", (req, res) => {
              res.processAsync();

              // Verify auth token on ALL requests (including non-POST) to
              // prevent unauthenticated probing of the server.
              let reqToken = "";
              try {
                reqToken = req.getHeader("Authorization") || "";
              } catch {
                // getHeader throws if header is missing in httpd.sys.mjs
              }
              if (!timingSafeEqual(reqToken, `Bearer ${authToken}`)) {
                res.setStatusLine("1.1", 403, "Forbidden");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Invalid or missing auth token" }
                }));
                res.finish();
                return;
              }

              if (req.method !== "POST") {
                res.setStatusLine("1.1", 405, "Method Not Allowed");
                res.setHeader("Allow", "POST", false);
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Method not allowed" }
                }));
                res.finish();
                return;
              }

              // Reject oversized request bodies to prevent memory exhaustion
              let contentLength = 0;
              try {
                contentLength = parseInt(req.getHeader("Content-Length"), 10) || 0;
              } catch {
                // Header missing — will be 0
              }
              if (contentLength > MAX_REQUEST_BODY) {
                res.setStatusLine("1.1", 413, "Payload Too Large");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Request body too large" }
                }));
                res.finish();
                return;
              }

              let message;
              try {
                message = JSON.parse(readRequestBody(req));
              } catch {
                res.setStatusLine("1.1", 200, "OK");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32700, message: "Parse error" }
                }));
                res.finish();
                return;
              }

              if (!message || typeof message !== "object" || Array.isArray(message)) {
                res.setStatusLine("1.1", 200, "OK");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Invalid Request" }
                }));
                res.finish();
                return;
              }

              const { id, method, params } = message;

              // Streamable HTTP notifications are accepted without a JSON-RPC body.
              // BEGIN MCP NOTIFICATION HTTP RESPONSE
              if (typeof method === "string" && method.startsWith("notifications/")) {
                res.setStatusLine("1.1", 202, "Accepted");
                res.finish();
                return;
              }
              // END MCP NOTIFICATION HTTP RESPONSE

              (async () => {
                try {
                  let result;
                  switch (method) {
                    case "initialize": {
                      // Per MCP lifecycle: respond with the requested version if
                      // we support it, otherwise the latest version we know.
                      // Behavior never depends on the version inside Thunderbird,
                      // so we accept any well-known protocol version.
                      const requested = params?.protocolVersion;
                      if (typeof requested !== "string") {
                        res.setStatusLine("1.1", 200, "OK");
                        res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                        res.write(JSON.stringify({
                          jsonrpc: "2.0",
                          id: id ?? null,
                          error: { code: -32602, message: "Invalid params: protocolVersion must be a string" }
                        }));
                        res.finish();
                        return;
                      }
                      const negotiated = MCP_SUPPORTED_PROTOCOL_VERSIONS.has(requested)
                        ? requested
                        : MCP_LATEST_PROTOCOL_VERSION;
                      result = {
                        protocolVersion: negotiated,
                        capabilities: { tools: {} },
                        serverInfo: { name: "thunderbird-mcp", version: getExtVersion() }
                      };
                      break;
                    }
                    case "resources/list":
                      result = { resources: [] };
                      break;
                    case "prompts/list":
                      result = { prompts: [] };
                      break;
                    case "tools/list":
                      // Strip internal metadata (group, crud, title) — only expose MCP-spec fields
                      result = { tools: buildTools().filter(t => isToolEnabled(t.name)).map(({ name, description, inputSchema }) => ({ name, description, inputSchema })) };
                      break;
                    case "tools/call":
                      if (!params?.name) {
                        throw new Error("Missing tool name");
                      }
                      if (!isToolEnabled(params.name)) {
                        throw new Error(`Tool is disabled: ${params.name}`);
                      }
                      {
                        const toolArgs = coerceToolArgs(params.name, params.arguments || {});
                        const validationErrors = validateToolArgs(params.name, toolArgs);
                        if (validationErrors.length > 0) {
                          throw new Error(`Invalid parameters for '${params.name}': ${validationErrors.join("; ")}`);
                        }
                        result = {
                          content: buildToolResultContent(await callTool(params.name, toolArgs))
                        };
                      }
                      break;
                    default:
                      res.setStatusLine("1.1", 200, "OK");
                      res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                      res.write(JSON.stringify({
                        jsonrpc: "2.0",
                        id: id ?? null,
                        error: { code: -32601, message: "Method not found" }
                      }));
                      res.finish();
                      return;
                  }
                  res.setStatusLine("1.1", 200, "OK");
                  // charset=utf-8 is critical for proper emoji handling in responses
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({ jsonrpc: "2.0", id: id ?? null, result }));
                } catch (e) {
                  res.setStatusLine("1.1", 200, "OK");
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({
                    jsonrpc: "2.0",
                    id: id ?? null,
                    error: { code: -32000, message: e.toString() }
                  }));
                }
                res.finish();
              })().catch((e) => {
                console.error("thunderbird-mcp: unhandled dispatch error:", e);
                try { res.finish(); } catch {}
              });
            });

            // Try the default port first, then fall back to nearby ports
            let boundPort = null;
            const listenAll = isListenAllEnabled();
            for (let attempt = 0; attempt < MCP_MAX_PORT_ATTEMPTS; attempt++) {
              const tryPort = MCP_DEFAULT_PORT + attempt;
              try {
                if (listenAll) {
                  server.startAll(tryPort);
                } else {
                  server.start(tryPort);
                }
                boundPort = tryPort;
                break;
              } catch (portErr) {
                if (attempt === MCP_MAX_PORT_ATTEMPTS - 1) {
                  throw new Error(`Could not bind to any port in range ${MCP_DEFAULT_PORT}-${tryPort}: ${portErr}`, { cause: portErr });
                }
                console.warn(`Port ${tryPort} in use, trying ${tryPort + 1}...`);
              }
            }

            globalThis.__tbMcpServer = server;
            let connFilePath;
            try {
              // Write the connection file fresh on initial start so the secure
              // create path (0600 perms, directory checks) always runs. The
              // refresh timer self-heals it afterward via ensureConnectionInfo
              // if the OS deletes it while the server keeps running.
              connFilePath = writeConnectionInfo(boundPort, authToken);
              startConnectionInfoRefresh(boundPort, authToken);
            } catch (writeErr) {
              // Connection file write failed -- stop the orphaned server
              try { server.stop(() => {}); } catch (e) { console.error("thunderbird-mcp: server.stop failed:", e); }
              globalThis.__tbMcpServer = null;
              stopConnectionInfoRefreshTimer();
              throw writeErr;
            }
            console.log(`Thunderbird MCP server listening on port ${boundPort}`);
            console.log(`Connection info written to ${connFilePath}`);
            if (listenAll) {
              console.error(`thunderbird-mcp: WARNING - server is listening on all interfaces (0.0.0.0/[::]). This exposes the MCP server to your local network. Only enable on trusted networks.`);
            }
            return { success: true, port: boundPort };
          } catch (e) {
            console.error("Failed to start MCP server:", e);
            // Stop server if it was started but something else failed
            if (globalThis.__tbMcpServer) {
              try { globalThis.__tbMcpServer.stop(() => {}); } catch (e) { console.error("thunderbird-mcp: server.stop failed:", e); }
              globalThis.__tbMcpServer = null;
            }
            stopConnectionInfoRefreshTimer();
            // Clear cached promise so a retry can attempt to bind again
            globalThis.__tbMcpStartPromise = null;
            removeConnectionInfo();
            return { success: false, error: e.toString() };
          }
          })();
          // Set sentinel BEFORE awaiting to prevent race with concurrent start() calls
          globalThis.__tbMcpStartPromise = startPromise;
          return await startPromise;
        },

        getServerInfo: async function() {
          let port = null;
          let connectionFile = null;
          let buildVersion = null;
          let buildDate = null;

          // Read build info from bundled file via resource: protocol
          try {
            const uri = Services.io.newURI("resource://thunderbird-mcp/buildinfo.json");
            const channel = Services.io.newChannelFromURI(uri, null,
              Services.scriptSecurityManager.getSystemPrincipal(), null,
              Ci.nsILoadInfo.SEC_ALLOW_CROSS_ORIGIN_SEC_CONTEXT_IS_NULL,
              Ci.nsIContentPolicy.TYPE_OTHER);
            const sis = Cc["@mozilla.org/scriptableinputstream;1"]
              .createInstance(Ci.nsIScriptableInputStream);
            sis.init(channel.open());
            const text = sis.read(sis.available());
            sis.close();
            const bi = JSON.parse(text);
            buildVersion = bi.version || bi.commit || null;
            buildDate = bi.builtAt || null;
          } catch (e) {
            // buildinfo.json may legitimately be absent in dev builds; surface anything else.
            if (e?.name !== "NS_ERROR_FILE_NOT_FOUND") {
              console.warn("thunderbird-mcp: read buildinfo failed:", e);
            }
          }

          // Read connection info from temp file using XPCOM file I/O
          try {
            const connInfo = readConnectionInfo();
            connectionFile = connInfo.path;
            if (connInfo.data) {
              port = connInfo.data.port || null;
            }
          } catch (e) {
            // Connection file is absent before the server first binds; log other faults.
            if (e?.name !== "NS_ERROR_FILE_NOT_FOUND") {
              console.warn("thunderbird-mcp: read connection info failed:", e);
            }
          }

          return {
            running: !!globalThis.__tbMcpStartPromise,
            port,
            connectionFile,
            buildVersion,
            buildDate,
          };
        },

        getCurrentAuthToken: async function() {
          let authToken = "";
          try {
            const connInfo = readConnectionInfo();
            if (connInfo.data && typeof connInfo.data.token === "string") {
              authToken = connInfo.data.token;
            }
          } catch (e) {
            if (e?.name !== "NS_ERROR_FILE_NOT_FOUND") {
              console.warn("thunderbird-mcp: read auth token failed:", e);
            }
          }
          return { authToken };
        },

        getAccountAccessConfig: async function() {
          const { MailServices } = ChromeUtils.importESModule(
            "resource:///modules/MailServices.sys.mjs"
          );
          let allowed = [];
          try {
            const pref = Services.prefs.getStringPref(PREF_ALLOWED_ACCOUNTS, "");
            if (pref) allowed = JSON.parse(pref);
          } catch (e) {
            // Falls back to "all accounts allowed"; surface the corruption.
            console.warn("thunderbird-mcp: account-access pref is not valid JSON:", e.message);
          }

          const accounts = [];
          for (const account of MailServices.accounts.accounts) {
            const server = account.incomingServer;
            accounts.push({
              id: account.key,
              name: server.prettyName,
              type: server.type,
              allowed: allowed.length === 0 || allowed.includes(account.key),
            });
          }
          return {
            mode: allowed.length === 0 ? "all" : "restricted",
            allowedAccountIds: allowed,
            accounts,
          };
        },

        getToolAccessConfig: async function() {
          // Use same fail-closed parsing as getDisabledTools() so the UI
          // accurately reflects the server's actual state on corrupt prefs
          let disabled = [];
          let corrupt = false;
          try {
            const pref = Services.prefs.getStringPref(PREF_DISABLED_TOOLS, "");
            if (pref) {
              const parsed = JSON.parse(pref);
              if (!Array.isArray(parsed)) {
                corrupt = true;
              } else {
                disabled = parsed;
              }
            }
          } catch {
            corrupt = true;
          }

          // Build tool list with group/crud metadata, sorted by group then CRUD order
          const getMessagesLimit = getConfiguredGetMessagesLimit();
          const toolList = buildTools()
            .map(t => ({
              name: t.name,
              group: t.group,
              crud: t.crud,
              enabled: corrupt ? UNDISABLEABLE_TOOLS.has(t.name) : !disabled.includes(t.name),
              undisableable: UNDISABLEABLE_TOOLS.has(t.name),
              ...(t.name === "getMessages" ? {
                getMessagesLimit,
                getMessagesLimitMin: 1,
                getMessagesLimitMax: MAX_GET_MESSAGES_LIMIT,
              } : {}),
            }))
            .sort((a, b) => {
              const gA = GROUP_ORDER[a.group] ?? 99;
              const gB = GROUP_ORDER[b.group] ?? 99;
              if (gA !== gB) return gA - gB;
              return (CRUD_ORDER[a.crud] ?? 99) - (CRUD_ORDER[b.crud] ?? 99);
            });
          const result = {
            mode: corrupt ? "error" : (disabled.length === 0 ? "all" : "restricted"),
            disabledTools: disabled,
            groups: GROUP_LABELS,
            getMessagesLimit,
            getMessagesLimitMin: 1,
            getMessagesLimitMax: MAX_GET_MESSAGES_LIMIT,
            tools: toolList,
          };
          if (corrupt) {
            result.error = "Disabled tools preference is corrupt. All non-infrastructure tools are blocked. Save to reset.";
          }
          return result;
        },

        setToolAccess: async function(disabledTools, getMessagesLimit) {
          if (!Array.isArray(disabledTools)) {
            return { error: "disabledTools must be an array" };
          }
          // Validate types first, then semantic checks
          if (!disabledTools.every(t => typeof t === "string")) {
            return { error: "All tool names must be strings" };
          }
          // Reject internal sentinel values
          if (disabledTools.includes("__all__")) {
            return { error: "Invalid tool name: __all__" };
          }
          // Validate: can't disable undisableable tools
          const blocked = disabledTools.filter(t => UNDISABLEABLE_TOOLS.has(t));
          if (blocked.length > 0) {
            return { error: `Cannot disable infrastructure tools: ${blocked.join(", ")}` };
          }

          let requestedLimit = null;
          if (getMessagesLimit !== undefined && getMessagesLimit !== null && getMessagesLimit !== "") {
            requestedLimit = Number(getMessagesLimit);
            if (!Number.isInteger(requestedLimit)) {
              return { error: "getMessagesLimit must be an integer" };
            }
            if (requestedLimit < 1 || requestedLimit > MAX_GET_MESSAGES_LIMIT) {
              return { error: `getMessagesLimit must be between 1 and ${MAX_GET_MESSAGES_LIMIT}` };
            }
          }

          if (disabledTools.length === 0) {
            try { Services.prefs.clearUserPref(PREF_DISABLED_TOOLS); } catch { /* ignore */ }
          } else {
            Services.prefs.setStringPref(PREF_DISABLED_TOOLS, JSON.stringify(disabledTools));
          }
          if (requestedLimit !== null) {
            if (requestedLimit === DEFAULT_GET_MESSAGES_LIMIT) {
              try { Services.prefs.clearUserPref(PREF_GET_MESSAGES_LIMIT); } catch { /* ignore */ }
            } else {
              Services.prefs.setIntPref(PREF_GET_MESSAGES_LIMIT, requestedLimit);
            }
          }
          return {
            success: true,
            mode: disabledTools.length === 0 ? "all" : "restricted",
            disabledTools,
            getMessagesLimit: getConfiguredGetMessagesLimit(),
          };
        },

        setAccountAccess: async function(allowedAccountIds) {
          if (!Array.isArray(allowedAccountIds)) {
            return { error: "allowedAccountIds must be an array" };
          }
          const { MailServices } = ChromeUtils.importESModule(
            "resource:///modules/MailServices.sys.mjs"
          );
          const validIds = new Set();
          for (const account of MailServices.accounts.accounts) {
            validIds.add(account.key);
          }
          const invalid = allowedAccountIds.filter(id => !validIds.has(id));
          if (invalid.length > 0) {
            return { error: `Unknown account IDs: ${invalid.join(", ")}` };
          }

          if (allowedAccountIds.length === 0) {
            try { Services.prefs.clearUserPref(PREF_ALLOWED_ACCOUNTS); } catch { /* ignore */ }
          } else {
            Services.prefs.setStringPref(PREF_ALLOWED_ACCOUNTS, JSON.stringify(allowedAccountIds));
          }
          return {
            success: true,
            mode: allowedAccountIds.length === 0 ? "all" : "restricted",
            allowedAccountIds,
          };
        },

        getBlockSkipReview: async function() {
          let blocked = true;
          try {
            blocked = Services.prefs.getBoolPref(PREF_BLOCK_SKIPREVIEW, true);
          } catch { /* ignore */ }
          return { blockSkipReview: blocked };
        },

        setBlockSkipReview: async function(blockSkipReview) {
          if (typeof blockSkipReview !== "boolean") {
            return { error: "blockSkipReview must be a boolean" };
          }
          // Default is true; persist the explicit value either way so the user's
          // choice survives independent of the default we ship.
          Services.prefs.setBoolPref(PREF_BLOCK_SKIPREVIEW, blockSkipReview);
          return { success: true, blockSkipReview };
        },

        getStableAuthToken: async function() {
          return { stableAuthToken: getStableAuthTokenPref() };
        },

        setStableAuthToken: async function(stableAuthToken) {
          if (typeof stableAuthToken !== "string") {
            return { error: "stableAuthToken must be a string" };
          }
          stableAuthToken = stableAuthToken.trim();
          if (stableAuthToken && !AUTH_TOKEN_PATTERN.test(stableAuthToken)) {
            return { error: "stableAuthToken must be empty or 64 lowercase hex characters" };
          }
          if (stableAuthToken) {
            Services.prefs.setStringPref(PREF_STABLE_AUTH_TOKEN, stableAuthToken);
          } else {
            try { Services.prefs.clearUserPref(PREF_STABLE_AUTH_TOKEN); } catch { /* ignore */ }
          }
          return { success: true, stableAuthToken };
        },

        generateAuthToken: async function() {
          return { authToken: generateAuthToken() };
        },

        getListenAll: async function() {
          let listenAll = false;
          try {
            listenAll = Services.prefs.getBoolPref(PREF_LISTEN_ALL, false);
          } catch { /* ignore */ }
          return { listenAll };
        },

        setListenAll: async function(listenAll) {
          if (typeof listenAll !== "boolean") {
            return { error: "listenAll must be a boolean" };
          }
          console.log(`[MCP] setListenAll called with: ${listenAll}`);
          if (listenAll) {
            Services.prefs.setBoolPref(PREF_LISTEN_ALL, true);
          } else {
            try { Services.prefs.clearUserPref(PREF_LISTEN_ALL); } catch { /* ignore */ }
          }

          // Stop existing server
          if (globalThis.__tbMcpServer) {
            const stopServer = globalThis.__tbMcpServer;
            globalThis.__tbMcpServer = null;
            stopConnectionInfoRefreshTimer();
            // Wait for the socket close callback before rebinding the port.
            try {
              await new Promise((resolve) => {
                try { stopServer.stop(resolve); } catch { resolve(); }
              });
            } catch { /* ignore */ }
          }
          // Clear sentinels so start() can reinitialize
          globalThis.__tbMcpStartPromise = null;

          // Remove stale connection file
          try {
            const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
            tmpDir.append("thunderbird-mcp");
            const connFile = tmpDir.clone();
            connFile.append("connection.json");
            if (connFile.exists()) connFile.remove(false);
          } catch { /* best-effort cleanup */ }

          // Restart server with new binding
          console.log(`[MCP] Restarting server...`);
          return await this.start();
        },
      }
    };
  }

  onShutdown(isAppShutdown) {
    // Stop the HTTP server so the port is released
    if (globalThis.__tbMcpServer) {
      try { globalThis.__tbMcpServer.stop(() => {}); } catch { /* ignore */ }
      globalThis.__tbMcpServer = null;
    }
    stopConnectionInfoRefreshTimer();
    // Clear the start promise so a fresh start can occur on reload
    globalThis.__tbMcpStartPromise = null;

    // Always clean up the connection info file so stale tokens don't linger
    // (Inlined because getAPI() helpers are not in scope in onShutdown().)
    try {
      const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
      tmpDir.append("thunderbird-mcp");
      const connFile = tmpDir.clone();
      connFile.append("connection.json");
      if (connFile.exists()) {
        connFile.remove(false);
      }
    } catch {
      // Best-effort cleanup
    }

    // Always clean up temp attachment files (even on app shutdown) to avoid
    // leaving sensitive decoded attachments on disk.
    for (const tmpPath of _tempAttachFiles) {
      try {
        const f = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
        f.initWithPath(tmpPath);
        if (f.exists()) f.remove(false);
      } catch {}
    }
    _tempAttachFiles.clear();
    if (isAppShutdown) return;
    resProto.setSubstitution("thunderbird-mcp", null);
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }
};
