"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const apiSource = fs.readFileSync(
  path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8"
);
const manifest = JSON.parse(fs.readFileSync(
  path.resolve(__dirname, "../extension/manifest.json"), "utf8"
));

// Use the marker assertions and VM loading pattern from validation.test.cjs.
function getMarkedApiSnippet(startMarker, endMarker) {
  const start = apiSource.indexOf(startMarker);
  const end = apiSource.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0, `api.js marker missing: ${startMarker}`);
  assert.ok(end > start, `api.js marker missing: ${endMarker}`);
  return apiSource.slice(start, end);
}

function getConstantDeclaration(name) {
  const declaration = apiSource.match(new RegExp(`^\\s*const ${name} = [^\\n]+;`, "m"));
  assert.ok(declaration, `api.js constant missing: ${name}`);
  return declaration[0];
}

function loadPrivacyRuntime({ initial = {}, unreadable = [], legacyAddonManager = false } = {}) {
  const values = new Map();
  const readErrors = new Set();
  const writes = [];
  const calls = [];
  const handlerResults = new Map();
  const addonListeners = new Set();
  const addedListeners = [];
  const removedListeners = [];
  const warnings = [];
  const prefTypes = { PREF_INVALID: 0, PREF_STRING: 32, PREF_INT: 64, PREF_BOOL: 128 };
  const prefs = {
    getPrefType(name) {
      if (readErrors.has(name)) throw new Error("Preference is unreadable");
      if (!values.has(name)) return prefTypes.PREF_INVALID;
      const value = values.get(name);
      if (typeof value === "string") return prefTypes.PREF_STRING;
      if (typeof value === "boolean") return prefTypes.PREF_BOOL;
      if (Number.isInteger(value)) return prefTypes.PREF_INT;
      throw new Error("Unsupported preference value");
    },
    getStringPref(name, fallback) {
      if (readErrors.has(name)) throw new Error("Preference is unreadable");
      const value = values.get(name);
      if (typeof value === "string") return value;
      if (arguments.length > 1) return fallback;
      throw new Error("Preference is not a string");
    },
    getBoolPref(name, fallback) {
      if (readErrors.has(name)) throw new Error("Preference is unreadable");
      const value = values.get(name);
      if (typeof value === "boolean") return value;
      if (arguments.length > 1) return fallback;
      throw new Error("Preference is not a boolean");
    },
    getIntPref(name, fallback) {
      if (readErrors.has(name)) throw new Error("Preference is unreadable");
      const value = values.get(name);
      if (Number.isInteger(value)) return value;
      if (arguments.length > 1) return fallback;
      throw new Error("Preference is not an integer");
    },
    setStringPref(name, value) {
      writes.push({ name, value });
      values.set(name, value);
    },
    setBoolPref(name, value) {
      writes.push({ name, value });
      values.set(name, value);
    },
    setIntPref(name, value) {
      writes.push({ name, value });
      values.set(name, value);
    },
    clearUserPref(name) {
      writes.push({ name });
      values.delete(name);
    },
  };
  const AddonManager = {
    addAddonListener(listener) {
      addedListeners.push(listener);
      addonListeners.add(listener);
    },
    removeAddonListener(listener) {
      removedListeners.push(listener);
      addonListeners.delete(listener);
    },
  };
  const accounts = [
    { key: "account1", incomingServer: { prettyName: "First account", type: "imap" } },
    { key: "account2", incomingServer: { prettyName: "Second account", type: "pop3" } },
  ];
  const sandbox = {
    Services: { prefs },
    Ci: { nsIPrefBranch: prefTypes },
    ChromeUtils: {
      importESModule(uri) {
        if (uri === "resource:///modules/MailServices.sys.mjs") {
          return { MailServices: { accounts: { accounts } } };
        }
        assert.equal(uri, "resource://gre/modules/AddonManager.sys.mjs");
        if (legacyAddonManager) throw new Error("ES module unavailable");
        return { AddonManager };
      },
      import(uri) {
        assert.equal(uri, "resource://gre/modules/AddonManager.jsm");
        return { AddonManager };
      },
    },
    console: { warn: (...args) => warnings.push(args), error() {} },
  };
  const prefConstants = [
    "PREF_ALLOWED_ACCOUNTS", "PREF_DISABLED_TOOLS", "PREF_BLOCK_SKIPREVIEW",
    "PREF_STABLE_AUTH_TOKEN", "PREF_GET_MESSAGES_LIMIT", "PREF_ALLOW_ENCRYPTED_MESSAGES",
    "PREF_ALLOW_ALL_CALENDARS", "PREF_ALLOW_ALL_ADDRESS_BOOKS",
  ];
  const constants = [
    ...prefConstants, "DEFAULT_GET_MESSAGES_LIMIT", "MAX_GET_MESSAGES_LIMIT",
    "UNDISABLEABLE_TOOLS", "GROUP_ORDER", "GROUP_LABELS", "CRUD_ORDER",
  ];
  vm.createContext(sandbox);
  vm.runInContext([
    ...constants.map(getConstantDeclaration),
    getMarkedApiSnippet("// BEGIN CONTACT FIELD CONSTANTS", "// END CONTACT FIELD CONSTANTS"),
    getMarkedApiSnippet("// BEGIN OUTBOUND ATTACHMENT LIMITS", "// END OUTBOUND ATTACHMENT LIMITS"),
    getMarkedApiSnippet("function normalizeGetMessagesLimit(", "// BEGIN TOOL SCHEMA BUILDER"),
    getMarkedApiSnippet("// BEGIN PRIVACY PREFERENCE HELPERS", "// END PRIVACY PREFERENCE HELPERS"),
    getMarkedApiSnippet("// BEGIN SERVER ACCESS HELPERS", "// END SERVER ACCESS HELPERS"),
    getMarkedApiSnippet("// BEGIN MCP TEXT SANITIZATION", "// END MCP TEXT SANITIZATION"),
    getMarkedApiSnippet("// BEGIN FILTER SEARCH TERM HELPERS", "// END FILTER SEARCH TERM HELPERS"),
    getMarkedApiSnippet("// BEGIN TOOL SCHEMA BUILDER", "// END TOOL SCHEMA BUILDER"),
    getMarkedApiSnippet("// BEGIN TOOL DISPATCH", "// END TOOL DISPATCH"),
    "this.options = {",
    getMarkedApiSnippet("// BEGIN OPTIONS ACCESS API", "// END OPTIONS ACCESS API"),
    getMarkedApiSnippet("// BEGIN PRIVACY OPTIONS API", "// END PRIVACY OPTIONS API"),
    "};",
    `this.prefNames = { ${prefConstants.join(", ")} };`,
    "this.tools = buildTools();",
    "this.readAccessListPref = readAccessListPref;",
    "this.isAccountAllowed = isAccountAllowed;",
    "this.isToolEnabled = isToolEnabled;",
    "this.isPrivacyOptInEnabled = isPrivacyOptInEnabled;",
    "this.callTool = callTool;",
    "this.stripInvisibleCharacters = stripInvisibleCharacters;",
    "this.sanitizeToolResultText = sanitizeToolResultText;",
    "this.registerUninstallListener = function(context) {",
    getMarkedApiSnippet("// BEGIN UNINSTALL LISTENER REGISTRATION", "// END UNINSTALL LISTENER REGISTRATION"),
    "};",
    "this.removeUninstallListener = function() {",
    getMarkedApiSnippet("// BEGIN UNINSTALL LISTENER REMOVAL", "// END UNINSTALL LISTENER REMOVAL"),
    "};",
  ].join("\n"), sandbox);

  for (const [name, value] of Object.entries(initial)) {
    assert.ok(sandbox.prefNames[name], `Unknown preference constant: ${name}`);
    values.set(sandbox.prefNames[name], value);
  }
  for (const name of unreadable) {
    assert.ok(sandbox.prefNames[name], `Unknown preference constant: ${name}`);
    readErrors.add(sandbox.prefNames[name]);
  }
  for (const tool of sandbox.tools) {
    const handler = tool.name === "sendMail" ? "composeMail" : tool.name;
    sandbox[handler] = (...args) => {
      calls.push({ name: tool.name, args });
      return handlerResults.has(tool.name)
        ? handlerResults.get(tool.name) : { handled: tool.name };
    };
  }
  return {
    ...sandbox, values, writes, calls, handlerResults, addonListeners,
    addedListeners, removedListeners, warnings,
  };
}

describe("Access preferences use the same production parser in options and server", () => {
  for (const raw of ["{", "null", "{}", "true", '"account1"', "[1]", '["account1",false]', "[{}]", true, false, 0, 1, -1]) {
    it(`blocks invalid account and tool configuration (${typeof raw}): ${JSON.stringify(raw)}`, async () => {
      const runtime = loadPrivacyRuntime({ initial: {
        PREF_ALLOWED_ACCOUNTS: raw, PREF_DISABLED_TOOLS: raw,
      } });
      const accountConfig = await runtime.options.getAccountAccessConfig();
      assert.equal(accountConfig.mode, "error");
      assert.match(accountConfig.error, /blocked/i);
      for (const account of accountConfig.accounts) {
        assert.equal(account.allowed, false);
        assert.equal(runtime.isAccountAllowed(account.id), account.allowed);
      }
      const toolConfig = await runtime.options.getToolAccessConfig();
      assert.equal(toolConfig.mode, "error");
      assert.match(toolConfig.error, /blocked/i);
      for (const tool of toolConfig.tools) {
        assert.equal(tool.enabled, tool.undisableable);
        assert.equal(runtime.isToolEnabled(tool.name), tool.enabled);
      }
      assert.equal(runtime.writes.length, 0, "reading corrupt settings must not repair or clear them");
      assert.equal(runtime.values.get(runtime.prefNames.PREF_ALLOWED_ACCOUNTS), raw);
      assert.equal(runtime.values.get(runtime.prefNames.PREF_DISABLED_TOOLS), raw);
    });
  }

  it("blocks unreadable account and tool preferences in both APIs", async () => {
    const runtime = loadPrivacyRuntime({ unreadable: ["PREF_ALLOWED_ACCOUNTS", "PREF_DISABLED_TOOLS"] });
    const accounts = await runtime.options.getAccountAccessConfig();
    const tools = await runtime.options.getToolAccessConfig();
    assert.equal(accounts.mode, "error");
    assert.equal(tools.mode, "error");
    assert.ok(accounts.accounts.every(account => !account.allowed && !runtime.isAccountAllowed(account.id)));
    assert.ok(tools.tools.every(tool => tool.enabled === runtime.isToolEnabled(tool.name)));
    assert.equal(runtime.writes.length, 0);
  });

  for (const raw of [undefined, "", "[]"]) {
    it(`keeps unrestricted behavior for ${String(raw)}`, async () => {
      const runtime = loadPrivacyRuntime({ initial: raw === undefined ? {} : {
        PREF_ALLOWED_ACCOUNTS: raw, PREF_DISABLED_TOOLS: raw,
      } });
      const accounts = await runtime.options.getAccountAccessConfig();
      const tools = await runtime.options.getToolAccessConfig();
      assert.equal(accounts.mode, "all");
      assert.equal(tools.mode, "all");
      assert.equal(tools.getMessagesLimit, 10);
      assert.ok(accounts.accounts.every(account => account.allowed && runtime.isAccountAllowed(account.id)));
      assert.ok(tools.tools.every(tool => tool.enabled && runtime.isToolEnabled(tool.name)));
      if (raw === undefined) {
        for (const name of ["PREF_ALLOWED_ACCOUNTS", "PREF_DISABLED_TOOLS"]) {
          assert.equal(runtime.Services.prefs.getPrefType(runtime.prefNames[name]), runtime.Ci.nsIPrefBranch.PREF_INVALID);
          assert.equal(runtime.values.has(runtime.prefNames[name]), false);
        }
        assert.equal(runtime.writes.length, 0, "missing access preferences retain the unrestricted defaults without writes");
      }
    });
  }

  it("matches explicit account grants and tool exclusions", async () => {
    const runtime = loadPrivacyRuntime({ initial: {
      PREF_ALLOWED_ACCOUNTS: '["account1"]',
      PREF_DISABLED_TOOLS: '["getMessage","deleteContact"]',
    } });
    const accounts = await runtime.options.getAccountAccessConfig();
    const tools = await runtime.options.getToolAccessConfig();
    assert.equal(accounts.mode, "restricted");
    assert.equal(tools.mode, "restricted");
    assert.equal(runtime.isAccountAllowed("account1"), true);
    assert.equal(runtime.isAccountAllowed("account2"), false);
    for (const account of accounts.accounts) {
      assert.equal(account.allowed, runtime.isAccountAllowed(account.id));
    }
    for (const tool of tools.tools) {
      assert.equal(tool.enabled, runtime.isToolEnabled(tool.name));
    }
    assert.equal(runtime.isToolEnabled("getMessage"), false);
    assert.equal(runtime.isToolEnabled("deleteContact"), false);
    assert.equal(runtime.isToolEnabled("listCalendars"), true);
  });

  it("shows the internal all-tools sentinel as blocked in options too", async () => {
    const runtime = loadPrivacyRuntime({ initial: { PREF_DISABLED_TOOLS: '["__all__"]' } });
    const config = await runtime.options.getToolAccessConfig();
    assert.equal(config.mode, "error");
    assert.ok(config.tools.every(tool => tool.enabled === runtime.isToolEnabled(tool.name)));
    assert.ok(config.tools.some(tool => !tool.enabled));
  });

  it("repairs corrupt preferences only through explicit valid selections", async () => {
    const runtime = loadPrivacyRuntime({ initial: {
      PREF_ALLOWED_ACCOUNTS: "{", PREF_DISABLED_TOOLS: "{",
    } });
    assert.ok((await runtime.options.setAccountAccess([false])).error);
    assert.ok((await runtime.options.setToolAccess([false])).error);
    assert.equal(runtime.writes.length, 0);
    assert.equal((await runtime.options.setAccountAccess(["account1"])).success, true);
    assert.equal((await runtime.options.setToolAccess(["getMessage"])).success, true);
    assert.equal((await runtime.options.getAccountAccessConfig()).mode, "restricted");
    assert.equal((await runtime.options.getToolAccessConfig()).mode, "restricted");
    assert.equal(runtime.isAccountAllowed("account2"), false);
    assert.equal(runtime.isToolEnabled("getMessage"), false);
  });
});

describe("Calendar and address-book access at the production tool dispatcher", () => {
  const directArgs = {
    calendarId: "private-calendar", eventId: "private-event", taskId: "private-task",
    addressBookId: "private-book", contactId: "private-contact", query: "contact",
    title: "New item", email: "contact@example.com", skipReview: true,
  };

  for (const raw of ['["account1"]', '["account1","account2"]', "{", '["account1",null]']) {
    it(`blocks every resource tool before its handler runs with restriction ${raw}`, async () => {
      const runtime = loadPrivacyRuntime({ initial: { PREF_ALLOWED_ACCOUNTS: raw } });
      const resources = runtime.tools.filter(tool => ["calendar", "contacts"].includes(tool.group));
      for (const name of ["listCalendars", "getContact", "createContact", "updateContact", "deleteContact", "createEvent", "updateEvent", "deleteEvent", "createTask", "updateTask"]) {
        assert.ok(resources.some(tool => tool.name === name), `Missing resource tool coverage: ${name}`);
      }
      for (const tool of resources) {
        const result = await runtime.callTool(tool.name, directArgs);
        assert.match(result.error, /Account restrictions block/, tool.name);
        assert.match(result.error, tool.group === "calendar" ? /Allow all calendars/ : /Allow all address books/);
      }
      assert.equal(runtime.calls.length, 0, "no listing, direct lookup or write may reach a handler");
    });
  }

  for (const group of ["calendar", "contacts"]) {
    it(`allows only ${group} after its options checkbox is enabled`, async () => {
      const runtime = loadPrivacyRuntime({ initial: { PREF_ALLOWED_ACCOUNTS: '["account1"]' } });
      await runtime.options.setPrivacySettings(false, group === "calendar", group === "contacts");
      const resources = runtime.tools.filter(tool => ["calendar", "contacts"].includes(tool.group));
      for (const tool of resources) {
        const result = await runtime.callTool(tool.name, directArgs);
        if (tool.group === group) assert.equal(result.handled, tool.name);
        else assert.match(result.error, /Account restrictions block/);
      }
      assert.equal(runtime.calls.length, resources.filter(tool => tool.group === group).length);
      assert.ok(runtime.calls.some(call => call.args.includes(
        group === "calendar" ? directArgs.calendarId : directArgs.contactId
      )), "direct resource IDs reach handlers only after the grant");
    });
  }

  for (const raw of [undefined, "", "[]"]) {
    it(`allows all resource tools without configured restrictions: ${String(raw)}`, async () => {
      const runtime = loadPrivacyRuntime({ initial: raw === undefined ? {} : { PREF_ALLOWED_ACCOUNTS: raw } });
      const resources = runtime.tools.filter(tool => ["calendar", "contacts"].includes(tool.group));
      for (const tool of resources) {
        assert.equal((await runtime.callTool(tool.name, directArgs)).handled, tool.name);
      }
      assert.equal(runtime.calls.length, resources.length);
    });
  }

  it("fails closed when resource grants or account configuration cannot be read", async () => {
    for (const unreadable of [
      ["PREF_ALLOWED_ACCOUNTS"],
      ["PREF_ALLOW_ALL_CALENDARS", "PREF_ALLOW_ALL_ADDRESS_BOOKS"],
    ]) {
      const runtime = loadPrivacyRuntime({
        initial: { PREF_ALLOWED_ACCOUNTS: '["account1"]' }, unreadable,
      });
      const resources = runtime.tools.filter(tool => ["calendar", "contacts"].includes(tool.group));
      for (const tool of resources) {
        assert.match((await runtime.callTool(tool.name, directArgs)).error, /Account restrictions block/);
      }
      assert.equal(runtime.calls.length, 0);
    }
  });
});

describe("Options-only privacy preferences", () => {
  const preferenceNames = [
    "PREF_ALLOW_ENCRYPTED_MESSAGES", "PREF_ALLOW_ALL_CALENDARS", "PREF_ALLOW_ALL_ADDRESS_BOOKS",
  ];

  it("defaults every opt-in to false, including unreadable preferences", async () => {
    for (const unreadable of [[], preferenceNames]) {
      const runtime = loadPrivacyRuntime({ unreadable });
      const settings = await runtime.options.getPrivacySettings();
      assert.deepEqual(JSON.parse(JSON.stringify(settings)), {
        allowEncryptedMessages: false, allowAllCalendars: false, allowAllAddressBooks: false,
      });
      for (const name of preferenceNames) {
        assert.equal(runtime.isPrivacyOptInEnabled(runtime.prefNames[name]), false);
      }
      assert.equal(runtime.writes.length, 0);
    }
  });

  it("persists explicit booleans and supports revoking each opt-in", async () => {
    const runtime = loadPrivacyRuntime();
    assert.equal((await runtime.options.setPrivacySettings(true, true, true)).success, true);
    assert.ok(Object.values(await runtime.options.getPrivacySettings()).every(value => value === true));
    for (const name of preferenceNames) {
      assert.equal(runtime.values.get(runtime.prefNames[name]), true);
    }
    assert.equal((await runtime.options.setPrivacySettings(false, false, false)).success, true);
    assert.ok(Object.values(await runtime.options.getPrivacySettings()).every(value => value === false));
    assert.equal(runtime.writes.length, 6);
    assert.ok(runtime.tools.every(tool => !/PrivacySettings/.test(tool.name)), "MCP tools cannot grant privacy permissions");
  });

  it("rejects non-booleans without partially changing settings", async () => {
    const runtime = loadPrivacyRuntime();
    for (const invalid of [undefined, null, "true", 1, {}, []]) {
      for (let index = 0; index < 3; index++) {
        const args = [true, true, true];
        args[index] = invalid;
        assert.match((await runtime.options.setPrivacySettings(...args)).error, /booleans/);
      }
    }
    assert.equal(runtime.writes.length, 0);
  });
});

describe("Text returned by the production dispatcher", () => {
  it("strips every tag, bidi embedding/override/isolate, and selected invisible character", () => {
    const runtime = loadPrivacyRuntime();
    const ranges = [
      [0xE0000, 0xE007F], [0x202A, 0x202E], [0x2066, 0x2069],
      [0x200B, 0x200B], [0x2060, 0x2060], [0xFEFF, 0xFEFF],
    ];
    for (const [first, last] of ranges) {
      for (let point = first; point <= last; point++) {
        const input = `A${String.fromCodePoint(point)}B`;
        assert.equal(runtime.stripInvisibleCharacters(input), "AB", `U+${point.toString(16).toUpperCase()}`);
        assert.equal(runtime.sanitizeToolResultText({ body: input }).body, "AB");
      }
    }
  });

  it("keeps ZWJ, ZWNJ, emoji sequences, and normal script text", () => {
    const runtime = loadPrivacyRuntime();
    const text = "A\u200CB\u200DC 👩‍👩‍👧‍👦 می‌خواهم";
    assert.equal(runtime.stripInvisibleCharacters(text), text);
    assert.equal(runtime.sanitizeToolResultText({ subject: text }).subject, text);
  });

  it("cleans nested bodies, previews, subjects, and names while preserving JSON shape and identifiers", async () => {
    const runtime = loadPrivacyRuntime();
    const hidden = "vis\u202Eible\u200B\u{E0061}";
    const original = {
      messages: [{
        id: hidden, messageId: hidden, folderPath: hidden, url: hidden,
        subject: hidden, author: hidden, recipients: hidden, ccList: hidden,
        body: hidden, bodyIsHtml: false, preview: hidden,
        attachments: [{ name: hidden, contentId: hidden, partName: hidden, base64: hidden }],
      }],
      contacts: [{ name: hidden, displayName: hidden, firstName: hidden, lastName: hidden,
        title: hidden, description: hidden, note: hidden, organization: hidden, addressBook: hidden }],
      message: hidden, error: hidden, count: 1, empty: null, active: true,
    };
    const expected = JSON.parse(JSON.stringify(original));
    runtime.handlerResults.set("getMessages", original);
    const result = await runtime.callTool("getMessages", { messages: [] });
    const clean = JSON.parse(JSON.stringify(result));
    for (const key of ["subject", "author", "recipients", "ccList", "body", "preview"]) {
      expected.messages[0][key] = "visible";
    }
    expected.messages[0].attachments[0].name = "visible";
    for (const key of Object.keys(expected.contacts[0])) expected.contacts[0][key] = "visible";
    expected.message = "visible";
    expected.error = "visible";
    assert.deepEqual(clean, expected);
  });

  it("preserves explicit HTML and raw source while cleaning surrounding display text", async () => {
    const runtime = loadPrivacyRuntime();
    const raw = "Subject: A\u200BB\r\n\r\nraw\u202Econtent";
    const html = "<p>raw\u200Bhtml\u{E0061}</p>";
    runtime.handlerResults.set("getMessage", {
      subject: "A\u200BB", body: html, bodyIsHtml: true, rawSource: raw,
      attachments: [{ name: "f\u2060ile", url: "cid:x\u2060y" }],
    });
    const result = await runtime.callTool("getMessage", {});
    assert.equal(result.subject, "AB");
    assert.equal(result.body, html);
    assert.equal(result.rawSource, raw);
    assert.equal(result.attachments[0].name, "file");
    assert.equal(result.attachments[0].url, "cid:x\u2060y");
  });

  it("cleans calendar/account/folder names, category lists, and columnar names while preserving paths", async () => {
    const runtime = loadPrivacyRuntime();
    const hidden = "na\u202Eme\u{E0041}";
    const names = runtime.sanitizeToolResultText({ calendarName: hidden, accountName: hidden, folder: hidden });
    assert.deepEqual(JSON.parse(JSON.stringify(names)), { calendarName: "name", accountName: "name", folder: "name" });
    runtime.handlerResults.set("listFolders", { columns: ["name", "path", "accountId"], rows: [[hidden, hidden, hidden]] });
    const result = await runtime.callTool("listFolders", { format: "table" });
    assert.deepEqual(JSON.parse(JSON.stringify(result.rows)), [["name", hidden, hidden]]);
    runtime.handlerResults.set("listCategories", [hidden]);
    assert.deepEqual(JSON.parse(JSON.stringify(await runtime.callTool("listCategories", {}))), ["name"]);
  });

  it("retains non-enumerable metadata used for extra MCP content blocks", async () => {
    const runtime = loadPrivacyRuntime();
    const metadataKey = Symbol("extraMcpContent");
    const blocks = [{ type: "image", data: "aW1hZ2U=", mimeType: "image/png" }];
    const message = { subject: "A\u200BB", body: "C\u2060D", bodyIsHtml: false };
    Object.defineProperty(message, metadataKey, { value: blocks, enumerable: false });
    runtime.handlerResults.set("getMessage", message);
    const result = await runtime.callTool("getMessage", {});
    assert.equal(result.subject, "AB");
    assert.equal(result.body, "CD");
    assert.equal(result[metadataKey], blocks);
    assert.equal(Object.getOwnPropertyDescriptor(result, metadataKey).enumerable, false);
    assert.equal(JSON.stringify(result).includes("aW1hZ2U="), false);
  });
});

describe("Stable-token cleanup during enabled uninstall", () => {
  for (const legacyAddonManager of [false, true]) {
    it(`clears only this add-on's stable token (legacy import: ${legacyAddonManager})`, () => {
      const runtime = loadPrivacyRuntime({ legacyAddonManager, initial: {
        PREF_STABLE_AUTH_TOKEN: "a".repeat(64),
        PREF_ALLOWED_ACCOUNTS: '["account1"]',
        PREF_ALLOW_ENCRYPTED_MESSAGES: true,
      } });
      const owner = {};
      const id = manifest.browser_specific_settings.gecko.id;
      runtime.registerUninstallListener.call(owner, { extension: { id } });
      runtime.registerUninstallListener.call(owner, { extension: { id } });
      assert.equal(runtime.addedListeners.length, 1, "repeated getAPI calls must reuse the listener");
      const listener = runtime.addedListeners[0];
      listener.onUninstalling({ id: "another-addon@example.com" });
      assert.equal(runtime.values.get(runtime.prefNames.PREF_STABLE_AUTH_TOKEN), "a".repeat(64));
      assert.equal(runtime.writes.length, 0);
      listener.onUninstalling({ id });
      assert.equal(runtime.values.has(runtime.prefNames.PREF_STABLE_AUTH_TOKEN), false);
      assert.equal(runtime.values.get(runtime.prefNames.PREF_ALLOWED_ACCOUNTS), '["account1"]');
      assert.equal(runtime.values.get(runtime.prefNames.PREF_ALLOW_ENCRYPTED_MESSAGES), true);
      assert.deepEqual(runtime.writes, [{ name: runtime.prefNames.PREF_STABLE_AUTH_TOKEN }]);
    });
  }

  it("removes the registered listener on shutdown without clearing the token", () => {
    const runtime = loadPrivacyRuntime({ initial: { PREF_STABLE_AUTH_TOKEN: "b".repeat(64) } });
    const owner = {};
    runtime.registerUninstallListener.call(owner, { extension: { id: manifest.browser_specific_settings.gecko.id } });
    const listener = runtime.addedListeners[0];
    runtime.removeUninstallListener.call(owner);
    runtime.removeUninstallListener.call(owner);
    assert.deepEqual(runtime.removedListeners, [listener]);
    assert.equal(runtime.addonListeners.size, 0);
    assert.equal(owner._uninstallListener, null);
    assert.equal(owner._addonManager, null);
    assert.equal(runtime.values.get(runtime.prefNames.PREF_STABLE_AUTH_TOKEN), "b".repeat(64));
    assert.equal(runtime.writes.length, 0);
  });
});
