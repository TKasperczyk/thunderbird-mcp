"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { loadContactRuntime } = require("./helpers/contact-runtime.cjs");
const { loadPrivacyRuntime } = require("./helpers/privacy-access.cjs");

const optionsSource = fs.readFileSync(path.resolve(__dirname, "../extension/options.js"), "utf8");

function markedOptionsSnippet(name) {
  const startMarker = `// BEGIN ${name}`;
  const endMarker = `// END ${name}`;
  const start = optionsSource.indexOf(startMarker);
  const end = optionsSource.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0, `${startMarker} missing`);
  assert.ok(end > start, `${endMarker} missing`);
  return optionsSource.slice(start, end);
}

const restrictedAccounts = {
  mode: "restricted",
  accounts: [
    { id: "account1", name: "Work", type: "imap", allowed: true },
    { id: "account2", name: "Personal", type: "imap", allowed: false },
  ],
};
const corruptAccounts = {
  mode: "error",
  error: "Invalid account access configuration",
  accounts: restrictedAccounts.accounts.map(account => ({ ...account, allowed: false })),
};
const restrictedTools = {
  mode: "restricted",
  revision: "fixture-revision",
  groups: { system: "System", messages: "Messages" },
  getMessagesLimit: 10,
  tools: [
    { name: "listAccounts", group: "system", crud: "read", enabled: true, undisableable: true },
    { name: "getMessage", group: "messages", crud: "read", enabled: false },
    { name: "getMessages", group: "messages", crud: "read", enabled: true },
  ],
};
const corruptTools = {
  ...restrictedTools,
  mode: "error",
  error: "Invalid tool access configuration",
  tools: restrictedTools.tools.map(tool => ({ ...tool, enabled: !!tool.undisableable })),
};

// Only the DOM and browser API boundaries are stubbed. Access decisions and
// event listeners come from the production options script's marked sections.
function loadOptions(overrides = {}, apiOverrides = {}) {
  const byId = new Map();
  class Element {
    constructor(tagName) {
      this.tagName = tagName;
      this.children = [];
      this.listeners = new Map();
      this.attributes = new Map();
      this.dataset = {};
      this.checked = false;
      this.disabled = false;
      this.value = "";
      this.textContent = "";
      this.className = "";
    }
    set id(value) {
      this.elementId = value;
      byId.set(value, this);
    }
    get id() { return this.elementId; }
    set innerHTML(_value) { this.children = []; }
    appendChild(child) {
      this.children.push(child);
      return child;
    }
    addEventListener(type, listener) {
      const listeners = this.listeners.get(type) || [];
      listeners.push(listener);
      this.listeners.set(type, listeners);
    }
    async fire(type) {
      const listeners = this.listeners.get(type) || [];
      assert.ok(listeners.length, `No ${type} listener on ${this.id}`);
      for (const listener of listeners) await listener({ target: this });
    }
    querySelectorAll(selector) {
      assert.equal(selector, 'input[type="checkbox"]');
      const matches = [];
      for (const child of this.children) {
        if (child.tagName === "input" && child.type === "checkbox") matches.push(child);
        matches.push(...child.querySelectorAll(selector));
      }
      return matches;
    }
    setAttribute(name, value) { this.attributes.set(name, value); }
    removeAttribute(name) { this.attributes.delete(name); }
  }
  const document = {
    createElement: tagName => new Element(tagName),
    getElementById: id => byId.get(id) || null,
  };
  const elements = {};
  for (const [id, tagName] of Object.entries({
    accountList: "ul", saveBtn: "button", saveStatus: "span",
    toolList: "ul", saveToolsBtn: "button", saveToolsStatus: "span",
    allowEncryptedMessages: "input", allowAllCalendars: "input", allowAllAddressBooks: "input",
    savePrivacyBtn: "button", savePrivacyStatus: "span",
    allowContactWrites: "input", saveContactWritesBtn: "button", saveContactWritesStatus: "span",
  })) {
    elements[id] = document.createElement(tagName);
    elements[id].id = id;
    if (tagName === "button") elements[id].disabled = true;
  }

  const state = {
    accountConfig: structuredClone(restrictedAccounts),
    toolConfig: structuredClone(restrictedTools),
    privacySettings: { allowEncryptedMessages: false, allowAllCalendars: false, allowAllAddressBooks: false },
    contactWrites: { allowContactWrites: false },
    contactWritesResult: { success: true },
    accountResult: { success: true },
    toolResult: { success: true },
    privacyResult: { success: true },
    ...overrides,
  };
  const calls = { accounts: [], tools: [], privacy: [], contactWrites: [] };
  const api = {
    async getAccountAccessConfig() {
      if (state.accountConfig instanceof Error) throw state.accountConfig;
      return structuredClone(state.accountConfig);
    },
    async getToolAccessConfig() {
      if (state.accessApi) return state.accessApi.getToolAccessConfig();
      if (state.toolConfig instanceof Error) throw state.toolConfig;
      return structuredClone(state.toolConfig);
    },
    async getPrivacySettings() {
      if (state.privacySettings instanceof Error) throw state.privacySettings;
      return structuredClone(state.privacySettings);
    },
    async getAllowContactWrites() {
      if (state.contactWrites instanceof Error) throw state.contactWrites;
      return structuredClone(state.contactWrites);
    },
    async setAllowContactWrites(value) {
      calls.contactWrites.push(value);
      if (state.contactWritesResult instanceof Error) throw state.contactWritesResult;
      return structuredClone(state.contactWritesResult);
    },
    async setAccountAccess(ids) {
      calls.accounts.push(Array.from(ids));
      if (state.accountResult instanceof Error) throw state.accountResult;
      if (state.accountResult.success && state.nextAccountConfig) state.accountConfig = state.nextAccountConfig;
      return structuredClone(state.accountResult);
    },
    async setToolAccess(disabledTools, getMessagesLimit, revision) {
      calls.tools.push({ disabledTools: Array.from(disabledTools), getMessagesLimit });
      if (state.accessApi) return state.accessApi.setToolAccess(disabledTools, getMessagesLimit, revision);
      if (state.toolResult instanceof Error) throw state.toolResult;
      if (state.toolResult.success && state.nextToolConfig) state.toolConfig = state.nextToolConfig;
      return structuredClone(state.toolResult);
    },
    async setPrivacySettings(...settings) {
      calls.privacy.push(settings);
      if (state.privacyResult instanceof Error) throw state.privacyResult;
      return structuredClone(state.privacyResult);
    },
  };
  // Reuse this DOM/listener harness for persistence checks with the production
  // contact preference API; ordinary UI error tests keep their boundary mocks.
  const sandbox = { ...elements, document, browser: { mcpServer: { ...api, ...apiOverrides } } };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.resolve(__dirname, "../extension/tool-access.js"), "utf8"), sandbox);
  vm.runInContext([
    markedOptionsSnippet("OPTIONS ACCESS STATE"),
    markedOptionsSnippet("OPTIONS ACCOUNT ACCESS"),
    markedOptionsSnippet("OPTIONS TOOL ACCESS"),
    markedOptionsSnippet("OPTIONS CONTACT WRITES"),
    markedOptionsSnippet("OPTIONS PRIVACY SETTINGS"),
  ].join("\n"), sandbox);
  return {
    ...elements, document, state, calls,
    loadAccountAccess: sandbox.loadAccountAccess,
    loadToolAccess: sandbox.loadToolAccess,
    loadPrivacySettings: sandbox.loadPrivacySettings,
    loadContactWritesPref: sandbox.loadContactWritesPref,
  };
}

describe("Options account access", () => {
  it("shows corrupt configuration as blocked and ignores an unchanged Save", async () => {
    const ui = loadOptions({ accountConfig: corruptAccounts });
    await ui.loadAccountAccess();
    assert.equal(ui.saveBtn.disabled, true);
    assert.match(ui.saveStatus.textContent, /Invalid account access configuration.*blocked/);
    assert.equal(ui.saveStatus.className, "save-status error");
    assert.ok(ui.accountList.querySelectorAll('input[type="checkbox"]').every(checkbox => !checkbox.checked));
    await ui.saveBtn.fire("click");
    assert.deepEqual(ui.calls.accounts, []);
  });

  for (const allCurrentAccountsAllowed of [false, true]) {
    it(`preserves an unchanged allowlist with all current accounts allowed=${allCurrentAccountsAllowed}`, async () => {
      const accountConfig = structuredClone(restrictedAccounts);
      if (allCurrentAccountsAllowed) accountConfig.accounts[1].allowed = true;
      const ui = loadOptions({ accountConfig });
      await ui.loadAccountAccess();
      assert.equal(ui.saveBtn.disabled, true);
      await ui.saveBtn.fire("click");
      assert.deepEqual(ui.calls.accounts, []);

      const checkbox = ui.document.getElementById("acct-account2");
      checkbox.checked = !checkbox.checked;
      await checkbox.fire("change");
      assert.equal(ui.saveBtn.disabled, false);
      checkbox.checked = !checkbox.checked;
      await checkbox.fire("change");
      assert.equal(ui.saveBtn.disabled, true);
      await ui.saveBtn.fire("click");
      assert.deepEqual(ui.calls.accounts, []);
    });
  }

  it("rejects an empty selection without sending the unrestricted empty array", async () => {
    const ui = loadOptions();
    await ui.loadAccountAccess();
    const checkbox = ui.document.getElementById("acct-account1");
    checkbox.checked = false;
    await checkbox.fire("change");
    assert.equal(ui.saveBtn.disabled, false);
    await ui.saveBtn.fire("click");
    assert.deepEqual(ui.calls.accounts, []);
    assert.match(ui.saveStatus.textContent, /Select at least one account.*has not changed/);
    assert.equal(ui.saveStatus.className, "save-status error");
  });

  it("repairs a corrupt configuration only after an explicit account selection", async () => {
    const ui = loadOptions({ accountConfig: corruptAccounts, nextAccountConfig: restrictedAccounts });
    await ui.loadAccountAccess();
    const checkbox = ui.document.getElementById("acct-account1");
    checkbox.checked = true;
    await checkbox.fire("change");
    assert.equal(ui.saveBtn.disabled, false);
    await ui.saveBtn.fire("click");
    assert.deepEqual(ui.calls.accounts, [["account1"]]);
    assert.equal(ui.saveBtn.disabled, true);
    assert.equal(ui.saveStatus.className, "save-status");
    assert.equal(ui.document.getElementById("acct-account1").checked, true);
    assert.equal(ui.document.getElementById("acct-account2").checked, false);
  });
});

describe("Options tool access", () => {
  it("keeps corrupt tool configuration blocked on unchanged and limit-only saves", async () => {
    const ui = loadOptions({ toolConfig: corruptTools });
    await ui.loadToolAccess();
    assert.equal(ui.saveToolsBtn.disabled, true);
    assert.match(ui.saveToolsStatus.textContent, /Invalid tool access configuration.*blocked/);
    assert.equal(ui.saveToolsStatus.className, "save-status error");
    assert.equal(ui.document.getElementById("tool-listAccounts").checked, true);
    assert.equal(ui.document.getElementById("tool-listAccounts").disabled, true);
    assert.equal(ui.document.getElementById("tool-getMessage").checked, false);
    assert.equal(ui.document.getElementById("tool-getMessages").checked, false);
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(ui.calls.tools, []);

    const limit = ui.document.getElementById("getMessagesLimit");
    limit.value = "5";
    await limit.fire("input");
    assert.equal(ui.saveToolsBtn.disabled, true);
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(ui.calls.tools, []);
    assert.match(ui.saveToolsStatus.textContent, /blocked/);

    const checkbox = ui.document.getElementById("tool-getMessage");
    checkbox.checked = true;
    await checkbox.fire("change");
    assert.equal(ui.saveToolsBtn.disabled, false);
    checkbox.checked = false;
    await checkbox.fire("change");
    assert.equal(ui.saveToolsBtn.disabled, true);
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(ui.calls.tools, []);
  });

  it("preserves valid restrictions on unchanged and reverted tool selections", async () => {
    const ui = loadOptions();
    await ui.loadToolAccess();
    assert.equal(ui.saveToolsBtn.disabled, true);
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(ui.calls.tools, []);
    const checkbox = ui.document.getElementById("tool-getMessage");
    checkbox.checked = true;
    await checkbox.fire("change");
    assert.equal(ui.saveToolsBtn.disabled, false);
    checkbox.checked = false;
    await checkbox.fire("change");
    assert.equal(ui.saveToolsBtn.disabled, true);
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(ui.calls.tools, []);
  });

  it("repairs corrupt tool configuration after explicit tool access edits", async () => {
    const nextToolConfig = structuredClone(restrictedTools);
    nextToolConfig.tools[1].enabled = true;
    nextToolConfig.tools[2].enabled = false;
    const ui = loadOptions({ toolConfig: corruptTools, nextToolConfig });
    await ui.loadToolAccess();
    const checkbox = ui.document.getElementById("tool-getMessage");
    checkbox.checked = true;
    await checkbox.fire("change");
    assert.equal(ui.saveToolsBtn.disabled, false);
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(ui.calls.tools, [{ disabledTools: ["getMessages"], getMessagesLimit: 10 }]);
    assert.equal(ui.saveToolsBtn.disabled, true);
    assert.equal(ui.saveToolsStatus.className, "save-status");
    assert.equal(ui.document.getElementById("tool-getMessage").checked, true);
    assert.equal(ui.document.getElementById("tool-getMessages").checked, false);
  });

  it("allows a batch-limit-only change when the access configuration is valid", async () => {
    const ui = loadOptions();
    await ui.loadToolAccess();
    const limit = ui.document.getElementById("getMessagesLimit");
    limit.value = "7";
    await limit.fire("input");
    assert.equal(ui.saveToolsBtn.disabled, false);
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(ui.calls.tools, [{ disabledTools: ["getMessage"], getMessagesLimit: 7 }]);
  });
});

describe("Options privacy preferences", () => {
  it("loads and saves all privacy opt-ins as off by default", async () => {
    const ui = loadOptions();
    await ui.loadPrivacySettings();
    assert.equal(ui.allowEncryptedMessages.checked, false);
    assert.equal(ui.allowAllCalendars.checked, false);
    assert.equal(ui.allowAllAddressBooks.checked, false);
    assert.equal(ui.savePrivacyBtn.disabled, false);
    await ui.savePrivacyBtn.fire("click");
    assert.deepEqual(ui.calls.privacy, [[false, false, false]]);
    assert.equal(ui.savePrivacyStatus.textContent, "Saved.");
  });

  it("loads enabled flags and sends the user's selection in API parameter order", async () => {
    const ui = loadOptions({
      privacySettings: { allowEncryptedMessages: true, allowAllCalendars: false, allowAllAddressBooks: true },
    });
    await ui.loadPrivacySettings();
    assert.equal(ui.allowEncryptedMessages.checked, true);
    assert.equal(ui.allowAllCalendars.checked, false);
    assert.equal(ui.allowAllAddressBooks.checked, true);
    ui.allowEncryptedMessages.checked = false;
    ui.allowAllCalendars.checked = true;
    ui.allowAllAddressBooks.checked = false;
    await ui.savePrivacyBtn.fire("click");
    assert.deepEqual(ui.calls.privacy, [[false, true, false]]);
    assert.equal(ui.savePrivacyStatus.className, "save-status");
  });

  it("does not enable access from missing or non-boolean preference values", async () => {
    const ui = loadOptions({ privacySettings: { allowEncryptedMessages: "true", allowAllCalendars: 1 } });
    await ui.loadPrivacySettings();
    assert.equal(ui.allowEncryptedMessages.checked, false);
    assert.equal(ui.allowAllCalendars.checked, false);
    assert.equal(ui.allowAllAddressBooks.checked, false);
  });

  for (const errorKind of ["returned", "thrown"]) {
    it(`shows ${errorKind} privacy load errors and keeps Save disabled`, async () => {
      const message = "Cannot read privacy preferences";
      const privacySettings = errorKind === "thrown" ? new Error(message) : { error: message };
      const ui = loadOptions({ privacySettings });
      await ui.loadPrivacySettings();
      assert.equal(ui.savePrivacyBtn.disabled, true);
      assert.equal(ui.allowEncryptedMessages.checked, false);
      assert.equal(ui.allowAllCalendars.checked, false);
      assert.equal(ui.allowAllAddressBooks.checked, false);
      assert.match(ui.savePrivacyStatus.textContent, /Cannot read privacy preferences/);
      assert.equal(ui.savePrivacyStatus.className, "save-status error");
      assert.deepEqual(ui.calls.privacy, []);
    });

    it(`shows ${errorKind} privacy save errors and permits retry`, async () => {
      const message = "Cannot save privacy preferences";
      const privacyResult = errorKind === "thrown" ? new Error(message) : { error: message };
      const ui = loadOptions({ privacyResult });
      await ui.loadPrivacySettings();
      ui.allowEncryptedMessages.checked = true;
      await ui.savePrivacyBtn.fire("click");
      assert.deepEqual(ui.calls.privacy, [[true, false, false]]);
      assert.match(ui.savePrivacyStatus.textContent, /Cannot save privacy preferences/);
      assert.equal(ui.savePrivacyStatus.className, "save-status error");
      assert.equal(ui.savePrivacyBtn.disabled, false);
      ui.state.privacyResult = { success: true };
      await ui.savePrivacyBtn.fire("click");
      assert.equal(ui.calls.privacy.length, 2);
      assert.equal(ui.savePrivacyStatus.textContent, "Saved.");
    });
  }
});

describe("Options contact writes", () => {
  it("starts unchecked and locked until the stored value has been read", async () => {
    const ui = loadOptions();
    ui.allowContactWrites.checked = true;
    ui.allowContactWrites.disabled = false;
    const loading = ui.loadContactWritesPref();
    assert.equal(ui.allowContactWrites.checked, false);
    assert.equal(ui.allowContactWrites.disabled, true);
    assert.equal(ui.saveContactWritesBtn.disabled, true);
    await loading;
    assert.equal(ui.allowContactWrites.checked, false);
    assert.equal(ui.allowContactWrites.disabled, false);
    assert.equal(ui.saveContactWritesBtn.disabled, false);
  });

  it("loads an enabled value and saves the user's choice", async () => {
    const ui = loadOptions({ contactWrites: { allowContactWrites: true } });
    await ui.loadContactWritesPref();
    assert.equal(ui.allowContactWrites.checked, true);
    ui.allowContactWrites.checked = false;
    await ui.saveContactWritesBtn.fire("click");
    assert.deepEqual(ui.calls.contactWrites, [false]);
    assert.equal(ui.saveContactWritesStatus.textContent, "Saved.");
    assert.equal(ui.saveContactWritesBtn.disabled, false);
  });

  it("persists consent and revocation across Options reloads through the production preference API", async () => {
    const runtime = loadContactRuntime();
    const fields = { emails: [{ address: "roundtrip@example.com", type: "work" }], displayName: "Round Trip" };
    const first = loadOptions({}, runtime.options);
    await first.loadContactWritesPref();
    assert.equal(first.allowContactWrites.checked, false);
    assert.match(runtime.create(fields).error, /Allow contact writes/);

    first.allowContactWrites.checked = true;
    await first.saveContactWritesBtn.fire("click");
    assert.equal(first.saveContactWritesStatus.textContent, "Saved.");
    assert.equal((await runtime.options.getAllowContactWrites()).allowContactWrites, true);
    const created = runtime.create(fields);
    assert.equal(created.success, true);

    await first.loadPrivacySettings();
    first.allowAllAddressBooks.checked = true;
    await first.savePrivacyBtn.fire("click");
    assert.equal((await runtime.options.getPrivacySettings()).allowAllAddressBooks, true);
    assert.equal((await runtime.options.getAllowContactWrites()).allowContactWrites, true);

    const reopened = loadOptions({}, runtime.options);
    await reopened.loadContactWritesPref();
    assert.equal(reopened.allowContactWrites.checked, true);
    reopened.allowContactWrites.checked = false;
    await reopened.saveContactWritesBtn.fire("click");
    assert.equal(reopened.saveContactWritesStatus.textContent, "Saved.");
    assert.equal((await runtime.options.getAllowContactWrites()).allowContactWrites, false);

    const revoked = loadOptions({}, runtime.options);
    await revoked.loadContactWritesPref();
    assert.equal(revoked.allowContactWrites.checked, false);
    await revoked.loadPrivacySettings();
    assert.equal(revoked.allowAllAddressBooks.checked, true);
    await revoked.savePrivacyBtn.fire("click");
    assert.equal((await runtime.options.getAllowContactWrites()).allowContactWrites, false);
    assert.match(runtime.create(fields).error, /Allow contact writes/);
    assert.match(runtime.update(created.id, { displayName: "Blocked" }).error, /Allow contact writes/);
    assert.match(runtime.deleteContact(created.id).error, /Allow contact writes/);
    assert.equal(runtime.getContact(created.id).displayName, fields.displayName);
  });

  it("treats anything but true as off", async () => {
    for (const value of ["true", 1, null, undefined]) {
      const ui = loadOptions({ contactWrites: { allowContactWrites: value } });
      await ui.loadContactWritesPref();
      assert.equal(ui.allowContactWrites.checked, false);
    }
  });

  it("keeps the setting locked when loading fails", async () => {
    const ui = loadOptions({ contactWrites: new Error("unavailable") });
    await ui.loadContactWritesPref();
    assert.equal(ui.allowContactWrites.checked, false);
    assert.equal(ui.allowContactWrites.disabled, true);
    assert.equal(ui.saveContactWritesBtn.disabled, true);
    assert.match(ui.saveContactWritesStatus.textContent, /Error loading setting/);
    assert.equal(ui.saveContactWritesStatus.className, "save-status error");
    assert.deepEqual(ui.calls.contactWrites, []);
  });

  for (const errorKind of ["thrown", "returned"]) {
    it(`shows ${errorKind} save errors and permits retry`, async () => {
      const contactWritesResult = errorKind === "thrown" ? new Error("cannot save") : { error: "cannot save" };
      const ui = loadOptions({ contactWritesResult });
      await ui.loadContactWritesPref();
      ui.allowContactWrites.checked = true;
      await ui.saveContactWritesBtn.fire("click");
      assert.deepEqual(ui.calls.contactWrites, [true]);
      assert.match(ui.saveContactWritesStatus.textContent, /cannot save/);
      assert.equal(ui.saveContactWritesStatus.className, "save-status error");
      assert.equal(ui.saveContactWritesBtn.disabled, false);
      ui.state.contactWritesResult = { success: true };
      await ui.saveContactWritesBtn.fire("click");
      assert.equal(ui.saveContactWritesStatus.textContent, "Saved.");
    });
  }
});

describe("Options dependencies with the production preference API", () => {
  for (const childDisabled of [false, true]) {
    for (const parentDisabled of [false, true]) {
      it(`preserves child=${!childDisabled} across unsaved parent toggles, unrelated saves and reloads, parent=${!parentDisabled}`, async () => {
        const disabled = [childDisabled && "dryRunCompose", parentDisabled && "sendMail"].filter(Boolean);
        const runtime = loadPrivacyRuntime({ initial: { PREF_DISABLED_TOOLS: JSON.stringify(disabled) } });
        const ui = loadOptions({ accessApi: runtime.options });
        await ui.loadToolAccess();
        const get = name => ui.document.getElementById("tool-" + name);
        const labels = ui.toolList.children.flatMap(row => row.children);
        assert.ok(labels.some(label => label.children.some(note => note.textContent === "requires sendMail")));
        assert.equal(get("dryRunCompose").checked, !childDisabled && !parentDisabled);
        assert.equal(get("dryRunCompose").disabled, parentDisabled);
        assert.equal(get("listAccounts").checked, true);
        assert.equal(get("listAccounts").disabled, true);
        // The child's preference is independent even before any save.
        get("sendMail").checked = parentDisabled;
        await get("sendMail").fire("change");
        assert.equal(get("dryRunCompose").disabled, !parentDisabled);
        assert.equal(get("dryRunCompose").checked, !childDisabled && parentDisabled);
        get("sendMail").checked = !parentDisabled;
        await get("sendMail").fire("change");
        assert.equal(ui.saveToolsBtn.disabled, true);
        const limit = ui.document.getElementById("getMessagesLimit");
        limit.value = "7";
        await limit.fire("input");
        await ui.saveToolsBtn.fire("click");
        assert.deepEqual(Array.from((await runtime.options.getToolAccessConfig()).disabledTools).sort(), disabled.sort());
        assert.equal(get("dryRunCompose").checked, !childDisabled && !parentDisabled);
        assert.equal(get("dryRunCompose").disabled, parentDisabled);
        // Save the opposite parent state, then reload the real persisted state.
        get("sendMail").checked = parentDisabled;
        await get("sendMail").fire("change");
        await ui.saveToolsBtn.fire("click");
        const config = await runtime.options.getToolAccessConfig();
        assert.equal(config.disabledTools.includes("dryRunCompose"), childDisabled);
        assert.equal(config.disabledTools.includes("sendMail"), !parentDisabled);
        assert.equal(get("dryRunCompose").checked, !childDisabled && parentDisabled);
      });
    }
  }

  it("keeps an explicitly unchecked child off after disabling and re-enabling its parent", async () => {
    const runtime = loadPrivacyRuntime();
    const ui = loadOptions({ accessApi: runtime.options });
    await ui.loadToolAccess();
    const child = ui.document.getElementById("tool-dryRunCompose");
    const parent = ui.document.getElementById("tool-sendMail");
    child.checked = false;
    await child.fire("change");
    for (const enabled of [false, true]) {
      parent.checked = enabled;
      await parent.fire("change");
      assert.equal(child.checked, false);
    }
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(Array.from((await runtime.options.getToolAccessConfig()).disabledTools), ["dryRunCompose"]);
    assert.equal(ui.document.getElementById("tool-dryRunCompose").checked, false);
  });
});

describe("Options Tool Access upgrades with the production preference API", () => {
  for (const exposeLocks of [true, false]) {
    it(`shows a policy note without disabling existing allowed tools, lock API=${exposeLocks}`, async () => {
      const pref = "extensions.thunderbird-mcp.disabledTools";
      const runtime = loadPrivacyRuntime({ legacy: true, exposeLocks,
        defaults: new Map([[pref, '["getMessage"]']]), locked: new Set([pref]),
      });
      const ui = loadOptions({ accessApi: runtime.options });
      await ui.loadToolAccess();
      assert.match(ui.saveToolsStatus.textContent, /managed by policy.*new tools stay disabled/i);
      assert.equal(ui.saveToolsStatus.className, "save-status");
      assert.equal(ui.document.getElementById("tool-getMessage").checked, false);
      assert.equal(ui.document.getElementById("tool-getMessages").checked, true);
      for (const name of Object.keys(runtime.policy.introducedIn)) {
        assert.equal(ui.document.getElementById("tool-" + name).checked, false, name);
      }
      const newTool = ui.document.getElementById("tool-refreshFolder");
      newTool.checked = true;
      await newTool.fire("change");
      assert.match(ui.saveToolsStatus.textContent, /managed by policy/i);
      await ui.saveToolsBtn.fire("click");
      assert.match(ui.saveToolsStatus.textContent, /managed by policy/i);
      assert.equal(ui.saveToolsStatus.className, "save-status error");
      const observed = JSON.parse(runtime.values.get(runtime.prefNames.PREF_TOOL_ACCESS_STATE));
      assert.equal(observed.customized, true);
      assert.equal(observed.consideredRelease, "0.9.3");
      assert.equal(runtime.isToolEnabled("refreshFolder"), false);
      runtime.locked.clear();
      await ui.saveToolsBtn.fire("click");
      assert.match(ui.saveToolsStatus.textContent, /Reload/);
      await ui.loadToolAccess();
      assert.equal(ui.saveToolsStatus.textContent, "");
      assert.equal(ui.document.getElementById("tool-refreshFolder").checked, false);
      const enabled = ui.document.getElementById("tool-refreshFolder");
      enabled.checked = true;
      await enabled.fire("change");
      await ui.saveToolsBtn.fire("click");
      assert.equal(ui.saveToolsStatus.className, "save-status");
      assert.equal(runtime.isToolEnabled("refreshFolder"), true);
    });
  }

  it("shows migrated tools unchecked and keeps an explicit enable across Options and runtime reloads", async () => {
    const runtime = loadPrivacyRuntime({ legacy: true, initial: { PREF_DISABLED_TOOLS: '["getMessage"]' } });
    const ui = loadOptions({ accessApi: runtime.options });
    await ui.loadToolAccess();
    for (const name of [
      "dryRunCompose", "refreshFolder", "refreshCalendar", "getMessageHeaders", "batchGetMessageHeaders",
      "createVirtualFolder", "listVirtualFolders", "deleteVirtualFolder", "getUnsubscribeInfo",
    ]) {
      assert.equal(ui.document.getElementById("tool-" + name).checked, false, name);
    }
    assert.equal(ui.document.getElementById("tool-getMessage").checked, false);
    assert.equal(ui.document.getElementById("tool-sendMail").checked, true);
    assert.equal(ui.saveToolsBtn.disabled, true);

    const checkbox = ui.document.getElementById("tool-refreshFolder");
    checkbox.checked = true;
    await checkbox.fire("change");
    await ui.saveToolsBtn.fire("click");
    assert.equal(ui.saveToolsStatus.className, "save-status");
    const restarted = loadPrivacyRuntime({ legacy: true, values: runtime.values });
    const reopened = loadOptions({ accessApi: restarted.options });
    await reopened.loadToolAccess();
    assert.equal(reopened.document.getElementById("tool-refreshFolder").checked, true);
    assert.equal(reopened.document.getElementById("tool-refreshCalendar").checked, false);
    assert.equal(reopened.document.getElementById("tool-getMessage").checked, false);
  });

  it("preserves restrictions for tools absent from the page's catalog", async () => {
    const runtime = loadPrivacyRuntime({ initial: { PREF_DISABLED_TOOLS: '["futureTool","getMessage"]' } });
    const ui = loadOptions({ accessApi: runtime.options });
    await ui.loadToolAccess();
    assert.equal(ui.document.getElementById("tool-futureTool"), null);
    const limit = ui.document.getElementById("getMessagesLimit");
    limit.value = "7";
    await limit.fire("input");
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(Array.from((await runtime.options.getToolAccessConfig()).disabledTools).sort(), ["futureTool", "getMessage"]);

    const checkbox = ui.document.getElementById("tool-getMessage");
    checkbox.checked = true;
    await checkbox.fire("change");
    await ui.saveToolsBtn.fire("click");
    assert.deepEqual(Array.from((await runtime.options.getToolAccessConfig()).disabledTools), ["futureTool"]);
  });

  for (const limitOnly of [false, true]) {
    it(`rejects a stale Options save without overwriting newer choices, limitOnly=${limitOnly}`, async () => {
      const runtime = loadPrivacyRuntime();
      const first = loadOptions({ accessApi: runtime.options });
      const stale = loadOptions({ accessApi: runtime.options });
      await first.loadToolAccess();
      await stale.loadToolAccess();

      const changed = first.document.getElementById("tool-getMessage");
      changed.checked = false;
      await changed.fire("change");
      await first.saveToolsBtn.fire("click");
      const saved = await runtime.options.getToolAccessConfig();
      assert.deepEqual(Array.from(saved.disabledTools), ["getMessage"]);

      if (limitOnly) {
        const limit = stale.document.getElementById("getMessagesLimit");
        limit.value = "7";
        await limit.fire("input");
      } else {
        const checkbox = stale.document.getElementById("tool-dryRunCompose");
        checkbox.checked = false;
        await checkbox.fire("change");
      }
      const writeCount = runtime.writes.length;
      await stale.saveToolsBtn.fire("click");
      assert.match(stale.saveToolsStatus.textContent, /reload/i);
      assert.equal(stale.saveToolsStatus.className, "save-status error");
      assert.equal(runtime.writes.length, writeCount);
      const current = await runtime.options.getToolAccessConfig();
      assert.equal(current.revision, saved.revision);
      assert.deepEqual(Array.from(current.disabledTools), ["getMessage"]);
      assert.equal(current.getMessagesLimit, 10);
      await stale.loadToolAccess();
      assert.equal(stale.document.getElementById("tool-getMessage").checked, false);
      assert.equal(stale.saveToolsStatus.className, "save-status");
    });
  }

  for (const corruptPref of ["PREF_TOOL_ACCESS_STATE", "PREF_DISABLED_TOOLS"]) {
    it(`leaves corrupt ${corruptPref} untouched until the user changes a tool and saves`, async () => {
      const runtime = loadPrivacyRuntime({ initial: { [corruptPref]: "not-json" } });
      const ui = loadOptions({ accessApi: runtime.options });
      await ui.loadToolAccess();
      assert.match(ui.saveToolsStatus.textContent, /blocked/);
      assert.equal(ui.document.getElementById("tool-listAccounts").checked, true);
      assert.equal(ui.document.getElementById("tool-getMessage").checked, false);
      const writeCount = runtime.writes.length;
      const limit = ui.document.getElementById("getMessagesLimit");
      limit.value = "7";
      await limit.fire("input");
      await ui.saveToolsBtn.fire("click");
      assert.equal(runtime.writes.length, writeCount);
      assert.equal(runtime.values.get(runtime.prefNames[corruptPref]), "not-json");

      const checkbox = ui.document.getElementById("tool-getMessage");
      checkbox.checked = true;
      await checkbox.fire("change");
      await ui.saveToolsBtn.fire("click");
      assert.equal(ui.saveToolsStatus.className, "save-status");
      assert.equal(ui.document.getElementById("tool-getMessage").checked, true);
      assert.equal(ui.document.getElementById("tool-getMessages").checked, false);
      assert.notEqual(runtime.values.get(runtime.prefNames[corruptPref]), "not-json");
      assert.equal((await runtime.options.getToolAccessConfig()).mode, "restricted");
    });
  }

  it("keeps a migrated child's own denial when the parent is re-enabled", async () => {
    const runtime = loadPrivacyRuntime({ legacy: true, initial: { PREF_DISABLED_TOOLS: '["sendMail"]' } });
    const ui = loadOptions({ accessApi: runtime.options });
    await ui.loadToolAccess();
    const parent = ui.document.getElementById("tool-sendMail");
    const child = ui.document.getElementById("tool-dryRunCompose");
    assert.equal(child.checked, false);
    assert.equal(child.disabled, true);
    parent.checked = true;
    await parent.fire("change");
    assert.equal(child.disabled, false);
    assert.equal(child.checked, false);
    await ui.saveToolsBtn.fire("click");
    const config = await runtime.options.getToolAccessConfig();
    assert.equal(config.disabledTools.includes("sendMail"), false);
    assert.equal(config.disabledTools.includes("dryRunCompose"), true);
    assert.equal(ui.document.getElementById("tool-dryRunCompose").checked, false);
  });
});
