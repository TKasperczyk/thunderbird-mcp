"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

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
function loadOptions(overrides = {}) {
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
  })) {
    elements[id] = document.createElement(tagName);
    elements[id].id = id;
    if (tagName === "button") elements[id].disabled = true;
  }

  const state = {
    accountConfig: structuredClone(restrictedAccounts),
    toolConfig: structuredClone(restrictedTools),
    privacySettings: { allowEncryptedMessages: false, allowAllCalendars: false, allowAllAddressBooks: false },
    accountResult: { success: true },
    toolResult: { success: true },
    privacyResult: { success: true },
    ...overrides,
  };
  const calls = { accounts: [], tools: [], privacy: [] };
  const api = {
    async getAccountAccessConfig() {
      if (state.accountConfig instanceof Error) throw state.accountConfig;
      return structuredClone(state.accountConfig);
    },
    async getToolAccessConfig() {
      if (state.toolConfig instanceof Error) throw state.toolConfig;
      return structuredClone(state.toolConfig);
    },
    async getPrivacySettings() {
      if (state.privacySettings instanceof Error) throw state.privacySettings;
      return structuredClone(state.privacySettings);
    },
    async setAccountAccess(ids) {
      calls.accounts.push(Array.from(ids));
      if (state.accountResult instanceof Error) throw state.accountResult;
      if (state.accountResult.success && state.nextAccountConfig) state.accountConfig = state.nextAccountConfig;
      return structuredClone(state.accountResult);
    },
    async setToolAccess(disabledTools, getMessagesLimit) {
      calls.tools.push({ disabledTools: Array.from(disabledTools), getMessagesLimit });
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
  const sandbox = { ...elements, document, browser: { mcpServer: api } };
  vm.createContext(sandbox);
  vm.runInContext([
    markedOptionsSnippet("OPTIONS ACCESS STATE"),
    markedOptionsSnippet("OPTIONS ACCOUNT ACCESS"),
    markedOptionsSnippet("OPTIONS TOOL ACCESS"),
    markedOptionsSnippet("OPTIONS PRIVACY SETTINGS"),
  ].join("\n"), sandbox);
  return {
    ...elements, document, state, calls,
    loadAccountAccess: sandbox.loadAccountAccess,
    loadToolAccess: sandbox.loadToolAccess,
    loadPrivacySettings: sandbox.loadPrivacySettings,
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
