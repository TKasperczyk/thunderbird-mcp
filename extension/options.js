// `browser` is declared as a global in eslint.config.mjs for the
// extension/ file group, so we don't repeat the /* global */ comment
// here (it triggered no-redeclare).
"use strict";

const statusDot = document.getElementById("statusDot");
const statusText = document.getElementById("statusText");
const serverPort = document.getElementById("serverPort");
const connFile = document.getElementById("connFile");
const buildInfo = document.getElementById("buildInfo");
const accountList = document.getElementById("accountList");
const saveBtn = document.getElementById("saveBtn");
const saveStatus = document.getElementById("saveStatus");

const toolList = document.getElementById("toolList");
const saveToolsBtn = document.getElementById("saveToolsBtn");
const saveToolsStatus = document.getElementById("saveToolsStatus");

let currentAccounts = [];
let currentTools = [];

// CRUD labels for sub-group headers
const CRUD_LABELS = { read: "Read", create: "Create", update: "Update", delete: "Delete" };

async function loadServerInfo() {
  try {
    const info = await browser.mcpServer.getServerInfo();
    if (info.running) {
      statusDot.className = "status-dot running";
      statusText.textContent = "Running";
      serverPort.textContent = info.port || "--";
      connFile.textContent = info.connectionFile || "--";
    } else {
      statusDot.className = "status-dot stopped";
      statusText.textContent = "Not running";
      serverPort.textContent = "--";
      connFile.textContent = "--";
    }
    if (info.buildVersion) {
      // Parse git describe: "v0.2.0-7-g1461f1a+dirty" → tag, commits, hash, dirty
      const m = info.buildVersion.match(/^(v[\d.]+)(?:-(\d+)-g([0-9a-f]+))?(\+dirty)?$/);
      let display;
      if (m) {
        const [, tag, commits, hash, dirty] = m;
        display = tag;
        if (commits && commits !== "0") display += ` +${commits}`;
        display += ` (${hash || tag})`;
        if (dirty) {
          display += " +dirty";
          if (info.buildDate) {
            display += " " + info.buildDate.replace("T", " ").replace(/\.\d+Z$/, " UTC");
          }
        }
      } else {
        display = info.buildVersion;
      }
      buildInfo.textContent = display;
    } else {
      buildInfo.textContent = "--";
    }
  } catch (e) {
    statusDot.className = "status-dot stopped";
    statusText.textContent = "Error: " + e.message;
  }
}

async function loadAccountAccess() {
  try {
    const data = await browser.mcpServer.getAccountAccessConfig();
    currentAccounts = data.accounts || [];

    if (currentAccounts.length === 0) {
      accountList.innerHTML = "<li>No accounts found.</li>";
      return;
    }

    accountList.innerHTML = "";
    for (const acct of currentAccounts) {
      const li = document.createElement("li");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.id = "acct-" + acct.id;
      checkbox.value = acct.id;
      checkbox.checked = acct.allowed;
      checkbox.addEventListener("change", onAccountChange);

      const label = document.createElement("label");
      label.htmlFor = checkbox.id;
      label.textContent = acct.name;

      const typeSpan = document.createElement("span");
      typeSpan.className = "account-type";
      typeSpan.textContent = acct.type;
      label.appendChild(typeSpan);

      li.appendChild(checkbox);
      li.appendChild(label);
      accountList.appendChild(li);
    }

    saveBtn.disabled = false;
    saveStatus.textContent = "";
  } catch (e) {
    accountList.innerHTML = "";
    const li = document.createElement("li");
    li.textContent = "Error loading accounts: " + e.message;
    accountList.appendChild(li);
  }
}

function onAccountChange() {
  saveStatus.textContent = "";
}

saveBtn.addEventListener("click", async () => {
  saveBtn.disabled = true;
  saveStatus.textContent = "";
  saveStatus.className = "save-status";

  const checkboxes = accountList.querySelectorAll('input[type="checkbox"]');
  const checked = [];
  let allChecked = true;
  for (const cb of checkboxes) {
    if (cb.checked) {
      checked.push(cb.value);
    } else {
      allChecked = false;
    }
  }

  // If all are checked, send empty array (= allow all)
  const allowedIds = allChecked ? [] : checked;

  try {
    const result = await browser.mcpServer.setAccountAccess(allowedIds);
    if (result.error) {
      saveStatus.textContent = result.error;
      saveStatus.className = "save-status error";
    } else {
      saveStatus.textContent = "Saved.";
      // Reload to reflect updated state
      await loadAccountAccess();
    }
  } catch (e) {
    saveStatus.textContent = "Error: " + e.message;
    saveStatus.className = "save-status error";
  }
  saveBtn.disabled = false;
});

async function loadToolAccess() {
  try {
    const data = await browser.mcpServer.getToolAccessConfig();
    // Engine returns a structured error object on failure now (mode: "error"
    // + error: <message>) rather than throwing into the WebExtension boundary
    // where it gets wrapped to "An unexpected error occurred".
    if (data && data.error) {
      toolList.innerHTML = "";
      const li = document.createElement("li");
      li.textContent = "Error loading tools: " + data.error;
      toolList.appendChild(li);
      return;
    }
    currentTools = (data && data.tools) || [];
    const groupLabels = (data && data.groups) || {};
    const usingDefaults = !!(data && data.usingDefaults);

    if (currentTools.length === 0) {
      toolList.innerHTML = "<li>No tools found.</li>";
      return;
    }

    toolList.innerHTML = "";

    // Banner at the top when the user is still on the safety defaults --
    // destructive ops (send, delete) are off until the user opts in.
    if (usingDefaults) {
      const banner = document.createElement("li");
      banner.className = "tool-defaults-banner";
      banner.textContent =
        "Safe defaults active: send/forward/reply, all delete tools, and " +
        "batch tools (bulk_move_by_query, createDrafts) are off. " +
        "Check the boxes below for the ones you want to allow, then Save.";
      toolList.appendChild(banner);
    }

    // Tools arrive pre-sorted by group then CRUD order from the server.
    let currentGroup = null;
    let currentCrud = null;

    for (const tool of currentTools) {
      const group = tool.group || "other";
      const crud = tool.crud || "other";

      if (group !== currentGroup) {
        currentGroup = group;
        currentCrud = null;
        const header = document.createElement("li");
        header.className = "tool-group-header";
        header.textContent = groupLabels[group] || group.charAt(0).toUpperCase() + group.slice(1);
        toolList.appendChild(header);
      }

      if (crud !== currentCrud) {
        currentCrud = crud;
        const subHeader = document.createElement("li");
        subHeader.className = "tool-crud-header";
        subHeader.textContent = CRUD_LABELS[crud] || crud.charAt(0).toUpperCase() + crud.slice(1);
        toolList.appendChild(subHeader);
      }

      const li = document.createElement("li");
      if (tool.defaultDisabled) {
        li.classList.add("tool-default-off");
      }
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.id = "tool-" + tool.name;
      checkbox.value = tool.name;
      checkbox.checked = tool.enabled;
      if (tool.undisableable) {
        checkbox.disabled = true;
      }
      checkbox.addEventListener("change", () => {
        saveToolsStatus.textContent = "";
      });

      const label = document.createElement("label");
      label.htmlFor = checkbox.id;
      label.textContent = tool.name;
      if (tool.undisableable) {
        const lockSpan = document.createElement("span");
        lockSpan.className = "account-type";
        lockSpan.textContent = "required";
        label.appendChild(lockSpan);
      } else if (tool.defaultDisabled) {
        const defSpan = document.createElement("span");
        defSpan.className = "account-type tool-default-off-badge";
        defSpan.textContent = "off by default";
        label.appendChild(defSpan);
      }

      li.appendChild(checkbox);
      li.appendChild(label);
      toolList.appendChild(li);
    }

    saveToolsBtn.disabled = false;
    saveToolsStatus.textContent = "";
  } catch (e) {
    toolList.innerHTML = "";
    const li = document.createElement("li");
    li.textContent = "Error loading tools: " + e.message;
    toolList.appendChild(li);
  }
}

saveToolsBtn.addEventListener("click", async () => {
  saveToolsBtn.disabled = true;
  saveToolsStatus.textContent = "";
  saveToolsStatus.className = "save-status";

  const checkboxes = toolList.querySelectorAll('input[type="checkbox"]');
  const disabled = [];
  for (const cb of checkboxes) {
    if (!cb.checked && !cb.disabled) {
      disabled.push(cb.value);
    }
  }

  try {
    const result = await browser.mcpServer.setToolAccess(disabled);
    if (result.error) {
      saveToolsStatus.textContent = result.error;
      saveToolsStatus.className = "save-status error";
    } else {
      saveToolsStatus.textContent = "Saved.";
      await loadToolAccess();
    }
  } catch (e) {
    saveToolsStatus.textContent = "Error: " + e.message;
    saveToolsStatus.className = "save-status error";
  }
  saveToolsBtn.disabled = false;
});

const blockSkipReviewCheckbox = document.getElementById("blockSkipReview");
const saveSkipReviewBtn = document.getElementById("saveSkipReviewBtn");
const saveSkipReviewStatus = document.getElementById("saveSkipReviewStatus");

async function loadSkipReviewPref() {
  try {
    const { blockSkipReview } = await browser.mcpServer.getBlockSkipReview();
    blockSkipReviewCheckbox.checked = !!blockSkipReview;
    saveSkipReviewBtn.disabled = false;
    saveSkipReviewStatus.textContent = "";
  } catch (e) {
    saveSkipReviewStatus.textContent = "Error loading setting: " + e.message;
    saveSkipReviewStatus.className = "save-status error";
  }
}

saveSkipReviewBtn.addEventListener("click", async () => {
  saveSkipReviewBtn.disabled = true;
  saveSkipReviewStatus.textContent = "Saving...";
  saveSkipReviewStatus.className = "save-status";
  try {
    const result = await browser.mcpServer.setBlockSkipReview(blockSkipReviewCheckbox.checked);
    if (result.error) {
      saveSkipReviewStatus.textContent = result.error;
      saveSkipReviewStatus.className = "save-status error";
    } else {
      saveSkipReviewStatus.textContent = "Saved.";
    }
  } catch (e) {
    saveSkipReviewStatus.textContent = "Error: " + e.message;
    saveSkipReviewStatus.className = "save-status error";
  }
  saveSkipReviewBtn.disabled = false;
});

// ── Advanced Permissions ─────────────────────────────────────────────────

const permissionsJson = document.getElementById("permissionsJson");
const savePermissionsBtn = document.getElementById("savePermissionsBtn");
const clearPermissionsBtn = document.getElementById("clearPermissionsBtn");
const savePermissionsStatus = document.getElementById("savePermissionsStatus");

const PERMISSION_PRESETS = {
  empty: { version: 1, tools: {}, accounts: {}, folders: {} },
  readOnly: {
    version: 1,
    tools: {
      "*": { enabled: false }, // not honored by engine but documents intent
      sendMail: { enabled: false }, replyToMessage: { enabled: false }, forwardMessage: { enabled: false },
      createContact: { enabled: false }, updateContact: { enabled: false }, deleteContact: { enabled: false },
      createEvent: { enabled: false }, updateEvent: { enabled: false }, deleteEvent: { enabled: false },
      createTask: { enabled: false }, updateTask: { enabled: false },
      createFilter: { enabled: false }, updateFilter: { enabled: false }, deleteFilter: { enabled: false },
      reorderFilters: { enabled: false }, applyFilters: { enabled: false },
      createFolder: { enabled: false }, renameFolder: { enabled: false }, deleteFolder: { enabled: false },
      moveFolder: { enabled: false }, emptyTrash: { enabled: false }, emptyJunk: { enabled: false },
      updateMessage: { enabled: false }, deleteMessages: { enabled: false },
      bulk_move_by_query: { enabled: false }, createDrafts: { enabled: false },
    },
    accounts: {}, folders: {}
  },
  sendDomain: {
    version: 1,
    tools: {
      sendMail:        { enabled: true, alwaysReview: true, argRules: { to: { anyOf: ["*@example.com"] }, cc: { anyOf: ["*@example.com"] }, bcc: { anyOf: ["*@example.com"] } } },
      replyToMessage:  { enabled: true, alwaysReview: true, argRules: { to: { anyOf: ["*@example.com"] }, cc: { anyOf: ["*@example.com"] }, bcc: { anyOf: ["*@example.com"] } } },
      forwardMessage:  { enabled: true, alwaysReview: true, argRules: { to: { anyOf: ["*@example.com"] }, cc: { anyOf: ["*@example.com"] }, bcc: { anyOf: ["*@example.com"] } } }
    },
    accounts: {}, folders: {}
  },
  trashFolderOnly: {
    version: 1,
    tools: {
      deleteMessages: { enabled: false },
      deleteFolder:   { enabled: false }
    },
    accounts: {},
    folders: {
      "imap://user%40example.com@imap.example.com/Trash": {
        tools: {
          deleteMessages: { enabled: true, maxBatchSize: 500 }
        }
      }
    }
  },
  // Inbox-triage workflow: enable the aggregation + bulk tools so the AI can
  // (a) count-by-domain, (b) dry-run a batch move, (c) execute after the user
  // eyeballs the count. maxBatchSize caps a single bulk call at 500.
  inventoryAndTriage: {
    version: 1,
    tools: {
      inbox_inventory:     { enabled: true },
      bulk_move_by_query:  { enabled: true, maxBatchSize: 500, requireFolderArg: true }
    },
    accounts: {},
    folders: {}
  },
  // Parallel-correspondence workflow: enable batch draft creation but keep
  // every send-to-recipient path review-gated (alwaysReview forces the
  // compose window, max 10 drafts per call to prevent a runaway loop).
  batchDrafts: {
    version: 1,
    tools: {
      createDrafts:   { enabled: true, maxBatchSize: 10 },
      sendMail:       { enabled: true, alwaysReview: true },
      replyToMessage: { enabled: true, alwaysReview: true },
      forwardMessage: { enabled: true, alwaysReview: true }
    },
    accounts: {},
    folders: {}
  }
};

async function loadPermissionsConfig() {
  try {
    const data = await browser.mcpServer.getPermissionsConfig();
    if (data.error) {
      savePermissionsStatus.textContent = "Error: " + data.error;
      savePermissionsStatus.className = "save-status error";
      return;
    }
    permissionsJson.value = JSON.stringify(data.policy, null, 2);
    savePermissionsBtn.disabled = false;
    savePermissionsStatus.textContent = "";
  } catch (e) {
    savePermissionsStatus.textContent = "Error loading: " + e.message;
    savePermissionsStatus.className = "save-status error";
  }
}

savePermissionsBtn.addEventListener("click", async () => {
  savePermissionsBtn.disabled = true;
  savePermissionsStatus.textContent = "Saving...";
  savePermissionsStatus.className = "save-status";
  let parsed;
  try {
    parsed = JSON.parse(permissionsJson.value || "{}");
  } catch (e) {
    savePermissionsStatus.textContent = "Invalid JSON: " + e.message;
    savePermissionsStatus.className = "save-status error";
    savePermissionsBtn.disabled = false;
    return;
  }
  try {
    const res = await browser.mcpServer.setPermissionsConfig(parsed);
    if (res.error) {
      savePermissionsStatus.textContent = "Error: " + res.error;
      savePermissionsStatus.className = "save-status error";
    } else {
      savePermissionsStatus.textContent = "Saved.";
      permissionsJson.value = JSON.stringify(res.policy, null, 2);
    }
  } catch (e) {
    savePermissionsStatus.textContent = "Error: " + e.message;
    savePermissionsStatus.className = "save-status error";
  }
  savePermissionsBtn.disabled = false;
});

clearPermissionsBtn.addEventListener("click", async () => {
  savePermissionsStatus.textContent = "";
  try {
    const res = await browser.mcpServer.clearPermissionsConfig();
    if (res.error) {
      savePermissionsStatus.textContent = "Error: " + res.error;
      savePermissionsStatus.className = "save-status error";
      return;
    }
    permissionsJson.value = JSON.stringify(PERMISSION_PRESETS.empty, null, 2);
    savePermissionsStatus.textContent = "Cleared.";
    savePermissionsBtn.disabled = false;
  } catch (e) {
    savePermissionsStatus.textContent = "Error: " + e.message;
    savePermissionsStatus.className = "save-status error";
  }
});

document.querySelectorAll(".preset-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    const preset = PERMISSION_PRESETS[btn.dataset.preset];
    if (preset) {
      permissionsJson.value = JSON.stringify(preset, null, 2);
      savePermissionsStatus.textContent = "Preset loaded -- click Save policy to apply.";
      savePermissionsStatus.className = "save-status";
      // Re-enable Save regardless of prior state. If the initial
      // loadPermissionsConfig() errored and left the button disabled,
      // the user still needs to be able to commit the preset they just
      // picked -- the preset is local data, not dependent on a
      // successful load round-trip.
      savePermissionsBtn.disabled = false;
    }
  });
});

loadServerInfo().catch(e => console.error("thunderbird-mcp options:", "loadServerInfo failed:", e));
loadAccountAccess().catch(e => console.error("thunderbird-mcp options:", "loadAccountAccess failed:", e));
loadToolAccess().catch(e => console.error("thunderbird-mcp options:", "loadToolAccess failed:", e));
loadSkipReviewPref().catch(e => console.error("thunderbird-mcp options:", "loadSkipReviewPref failed:", e));
loadPermissionsConfig().catch(e => console.error("thunderbird-mcp options:", "loadPermissionsConfig failed:", e));
