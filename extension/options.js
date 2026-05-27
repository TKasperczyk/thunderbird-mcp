/* global browser */
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

const currentAuthTokenInput = document.getElementById("currentAuthToken");
const copyAuthTokenBtn = document.getElementById("copyAuthTokenBtn");
const copyAuthTokenStatus = document.getElementById("copyAuthTokenStatus");
const useStableAuthTokenCheckbox = document.getElementById("useStableAuthToken");
const stableAuthTokenControls = document.getElementById("stableAuthTokenControls");
const stableAuthTokenInput = document.getElementById("stableAuthToken");
const generateStableAuthTokenBtn = document.getElementById("generateStableAuthTokenBtn");
const regenerateStableAuthTokenBtn = document.getElementById("regenerateStableAuthTokenBtn");
const stableAuthTokenStatus = document.getElementById("stableAuthTokenStatus");

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
    } else if (info.startError) {
      // Show the real error instead of a misleading "Running" label.
      statusDot.className = "status-dot stopped";
      statusText.textContent = "Failed: " + info.startError;
      serverPort.textContent = "--";
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

function updateStableAuthTokenControls() {
  stableAuthTokenControls.hidden = !useStableAuthTokenCheckbox.checked;
}

function setStableAuthTokenBusy(busy) {
  useStableAuthTokenCheckbox.disabled = busy;
  stableAuthTokenInput.disabled = busy;
  generateStableAuthTokenBtn.disabled = busy;
  regenerateStableAuthTokenBtn.disabled = busy;
}

function setStableAuthTokenStatus(message, error = false) {
  stableAuthTokenStatus.textContent = message;
  stableAuthTokenStatus.className = error ? "save-status error" : "save-status";
}

async function requestGeneratedAuthToken() {
  const result = await browser.mcpServer.generateAuthToken();
  if (result.error) {
    throw new Error(result.error);
  }
  if (!result.authToken) {
    throw new Error("Generated token was empty");
  }
  return result.authToken;
}

async function saveStableAuthTokenValue(value, successMessage = "Saved.") {
  const stableAuthToken = value.trim();
  setStableAuthTokenStatus("Saving...");
  const result = await browser.mcpServer.setStableAuthToken(stableAuthToken);
  if (result.error) {
    setStableAuthTokenStatus(result.error, true);
    return false;
  }
  stableAuthTokenInput.value = result.stableAuthToken || "";
  setStableAuthTokenStatus(successMessage);
  return true;
}

async function updateStoredStableAuthToken(value, successMessage) {
  setStableAuthTokenBusy(true);
  try {
    await saveStableAuthTokenValue(value, successMessage);
  } catch (e) {
    setStableAuthTokenStatus("Error: " + e.message, true);
  }
  setStableAuthTokenBusy(false);
  updateStableAuthTokenControls();
}

async function generateAndStoreStableAuthToken(successMessage) {
  setStableAuthTokenBusy(true);
  try {
    const token = await requestGeneratedAuthToken();
    stableAuthTokenInput.value = token;
    await saveStableAuthTokenValue(token, successMessage);
  } catch (e) {
    setStableAuthTokenStatus("Error: " + e.message, true);
  }
  setStableAuthTokenBusy(false);
  updateStableAuthTokenControls();
}

async function loadAuthenticationConfig() {
  try {
    const [current, stable] = await Promise.all([
      browser.mcpServer.getCurrentAuthToken(),
      browser.mcpServer.getStableAuthToken(),
    ]);
    currentAuthTokenInput.value = current.authToken || "";
    copyAuthTokenBtn.disabled = !currentAuthTokenInput.value;
    copyAuthTokenStatus.textContent = "";
    copyAuthTokenStatus.className = "save-status";

    stableAuthTokenInput.value = stable.stableAuthToken || "";
    useStableAuthTokenCheckbox.checked = !!stableAuthTokenInput.value;
    updateStableAuthTokenControls();
    setStableAuthTokenStatus("");
  } catch (e) {
    copyAuthTokenStatus.textContent = "Error loading token: " + e.message;
    copyAuthTokenStatus.className = "save-status error";
    setStableAuthTokenStatus("Error loading setting: " + e.message, true);
  }
}

copyAuthTokenBtn.addEventListener("click", async () => {
  copyAuthTokenStatus.textContent = "";
  copyAuthTokenStatus.className = "save-status";
  copyAuthTokenBtn.disabled = true;
  try {
    await navigator.clipboard.writeText(currentAuthTokenInput.value);
    copyAuthTokenStatus.textContent = "Copied.";
  } catch (e) {
    copyAuthTokenStatus.textContent = "Error: " + e.message;
    copyAuthTokenStatus.className = "save-status error";
  }
  copyAuthTokenBtn.disabled = !currentAuthTokenInput.value;
});

useStableAuthTokenCheckbox.addEventListener("change", async () => {
  updateStableAuthTokenControls();
  if (useStableAuthTokenCheckbox.checked) {
    if (!stableAuthTokenInput.value.trim()) {
      await generateAndStoreStableAuthToken("Generated and saved.");
    } else {
      await updateStoredStableAuthToken(stableAuthTokenInput.value, "Saved.");
    }
  } else {
    stableAuthTokenInput.value = "";
    await updateStoredStableAuthToken("", "Stable token cleared.");
  }
});

stableAuthTokenInput.addEventListener("change", async () => {
  if (!useStableAuthTokenCheckbox.checked) {
    return;
  }
  if (!stableAuthTokenInput.value.trim()) {
    await generateAndStoreStableAuthToken("Generated and saved.");
  } else {
    await updateStoredStableAuthToken(stableAuthTokenInput.value, "Saved.");
  }
});

generateStableAuthTokenBtn.addEventListener("click", async () => {
  await generateAndStoreStableAuthToken("Generated and saved.");
});

regenerateStableAuthTokenBtn.addEventListener("click", async () => {
  await generateAndStoreStableAuthToken("Regenerated and saved.");
});

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
    currentTools = data.tools || [];
    const groupLabels = data.groups || {};

    if (currentTools.length === 0) {
      toolList.innerHTML = "<li>No tools found.</li>";
      return;
    }

    toolList.innerHTML = "";

    // Tools arrive pre-sorted by group then CRUD order from the server.
    // Build grouped structure from tool metadata.
    let currentGroup = null;
    let currentCrud = null;

    for (const tool of currentTools) {
      const group = tool.group || "other";
      const crud = tool.crud || "other";

      // New group header
      if (group !== currentGroup) {
        currentGroup = group;
        currentCrud = null;
        const header = document.createElement("li");
        header.className = "tool-group-header";
        header.textContent = groupLabels[group] || group.charAt(0).toUpperCase() + group.slice(1);
        toolList.appendChild(header);
      }

      // New CRUD sub-header within group
      if (crud !== currentCrud) {
        currentCrud = crud;
        const subHeader = document.createElement("li");
        subHeader.className = "tool-crud-header";
        subHeader.textContent = CRUD_LABELS[crud] || crud.charAt(0).toUpperCase() + crud.slice(1);
        toolList.appendChild(subHeader);
      }

      const li = document.createElement("li");
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
const blockFilterForwardReplyCheckbox = document.getElementById("blockFilterForwardReply");
const blockContactWritesCheckbox = document.getElementById("blockContactWrites");
const blockMailboxExportCheckbox = document.getElementById("blockMailboxExport");
const saveSkipReviewBtn = document.getElementById("saveSkipReviewBtn");
const saveSkipReviewStatus = document.getElementById("saveSkipReviewStatus");

async function loadSafeguardPrefs() {
  try {
    const [skip, filter, contacts, exportPref] = await Promise.all([
      browser.mcpServer.getBlockSkipReview(),
      browser.mcpServer.getBlockFilterForwardReply(),
      browser.mcpServer.getBlockContactWrites(),
      browser.mcpServer.getBlockMailboxExport(),
    ]);
    blockSkipReviewCheckbox.checked = !!skip.blockSkipReview;
    blockFilterForwardReplyCheckbox.checked = !!filter.blockFilterForwardReply;
    blockContactWritesCheckbox.checked = !!contacts.blockContactWrites;
    blockMailboxExportCheckbox.checked = !!exportPref.blockMailboxExport;
    saveSkipReviewBtn.disabled = false;
    saveSkipReviewStatus.textContent = "";
  } catch (e) {
    saveSkipReviewStatus.textContent = "Error loading settings: " + e.message;
    saveSkipReviewStatus.className = "save-status error";
  }
}

saveSkipReviewBtn.addEventListener("click", async () => {
  saveSkipReviewBtn.disabled = true;
  saveSkipReviewStatus.textContent = "Saving...";
  saveSkipReviewStatus.className = "save-status";
  try {
    // Persist all four in parallel so one click writes a consistent state.
    // If any individual setter returns an error, surface it but continue
    // saving the others -- partial application is better than total revert.
    const results = await Promise.all([
      browser.mcpServer.setBlockSkipReview(blockSkipReviewCheckbox.checked),
      browser.mcpServer.setBlockFilterForwardReply(blockFilterForwardReplyCheckbox.checked),
      browser.mcpServer.setBlockContactWrites(blockContactWritesCheckbox.checked),
      browser.mcpServer.setBlockMailboxExport(blockMailboxExportCheckbox.checked),
    ]);
    const errors = results.filter(r => r && r.error).map(r => r.error);
    if (errors.length > 0) {
      saveSkipReviewStatus.textContent = errors.join("; ");
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

// ── Audit log viewer ─────────────────────────────────────────────────────────

const auditToolFilter = document.getElementById("auditToolFilter");
const refreshAuditBtn = document.getElementById("refreshAuditBtn");
const clearAuditBtn = document.getElementById("clearAuditBtn");
const auditEntriesEl = document.getElementById("auditEntries");
const auditStatusEl = document.getElementById("auditStatus");

function formatAuditEntry(entry) {
  const ts = entry.ts || "(no ts)";
  const tool = entry.tool || "(unknown tool)";
  // Strip the well-known top-level fields and JSON-stringify the rest for
  // free-form display. We deliberately do not deep-walk attempts to redact;
  // appendComposeAudit already keeps fields metadata-only.
  const { ts: _ts, tool: _tool, ...rest } = entry;
  const detail = Object.keys(rest).length ? " " + JSON.stringify(rest) : "";
  const div = document.createElement("div");
  div.className = "audit-entry";
  const tsSpan = document.createElement("span");
  tsSpan.className = "audit-ts";
  tsSpan.textContent = ts + "  ";
  const toolSpan = document.createElement("span");
  toolSpan.className = "audit-tool";
  toolSpan.textContent = tool;
  const detSpan = document.createElement("span");
  detSpan.textContent = detail;
  div.appendChild(tsSpan);
  div.appendChild(toolSpan);
  div.appendChild(detSpan);
  return div;
}

async function loadAuditLog() {
  auditEntriesEl.textContent = "Loading...";
  auditStatusEl.textContent = "";
  auditStatusEl.className = "save-status";
  const tool = auditToolFilter.value || "";
  try {
    // Avoid passing `undefined` for the second arg -- some WebExtensions
    // schema validators reject explicit-undefined even when the parameter
    // is declared optional. Call with one arg when there's no filter, two
    // args when there is.
    const result = tool
      ? await browser.mcpServer.readAuditLog(200, { tool })
      : await browser.mcpServer.readAuditLog(200);
    auditEntriesEl.innerHTML = "";
    if (!result || !Array.isArray(result.entries) || result.entries.length === 0) {
      const empty = document.createElement("div");
      empty.className = "audit-empty";
      empty.textContent = "(no entries)";
      auditEntriesEl.appendChild(empty);
      return;
    }
    for (const e of result.entries) {
      auditEntriesEl.appendChild(formatAuditEntry(e));
    }
    if (result.truncated) {
      const trunc = document.createElement("div");
      trunc.className = "audit-empty";
      trunc.textContent = "(more entries exist; showing newest " + result.entries.length + ")";
      auditEntriesEl.appendChild(trunc);
    }
  } catch (e) {
    auditEntriesEl.textContent = "";
    auditStatusEl.textContent = "Error: " + e.message;
    auditStatusEl.className = "save-status error";
  }
}

refreshAuditBtn.addEventListener("click", () => {
  loadAuditLog().catch(e => console.error("thunderbird-mcp options:", "audit refresh failed:", e));
});
auditToolFilter.addEventListener("change", () => {
  loadAuditLog().catch(e => console.error("thunderbird-mcp options:", "audit refresh failed:", e));
});
clearAuditBtn.addEventListener("click", async () => {
  if (!confirm("Delete the entire audit log? This cannot be undone.")) return;
  clearAuditBtn.disabled = true;
  auditStatusEl.textContent = "Clearing...";
  auditStatusEl.className = "save-status";
  try {
    const result = await browser.mcpServer.clearAuditLog();
    if (result && result.error) {
      auditStatusEl.textContent = "Error: " + result.error;
      auditStatusEl.className = "save-status error";
    } else {
      auditStatusEl.textContent = "Cleared (" + (result.bytesRemoved || 0) + " bytes removed).";
    }
  } catch (e) {
    auditStatusEl.textContent = "Error: " + e.message;
    auditStatusEl.className = "save-status error";
  }
  clearAuditBtn.disabled = false;
  await loadAuditLog();
});

loadServerInfo().catch(e => console.error("thunderbird-mcp options:", "loadServerInfo failed:", e));
loadAuthenticationConfig().catch(e => console.error("thunderbird-mcp options:", "loadAuthenticationConfig failed:", e));
loadAccountAccess().catch(e => console.error("thunderbird-mcp options:", "loadAccountAccess failed:", e));
loadToolAccess().catch(e => console.error("thunderbird-mcp options:", "loadToolAccess failed:", e));
loadSafeguardPrefs().catch(e => console.error("thunderbird-mcp options:", "loadSafeguardPrefs failed:", e));
loadAuditLog().catch(e => console.error("thunderbird-mcp options:", "loadAuditLog failed:", e));
