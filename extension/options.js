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

let currentAccounts = [];
let currentTools = [];

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
    if (info.buildCommit) {
      const date = info.buildDate ? " (" + info.buildDate.replace("T", " ").replace(/\.\d+Z$/, " UTC") + ")" : "";
      buildInfo.textContent = info.buildCommit + date;
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
    accountList.innerHTML = "<li>Error loading accounts: " + e.message + "</li>";
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

    if (currentTools.length === 0) {
      toolList.innerHTML = "<li>No tools found.</li>";
      return;
    }

    toolList.innerHTML = "";
    for (const tool of currentTools) {
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
    toolList.innerHTML = "<li>Error loading tools: " + e.message + "</li>";
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

loadServerInfo();
loadAccountAccess();
loadToolAccess();
