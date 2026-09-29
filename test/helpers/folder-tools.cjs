"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../../extension/mcp_server/api.js"), "utf8");

function snippet(startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0, `Missing api.js marker: ${startMarker}`);
  assert.ok(end > start, `Missing api.js marker: ${endMarker}`);
  return source.slice(start, end);
}

function extractFunction(name) {
  const marker = `function ${name}(`;
  const start = source.indexOf(marker);
  assert.ok(start >= 0, `${name} missing`);
  const brace = source.indexOf("{", start);
  let depth = 0;
  for (let i = brace; i < source.length; i++) {
    if (source[i] === "{") depth++;
    else if (source[i] === "}") {
      depth--;
      if (depth === 0) return source.slice(start, i + 1);
    }
  }
  assert.fail(`unterminated ${name}`);
}

// Load real folder access, handlers, dispatch and sanitization; mock only the
// account preference and native Thunderbird services they call.
function loadFolderTools({ accounts, folders, isAccountAllowed = () => true, copyMessages = () => {}, copyFolder = () => {} }) {
  const runtime = vm.createContext({
    console,
    Ci: { nsMsgFolderFlags: { Trash: 0x100 } },
    isAccountAllowed,
    buildTools: () => [{ name: "listFolders", group: "system" }, { name: "updateMessage", group: "messages" }],
    MailServices: {
      accounts: {
        accounts,
        findAccountForServer: server => accounts.find(account => account.incomingServer === server),
      },
      folderLookup: { getFolderForURL: uri => folders.find(folder => folder.URI === uri) || null },
      copy: { copyMessages, copyFolder },
    },
  });
  vm.runInContext([
    extractFunction("normalizeRfcMessageId"),
    extractFunction("messageIdForHeaderLookup"),
    extractFunction("lookupHeaderInDatabase"),
    snippet("// BEGIN OUTBOX DESTINATION GUARD", "// END OUTBOX DESTINATION GUARD"),
    snippet("function isFolderAccessible(", "function toColumnarTable("),
    snippet("function toColumnarTable(", "function listAccounts("),
    snippet("function listFolders(", "function findIdentityIn("),
    snippet("function openFolder(", "function findMessage("),
    snippet("// BEGIN UPDATE MESSAGE TOOL", "// END UPDATE MESSAGE TOOL"),
    snippet("function createFolder(", "function deleteFolder("),
    snippet("function moveFolder(", "// BEGIN FILTER TOOL HANDLERS"),
    snippet("// BEGIN MCP TEXT SANITIZATION", "// END MCP TEXT SANITIZATION"),
    snippet("// BEGIN TOOL DISPATCH", "// END TOOL DISPATCH"),
  ].join("\n"), runtime);
  return runtime;
}

module.exports = { loadFolderTools };
