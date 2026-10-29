"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { makeCi, makeSearchSession } = require("./search-terms.cjs");

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
function loadFolderTools({
  accounts, folders, isAccountAllowed = () => true,
  copyMessages = () => {}, copyFolder = () => {},
  getFolderForURL = uri => folders.find(folder => folder.URI === uri) || null,
  virtualFolderHelper = {}, saveVirtualFolders = () => {},
  isAddressBookAccessRestricted = () => false,
  openFolderDB = folder => folder.msgDatabase,
  openDBs = [], cachedDBForFolder = () => null, cachedDBForFilePath = () => null,
}) {
  // Thunderbird first escapes a UTF-8 byte string as a path (128/140), then
  // escapes slash as a filename component (153/158). Preserve both URI forms.
  const netUtil = {
    ESCAPE_URL_PATH: 1,
    ESCAPE_URL_FILE_BASENAME: 2,
    ESCAPE_URL_FILE_EXTENSION: 4,
    escapeString(bytes) {
      return Array.from(bytes, byte => /[A-Za-z0-9*+\-._/]/.test(byte)
        ? byte : `%${byte.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`).join("");
    },
    escapeURL(value) { return value.replace(/\//g, "%2F"); },
  };
  const runtime = vm.createContext({
    console,
    TextEncoder,
    TextDecoder,
    Ci: {
      ...makeCi(),
      nsMsgFolderFlags: { Virtual: 0x20, Trash: 0x100, AllMail: 0x8000, SpecialUse: 0x40405f00 },
      nsIMsgSearchSession: {}, nsINetUtil: netUtil, nsIMsgDBService: {},
    },
    Cc: {
      "@mozilla.org/messenger/searchSession;1": { createInstance: makeSearchSession },
      "@mozilla.org/network/io-service;1": { getService: () => netUtil },
      "@mozilla.org/msgDatabase/msgDBService;1": { getService: () => ({ openFolderDB, openDBs, cachedDBForFolder, cachedDBForFilePath }) },
    },
    ChromeUtils: { importESModule: () => ({ VirtualFolderHelper: virtualFolderHelper }) },
    Services: { prefs: {
      getCharPref: (_name, fallback) => fallback,
      getBoolPref: (_name, fallback) => fallback,
    } },
    isAccountAllowed,
    isAddressBookAccessRestricted,
    getConfiguredGetMessagesLimit: () => 10,
    MailServices: {
      accounts: {
        accounts,
        allFolders: folders,
        getAccount: key => accounts.find(account => account.key === key) || null,
        findAccountForServer: server => accounts.find(account => account.incomingServer === server),
        saveVirtualFolders,
      },
      folderLookup: { getFolderForURL },
      copy: { copyMessages, copyFolder },
    },
  });
  vm.runInContext([
    extractFunction("normalizeRfcMessageId"),
    extractFunction("messageIdForHeaderLookup"),
    extractFunction("lookupHeaderInDatabase"),
    snippet("// BEGIN FILTER SEARCH TERM HELPERS", "// END FILTER SEARCH TERM HELPERS"),
    snippet("// BEGIN CONTACT FIELD CONSTANTS", "// END CONTACT FIELD CONSTANTS"),
    snippet("// BEGIN OUTBOUND ATTACHMENT LIMITS", "// END OUTBOUND ATTACHMENT LIMITS"),
    snippet("// BEGIN MESSAGE HEADER LIMITS", "// END MESSAGE HEADER LIMITS"),
    snippet("// BEGIN SEARCH RESULT LIMITS", "// END SEARCH RESULT LIMITS"),
    snippet("// BEGIN BOUNDED REFRESH HELPERS", "// END BOUNDED REFRESH HELPERS"),
    snippet("// BEGIN TOOL SCHEMA BUILDER", "// END TOOL SCHEMA BUILDER"),
    snippet("function isFolderAccessible(", "function toColumnarTable("),
    snippet("function toColumnarTable(", "function listFolders("),
    snippet("function listFolders(", "function findIdentityIn("),
    snippet("function openFolder(", "function findMessage("),
    snippet("// BEGIN UPDATE MESSAGE TOOL", "// END UPDATE MESSAGE TOOL"),
    snippet("// BEGIN RAW MIME ATTACHMENT HELPERS", "// END RAW MIME ATTACHMENT HELPERS"),
    snippet("function createFolder(", "function deleteFolder("),
    snippet("function moveFolder(", "// BEGIN FILTER TOOL HANDLERS"),
    snippet("// BEGIN MCP TEXT SANITIZATION", "// END MCP TEXT SANITIZATION"),
    snippet("// BEGIN TOOL DISPATCH", "// END TOOL DISPATCH"),
  ].join("\n"), runtime);
  return runtime;
}

module.exports = { loadFolderTools };
