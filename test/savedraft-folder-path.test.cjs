"use strict";

/**
 * saveDraft reports the folder the draft went to.
 *
 * The caller only passes `from`; which folder that identity uses for
 * drafts is Thunderbird's business. These tests pin that the folder is
 * read off the identity, that the property probing copes with the
 * different shapes across versions, and that a missing property never
 * turns a successful save into an error.
 */

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const Ci = {
  nsIMsgComposeParams: Symbol("nsIMsgComposeParams"),
  nsIMsgCompFields: Symbol("nsIMsgCompFields"),
  nsIMsgCompType: { New: 0 },
  nsIMsgCompDeliverMode: { SaveAsDraft: 4 },
};

const FOLDER = "imap://user@example.com/Drafts";

function loadSaveDraft(identity) {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const start = source.indexOf("function saveDraft(to, subject,");
  const end = source.indexOf("function replyToMessage(", start);
  assert.ok(start >= 0, "saveDraft start marker missing");
  assert.ok(end > start, "saveDraft end marker missing");

  const composeFields = { to: "", cc: "", bcc: "", subject: "", body: "", setHeader() {} };
  const params = {};

  const sandbox = {
    Ci,
    Cc: {
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => params },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => composeFields },
    },
    setComposeIdentity: (p) => { p.identity = identity; return null; },
    resolveComposeFormat: () => ({ useHtml: false, format: 1 }),
    formatBodyHtml: (body) => body || "",
    filePathsToAttachDescs: () => ({ descs: [], failed: [] }),
    sendMessageDirectly: () => Promise.resolve({ success: true }),
    console,
  };

  vm.createContext(sandbox);
  vm.runInContext(`${source.slice(start, end)}\nthis.saveDraft = saveDraft;`, sandbox);
  return sandbox.saveDraft;
}

describe("saveDraft reports the drafts folder", () => {
  it("reads a URI string off the identity", async () => {
    const saveDraft = loadSaveDraft({ draftsFolderURI: FOLDER });
    assert.equal((await saveDraft("a@example.com", "Hi", "body")).folderPath, FOLDER);
  });

  it("accepts a folder object with a URI property", async () => {
    const saveDraft = loadSaveDraft({ draftFolder: { URI: FOLDER } });
    assert.equal((await saveDraft("a@example.com", "Hi", "body")).folderPath, FOLDER);
  });

  it("moves on when a property throws instead of existing", async () => {
    const identity = {
      get draftsFolderURI() { throw new Error("not on this version"); },
      draftFolder: FOLDER,
    };
    assert.equal((await loadSaveDraft(identity)("a@example.com", "Hi", "body")).folderPath, FOLDER);
  });

  it("ignores values that are not folder URIs", async () => {
    const saveDraft = loadSaveDraft({ draftsFolderURI: "Drafts", draftFolder: FOLDER });
    assert.equal((await saveDraft("a@example.com", "Hi", "body")).folderPath, FOLDER);
  });

  it("still saves when the identity names no drafts folder", async () => {
    const result = await loadSaveDraft({})("a@example.com", "Hi", "body");
    assert.equal(result.success, true);
    assert.equal(result.message, "Draft saved");
    assert.equal(result.folderPath, undefined);
  });
});
