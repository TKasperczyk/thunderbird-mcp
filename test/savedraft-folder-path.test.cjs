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

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");

function snippet(name) {
  const startMarker = `// BEGIN ${name}`;
  const endMarker = `// END ${name}`;
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

const Ci = {
  nsIMsgComposeParams: Symbol("nsIMsgComposeParams"),
  nsIMsgCompFields: Symbol("nsIMsgCompFields"),
  nsIMsgCompType: { New: 0 },
  nsIMsgCompDeliverMode: { SaveAsDraft: 4 },
};

const FOLDER = "imap://user@example.com/Drafts";

function loadSaveDraft(identity, options = {}) {
  const composeFields = { to: "", cc: "", bcc: "", subject: "", body: "", setHeader() {} };
  const params = {};
  const lookups = [];

  const sandbox = {
    Ci,
    Cc: {
      "@mozilla.org/messengercompose/composeparams;1": { createInstance: () => params },
      "@mozilla.org/messengercompose/composefields;1": { createInstance: () => composeFields },
    },
    setComposeIdentity: (p) => { p.identity = identity; return null; },
    resolveComposeFormat: () => ({ useHtml: false, format: 1 }),
    buildBodyWithSignature: body => body || "",
    filePathsToAttachDescs: () => ({ descs: [], failed: [] }),
    getAccessibleFolder(uri) {
      lookups.push(uri);
      return options.getAccessibleFolder
        ? options.getAccessibleFolder(uri)
        : { folder: { URI: uri } };
    },
    sendMessageDirectly() {
      if (options.onSave) options.onSave();
      return Promise.resolve(options.result || { success: true });
    },
    console,
  };

  vm.createContext(sandbox);
  vm.runInContext([
    snippet("DRAFT HELPERS"),
    snippet("SAVE DRAFT TOOL"),
    "this.saveDraft = saveDraft;",
  ].join("\n"), sandbox);
  return { saveDraft: sandbox.saveDraft, lookups };
}

describe("saveDraft reports the drafts folder", () => {
  for (const property of ["draftsFolderURI", "draftFolder"]) {
    for (const value of [FOLDER, { URI: FOLDER }]) {
      it(`resolves an accessible ${typeof value} ${property}`, async () => {
        const { saveDraft, lookups } = loadSaveDraft({ [property]: value });
        const result = await saveDraft("a@example.com", "Hi", "body");

        assert.equal(result.success, true);
        assert.equal(result.folderPath, FOLDER);
        assert.deepEqual(lookups, [FOLDER]);
      });
    }
  }

  it("moves on when a property throws instead of existing", async () => {
    const identity = {
      get draftsFolderURI() { throw new Error("not on this version"); },
      draftFolder: FOLDER,
    };
    const { saveDraft } = loadSaveDraft(identity);
    assert.equal((await saveDraft("a@example.com", "Hi", "body")).folderPath, FOLDER);
  });

  it("ignores values that are not folder URIs", async () => {
    const { saveDraft } = loadSaveDraft({ draftsFolderURI: "Drafts", draftFolder: FOLDER });
    assert.equal((await saveDraft("a@example.com", "Hi", "body")).folderPath, FOLDER);
  });

  it("still saves when the identity names no drafts folder", async () => {
    const { saveDraft, lookups } = loadSaveDraft({});
    const result = await saveDraft("a@example.com", "Hi", "body");
    assert.equal(result.success, true);
    assert.equal(result.message, "Draft saved");
    assert.equal(result.folderPath, undefined);
    assert.deepEqual(lookups, []);
  });

  it("does not probe the Sent-folder or nonexistent identity properties", async () => {
    let forbiddenReads = 0;
    const identity = {
      get draftsFolder() { forbiddenReads++; return FOLDER; },
      get fccFolder() { forbiddenReads++; return "imap://user@example.com/Sent"; },
    };
    const { saveDraft, lookups } = loadSaveDraft(identity);
    const result = await saveDraft("a@example.com", "Hi", "body");

    assert.equal(result.success, true);
    assert.equal(result.folderPath, undefined);
    assert.equal(forbiddenReads, 0);
    assert.deepEqual(lookups, []);
  });

  it("uses the identity folder after the save completes", async () => {
    const oldFolder = "imap://user@example.com/OldDrafts";
    const identity = { draftFolder: oldFolder };
    const { saveDraft, lookups } = loadSaveDraft(identity, {
      onSave() { identity.draftFolder = FOLDER; },
    });
    const result = await saveDraft("a@example.com", "Hi", "body");

    assert.equal(result.success, true);
    assert.equal(result.folderPath, FOLDER);
    assert.deepEqual(lookups, [FOLDER]);
  });

  for (const identity of [
    null,
    { draftsFolderURI: null, draftFolder: undefined },
    { draftsFolderURI: "", draftFolder: "Drafts" },
    { draftsFolderURI: { URI: "Drafts" }, draftFolder: { URI: 42 } },
    { get draftsFolderURI() { throw new Error("missing property"); }, get draftFolder() { throw new Error("missing property"); } },
    { draftsFolderURI: { get URI() { throw new Error("missing URI"); } }, draftFolder: {} },
  ]) {
    it("keeps successful saves without a usable identity folder", async () => {
      const { saveDraft, lookups } = loadSaveDraft(identity);
      const result = await saveDraft("a@example.com", "Hi", "body");

      assert.equal(result.success, true);
      assert.equal(result.folderPath, undefined);
      assert.deepEqual(lookups, []);
    });
  }

  it("omits restricted Drafts URIs and the access error containing them", async () => {
    const restricted = "imap://restricted@example.com/Drafts";
    const { saveDraft, lookups } = loadSaveDraft({ draftFolder: restricted }, {
      getAccessibleFolder(uri) { return { error: `Account not accessible for folder: ${uri}` }; },
    });
    const result = await saveDraft("a@example.com", "Hi", "body");

    assert.equal(result.success, true);
    assert.equal(result.folderPath, undefined);
    assert.equal(result.error, undefined);
    assert.equal(JSON.stringify(result).includes(restricted), false);
    assert.deepEqual(lookups, [restricted]);
  });

  it("omits folder details when accessibility lookup throws", async () => {
    const { saveDraft } = loadSaveDraft({ draftFolder: FOLDER }, {
      getAccessibleFolder(uri) { throw new Error(`Lookup failed for ${uri}`); },
    });
    const result = await saveDraft("a@example.com", "Hi", "body");

    assert.equal(result.success, true);
    assert.equal(result.folderPath, undefined);
    assert.equal(result.error, undefined);
    assert.equal(JSON.stringify(result).includes(FOLDER), false);
  });

  it("does not report a folder after a failed save", async () => {
    const { saveDraft, lookups } = loadSaveDraft({ draftFolder: FOLDER }, {
      result: { error: "Save failed" },
    });
    const result = await saveDraft("a@example.com", "Hi", "body");

    assert.equal(result.error, "Save failed");
    assert.equal(result.folderPath, undefined);
    assert.deepEqual(lookups, []);
  });
});
