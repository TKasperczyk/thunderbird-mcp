"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { loadRefreshRuntime } = require("./helpers/refresh-runtime.cjs");
const { makeImapFolder } = require("./helpers/imap-folder.cjs");

const WAIT_MS = 1000;
const UPDATE_MS = 50;
const MIXED_CALLER_WAIT_MS = 75;
const BUSY_UPDATE_MS = 300;
const BUSY_START_DELAY_MS = 10;
const STORE_MS = 40;
const FAILURE = 0x80004005;
const FLAGS = { Virtual: 0x20, ImapNoselect: 0x01000000 };
const FOLDER_URI = "imap://account/Inbox";
const FRESH_UPDATE = { requireFreshUpdate: true };

// Keep the original count fixture on the shared IMAP event-order model.
function loadRefresh({ type = "imap", offline = false, denied = false, delay = UPDATE_MS, exitCode = 0, mfn } = {}) {
  let harness;
  const folder = {
    URI: FOLDER_URI, server: { type }, flags: 0, isServer: false,
    total: 10, unread: 2, arrivals: 2, countReads: 0,
    getTotalMessages() { this.countReads++; return this.total; },
    getNumUnread() { this.countReads++; return this.unread; },
    QueryInterface() { return this; },
    getNewMessages() { assert.fail("account-level Get Mail must not be used"); },
  };
  const { listeners, mailSession } = makeImapFolder({
    folder, clock: () => harness.clock, runtime: () => harness.runtime, delay, exitCode,
    onUpdateComplete(current, status) {
      if (status === 0) { current.total += current.arrivals; current.unread += current.arrivals; }
    },
  });
  harness = loadRefreshRuntime({
    blocks: ["REFRESH FOLDER TOOL"],
    expose: ["refreshFolder", "waitForFolderRefresh"],
    globals: {
      Services: { io: { offline } },
      Ci: { nsMsgFolderFlags: FLAGS, nsIFolderListener: { event: 0x80 }, nsIMsgImapMailFolder: "imap", nsIMsgMailNewsUrl: "mail-url" },
      MailServices: { mailSession, ...(mfn ? { mfn } : {}) },
      getAccessibleFolder: uri => denied ? { error: "Account not accessible" }
        : uri === folder.URI ? { folder } : { error: "Folder not found" },
    },
  });
  return { ...harness, folder, listeners, mailSession, refresh: timeout => harness.runtime.refreshFolder(folder.URI, timeout ?? WAIT_MS) };
}

describe("refreshFolder IMAP completion", () => {
  it("waits for the update, confirms its successful URL, and reports net counts", async () => {
    const h = loadRefresh();
    const pending = h.refresh();
    assert.equal(h.listeners.size, 1);
    await h.clock.advance(UPDATE_MS);
    const result = await pending;
    assert.equal(result.success, true);
    assert.equal(result.status, "loaded");
    assert.equal(result.updateConfirmed, true);
    assert.equal(result.warning, undefined);
    assert.equal(result.totalBefore, 10);
    assert.equal(result.totalAfter, 12);
    assert.equal(result.unreadAfter, 4);
    assert.equal(result.netMessageDelta, 2);
    assert.equal(result.elapsedMs, UPDATE_MS);
    assert.equal(h.folder.imapUpdateCalls, 1);
    assert.equal(h.folder.updateCalls, 0);
    assert.equal(h.listeners.size, 0);
    assert.equal(h.clock.activeTimers.length, 0);
  });

  it("lets a failing update exit code beat the preceding FolderLoaded", async () => {
    const h = loadRefresh({ exitCode: FAILURE });
    const pending = h.refresh();
    await h.clock.advance(UPDATE_MS);
    const result = await pending;
    assert.equal(result.status, "error");
    assert.equal(result.success, false);
    assert.equal(result.updateConfirmed, false);
    assert.match(result.error, /0x80004005/);
    assert.equal(result.netMessageDelta, 0);
    assert.equal(h.listeners.size, 0);
  });

  for (const exitCode of [0, FAILURE]) {
    it(`follows a pre-existing update with an owned update (exit ${exitCode})`, async () => {
      const h = loadRefresh({ delay: BUSY_UPDATE_MS, exitCode });
      h.folder.updateFolder(); // The existing fire-and-forget folder-read path.
      await h.clock.advance(BUSY_START_DELAY_MS);
      const pending = h.refresh();
      assert.equal(h.folder.urlListener, null, "busy update must discard the new listener");
      await h.clock.advance(BUSY_UPDATE_MS - BUSY_START_DELAY_MS);
      assert.equal(h.listeners.size, 1, "wait for the trailing update");
      await h.clock.advance(BUSY_UPDATE_MS);
      const result = await pending;
      assert.equal(result.status, exitCode === 0 ? "loaded" : "error");
      assert.equal(result.updateConfirmed, exitCode === 0);
      assert.equal(result.warning, undefined);
      assert.equal(result.netMessageDelta, exitCode === 0 ? 4 : 0);
      assert.equal(result.elapsedMs, 2 * BUSY_UPDATE_MS - BUSY_START_DELAY_MS);
      assert.equal(h.folder.imapUpdateCalls, 3, "native update, busy call, one trailing update");
      assert.equal(h.listeners.size, 0);
    });
  }

  it("reports synchronous FolderLoaded during a body fetch as no update ran", async () => {
    const h = loadRefresh();
    h.folder.urlRunning = true;
    const pending = h.refresh();
    await h.clock.flush();
    const result = await pending;
    assert.equal(result.status, "loaded");
    assert.equal(result.updateConfirmed, false);
    assert.match(result.warning, /No update ran/);
    assert.equal(result.elapsedMs, 0);
    assert.equal(h.folder.updateUrl, null);
    assert.equal(h.folder.imapUpdateCalls, 2, "no more than one trailing attempt");
    assert.equal(h.listeners.size, 0);
  });

  for (const exitCode of [0, FAILURE]) {
    it(`ignores a non-update URL consuming the listener (exit ${exitCode})`, async () => {
      const h = loadRefresh({ delay: BUSY_UPDATE_MS });
      // A STORE queued ahead of SELECT can finish first. A STORE issued after
      // SELECT normally waits, so this deliberately models the earlier queue.
      const storeUrl = { updatingFolder: false, QueryInterface() { return this; } };
      h.clock.schedule(STORE_MS, () => h.folder.finishUrl(storeUrl, exitCode));
      let completed = false;
      const pending = h.refresh().then(result => { completed = true; return result; });
      await h.clock.advance(STORE_MS);
      assert.equal(completed, false);
      assert.equal(h.folder.urlListener, null);
      await h.clock.advance(BUSY_UPDATE_MS - STORE_MS);
      assert.equal(completed, false);
      await h.clock.advance(BUSY_UPDATE_MS);
      const result = await pending;
      assert.equal(result.status, "loaded");
      assert.equal(result.updateConfirmed, true);
      assert.equal(result.netMessageDelta, 4);
      assert.equal(result.warning, undefined);
    });
  }

  it("ignores an unidentifiable URL instead of treating it as a successful update", async () => {
    const h = loadRefresh();
    let completed = false;
    const pending = h.refresh().then(result => { completed = true; return result; });
    h.folder.urlListener.OnStopRunningUrl(null, 0);
    await h.clock.flush();
    assert.equal(completed, false);
    await h.clock.advance(UPDATE_MS);
    assert.equal((await pending).updateConfirmed, true);
  });

  it("cleans up when starting an update throws", async () => {
    const h = loadRefresh();
    h.folder.updateFolderWithListener = () => { throw new Error("Cannot start update"); };
    const result = await h.refresh();
    assert.equal(result.status, "error");
    assert.match(result.error, /Cannot start update/);
    assert.equal(h.listeners.size, 0);
    assert.equal(h.clock.activeTimers.length, 0);
  });
});

describe("refreshFolder lifecycle and access", () => {
  it("keeps listening after a caller timeout and lets a fresh caller join", async () => {
    const h = loadRefresh({ delay: null });
    const first = h.refresh();
    await h.clock.advance(WAIT_MS);
    const timeout = await first;
    assert.equal(timeout.status, "timeout");
    assert.equal(timeout.success, false);
    assert.equal(timeout.updateConfirmed, false);
    assert.match(timeout.error, /not confirmed/);
    assert.equal(h.listeners.size, 1);
    const second = h.refresh();
    h.folder.finishUrl(h.folder.updateUrl, 0);
    await h.clock.flush();
    const result = await second;
    assert.equal(result.status, "loaded");
    assert.equal(result.joinedInFlight, true);
    assert.equal(h.folder.imapUpdateCalls, 1);
    assert.equal(h.listeners.size, 0);
  });

  it("retries a stale folder wait and retires its observer without renewing age", async () => {
    const h = loadRefresh({ delay: null });
    const first = h.refresh();
    await h.clock.advance(WAIT_MS);
    await first;
    await h.clock.advance(h.runtime.REFRESH_MAX_TIMEOUT_MS - WAIT_MS);
    const second = h.refresh();
    assert.equal(h.folder.imapUpdateCalls, 2);
    assert.equal(h.listeners.size, 1, "the old observer was retired");
    h.folder.finishUrl(h.folder.updateUrl);
    await h.clock.flush();
    assert.equal(h.folder.imapUpdateCalls, 3);
    h.folder.finishUrl(h.folder.updateUrl);
    await h.clock.flush();
    const result = await second;
    assert.equal(result.joinedInFlight, undefined);
    assert.equal(result.status, "loaded");
    assert.equal(h.listeners.size, 0);
  });

  it("provides a 5-second structured adapter without reading message counts", async () => {
    const h = loadRefresh({ delay: null });
    const searchWaitMs = 5000;
    const pending = h.runtime.waitForFolderRefresh(h.folder, searchWaitMs);
    await h.clock.advance(searchWaitMs);
    const result = await pending;
    assert.equal(result.status, "timeout");
    assert.equal(result.joinedInFlight, false);
    assert.equal(result.folderPath, undefined);
    assert.equal(h.folder.countReads, 0);
    h.folder.finishUrl(h.folder.updateUrl, 0);
    await h.clock.flush();
  });

  it("refuses offline calls without starting or joining a provider", async () => {
    const h = loadRefresh({ offline: true });
    for (const call of [() => h.refresh(), () => h.runtime.waitForFolderRefresh(h.folder, WAIT_MS, FRESH_UPDATE)]) {
      const result = await call();
      assert.equal(result.status, "error");
      assert.match(result.error, /offline/);
    }
    assert.equal(h.folder.imapUpdateCalls, 0);
    assert.equal(h.folder.countReads, 0);
    assert.equal(h.listeners.size, 0);
    assert.equal(h.clock.activeTimers.length, 0);
  });

  it("does not join an existing refresh after Thunderbird goes offline", async () => {
    const h = loadRefresh({ delay: null });
    const first = h.refresh();
    await h.clock.advance(WAIT_MS);
    await first;
    h.runtime.Services.io.offline = true;
    const result = await h.refresh();
    assert.equal(result.status, "error");
    assert.match(result.error, /offline/);
    assert.equal(result.joinedInFlight, undefined);
    assert.equal(h.folder.imapUpdateCalls, 1);
    await h.clock.advance(h.runtime.REFRESH_ABANDON_MS);
    assert.equal(h.listeners.size, 0);
  });

  it("rejects inaccessible and unknown folders before provider or count access", async () => {
    const denied = loadRefresh({ denied: true });
    assert.match((await denied.refresh()).error, /not accessible/);
    assert.equal(denied.folder.countReads, 0);
    assert.equal(denied.folder.imapUpdateCalls, 0);
    const unknown = loadRefresh();
    assert.match((await unknown.runtime.refreshFolder("missing")).error, /not found/);
  });

  for (const type of ["imap", "owl", "none", "pop3"]) {
    it(`refuses a ${type} account root without walking descendants`, async () => {
      const h = loadRefresh({ type });
      h.folder.isServer = true;
      Object.defineProperty(h.folder, "descendants", { get() { return assert.fail("must not fan out"); } });
      assert.match((await h.refresh()).error, /Account roots/);
      assert.equal(h.folder.updateCalls + h.folder.imapUpdateCalls, 0);
      assert.equal(h.folder.countReads, 0);
    });
  }

  it("refuses saved searches and no-select folders, including local saved searches", async () => {
    for (const [type, flag, message] of [["imap", FLAGS.ImapNoselect, /no-select/], ["owl", FLAGS.Virtual, /virtual/], ["none", FLAGS.Virtual, /virtual/]]) {
      const h = loadRefresh({ type });
      h.folder.flags = flag;
      assert.match((await h.refresh()).error, message);
      assert.equal(h.folder.updateCalls + h.folder.imapUpdateCalls, 0);
    }
  });

  it("skips individual news folders", async () => {
    const h = loadRefresh({ type: "nntp" });
    const result = await h.refresh();
    assert.equal(result.success, true);
    assert.match(result.skipped, /not supported/);
    assert.equal(h.folder.updateCalls, 0);
  });

  it("refuses rebuildIndex on server folders without starting an update", async () => {
    for (const type of ["imap", "owl", "ews"]) {
      const h = loadRefresh({ type });
      const result = await h.runtime.refreshFolder(h.folder.URI, WAIT_MS, true);
      assert.match(result.error, /rebuildIndex applies only to Local Folders, POP and RSS folders/);
      assert.equal(h.folder.updateCalls + h.folder.imapUpdateCalls, 0);
      assert.equal(h.listeners.size, 0);
    }
  });
});

// A local store's folder with the calls Repair Folder makes, recorded in order.
function loadLocal(options = {}) {
  const calls = [];
  const mfn = { notifyFolderReindexTriggered(folder) { calls.push(["reindex", folder.URI]); } };
  const h = loadRefresh({ type: "none", ...options, mfn });
  let valid = true;
  const update = h.folder.updateFolder.bind(h.folder);
  Object.assign(h.folder, {
    msgDatabase: {
      get summaryValid() { return valid; },
      set summaryValid(value) { calls.push(["summaryValid", value]); valid = value; },
    },
    closeAndBackupFolderDB(name) { calls.push(["closeAndBackup", name]); },
    ForceDBClosed() { calls.push(["forceClosed"]); },
    msgStore: { ensureLocalStore(folder) { calls.push(["ensureLocalStore", folder.URI]); } },
    updateFolder(window) {
      calls.push(["updateFolder", window]);
      return update(window);
    },
  });
  return { ...h, calls, rebuild: timeout => h.runtime.refreshFolder(h.folder.URI, timeout ?? WAIT_MS, true) };
}

describe("refreshFolder local stores", () => {
  for (const type of ["none", "pop3", "rss"]) {
    it(`re-reads the summary of a ${type} folder without rebuilding it`, async () => {
      const h = loadLocal({ type });
      const pending = h.refresh();
      await h.clock.advance(UPDATE_MS);
      const result = await pending;
      assert.equal(result.success, true);
      assert.equal(result.status, "loaded");
      assert.equal(result.localStore, true);
      assert.equal(result.indexRebuilt, false);
      assert.equal(result.updateConfirmed, undefined);
      assert.equal(result.elapsedMs, UPDATE_MS);
      assert.equal(result.totalBefore, 10);
      assert.equal(result.netMessageDelta, 0);
      assert.deepEqual(h.calls, [["updateFolder", null]]);
      assert.equal(h.listeners.size, 0);
      assert.equal(h.clock.activeTimers.length, 0);
    });
  }

  it("rebuilds the index with the steps of Thunderbird's Repair Folder, in order", async () => {
    const h = loadLocal();
    const pending = h.rebuild();
    await h.clock.advance(UPDATE_MS);
    const result = await pending;
    assert.equal(result.success, true);
    assert.equal(result.status, "loaded");
    assert.equal(result.indexRebuilt, true);
    assert.deepEqual(h.calls, [
      ["reindex", FOLDER_URI], ["summaryValid", false], ["closeAndBackup", ""],
      ["ensureLocalStore", FOLDER_URI], ["updateFolder", null],
    ]);
    assert.equal(h.listeners.size, 0);
  });

  it("closes the database by force when closing it with a backup fails", async () => {
    const h = loadLocal();
    h.folder.closeAndBackupFolderDB = () => { h.calls.push(["closeAndBackup"]); throw new Error("NS_ERROR_FAILURE"); };
    const pending = h.rebuild();
    await h.clock.advance(UPDATE_MS);
    assert.equal((await pending).status, "loaded");
    assert.deepEqual(h.calls.slice(2), [["closeAndBackup"], ["forceClosed"], ["ensureLocalStore", FOLDER_URI], ["updateFolder", null]]);
  });

  it("works on local folders while Thunderbird is offline", async () => {
    const h = loadLocal({ offline: true });
    const pending = h.rebuild();
    await h.clock.advance(UPDATE_MS);
    const result = await pending;
    assert.equal(result.status, "loaded");
    assert.equal(result.indexRebuilt, true);
  });

  it("refuses a busy folder without touching its index", async () => {
    const h = loadLocal();
    h.folder.locked = true;
    const result = await h.rebuild();
    assert.equal(result.success, false);
    assert.match(result.error, /busy/);
    assert.deepEqual(h.calls, []);
    assert.equal(h.folder.countReads, 0);
    assert.equal(h.listeners.size, 0);
  });

  it("reports a store that fails to update and releases its listener", async () => {
    const h = loadLocal();
    h.folder.updateFolder = () => { throw new Error("Store unavailable"); };
    const result = await h.rebuild();
    assert.equal(result.success, false);
    assert.equal(result.status, "error");
    assert.match(result.error, /Store unavailable/);
    assert.equal(result.indexRebuilt, false);
    assert.equal(h.listeners.size, 0);
  });

  it("times out when the store never reports loaded", async () => {
    const h = loadLocal({ delay: null });
    const pending = h.rebuild(WAIT_MS);
    await h.clock.advance(WAIT_MS);
    const result = await pending;
    assert.equal(result.success, false);
    assert.equal(result.status, "timeout");
    assert.equal(result.indexRebuilt, false);
  });

  it("does not let a rebuild join a plain refresh of the same folder", async () => {
    const h = loadLocal();
    const plain = h.refresh();
    const rebuilt = h.rebuild();
    await h.clock.advance(UPDATE_MS);
    const [first, second] = await Promise.all([plain, rebuilt]);
    assert.equal(first.indexRebuilt, false);
    assert.equal(second.indexRebuilt, true);
    assert.equal(second.joinedInFlight, undefined);
    assert.equal(h.calls.filter(([name]) => name === "updateFolder").length, 2);
    assert.equal(h.listeners.size, 0);
  });
});

describe("refreshFolder trailing freshness update", () => {
  it("starts and confirms one trailing update after synchronous loading", async () => {
    const h = loadRefresh();
    h.folder.synchronousLoads = 1;
    const pending = h.refresh();
    await h.clock.advance(UPDATE_MS);
    const result = await pending;
    assert.equal(result.updateConfirmed, true);
    assert.equal(result.elapsedMs, UPDATE_MS);
    assert.equal(result.warning, undefined);
    assert.equal(h.folder.imapUpdateCalls, 2);
    assert.equal(h.listeners.size, 0);
  });

  it("keeps the adapter's default single-attempt behavior", async () => {
    const h = loadRefresh();
    h.folder.synchronousLoads = 1;
    const pending = h.runtime.waitForFolderRefresh(h.folder, WAIT_MS);
    await h.clock.flush();
    assert.equal((await pending).updateConfirmed, false);
    assert.equal(h.folder.imapUpdateCalls, 1);
  });

  it("shares a busy completion and its trailing update across concurrent callers", async () => {
    const h = loadRefresh();
    h.folder.updateFolder();
    const first = h.refresh();
    const second = h.refresh();
    await h.clock.advance(UPDATE_MS);
    assert.equal(h.folder.imapUpdateCalls, 3);
    const late = h.refresh();
    await h.clock.advance(UPDATE_MS);
    for (const pending of [first, second, late]) assert.equal((await pending).updateConfirmed, true);
    assert.equal((await first).joinedInFlight, undefined);
    assert.equal((await second).joinedInFlight, true);
    assert.equal((await late).joinedInFlight, true);
    assert.equal(h.folder.imapUpdateCalls, 3);
    assert.equal(h.listeners.size, 0);
  });

  for (const freshFirst of [false, true]) {
    it(`preserves the default caller's first completion when it arrives ${freshFirst ? "after" : "before"} a fresh caller`, async () => {
      const h = loadRefresh();
      h.folder.updateFolder();
      let legacyResult;
      let legacyCompletedAt;
      let freshResult;
      let freshCompletedAt;
      const pending = [];
      for (const fresh of [freshFirst, !freshFirst]) {
        pending.push(h.runtime.waitForFolderRefresh(h.folder, fresh ? WAIT_MS : MIXED_CALLER_WAIT_MS,
          fresh ? FRESH_UPDATE : undefined).then(result => {
          if (fresh) {
            freshResult = result;
            freshCompletedAt = h.clock.now;
          } else {
            legacyResult = result;
            legacyCompletedAt = h.clock.now;
          }
          return result;
        }));
        if (fresh === freshFirst) await h.clock.advance(BUSY_START_DELAY_MS);
      }
      await h.clock.advance(UPDATE_MS - BUSY_START_DELAY_MS);
      assert.equal(legacyResult?.status, "loaded", "default caller settles on the first completion");
      assert.equal(legacyResult.updateConfirmed, false);
      assert.equal(legacyResult.joinedInFlight, freshFirst);
      assert.match(legacyResult.warning, /could not be confirmed/);
      assert.equal(legacyCompletedAt, UPDATE_MS);
      assert.equal(freshResult, undefined, "fresh caller is still waiting for the trailing update");
      assert.equal(h.folder.imapUpdateCalls, 3);
      assert.equal(h.listeners.size, 1);
      await h.clock.advance(UPDATE_MS);
      await Promise.all(pending);
      assert.equal(legacyResult.status, "loaded", "the default deadline cannot replace its completed result");
      assert.equal(legacyResult.updateConfirmed, false);
      assert.equal(freshResult.status, "loaded");
      assert.equal(freshResult.updateConfirmed, true);
      assert.equal(freshResult.joinedInFlight, !freshFirst);
      assert.equal(freshCompletedAt, 2 * UPDATE_MS);
      assert.equal(h.folder.imapUpdateCalls, 3);
      assert.equal(h.listeners.size, 0);
      assert.equal(h.clock.activeTimers.length, 0);
    });
  }

  it("lets a default caller joining the trailing attempt wait for that attempt's completion", async () => {
    const h = loadRefresh();
    h.folder.updateFolder();
    const fresh = h.refresh();
    await h.clock.advance(UPDATE_MS);
    let legacyResult;
    const legacy = h.runtime.waitForFolderRefresh(h.folder, MIXED_CALLER_WAIT_MS)
      .then(result => { legacyResult = result; return result; });
    await h.clock.flush();
    assert.equal(legacyResult, undefined, "the earlier load predates this caller");
    await h.clock.advance(UPDATE_MS);
    assert.equal((await legacy).status, "loaded");
    assert.equal(legacyResult.updateConfirmed, true);
    assert.equal(legacyResult.joinedInFlight, true);
    assert.equal((await fresh).updateConfirmed, true);
    assert.equal(h.folder.imapUpdateCalls, 3);
    assert.equal(h.listeners.size, 0);
    assert.equal(h.clock.activeTimers.length, 0);
  });

  it("does not reset an earlier caller's budget for the trailing update", async () => {
    const h = loadRefresh();
    h.folder.updateFolder();
    const budget = UPDATE_MS + BUSY_START_DELAY_MS;
    const pending = h.runtime.waitForFolderRefresh(h.folder, budget, FRESH_UPDATE);
    await h.clock.advance(budget);
    const result = await pending;
    assert.equal(result.status, "timeout");
    assert.notEqual(result.updateConfirmed, true);
    assert.equal(h.clock.now, budget);
    assert.equal(h.folder.imapUpdateCalls, 3);
    await h.clock.advance(UPDATE_MS - BUSY_START_DELAY_MS);
    assert.equal(result.status, "timeout", "late confirmation cannot change the returned result");
    assert.equal(h.listeners.size, 0);
  });

  it("does not start trailing work after its only caller's budget expires", async () => {
    const h = loadRefresh();
    h.folder.updateFolder();
    const budget = UPDATE_MS / 2;
    const pending = h.runtime.waitForFolderRefresh(h.folder, budget, FRESH_UPDATE);
    await h.clock.advance(budget);
    assert.equal((await pending).status, "timeout");
    await h.clock.advance(budget);
    assert.equal(h.folder.imapUpdateCalls, 2);
    assert.equal(h.listeners.size, 0);
    assert.equal(h.clock.activeTimers.length, 0);
  });

  it("does not use an unflagged caller's budget after the fresh caller expires", async () => {
    const h = loadRefresh();
    h.folder.updateFolder();
    const fresh = h.runtime.waitForFolderRefresh(h.folder, UPDATE_MS / 2, FRESH_UPDATE);
    const legacy = h.runtime.waitForFolderRefresh(h.folder, WAIT_MS);
    await h.clock.advance(UPDATE_MS);
    assert.equal((await fresh).status, "timeout");
    assert.equal((await legacy).updateConfirmed, false);
    assert.equal(h.folder.imapUpdateCalls, 2);
    assert.equal(h.listeners.size, 0);
  });

  it("uses a remaining concurrent caller's budget after the first caller expires", async () => {
    const h = loadRefresh();
    h.folder.updateFolder();
    const first = h.runtime.waitForFolderRefresh(h.folder, UPDATE_MS / 2, FRESH_UPDATE);
    const second = h.runtime.waitForFolderRefresh(h.folder, WAIT_MS, FRESH_UPDATE);
    await h.clock.advance(UPDATE_MS);
    assert.equal((await first).status, "timeout");
    assert.equal(h.folder.imapUpdateCalls, 3);
    await h.clock.advance(UPDATE_MS);
    assert.equal((await second).updateConfirmed, true);
    assert.equal((await second).joinedInFlight, true);
  });

  it("refuses a trailing update if Thunderbird goes offline", async () => {
    const h = loadRefresh();
    h.folder.updateFolder();
    const pending = h.refresh();
    h.runtime.Services.io.offline = true;
    await h.clock.advance(UPDATE_MS);
    const result = await pending;
    assert.equal(result.status, "error");
    assert.match(result.error, /offline/);
    assert.equal(h.folder.imapUpdateCalls, 2);
    assert.equal(h.listeners.size, 0);
  });

  it("never starts a third attempt, including for callers joining the trailing attempt", async () => {
    const h = loadRefresh({ delay: null });
    h.folder.synchronousLoads = 1;
    const first = h.refresh();
    await h.clock.flush();
    const second = h.refresh();
    h.folder.urlListener = null; // A provider that supplies only FolderLoaded.
    h.folder.finishUrl(h.folder.updateUrl);
    await h.clock.flush();
    assert.equal((await first).updateConfirmed, false);
    assert.equal((await second).updateConfirmed, false);
    assert.equal(h.folder.imapUpdateCalls, 2);
    assert.equal(h.listeners.size, 0);
  });

  it("ignores queued loaded events and old URL callbacks after starting the trailing update", async () => {
    const h = loadRefresh({ delay: null });
    let completed = false;
    const pending = h.refresh().then(result => { completed = true; return result; });
    const oldListener = h.folder.urlListener;
    const oldFolderListener = [...h.listeners][0];
    const oldUrl = h.folder.updateUrl;
    h.folder.urlListener = null;
    h.folder.finishUrl(oldUrl);
    h.mailSession.fire(h.folder); // Queue a duplicate before the first one runs.
    await h.clock.flush();
    assert.equal(h.folder.imapUpdateCalls, 2);
    oldListener.OnStopRunningUrl(oldUrl, 0);
    oldFolderListener.onFolderEvent(h.folder, "FolderLoaded");
    await h.clock.flush();
    assert.equal(completed, false);
    assert.equal(h.listeners.size, 1);
    h.folder.finishUrl(h.folder.updateUrl);
    await h.clock.flush();
    assert.equal((await pending).updateConfirmed, true);
    assert.equal(h.listeners.size, 0);
  });

  it("reports a trailing startup failure and cleans the operation", async () => {
    const h = loadRefresh({ delay: null });
    const pending = h.refresh();
    h.folder.urlListener = null;
    h.folder.finishUrl(h.folder.updateUrl);
    h.folder.updateFolderWithListener = () => { throw new Error("Trailing update failed"); };
    await h.clock.flush();
    assert.match((await pending).error, /Trailing update failed/);
    assert.equal(h.listeners.size, 0);
    assert.equal(h.clock.activeTimers.length, 0);
  });

  it("does not renew operation age when the trailing attempt starts", async () => {
    const h = loadRefresh({ delay: null });
    h.folder.updateFolder();
    const first = h.refresh();
    await h.clock.advance(BUSY_UPDATE_MS);
    h.folder.finishUrl(h.folder.updateUrl);
    await h.clock.flush();
    const oldUrl = h.folder.updateUrl;
    const oldListener = h.folder.urlListener;
    await h.clock.advance(WAIT_MS - BUSY_UPDATE_MS);
    assert.equal((await first).status, "timeout");
    await h.clock.advance(h.runtime.REFRESH_MAX_TIMEOUT_MS - WAIT_MS);
    let completed = false;
    const replacement = h.refresh().then(result => { completed = true; return result; });
    assert.equal(h.folder.imapUpdateCalls, 4, "the original operation age retires the trailing attempt");
    oldListener.OnStopRunningUrl(oldUrl, 0);
    await h.clock.flush();
    assert.equal(completed, false);
    assert.equal(h.listeners.size, 1);
    h.folder.finishUrl(oldUrl);
    await h.clock.flush();
    assert.equal(h.folder.imapUpdateCalls, 5);
    h.folder.finishUrl(h.folder.updateUrl);
    await h.clock.flush();
    assert.equal((await replacement).updateConfirmed, true);
    assert.equal((await replacement).joinedInFlight, undefined);
    assert.equal(h.listeners.size, 0);
  });

  it("abandons an unobserved trailing attempt at the original operation deadline", async () => {
    const h = loadRefresh({ delay: null });
    h.folder.updateFolder();
    const pending = h.refresh();
    await h.clock.advance(BUSY_UPDATE_MS);
    h.folder.finishUrl(h.folder.updateUrl);
    await h.clock.flush();
    await h.clock.advance(WAIT_MS - BUSY_UPDATE_MS);
    assert.equal((await pending).status, "timeout");
    await h.clock.advance(h.runtime.REFRESH_ABANDON_MS - WAIT_MS - 1);
    assert.equal(h.listeners.size, 1);
    await h.clock.advance(1);
    assert.equal(h.listeners.size, 0);
    assert.equal(h.clock.activeTimers.length, 0);
  });
});

describe("refreshFolder generic providers", () => {
  for (const type of ["owl", "ews", "graph"]) {
    it(`waits for ${type} loaded without claiming a confirmed successful update`, async () => {
      const h = loadRefresh({ type });
      const pending = h.refresh();
      h.mailSession.fire({ URI: "other" });
      h.mailSession.fire(h.folder, "CompactCompleted");
      await h.clock.advance(2 * UPDATE_MS);
      const result = await pending;
      assert.equal(result.status, "loaded");
      assert.equal(result.updateConfirmed, false);
      assert.match(result.warning, /could not be confirmed/);
      assert.equal(result.elapsedMs, 2 * UPDATE_MS);
      assert.equal(h.folder.updateCalls, 2);
      assert.equal(h.listeners.size, 0);
    });
  }

  it("accepts synchronous generic-provider loading without the IMAP no-update claim", async () => {
    const h = loadRefresh({ type: "owl" });
    h.folder.updateFolder = () => h.mailSession.fire(h.folder);
    const pending = h.refresh();
    await h.clock.flush();
    const result = await pending;
    assert.equal(result.status, "loaded");
    assert.doesNotMatch(result.warning, /No update ran/);
    assert.equal(h.listeners.size, 0);
  });

  it("ignores an earlier generic-provider rejection while the trailing attempt is pending", async () => {
    const h = loadRefresh({ type: "owl", delay: null });
    let rejectFirst;
    let starts = 0;
    h.folder.updateFolder = () => {
      if (++starts === 1) return new Promise((_resolve, reject) => { rejectFirst = reject; });
    };
    let completed = false;
    const pending = h.refresh().then(result => { completed = true; return result; });
    h.mailSession.fire(h.folder);
    await h.clock.flush();
    assert.equal(starts, 2);
    rejectFirst(new Error("Old provider rejection"));
    await h.clock.flush();
    assert.equal(completed, false);
    h.mailSession.fire(h.folder);
    await h.clock.flush();
    assert.equal((await pending).status, "loaded");
    assert.equal((await pending).updateConfirmed, false);
    assert.equal((await pending).error, undefined);
    assert.doesNotMatch((await pending).warning, /Old provider/);
    assert.equal(h.listeners.size, 0);
  });
});
