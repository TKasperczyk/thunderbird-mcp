#!/usr/bin/env node
/**
 * Integration smoke test for Thunderbird MCP.
 *
 * Spawns mcp-bridge.cjs as a child process (exactly as Claude Desktop
 * does), pipes stdin/stdout, and drives a sequence of MCP JSON-RPC
 * calls against a LIVE Thunderbird running with the CI profile pointed
 * at Greenmail. Asserts behavior end-to-end -- not just "bridge
 * responds" but "tool actually did the work against an IMAP server".
 *
 * Prerequisites (set up by test/integration/run.sh in CI):
 *   - Greenmail is running on localhost:3143 (IMAP) / 3025 (SMTP)
 *   - A TB profile exists with our XPI installed + an IMAP account
 *     pointing at greenmail + seeded messages in the inbox
 *   - Thunderbird has been launched headlessly against that profile
 *   - /tmp/thunderbird-mcp/connection.json exists and is reachable
 *
 * Fixtures (seeded by scripts/ci/seed-greenmail.cjs):
 *   ~20 messages from 5 distinct senders across 3 domains:
 *     linkedin.com         (2 senders, 8 messages)
 *     github.com           (1 sender,  4 messages)
 *     stripe.com           (1 sender,  4 messages)
 *     newsletter.example.com (1 sender, 4 messages)
 *
 * Exit codes:
 *   0 = all asserts pass
 *   1 = at least one assertion failed (prints the failing assertion first,
 *       then a JSON dump of the full offending response)
 */

const { spawn } = require("child_process");
const path = require("path");
const assert = require("node:assert/strict");

const BRIDGE_PATH = path.resolve(__dirname, "..", "..", "mcp-bridge.cjs");
const TIMEOUT_MS = 30000;

// One persistent bridge process for the whole suite. Spawning per-call
// would work but adds ~1s overhead each.
let bridge;
let msgId = 0;
const pending = new Map();
let stdoutBuffer = "";

function startBridge() {
  bridge = spawn("node", [BRIDGE_PATH], {
    stdio: ["pipe", "pipe", "inherit"],
    env: { ...process.env, THUNDERBIRD_MCP_DEBUG: "" },
  });
  bridge.stdout.setEncoding("utf8");
  bridge.stdout.on("data", (chunk) => {
    stdoutBuffer += chunk;
    let idx;
    while ((idx = stdoutBuffer.indexOf("\n")) !== -1) {
      const line = stdoutBuffer.slice(0, idx).replace(/\r$/, "");
      stdoutBuffer = stdoutBuffer.slice(idx + 1);
      if (!line.trim()) continue;
      try {
        const msg = JSON.parse(line);
        const p = pending.get(msg.id);
        if (p) { pending.delete(msg.id); p.resolve(msg); }
      } catch (e) {
        console.error("non-JSON line from bridge:", JSON.stringify(line));
      }
    }
  });
  bridge.on("exit", (code) => {
    for (const { reject } of pending.values()) {
      reject(new Error(`bridge exited with code ${code} before response`));
    }
  });
}

function call(method, params) {
  const id = ++msgId;
  const payload = JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n";
  bridge.stdin.write(payload);
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pending.delete(id);
      reject(new Error(`timeout waiting for ${method} id=${id}`));
    }, TIMEOUT_MS);
    pending.set(id, {
      resolve: (r) => { clearTimeout(timer); resolve(r); },
      reject:  (e) => { clearTimeout(timer); reject(e); },
    });
  });
}

async function tool(name, args = {}) {
  const res = await call("tools/call", { name, arguments: args });
  if (res.error) throw new Error(`${name}: ${res.error.message}`);
  const content = res.result && res.result.content && res.result.content[0];
  if (!content || content.type !== "text") throw new Error(`${name}: no text content`);
  return JSON.parse(content.text);
}

async function notify(method, params) {
  bridge.stdin.write(JSON.stringify({ jsonrpc: "2.0", method, params }) + "\n");
}

// ── Assertions --------------------------------------------------------

let failures = 0;
function check(label, fn) {
  try {
    fn();
    console.log(`  PASS  ${label}`);
  } catch (e) {
    failures++;
    console.log(`  FAIL  ${label}`);
    console.log(`        ${e.message}`);
  }
}

// ── Main --------------------------------------------------------------

(async () => {
  startBridge();

  // 1. Initialize / tools/list
  const initRes = await call("initialize", {
    protocolVersion: "2025-06-18",
    capabilities: {},
    clientInfo: { name: "ci-smoke", version: "1.0" },
  });
  check("initialize returns serverInfo", () => {
    assert.equal(initRes.result.serverInfo.name, "thunderbird-mcp");
  });
  await notify("notifications/initialized");

  const listRes = await call("tools/list", {});
  const allNames = listRes.result.tools.map(t => t.name);
  check("tools/list includes registry-migrated tool", () => {
    assert.ok(allNames.includes("inbox_inventory"), `inbox_inventory missing from ${allNames.length} tools`);
  });
  check("tools/list includes bulk_move_by_query", () => {
    assert.ok(allNames.includes("bulk_move_by_query"));
  });
  check("tools/list includes createDrafts", () => {
    assert.ok(allNames.includes("createDrafts"));
  });
  check("applyFilters schema has dry_run", () => {
    const af = listRes.result.tools.find(t => t.name === "applyFilters");
    assert.ok(af && af.inputSchema.properties.dry_run, "applyFilters.inputSchema.properties.dry_run missing");
  });

  // 2. Accounts reachable (greenmail mock visible)
  const accounts = await tool("listAccounts");
  check("listAccounts returns at least one IMAP account", () => {
    assert.ok(Array.isArray(accounts));
    assert.ok(accounts.length >= 1);
    assert.ok(accounts.some(a => a.type === "imap"));
  });
  const imapAccountId = (accounts.find(a => a.type === "imap") || {}).id;

  // 3. Folder tree
  const folders = await tool("listFolders", { accountId: imapAccountId });
  const inbox = folders.find(f => f.type === "inbox" || /Inbox/i.test(f.name));
  check("listFolders includes an Inbox on the IMAP account", () => {
    assert.ok(inbox, `Inbox not found; got folders: ${folders.map(f => f.name).join(", ")}`);
  });

  // 4. Trigger IMAP sync via searchMessages (updateFolder inside), then
  //    inbox_inventory should see the seeded messages
  await tool("searchMessages", { query: "", folderPath: inbox.URI || inbox.folderPath, countOnly: true });

  const inv = await tool("inbox_inventory", {
    query: "",
    folderPath: inbox.URI || inbox.folderPath,
    groupBy: "from_domain",
  });
  check("inbox_inventory groups by from_domain with >=3 groups", () => {
    assert.ok(inv.totalMatched >= 10, `expected >=10 matched, got ${inv.totalMatched}`);
    assert.ok(inv.groupCount >= 3, `expected >=3 groups, got ${inv.groupCount}`);
    assert.ok(inv.groups.some(g => g.key === "linkedin.com"),
      `linkedin.com group missing; got: ${inv.groups.map(g => g.key).join(", ")}`);
  });
  check("inbox_inventory sample-subjects populated", () => {
    const lk = inv.groups.find(g => g.key === "linkedin.com");
    assert.ok(lk.sampleSubjects && lk.sampleSubjects.length > 0);
  });

  // 5. bulk_move_by_query dry-run then real -- smoke-tests the whole
  //    search-filter + move-via-copyMessages pipeline AND the suspicious
  //    tag-apply loop.
  const linkedInFolderName = "CI-linkedin-triage";
  await tool("createFolder", { parentFolderPath: inbox.URI || inbox.folderPath, name: linkedInFolderName });
  const foldersAfter = await tool("listFolders", { accountId: imapAccountId });
  const triageFolder = foldersAfter.find(f => f.name === linkedInFolderName);
  check("createFolder succeeded for triage target", () => {
    assert.ok(triageFolder, `triage folder ${linkedInFolderName} not found after create`);
  });

  const dry = await tool("bulk_move_by_query", {
    query: "from:linkedin.com",
    folderPath: inbox.URI || inbox.folderPath,
    to: triageFolder.URI || triageFolder.folderPath,
    dry_run: true,
  });
  check("bulk_move_by_query dry run reports matched>=6 linkedin messages", () => {
    assert.equal(dry.dryRun, true);
    assert.ok(dry.matched >= 6, `expected >=6 matched in dry run, got ${dry.matched}`);
  });

  const live = await tool("bulk_move_by_query", {
    query: "from:linkedin.com",
    folderPath: inbox.URI || inbox.folderPath,
    to: triageFolder.URI || triageFolder.folderPath,
    dry_run: false,
    addTags: ["$label3"],  // exercises the tag-application code path
  });
  check("bulk_move_by_query live run reports moved count", () => {
    assert.equal(live.dryRun, false);
    assert.ok(live.moved >= 1, `expected >=1 moved, got ${live.moved}`);
  });

  // After the move, the triage folder should have those messages
  // (IMAP is async; allow 3s for propagation)
  await new Promise(r => setTimeout(r, 3000));
  const triageContents = await tool("searchMessages", {
    query: "",
    folderPath: triageFolder.URI || triageFolder.folderPath,
    countOnly: true,
  });
  check("linkedin messages landed in triage folder", () => {
    assert.ok(triageContents.count >= 1,
      `expected >=1 msg in triage folder after move, got ${triageContents.count}`);
  });

  // 6. applyFilters dry-run returns per-filter report
  //    (We didn't create real filters so byFilter can be empty, but the
  //    shape MUST be correct.)
  const filtersReport = await tool("applyFilters", {
    accountId: imapAccountId,
    folderPath: inbox.URI || inbox.folderPath,
    dry_run: true,
  });
  check("applyFilters dry-run returns byFilter array + messagesProcessed", () => {
    assert.equal(filtersReport.dryRun, true);
    assert.ok(Array.isArray(filtersReport.byFilter));
    assert.ok(typeof filtersReport.messagesProcessed === "number");
  });

  // 7. createDrafts round-trip (minimal)
  const drafts = await tool("createDrafts", {
    drafts: [
      { to: "x@ci.local", subject: "CI smoke draft 1", body: "body 1" },
      { to: "y@ci.local", subject: "CI smoke draft 2", body: "body 2" },
    ],
  });
  check("createDrafts returns per-draft result with draftFolderURI", () => {
    assert.equal(drafts.total, 2);
    assert.equal(drafts.succeeded, 2, `expected 2 succeeded, got ${drafts.succeeded} -- drafts: ${JSON.stringify(drafts.drafts)}`);
    assert.ok(drafts.drafts[0].draftFolderURI);
  });

  // 8. Cleanup tempfiles/windows (createTask would open a GUI dialog -- skip)
  await tool("deleteFolder", { folderPath: triageFolder.URI || triageFolder.folderPath });

  bridge.stdin.end();

  if (failures > 0) {
    console.error(`\n${failures} assertion(s) failed`);
    process.exit(1);
  }
  console.log("\nAll integration asserts passed.");
  process.exit(0);
})().catch((e) => {
  console.error("Smoke harness error:", e);
  process.exit(2);
});
