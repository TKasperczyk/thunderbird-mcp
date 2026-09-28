"use strict";

const { describe, it, beforeEach, afterEach } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { execFileSync, spawnSync } = require("node:child_process");

const workflow = fs.readFileSync(path.resolve(__dirname, "../.github/workflows/build-and-release.yml"), "utf8");

// Exercise the actual shell steps without adding a YAML dependency to the suite.
function stepScript(name) {
  const step = workflow.split("      - name: " + name + "\n")[1]?.split("\n      - name: ")[0];
  assert.ok(step, "Missing workflow step: " + name);
  const script = step.match(/^ {8}run: \|\n((?: {10}.*\n|\n)+)/m);
  assert.ok(script, "Missing shell block for: " + name);
  return script[1].replace(/^ {10}/gm, "");
}

// These scripts run on the Ubuntu release runner and use GNU shell utilities.
describe("release workflow", { skip: process.platform !== "linux" }, () => {
  let root;

  beforeEach(() => {
    root = fs.mkdtempSync(path.join(os.tmpdir(), "tbmcp-release-"));
  });

  afterEach(() => {
    fs.rmSync(root, { recursive: true, force: true });
  });

  function runRelease({ state = "existing", assets, remoteBytes = "same XPI", downloadFails = false } = {}) {
    const asset = "thunderbird-mcp-v0.9.0.xpi";
    const bin = path.join(root, "bin");
    const calls = path.join(root, "gh-calls.jsonl");
    fs.mkdirSync(bin);
    fs.mkdirSync(path.join(root, "dist"));
    fs.writeFileSync(path.join(root, "dist", asset), "same XPI");
    fs.writeFileSync(path.join(root, "published.xpi"), remoteBytes);
    fs.writeFileSync(path.join(bin, "gh"), `#!/usr/bin/env node
const fs = require("node:fs");
const path = require("node:path");
const args = process.argv.slice(2);
fs.appendFileSync(process.env.P2_GH_CALLS, JSON.stringify(args) + "\\n");
if (args[0] !== "release") process.exit(2);
if (args[1] === "view") {
  if (process.env.P2_RELEASE_STATE === "missing") process.exit(1);
  process.stdout.write(process.env.P2_RELEASE_ASSETS);
} else if (args[1] === "download") {
  if (process.env.P2_DOWNLOAD_FAILS === "yes") process.exit(1);
  const dir = args[args.indexOf("--dir") + 1];
  const name = args[args.indexOf("--pattern") + 1];
  fs.copyFileSync(process.env.P2_REMOTE_XPI, path.join(dir, name));
} else if (!["upload", "create"].includes(args[1])) {
  process.exit(2);
}
`, { mode: 0o755 });
    const result = spawnSync("bash", ["-euo", "pipefail", "-c", stepScript("Create / update GitHub release")], {
      cwd: root,
      encoding: "utf8",
      env: {
        ...process.env,
        PATH: bin + path.delimiter + process.env.PATH,
        TMPDIR: root,
        TAG: "v0.9.0",
        GITHUB_REPOSITORY: "example/thunderbird-mcp",
        P2_GH_CALLS: calls,
        P2_RELEASE_STATE: state,
        P2_RELEASE_ASSETS: (assets ?? [asset]).join("\n") + "\n",
        P2_REMOTE_XPI: path.join(root, "published.xpi"),
        P2_DOWNLOAD_FAILS: downloadFails ? "yes" : "no",
      },
    });
    return {
      ...result,
      calls: fs.readFileSync(calls, "utf8").trim().split("\n").map(line => JSON.parse(line)),
    };
  }

  it("skips uploading an existing asset only when its bytes have the same sha256", () => {
    const result = runRelease();
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout, /already has sha256 [0-9a-f]{64}; skipping upload/);
    assert.deepEqual(result.calls.map(args => args[1]), ["view", "download"]);
  });

  it("fails without modifying a same-named asset with different bytes", () => {
    const result = runRelease({ remoteBytes: "different published XPI" });
    assert.notEqual(result.status, 0);
    assert.match(result.stdout, /::error::Release asset .* already exists with sha256 .*refusing to replace it/);
    assert.deepEqual(result.calls.map(args => args[1]), ["view", "download"]);
  });

  it("does not upload when downloading an existing asset fails", () => {
    const result = runRelease({ downloadFails: true });
    assert.notEqual(result.status, 0);
    assert.deepEqual(result.calls.map(args => args[1]), ["view", "download"]);
  });

  it("uploads a missing asset without attempting to replace another asset", () => {
    const result = runRelease({ assets: ["another.xpi"] });
    assert.equal(result.status, 0, result.stderr);
    assert.deepEqual(result.calls.map(args => args[1]), ["view", "upload"]);
    assert.ok(!result.calls.flat().includes("--clobber"));
  });

  it("creates a missing release with the XPI and verifies that the tag exists", () => {
    const result = runRelease({ state: "missing" });
    assert.equal(result.status, 0, result.stderr);
    assert.deepEqual(result.calls.map(args => args[1]), ["view", "create"]);
    assert.ok(result.calls[1].includes("--verify-tag"));
    assert.ok(result.calls[1].includes("dist/thunderbird-mcp-v0.9.0.xpi"));
  });

  function createRepo() {
    const repo = path.join(root, "repo");
    fs.mkdirSync(repo);
    function git(...args) {
      return execFileSync("git", args, {
        cwd: repo,
        encoding: "utf8",
        stdio: ["ignore", "pipe", "pipe"],
        env: {
          ...process.env,
          GIT_CONFIG_NOSYSTEM: "1",
          GIT_CONFIG_GLOBAL: "/dev/null",
          GIT_AUTHOR_NAME: "Release test",
          GIT_AUTHOR_EMAIL: "release@example.invalid",
          GIT_COMMITTER_NAME: "Release test",
          GIT_COMMITTER_EMAIL: "release@example.invalid",
        },
      }).trim();
    }
    git("init", "--initial-branch=main");
    git("commit", "--allow-empty", "-m", "Initial commit");
    // A local origin exercises the actual fetch and ancestry check without network access.
    git("remote", "add", "origin", repo);
    return { repo, git };
  }

  function runAncestryCheck(repo) {
    return spawnSync("bash", ["-euo", "pipefail", "-c", stepScript("Verify release commit is on main")], {
      cwd: repo,
      encoding: "utf8",
      env: { ...process.env, GITHUB_REF_NAME: "v0.9.0" },
    });
  }

  it("allows a tag whose commit is an ancestor of origin/main, even after main advances", () => {
    const { repo, git } = createRepo();
    const releaseCommit = git("rev-parse", "HEAD");
    git("commit", "--allow-empty", "-m", "Later main commit");
    git("checkout", "--detach", releaseCommit);
    const result = runAncestryCheck(repo);
    assert.equal(result.status, 0, result.stderr);
  });

  it("rejects a tag that has not been merged into origin/main", () => {
    const { repo, git } = createRepo();
    git("checkout", "-b", "unmerged");
    git("commit", "--allow-empty", "-m", "Unmerged release commit");
    const result = runAncestryCheck(repo);
    assert.notEqual(result.status, 0);
    assert.match(result.stdout, /::error::Tag v0\.9\.0 is not contained in origin\/main; refusing to release/);
  });
});
