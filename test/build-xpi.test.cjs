"use strict";

const { describe, it, beforeEach, afterEach } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { execFileSync, spawnSync } = require("node:child_process");

const project = path.resolve(__dirname, "..");
const hasZipTools = ["zip", "unzip"].every(command => spawnSync(command, ["-v"]).status === 0);

describe("XPI packaging", { skip: process.platform !== "linux" || !hasZipTools }, () => {
  let root;
  let xpi;

  beforeEach(() => {
    root = fs.mkdtempSync(path.join(os.tmpdir(), "tbmcp-build-"));
    fs.mkdirSync(path.join(root, "scripts"));
    fs.mkdirSync(path.join(root, "extension", "icons"), { recursive: true });
    for (const file of ["scripts/build.sh", "scripts/build-xpi.cjs", "LICENSE", "package.json", "extension/manifest.json"]) {
      fs.copyFileSync(path.join(project, file), path.join(root, file));
    }
    fs.writeFileSync(path.join(root, "extension", "icons", "fixture.txt"), "fixture");
    xpi = path.join(root, "dist", "thunderbird-mcp.xpi");
  });

  afterEach(() => fs.rmSync(root, { recursive: true, force: true }));

  for (const [command, script] of [["bash", "scripts/build.sh"], [process.execPath, "scripts/build-xpi.cjs"]]) {
    it(`ships the unchanged MIT notice at the XPI root with ${script}`, () => {
      execFileSync(command, [script], { cwd: root, stdio: "pipe" });
      assert.equal(execFileSync("unzip", ["-p", xpi, "LICENSE"], { encoding: "utf8" }),
        fs.readFileSync(path.join(project, "LICENSE"), "utf8"));
      const manifest = JSON.parse(execFileSync("unzip", ["-p", xpi, "manifest.json"], { encoding: "utf8" }));
      assert.equal(manifest.version, JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")).version);
    });
  }

  it("builds identical release bytes despite changed source mtimes and runner timezone", () => {
    const env = { ...process.env, SOURCE_DATE_EPOCH: "1700000000", TZ: "America/New_York" };
    execFileSync("bash", ["scripts/build.sh"], { cwd: root, env, stdio: "pipe" });
    const first = fs.readFileSync(xpi);
    const later = new Date("2026-09-29T12:34:56Z");
    fs.utimesSync(path.join(root, "extension", "icons", "fixture.txt"), later, later);
    fs.utimesSync(path.join(root, "LICENSE"), later, later);
    execFileSync("bash", ["scripts/build.sh"], { cwd: root, env: { ...env, TZ: "Asia/Tokyo" }, stdio: "pipe" });
    assert.deepEqual(fs.readFileSync(xpi), first);
    const info = JSON.parse(execFileSync("unzip", ["-p", xpi, "buildinfo.json"], { encoding: "utf8" }));
    assert.equal(info.builtAt, "2023-11-14T22:13:20Z");
  });
});
