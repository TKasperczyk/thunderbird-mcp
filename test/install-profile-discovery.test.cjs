"use strict";

const { describe, it, beforeEach, afterEach } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

const projectRoot = path.resolve(__dirname, "..");
const installScript = path.join(projectRoot, "scripts", "install.sh");

describe("Installer profile discovery (scripts/install.sh)", { skip: process.platform !== "linux" }, () => {
  let fakeHome;

  beforeEach(() => {
    fakeHome = fs.mkdtempSync(path.join(os.tmpdir(), "tbmcp-install-test-"));
  });

  afterEach(() => {
    fs.rmSync(fakeHome, { recursive: true, force: true });
  });

  function runInstaller(args = ["--print-profile"], options = {}) {
    const env = {
      ...process.env,
      HOME: fakeHome,
      ...(options.env || {}),
    };
    return spawnSync("bash", [installScript, ...args], {
      cwd: projectRoot,
      env,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
      ...options,
    });
  }

  it("detects modern Snap profile using 'current' symlink with .config/thunderbird", () => {
    const revDir = path.join(fakeHome, "snap", "thunderbird", "1312");
    const configDir = path.join(revDir, ".config", "thunderbird");
    const profileDir = path.join(configDir, "modern.default");
    fs.mkdirSync(profileDir, { recursive: true });
    fs.writeFileSync(path.join(profileDir, "prefs.js"), "// test prefs\n");

    const currentSymlink = path.join(fakeHome, "snap", "thunderbird", "current");
    fs.symlinkSync(revDir, currentSymlink);

    const iniContent = [
      "[General]",
      "StartWithLastProfile=1",
      "Version=2",
      "",
      "[Profile0]",
      "Name=default",
      "IsRelative=1",
      "Path=modern.default",
      "Default=1",
      "",
    ].join("\n");
    fs.writeFileSync(path.join(configDir, "profiles.ini"), iniContent);

    const result = runInstaller(["--print-profile"]);
    assert.equal(result.status, 0, `install.sh failed: ${result.stderr}`);

    const expected = path.join(fakeHome, "snap", "thunderbird", "current", ".config", "thunderbird", "modern.default");
    assert.equal(result.stdout.trim(), expected);
  });

  it("deduplicates canonical target between 'current' symlink and versioned directory", () => {
    const revDir = path.join(fakeHome, "snap", "thunderbird", "1312");
    const configDir = path.join(revDir, ".config", "thunderbird");
    const profileDir = path.join(configDir, "shared.default");
    fs.mkdirSync(profileDir, { recursive: true });
    fs.writeFileSync(path.join(profileDir, "prefs.js"), "// test prefs\n");

    const currentSymlink = path.join(fakeHome, "snap", "thunderbird", "current");
    fs.symlinkSync(revDir, currentSymlink);

    const iniContent = [
      "[Profile0]",
      "Name=default",
      "IsRelative=1",
      "Path=shared.default",
      "Default=1",
      "",
    ].join("\n");
    fs.writeFileSync(path.join(configDir, "profiles.ini"), iniContent);

    // In non-interactive mode, if duplicates are not canonicalized, install.sh would fail
    // reporting multiple profiles. With canonical deduplication, exactly one profile is found.
    const result = runInstaller(["--dry-run"]);
    assert.equal(result.status, 0, `install.sh failed with multiple profile error: ${result.stderr}`);
    assert.match(result.stdout, /Dry run: would install to profile:/);
    assert.match(result.stdout, /shared\.default/);
  });

  it("detects modern Snap profile using versioned numeric directory when current symlink is absent", () => {
    const configDir = path.join(fakeHome, "snap", "thunderbird", "2048", ".config", "thunderbird");
    const profileDir = path.join(configDir, "versioned.default");
    fs.mkdirSync(profileDir, { recursive: true });
    fs.writeFileSync(path.join(profileDir, "prefs.js"), "// test prefs\n");

    const iniContent = [
      "[Profile0]",
      "Name=default",
      "IsRelative=1",
      "Path=versioned.default",
      "Default=1",
      "",
    ].join("\n");
    fs.writeFileSync(path.join(configDir, "profiles.ini"), iniContent);

    const result = runInstaller(["--print-profile"]);
    assert.equal(result.status, 0, `install.sh failed: ${result.stderr}`);

    const expected = path.join(configDir, "versioned.default");
    assert.equal(result.stdout.trim(), expected);
  });

  it("detects legacy Snap profile at common/.thunderbird", () => {
    const configDir = path.join(fakeHome, "snap", "thunderbird", "common", ".thunderbird");
    const profileDir = path.join(configDir, "legacy.default");
    fs.mkdirSync(profileDir, { recursive: true });
    fs.writeFileSync(path.join(profileDir, "prefs.js"), "// test prefs\n");

    const iniContent = [
      "[Profile0]",
      "Name=default",
      "IsRelative=1",
      "Path=legacy.default",
      "Default=1",
      "",
    ].join("\n");
    fs.writeFileSync(path.join(configDir, "profiles.ini"), iniContent);

    const result = runInstaller(["--print-profile"]);
    assert.equal(result.status, 0, `install.sh failed: ${result.stderr}`);

    const expected = path.join(configDir, "legacy.default");
    assert.equal(result.stdout.trim(), expected);
  });

  it("detects traditional ~/.thunderbird profile", () => {
    const configDir = path.join(fakeHome, ".thunderbird");
    const profileDir = path.join(configDir, "traditional.default");
    fs.mkdirSync(profileDir, { recursive: true });
    fs.writeFileSync(path.join(profileDir, "prefs.js"), "// test prefs\n");

    const iniContent = [
      "[Profile0]",
      "Name=default",
      "IsRelative=1",
      "Path=traditional.default",
      "Default=1",
      "",
    ].join("\n");
    fs.writeFileSync(path.join(configDir, "profiles.ini"), iniContent);

    const result = runInstaller(["--print-profile"]);
    assert.equal(result.status, 0, `install.sh failed: ${result.stderr}`);

    const expected = path.join(configDir, "traditional.default");
    assert.equal(result.stdout.trim(), expected);
  });

  it("detects Flatpak profile at ~/.var/app/org.mozilla.Thunderbird/.thunderbird", () => {
    const configDir = path.join(fakeHome, ".var", "app", "org.mozilla.Thunderbird", ".thunderbird");
    const profileDir = path.join(configDir, "flatpak.default");
    fs.mkdirSync(profileDir, { recursive: true });
    fs.writeFileSync(path.join(profileDir, "prefs.js"), "// test prefs\n");

    const iniContent = [
      "[Profile0]",
      "Name=default",
      "IsRelative=1",
      "Path=flatpak.default",
      "Default=1",
      "",
    ].join("\n");
    fs.writeFileSync(path.join(configDir, "profiles.ini"), iniContent);

    const result = runInstaller(["--print-profile"]);
    assert.equal(result.status, 0, `install.sh failed: ${result.stderr}`);

    const expected = path.join(configDir, "flatpak.default");
    assert.equal(result.stdout.trim(), expected);
  });

  it("detects fallback profile directory directly when profiles.ini is absent", () => {
    const profileDir = path.join(fakeHome, "snap", "thunderbird", "current", ".config", "thunderbird", "fallback.default-release");
    fs.mkdirSync(profileDir, { recursive: true });
    fs.writeFileSync(path.join(profileDir, "prefs.js"), "// test prefs\n");

    const result = runInstaller(["--print-profile"]);
    assert.equal(result.status, 0, `install.sh failed: ${result.stderr}`);

    const expected = profileDir;
    assert.equal(result.stdout.trim(), expected);
  });

  it("fails gracefully in non-interactive mode when multiple distinct profiles exist", () => {
    const configDir = path.join(fakeHome, ".thunderbird");
    const profA = path.join(configDir, "profA.profile");
    const profB = path.join(configDir, "profB.profile");
    fs.mkdirSync(profA, { recursive: true });
    fs.mkdirSync(profB, { recursive: true });
    fs.writeFileSync(path.join(profA, "prefs.js"), "// test prefs\n");
    fs.writeFileSync(path.join(profB, "prefs.js"), "// test prefs\n");

    const iniContent = [
      "[Profile0]",
      "Name=A",
      "IsRelative=1",
      "Path=profA.profile",
      "",
      "[Profile1]",
      "Name=B",
      "IsRelative=1",
      "Path=profB.profile",
      "",
    ].join("\n");
    fs.writeFileSync(path.join(configDir, "profiles.ini"), iniContent);

    const result = runInstaller(["--print-profile"]);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /Found multiple Thunderbird profiles and cannot prompt/);
  });

  it("fails with informative error when no profile exists", () => {
    const result = runInstaller(["--print-profile"]);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /Error: No Thunderbird profile found for Linux/);
  });

  it("handles broken symlinks and invalid paths without crashing", () => {
    const snapThunderbird = path.join(fakeHome, "snap", "thunderbird");
    fs.mkdirSync(snapThunderbird, { recursive: true });
    fs.symlinkSync(path.join(snapThunderbird, "does-not-exist"), path.join(snapThunderbird, "current"));

    const result = runInstaller(["--print-profile"]);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /Error: No Thunderbird profile found for Linux/);
  });
});
