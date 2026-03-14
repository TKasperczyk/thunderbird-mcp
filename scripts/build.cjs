#!/usr/bin/env node
/**
 * Cross-platform build script for the Thunderbird MCP extension.
 * Packages the extension/ directory into dist/thunderbird-mcp.xpi.
 *
 * Usage:
 *   node scripts/build.cjs
 *   npm run build
 */

const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const PROJECT_DIR = path.resolve(__dirname, "..");
const EXTENSION_DIR = path.join(PROJECT_DIR, "extension");
const DIST_DIR = path.join(PROJECT_DIR, "dist");
const XPI_PATH = path.join(DIST_DIR, "thunderbird-mcp.xpi");

// Files/patterns to exclude from the XPI
const EXCLUDE = [".DS_Store", ".git"];

function shouldExclude(name) {
  return EXCLUDE.some((pattern) => name.includes(pattern));
}

function build() {
  // Ensure dist/ exists
  fs.mkdirSync(DIST_DIR, { recursive: true });

  // Remove old XPI for a clean build
  if (fs.existsSync(XPI_PATH)) {
    fs.unlinkSync(XPI_PATH);
  }

  // Write build version info (git hash + timestamp) into the extension
  let gitHash = "unknown";
  try {
    gitHash = execSync("git rev-parse --short HEAD", { encoding: "utf8" }).trim();
  } catch { /* not a git repo or git not available */ }
  const buildInfo = JSON.stringify({
    commit: gitHash,
    builtAt: new Date().toISOString(),
  });
  fs.writeFileSync(path.join(EXTENSION_DIR, "buildinfo.json"), buildInfo);

  // Collect files to package
  const files = [];
  function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (shouldExclude(entry.name)) continue;
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath);
      } else {
        const arcName = path.relative(EXTENSION_DIR, fullPath).split(path.sep).join("/");
        files.push({ fullPath, arcName });
      }
    }
  }
  walk(EXTENSION_DIR);

  // Try using the system zip command first (preserves compatibility with build.sh)
  // Fall back to a pure-Node approach using archiver-like logic
  try {
    // Check if zip is available
    execSync("zip --version", { stdio: "ignore" });
    execSync(
      `cd "${EXTENSION_DIR}" && zip -r "${XPI_PATH}" . -x "*.DS_Store" -x "*.git*"`,
      { stdio: "inherit", shell: true }
    );
  } catch {
    // No zip command available — use Node.js built-in zlib via JSZip-like approach
    // Since we have no dependencies, build a minimal valid ZIP file
    buildZipPure(files);
  }

  const stat = fs.statSync(XPI_PATH);
  console.log(`Built: ${XPI_PATH} (${stat.size} bytes)`);
}

/**
 * Build a ZIP file using only Node.js built-ins (no external dependencies).
 * Produces a valid ZIP (PKZIP 2.0) with DEFLATE compression.
 */
function buildZipPure(files) {
  const zlib = require("zlib");

  const localHeaders = [];
  const centralHeaders = [];
  let offset = 0;

  for (const { fullPath, arcName } of files) {
    const content = fs.readFileSync(fullPath);
    const compressed = zlib.deflateRawSync(content);
    const crc = crc32(content);

    const nameBuffer = Buffer.from(arcName, "utf8");
    const useCompressed = compressed.length < content.length;
    const storedData = useCompressed ? compressed : content;
    const method = useCompressed ? 8 : 0; // 8 = DEFLATE, 0 = STORE

    // Local file header (30 bytes + name + data)
    const local = Buffer.alloc(30 + nameBuffer.length);
    local.writeUInt32LE(0x04034b50, 0); // signature
    local.writeUInt16LE(20, 4); // version needed
    local.writeUInt16LE(0, 6); // flags
    local.writeUInt16LE(method, 8); // compression method
    local.writeUInt16LE(0, 10); // mod time
    local.writeUInt16LE(0, 12); // mod date
    local.writeUInt32LE(crc, 14); // crc-32
    local.writeUInt32LE(storedData.length, 18); // compressed size
    local.writeUInt32LE(content.length, 22); // uncompressed size
    local.writeUInt16LE(nameBuffer.length, 26); // filename length
    local.writeUInt16LE(0, 28); // extra field length
    nameBuffer.copy(local, 30);

    localHeaders.push(Buffer.concat([local, storedData]));

    // Central directory entry (46 bytes + name)
    const central = Buffer.alloc(46 + nameBuffer.length);
    central.writeUInt32LE(0x02014b50, 0); // signature
    central.writeUInt16LE(20, 4); // version made by
    central.writeUInt16LE(20, 6); // version needed
    central.writeUInt16LE(0, 8); // flags
    central.writeUInt16LE(method, 10); // compression method
    central.writeUInt16LE(0, 12); // mod time
    central.writeUInt16LE(0, 14); // mod date
    central.writeUInt32LE(crc, 16); // crc-32
    central.writeUInt32LE(storedData.length, 20); // compressed size
    central.writeUInt32LE(content.length, 24); // uncompressed size
    central.writeUInt16LE(nameBuffer.length, 28); // filename length
    central.writeUInt16LE(0, 30); // extra field length
    central.writeUInt16LE(0, 32); // comment length
    central.writeUInt16LE(0, 34); // disk number start
    central.writeUInt16LE(0, 36); // internal attributes
    central.writeUInt32LE(0, 38); // external attributes
    central.writeUInt32LE(offset, 42); // local header offset
    nameBuffer.copy(central, 46);

    centralHeaders.push(central);
    offset += local.length + storedData.length;
  }

  const centralDirOffset = offset;
  const centralDirSize = centralHeaders.reduce((s, b) => s + b.length, 0);

  // End of central directory (22 bytes)
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); // signature
  eocd.writeUInt16LE(0, 4); // disk number
  eocd.writeUInt16LE(0, 6); // disk with central dir
  eocd.writeUInt16LE(files.length, 8); // entries on this disk
  eocd.writeUInt16LE(files.length, 10); // total entries
  eocd.writeUInt32LE(centralDirSize, 12); // central dir size
  eocd.writeUInt32LE(centralDirOffset, 16); // central dir offset
  eocd.writeUInt16LE(0, 20); // comment length

  const zipBuffer = Buffer.concat([...localHeaders, ...centralHeaders, eocd]);
  fs.writeFileSync(XPI_PATH, zipBuffer);
}

/**
 * Compute CRC-32 for a buffer (standard ZIP/PNG polynomial).
 */
function crc32(buf) {
  // Build table on first call
  if (!crc32.table) {
    crc32.table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      crc32.table[i] = c;
    }
  }
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc = crc32.table[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

build();
