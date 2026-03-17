#!/usr/bin/env node
/**
 * Build the Thunderbird MCP extension XPI (cross-platform, no external deps).
 */
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const PROJECT_DIR = path.resolve(__dirname, '..');
const EXT_DIR = path.join(PROJECT_DIR, 'extension');
const DIST_DIR = path.join(PROJECT_DIR, 'dist');
const OUT_FILE = path.join(DIST_DIR, 'thunderbird-mcp.xpi');

function crc32(buf) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
    }
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

class ZipWriter {
  constructor() { this.files = []; this.offset = 0; this.buf = []; }

  addFile(name, data) {
    // Always use forward slashes in zip entries
    name = name.split(path.sep).join('/');
    const nameBuf = Buffer.from(name, 'utf8');
    const crc = crc32(data);
    const compressed = zlib.deflateRawSync(data);

    const lh = Buffer.alloc(30 + nameBuf.length);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4);
    lh.writeUInt16LE(0, 6);
    lh.writeUInt16LE(8, 8);
    lh.writeUInt16LE(0, 10);
    lh.writeUInt16LE(0, 12);
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(compressed.length, 18);
    lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26);
    lh.writeUInt16LE(0, 28);
    nameBuf.copy(lh, 30);

    this.files.push({ name: nameBuf, crc, compressedSize: compressed.length, uncompressedSize: data.length, offset: this.offset });
    this.buf.push(lh, compressed);
    this.offset += lh.length + compressed.length;
  }

  toBuffer() {
    const cd = [];
    let cdSize = 0;
    for (const f of this.files) {
      const e = Buffer.alloc(46 + f.name.length);
      e.writeUInt32LE(0x02014b50, 0);
      e.writeUInt16LE(20, 4);
      e.writeUInt16LE(20, 6);
      e.writeUInt16LE(0, 8);
      e.writeUInt16LE(8, 10);
      e.writeUInt16LE(0, 12);
      e.writeUInt16LE(0, 14);
      e.writeUInt32LE(f.crc, 16);
      e.writeUInt32LE(f.compressedSize, 20);
      e.writeUInt32LE(f.uncompressedSize, 24);
      e.writeUInt16LE(f.name.length, 28);
      e.writeUInt16LE(0, 30);
      e.writeUInt16LE(0, 32);
      e.writeUInt16LE(0, 34);
      e.writeUInt16LE(0, 36);
      e.writeUInt32LE(0, 38);
      e.writeUInt32LE(f.offset, 42);
      f.name.copy(e, 46);
      cd.push(e);
      cdSize += e.length;
    }

    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054b50, 0);
    eocd.writeUInt16LE(0, 4);
    eocd.writeUInt16LE(0, 6);
    eocd.writeUInt16LE(this.files.length, 8);
    eocd.writeUInt16LE(this.files.length, 10);
    eocd.writeUInt32LE(cdSize, 12);
    eocd.writeUInt32LE(this.offset, 16);
    eocd.writeUInt16LE(0, 20);

    return Buffer.concat([...this.buf, ...cd, eocd]);
  }
}

function addDir(zip, dir, prefix) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    const zipPath = prefix ? prefix + '/' + entry.name : entry.name;
    if (entry.isDirectory()) {
      addDir(zip, full, zipPath);
    } else {
      zip.addFile(zipPath, fs.readFileSync(full));
    }
  }
}

// Stamp buildinfo.json with git describe version and timestamp
const { execSync } = require('child_process');
const BUILDINFO_FILE = path.join(EXT_DIR, 'buildinfo.json');
try {
  // Produces e.g. "v1.2.0-3-g1461f1a" (tag + commits past tag + hash)
  let version;
  try {
    version = execSync('git describe --tags --always', { cwd: PROJECT_DIR, encoding: 'utf8' }).trim();
  } catch {
    // No tags exist — fall back to short hash
    version = execSync('git rev-parse --short HEAD', { cwd: PROJECT_DIR, encoding: 'utf8' }).trim();
  }
  // Append +dirty if there are uncommitted changes
  try {
    execSync('git diff --quiet && git diff --cached --quiet', { cwd: PROJECT_DIR });
  } catch {
    version += '+dirty';
  }
  const buildInfo = JSON.stringify({ version, builtAt: new Date().toISOString() });
  fs.writeFileSync(BUILDINFO_FILE, buildInfo);

  // Update manifest.json version from git tag (Thunderbird requires numeric version)
  const MANIFEST_FILE = path.join(EXT_DIR, 'manifest.json');
  const manifest = JSON.parse(fs.readFileSync(MANIFEST_FILE, 'utf8'));
  const tagMatch = version.match(/^v?(\d+\.\d+(?:\.\d+)?)/);
  if (tagMatch) {
    manifest.version = tagMatch[1];
    fs.writeFileSync(MANIFEST_FILE, JSON.stringify(manifest, null, 2) + '\n');
  }
} catch {
  console.warn('Warning: could not stamp buildinfo.json (git not available?)');
}

fs.mkdirSync(DIST_DIR, { recursive: true });
const zip = new ZipWriter();
addDir(zip, EXT_DIR, '');
fs.writeFileSync(OUT_FILE, zip.toBuffer());
console.log(`Built: ${OUT_FILE} (${(fs.statSync(OUT_FILE).size / 1024).toFixed(0)} KB)`);
