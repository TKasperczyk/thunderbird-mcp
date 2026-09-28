#!/bin/bash
# Build the Thunderbird MCP extension

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
EXTENSION_DIR="$PROJECT_DIR/extension"
DIST_DIR="$PROJECT_DIR/dist"
PACKAGE_JSON="$PROJECT_DIR/package.json"

echo "Building Thunderbird MCP extension..."

if command -v node > /dev/null 2>&1; then
  PACKAGE_VERSION=$(node -e "
    const fs = require('fs');
    const p = process.argv[1];
    try {
      const pkg = JSON.parse(fs.readFileSync(p, 'utf8'));
      if (typeof pkg.version !== 'string' || !pkg.version) {
        throw new Error('package.json does not contain a string \"version\" field');
      }
      process.stdout.write(pkg.version);
    } catch (err) {
      console.error('Error: could not read package.json version: ' + err.message);
      process.exit(1);
    }
  " "$PACKAGE_JSON")
else
  PACKAGE_VERSION=$(sed -nE 's/^[[:space:]]*"version"[[:space:]]*:[[:space:]]*"([^"]+)".*/\1/p' "$PACKAGE_JSON" | head -n 1)
  if [ -z "$PACKAGE_VERSION" ]; then
    echo "Error: could not read package.json version" >&2
    exit 1
  fi
fi

# Create dist directory
mkdir -p "$DIST_DIR"

# Remove old XPI to ensure a clean build
rm -f "$DIST_DIR/thunderbird-mcp.xpi"

# Stamp build version info (git-describe + timestamp) into buildinfo.json
VERSION="unknown"
if git -C "$PROJECT_DIR" describe --tags --always > /dev/null 2>&1; then
  VERSION=$(git -C "$PROJECT_DIR" describe --tags --always)
elif git -C "$PROJECT_DIR" rev-parse --short HEAD > /dev/null 2>&1; then
  VERSION=$(git -C "$PROJECT_DIR" rev-parse --short HEAD)
fi
# Append +dirty if there are uncommitted changes
if ! git -C "$PROJECT_DIR" diff --quiet 2>/dev/null || ! git -C "$PROJECT_DIR" diff --cached --quiet 2>/dev/null; then
  VERSION="${VERSION}+dirty"
fi
if [ -n "${SOURCE_DATE_EPOCH:-}" ]; then
  # The Ubuntu release job supplies the commit time so reruns produce the same XPI.
  export TZ=UTC
  BUILT_AT=$(date -u -d "@$SOURCE_DATE_EPOCH" +"%Y-%m-%dT%H:%M:%SZ")
else
  BUILT_AT=$(date -u +"%Y-%m-%dT%H:%M:%SZ")
fi
echo "{\"version\":\"$VERSION\",\"builtAt\":\"$BUILT_AT\"}" > "$EXTENSION_DIR/buildinfo.json"
echo "Build version: $VERSION"

# Update manifest.json version from package.json
if command -v node > /dev/null 2>&1; then
  node -e "
    const fs = require('fs');
    const p = '$EXTENSION_DIR/manifest.json';
    const m = JSON.parse(fs.readFileSync(p, 'utf8'));
    m.version = '$PACKAGE_VERSION';
    fs.writeFileSync(p, JSON.stringify(m, null, 2) + '\n');
  "
else
  sed -i.bak "s/\"version\": *\"[^\"]*\"/\"version\": \"$PACKAGE_VERSION\"/" "$EXTENSION_DIR/manifest.json"
  rm -f "$EXTENSION_DIR/manifest.json.bak"
fi
echo "Manifest version: $PACKAGE_VERSION"

# Package extension
cd "$EXTENSION_DIR"
if [ -n "${SOURCE_DATE_EPOCH:-}" ]; then
  find . -exec touch -d "@$SOURCE_DATE_EPOCH" {} +
  touch -d "@$SOURCE_DATE_EPOCH" "$PROJECT_DIR/LICENSE"
  find . -type f ! -name '*.DS_Store' ! -path '*.git*' | LC_ALL=C sort |
    zip -X "$DIST_DIR/thunderbird-mcp.xpi" -@
else
  zip -r "$DIST_DIR/thunderbird-mcp.xpi" . -x "*.DS_Store" -x "*.git*"
fi
zip -X -j "$DIST_DIR/thunderbird-mcp.xpi" "$PROJECT_DIR/LICENSE"

echo "Built: $DIST_DIR/thunderbird-mcp.xpi"
