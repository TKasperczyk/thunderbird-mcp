#!/bin/bash
# Build the Thunderbird MCP extension

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
EXTENSION_DIR="$PROJECT_DIR/extension"
DIST_DIR="$PROJECT_DIR/dist"
PACKAGE_JSON="$PROJECT_DIR/package.json"

echo "Building Thunderbird MCP extension..."

# BUILD_VERSION env var takes precedence (used by CI to avoid modifying package.json)
if [ -n "${BUILD_VERSION:-}" ]; then
  PACKAGE_VERSION="$BUILD_VERSION"
elif command -v node > /dev/null 2>&1; then
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

# Stamp build version info into buildinfo.json
# BUILD_VERSION (set by CI) takes precedence — avoids using the previous git tag
# which would be off-by-one since the new tag is only created after the build step.
if [ -n "${BUILD_VERSION:-}" ]; then
  VERSION="v${BUILD_VERSION}"
else
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
fi
BUILT_AT=$(date -u +"%Y-%m-%dT%H:%M:%SZ")
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
zip -r "$DIST_DIR/thunderbird-mcp.xpi" . -x "*.DS_Store" -x "*.git*"

echo "Built: $DIST_DIR/thunderbird-mcp.xpi"
