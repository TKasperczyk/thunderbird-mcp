#!/bin/bash
# Build the Thunderbird MCP extension

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
EXTENSION_DIR="$PROJECT_DIR/extension"
DIST_DIR="$PROJECT_DIR/dist"

echo "Building Thunderbird MCP extension..."

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
BUILT_AT=$(date -u +"%Y-%m-%dT%H:%M:%SZ")
echo "{\"version\":\"$VERSION\",\"builtAt\":\"$BUILT_AT\"}" > "$EXTENSION_DIR/buildinfo.json"
echo "Build version: $VERSION"

# Update manifest.json version from git tag (Thunderbird requires numeric version)
TAG_VERSION=$(echo "$VERSION" | grep -oE '^v?[0-9]+\.[0-9]+(\.[0-9]+)?' | sed 's/^v//')
if [ -n "$TAG_VERSION" ]; then
  # Use node if available for safe JSON editing, otherwise sed
  if command -v node > /dev/null 2>&1; then
    node -e "
      const fs = require('fs');
      const p = '$EXTENSION_DIR/manifest.json';
      const m = JSON.parse(fs.readFileSync(p, 'utf8'));
      m.version = '$TAG_VERSION';
      fs.writeFileSync(p, JSON.stringify(m, null, 2) + '\n');
    "
  else
    sed -i.bak "s/\"version\": *\"[^\"]*\"/\"version\": \"$TAG_VERSION\"/" "$EXTENSION_DIR/manifest.json"
    rm -f "$EXTENSION_DIR/manifest.json.bak"
  fi
  echo "Manifest version: $TAG_VERSION"
fi

# Package extension
cd "$EXTENSION_DIR"
zip -r "$DIST_DIR/thunderbird-mcp.xpi" . -x "*.DS_Store" -x "*.git*"

echo "Built: $DIST_DIR/thunderbird-mcp.xpi"
