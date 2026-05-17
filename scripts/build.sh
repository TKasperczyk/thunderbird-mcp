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

# Determine GitHub owner and repo name for addon ID and update URLs.
# In CI $GITHUB_REPOSITORY is set (e.g. "the78mole/thunderbird-mcp").
# Locally the value is parsed from the git remote URL.
if [ -n "${GITHUB_REPOSITORY:-}" ]; then
  GITHUB_OWNER="${GITHUB_REPOSITORY%%/*}"
  GITHUB_REPO_NAME="${GITHUB_REPOSITORY##*/}"
else
  REMOTE_URL=$(git -C "$PROJECT_DIR" remote get-url origin 2>/dev/null || echo "")
  if [[ "$REMOTE_URL" =~ github\.com[:/]([^/]+)/([^/.]+) ]]; then
    GITHUB_OWNER="${BASH_REMATCH[1]}"
    GITHUB_REPO_NAME="${BASH_REMATCH[2]}"
  else
    GITHUB_OWNER="unknown"
    GITHUB_REPO_NAME="thunderbird-mcp"
  fi
fi
ADDON_ID="thunderbird-mcp@${GITHUB_OWNER}.github.io"
BASE_URL="https://github.com/${GITHUB_OWNER}/${GITHUB_REPO_NAME}"
UPDATE_URL="${BASE_URL}/releases/latest/download/updates.json"
UPDATE_LINK="${BASE_URL}/releases/download/v${PACKAGE_VERSION}/thunderbird-mcp-${PACKAGE_VERSION}.xpi"
echo "Addon ID: $ADDON_ID"
echo "Update URL: $UPDATE_URL"

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

# Update manifest.json: version, addon ID and update_url
node -e "
  const fs = require('fs');
  const p = '$EXTENSION_DIR/manifest.json';
  const m = JSON.parse(fs.readFileSync(p, 'utf8'));
  m.version = '$PACKAGE_VERSION';
  m.browser_specific_settings = m.browser_specific_settings || {};
  m.browser_specific_settings.gecko = m.browser_specific_settings.gecko || {};
  m.browser_specific_settings.gecko.id = '$ADDON_ID';
  m.browser_specific_settings.gecko.update_url = '$UPDATE_URL';
  fs.writeFileSync(p, JSON.stringify(m, null, 2) + '\n');
"
echo "Manifest version: $PACKAGE_VERSION  id: $ADDON_ID"

# Generate updates.json for Thunderbird auto-update
node -e "
  const fs = require('fs');
  const updates = {
    addons: {
      '$ADDON_ID': {
        updates: [
          {
            version: '$PACKAGE_VERSION',
            target_platform: 'all',
            update_link: '$UPDATE_LINK'
          }
        ]
      }
    }
  };
  fs.writeFileSync('$DIST_DIR/updates.json', JSON.stringify(updates, null, 2) + '\n');
"
echo "Generated: $DIST_DIR/updates.json"

# Package extension
cd "$EXTENSION_DIR"
zip -r "$DIST_DIR/thunderbird-mcp.xpi" . -x "*.DS_Store" -x "*.git*"

echo "Built: $DIST_DIR/thunderbird-mcp.xpi"
