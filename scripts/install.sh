#!/bin/bash
# Install the Thunderbird MCP extension

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
DIST_DIR="$PROJECT_DIR/dist"
XPI_FILE="$DIST_DIR/thunderbird-mcp.xpi"

# Find Thunderbird profile directory
find_profile() {
    # Native profile is preferred when multiple install types coexist --
    # vestigial Flatpak profile dirs from prior experiments must not
    # hijack a native install that is actually in use.
    local profile_roots=(
        "$HOME/.thunderbird"
        "$HOME/.var/app/org.mozilla.Thunderbird/.thunderbird"
        "$HOME/.var/app/org.mozilla.thunderbird/.thunderbird"
        "$HOME/.var/app/eu.betterbird.Betterbird/.thunderbird"
    )
    local profiles_dir=""
    local profile=""

    for candidate in "${profile_roots[@]}"; do
        if [[ -d "$candidate" ]]; then
            profiles_dir="$candidate"
            break
        fi
    done

    if [[ -z "$profiles_dir" ]]; then
        echo "Error: Thunderbird profiles directory not found (checked ${profile_roots[*]})" >&2
        exit 1
    fi

    # Look for default-release profile first, then any .default profile
    profile=$(ls -d "$profiles_dir"/*.default-release 2>/dev/null | head -1)
    if [[ -z "$profile" ]]; then
        profile=$(ls -d "$profiles_dir"/*.default 2>/dev/null | head -1)
    fi

    if [[ -z "$profile" ]]; then
        echo "Error: No Thunderbird profile found in $profiles_dir" >&2
        exit 1
    fi

    echo "$profile"
}

# Build if needed
if [[ ! -f "$XPI_FILE" ]]; then
    echo "Building extension first..."
    "$SCRIPT_DIR/build.sh"
fi

PROFILE_DIR=$(find_profile)
EXTENSIONS_DIR="$PROFILE_DIR/extensions"

echo "Installing to profile: $PROFILE_DIR"

# Create extensions directory if needed
mkdir -p "$EXTENSIONS_DIR"

# Copy extension
cp "$XPI_FILE" "$EXTENSIONS_DIR/thunderbird-mcp@tkasperczyk.dev.xpi"

echo "Installed! Restart Thunderbird to activate."
echo ""
echo "To configure your MCP client, add to your MCP settings:"
echo "  thunderbird-mail: node $PROJECT_DIR/mcp-bridge.cjs"
