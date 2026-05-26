#!/bin/bash
# Install the Thunderbird MCP extension

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
DIST_DIR="$PROJECT_DIR/dist"
XPI_FILE="$DIST_DIR/thunderbird-mcp.xpi"
DEFAULT_PROFILE_FILE="$SCRIPT_DIR/default-profile"

format_mtime() {
    local profile_dir="$1"

    if [[ "$(uname -s)" == "Darwin" ]]; then
        stat -f "%Sm" -t "%Y-%m-%d %H:%M:%S" "$profile_dir"
    else
        stat -c "%y" "$profile_dir" | cut -d'.' -f1
    fi
}

profile_size() {
    local profile_dir="$1"

    du -sh "$profile_dir" 2>/dev/null | awk '{print $1}'
}

profile_name() {
    local profile_dir="$1"

    basename "$profile_dir"
}

platform_profile_roots() {
    case "$(uname -s)" in
        Darwin)
            printf '%s\n' \
                "$HOME/Library/Thunderbird/Profiles" \
                "$HOME/Library/Application Support/Thunderbird/Profiles"
            ;;
        Linux)
            printf '%s\n' \
                "$HOME/.thunderbird" \
                "$HOME/.var/app/org.mozilla.Thunderbird/.thunderbird" \
                "$HOME/.var/app/org.mozilla.thunderbird/.thunderbird" \
                "$HOME/.var/app/eu.betterbird.Betterbird/.thunderbird"
            ;;
        *)
            echo "Error: Unsupported platform: $(uname -s)" >&2
            exit 1
            ;;
    esac
}

find_profile_dirs() {
    local roots=()
    local root
    local profile

    while IFS= read -r root; do
        roots+=("$root")
    done < <(platform_profile_roots)

    for root in "${roots[@]}"; do
        if [[ ! -d "$root" ]]; then
            continue
        fi

        while IFS= read -r profile; do
            echo "$profile"
        done < <(
            find "$root" \
                -maxdepth 1 \
                -type d \
                \( -name "*.default-release" -o -name "*.default" \) \
                | sort
        )

        while IFS= read -r profile; do
            echo "$profile"
        done < <(
            find "$root" \
                -maxdepth 1 \
                -type d \
                ! -path "$root" \
                ! -name "*.default-release" \
                ! -name "*.default" \
                | sort
        )
    done
}

profile_is_available() {
    local candidate="$1"
    shift
    local profile

    for profile in "$@"; do
        if [[ "$profile" == "$candidate" ]]; then
            return 0
        fi
    done

    return 1
}

read_default_profile() {
    if [[ ! -f "$DEFAULT_PROFILE_FILE" ]]; then
        return 1
    fi

    head -n 1 "$DEFAULT_PROFILE_FILE"
}

save_default_profile() {
    local profile_dir="$1"

    printf '%s\n' "$profile_dir" > "$DEFAULT_PROFILE_FILE"
}

select_profile() {
    local profiles=("$@")
    local default_profile
    local index
    local choice

    if default_profile="$(read_default_profile)" \
        && profile_is_available "$default_profile" "${profiles[@]}"; then
        echo "$default_profile"
        return
    fi

    if [[ "${#profiles[@]}" -eq 1 ]]; then
        save_default_profile "${profiles[0]}"
        echo "${profiles[0]}"
        return
    fi

    echo "Found ${#profiles[@]} Thunderbird profiles:" >&2
    for index in "${!profiles[@]}"; do
        printf '  %d) %s  size=%s  modified=%s\n' \
            "$((index + 1))" \
            "$(profile_name "${profiles[$index]}")" \
            "$(profile_size "${profiles[$index]}")" \
            "$(format_mtime "${profiles[$index]}")" >&2
    done

    while true; do
        printf 'Select profile to install into [1-%d]: ' "${#profiles[@]}" >&2
        read -r choice

        if [[ "$choice" =~ ^[0-9]+$ ]] \
            && (( choice >= 1 )) \
            && (( choice <= ${#profiles[@]} )); then
            save_default_profile "${profiles[$((choice - 1))]}"
            echo "${profiles[$((choice - 1))]}"
            return
        fi

        echo "Invalid selection: $choice" >&2
    done
}

# Find Thunderbird profile directory for the current platform.
find_profile() {
    local profiles=()
    local profile

    while IFS= read -r profile; do
        profiles+=("$profile")
    done < <(find_profile_dirs)

    if [[ "${#profiles[@]}" -eq 0 ]]; then
        echo "Error: No Thunderbird profile found for $(uname -s)" >&2
        exit 1
    fi

    select_profile "${profiles[@]}"
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
