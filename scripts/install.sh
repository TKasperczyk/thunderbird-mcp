#!/bin/bash
# Install the Thunderbird MCP extension

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
DIST_DIR="$PROJECT_DIR/dist"
XPI_FILE="$DIST_DIR/thunderbird-mcp.xpi"
DEFAULT_PROFILE_FILE="$SCRIPT_DIR/default-profile"

DRY_RUN=0
PRINT_PROFILE=0
for arg in "$@"; do
    case "$arg" in
    --dry-run|-n)
        DRY_RUN=1
        ;;
    --print-profile)
        PRINT_PROFILE=1
        DRY_RUN=1
        ;;
    esac
done

canonical_path() {
    local target="$1"

    if command -v realpath >/dev/null 2>&1; then
        realpath "$target" 2>/dev/null || echo "$target"
    elif command -v readlink >/dev/null 2>&1; then
        readlink -f "$target" 2>/dev/null || echo "$target"
    else
        (cd "$target" 2>/dev/null && pwd -P) || echo "$target"
    fi
}

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

linux_profile_dirs() {
    # Traditional .thunderbird
    printf '%s\n' "$HOME/.thunderbird"

    # Modern Snap with current symlink (.config/thunderbird and .thunderbird)
    printf '%s\n' \
        "$HOME/snap/thunderbird/current/.config/thunderbird" \
        "$HOME/snap/thunderbird/current/.thunderbird"

    # Snap common directory
    printf '%s\n' \
        "$HOME/snap/thunderbird/common/.config/thunderbird" \
        "$HOME/snap/thunderbird/common/.thunderbird"

    # Versioned numeric Snap directories (newest revision first)
    if [[ -d "$HOME/snap/thunderbird" ]]; then
        local rev_dirs=()
        local d
        for d in "$HOME/snap/thunderbird/"[0-9]*; do
            if [[ -d "$d" ]]; then
                rev_dirs+=("$d")
            fi
        done
        if [[ "${#rev_dirs[@]}" -gt 0 ]]; then
            while IFS= read -r d; do
                if [[ -n "$d" ]]; then
                    printf '%s\n' "$d/.config/thunderbird" "$d/.thunderbird"
                fi
            done < <(printf '%s\n' "${rev_dirs[@]}" | sort -V -r)
        fi
    fi

    # Flatpak directories (.thunderbird and .config/thunderbird)
    printf '%s\n' \
        "$HOME/.var/app/org.mozilla.Thunderbird/.thunderbird" \
        "$HOME/.var/app/org.mozilla.Thunderbird/.config/thunderbird" \
        "$HOME/.var/app/org.mozilla.thunderbird/.thunderbird" \
        "$HOME/.var/app/org.mozilla.thunderbird/.config/thunderbird" \
        "$HOME/.var/app/eu.betterbird.Betterbird/.thunderbird" \
        "$HOME/.var/app/eu.betterbird.Betterbird/.config/thunderbird"
}

platform_profile_roots() {
    case "$(uname -s)" in
    Darwin)
        printf '%s\n' \
            "$HOME/Library/Thunderbird/Profiles" \
            "$HOME/Library/Application Support/Thunderbird/Profiles"
        ;;
    Linux)
        linux_profile_dirs
        ;;
    *)
        echo "Error: Unsupported platform: $(uname -s)" >&2
        exit 1
        ;;
    esac
}

platform_profile_config_dirs() {
    case "$(uname -s)" in
    Darwin)
        printf '%s\n' \
            "$HOME/Library/Thunderbird" \
            "$HOME/Library/Application Support/Thunderbird"
        ;;
    Linux)
        linux_profile_dirs
        ;;
    *)
        echo "Error: Unsupported platform: $(uname -s)" >&2
        exit 1
        ;;
    esac
}

emit_profile_if_valid() {
    local profile_dir="$1"

    if [[ -d "$profile_dir" && -f "$profile_dir/prefs.js" ]]; then
        echo "$profile_dir"
    fi
}

resolve_profile_path() {
    local config_dir="$1"
    local is_relative="$2"
    local profile_path="$3"

    if [[ "$is_relative" == "1" ]]; then
        echo "$config_dir/$profile_path"
    else
        echo "$profile_path"
    fi
}

find_profiles_from_ini() {
    local config_dir
    local ini_file

    while IFS= read -r config_dir; do
        ini_file="$config_dir/profiles.ini"
        if [[ -f "$ini_file" ]]; then
            parse_profiles_ini "$config_dir" "$ini_file"
        fi
    done < <(platform_profile_config_dirs)
}

parse_profiles_ini() {
    local config_dir="$1"
    local ini_file="$2"
    local in_profile=0
    local is_relative=""
    local profile_path=""
    local is_default=""
    local default_profiles=()
    local other_profiles=()
    local resolved_profile
    local line

    add_ini_profile() {
        if [[ "$in_profile" -eq 1 && -n "$profile_path" ]]; then
            resolved_profile="$(resolve_profile_path "$config_dir" "$is_relative" "$profile_path")"
            if [[ -d "$resolved_profile" && -f "$resolved_profile/prefs.js" ]]; then
                if [[ "$is_default" == "1" ]]; then
                    default_profiles+=("$resolved_profile")
                else
                    other_profiles+=("$resolved_profile")
                fi
            fi
        fi

        is_relative=""
        profile_path=""
        is_default=""
    }

    while IFS= read -r line || [[ -n "$line" ]]; do
        case "$line" in
        "[Profile"*"]")
            add_ini_profile
            in_profile=1
            ;;
        "["*"]")
            add_ini_profile
            in_profile=0
            ;;
        IsRelative=*)
            if [[ "$in_profile" -eq 1 ]]; then
                is_relative="${line#IsRelative=}"
            fi
            ;;
        Path=*)
            if [[ "$in_profile" -eq 1 ]]; then
                profile_path="${line#Path=}"
            fi
            ;;
        Default=*)
            if [[ "$in_profile" -eq 1 ]]; then
                is_default="${line#Default=}"
            fi
            ;;
        esac
    done <"$ini_file"

    add_ini_profile

    if [[ "${#default_profiles[@]}" -gt 0 ]]; then
        printf '%s\n' "${default_profiles[@]}"
    fi
    if [[ "${#other_profiles[@]}" -gt 0 ]]; then
        printf '%s\n' "${other_profiles[@]}"
    fi
}

find_profile_dirs() {
    local roots=()
    local root
    local profile
    local ini_profiles=()
    local seen_canons=()
    local canon

    is_already_seen() {
        local check="$1"
        local s
        for s in "${seen_canons[@]}"; do
            if [[ "$s" == "$check" ]]; then
                return 0
            fi
        done
        return 1
    }

    while IFS= read -r profile; do
        if [[ -z "$profile" ]]; then
            continue
        fi
        canon="$(canonical_path "$profile")"
        if ! is_already_seen "$canon"; then
            seen_canons+=("$canon")
            ini_profiles+=("$profile")
        fi
    done < <(find_profiles_from_ini)

    if [[ "${#ini_profiles[@]}" -gt 0 ]]; then
        printf '%s\n' "${ini_profiles[@]}"
        return
    fi

    while IFS= read -r root; do
        roots+=("$root")
    done < <(platform_profile_roots)

    for root in "${roots[@]}"; do
        if [[ ! -d "$root" ]]; then
            continue
        fi

        while IFS= read -r profile; do
            canon="$(canonical_path "$profile")"
            if ! is_already_seen "$canon"; then
                seen_canons+=("$canon")
                emit_profile_if_valid "$profile"
            fi
        done < <(
            find "$root" \
                -maxdepth 1 \
                -type d \
                \( -name "*.default-release" -o -name "*.default" \) 2>/dev/null |
                sort
        )

        while IFS= read -r profile; do
            canon="$(canonical_path "$profile")"
            if ! is_already_seen "$canon"; then
                seen_canons+=("$canon")
                emit_profile_if_valid "$profile"
            fi
        done < <(
            find "$root" \
                -maxdepth 1 \
                -type d \
                ! -path "$root" \
                ! -name "*.default-release" \
                ! -name "*.default" 2>/dev/null |
                sort
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

    if ! printf '%s\n' "$profile_dir" >"$DEFAULT_PROFILE_FILE"; then
        echo "Warning: Could not save default profile to $DEFAULT_PROFILE_FILE" >&2
    fi
}

select_profile() {
    local profiles=("$@")
    local default_profile
    local index
    local choice

    if default_profile="$(read_default_profile)" &&
        profile_is_available "$default_profile" "${profiles[@]}"; then
        echo "$default_profile"
        return
    fi

    if [[ "${#profiles[@]}" -eq 1 ]]; then
        if [[ "$DRY_RUN" -eq 0 ]]; then
            save_default_profile "${profiles[0]}"
        fi
        echo "${profiles[0]}"
        return
    fi

    if [[ ! -t 0 ]]; then
        echo "Error: Found multiple Thunderbird profiles and cannot prompt" \
            "in non-interactive mode." >&2
        echo "Run scripts/install.sh interactively once to choose and save a default profile." >&2
        exit 1
    fi

    echo "Found ${#profiles[@]} Thunderbird profiles:" >&2
    for index in "${!profiles[@]}"; do
        printf '  %d) %s  size=%s  modified=%s\n' \
            "$((index + 1))" \
            "${profiles[$index]}" \
            "$(profile_size "${profiles[$index]}")" \
            "$(format_mtime "${profiles[$index]}")" >&2
    done

    while true; do
        printf 'Select profile to install into [1-%d]: ' "${#profiles[@]}" >&2
        if ! read -r choice; then
            echo "Error: Could not read profile selection." >&2
            exit 1
        fi

        if [[ "$choice" =~ ^[0-9]+$ ]] &&
            ((choice >= 1)) &&
            ((choice <= ${#profiles[@]})); then
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

PROFILE_DIR=$(find_profile)

if [[ "$PRINT_PROFILE" -eq 1 ]]; then
    echo "$PROFILE_DIR"
    exit 0
fi

if [[ "$DRY_RUN" -eq 1 ]]; then
    echo "Dry run: would install to profile: $PROFILE_DIR"
    exit 0
fi

# Build if needed
if [[ ! -f "$XPI_FILE" ]]; then
    echo "Building extension first..."
    "$SCRIPT_DIR/build.sh"
fi

EXTENSIONS_DIR="$PROFILE_DIR/extensions"

echo "Installing to profile: $PROFILE_DIR"

# Create extensions directory if needed
mkdir -p "$EXTENSIONS_DIR"

# Copy extension
echo "Warning: Re-running install.sh overwrites the AddonManager-managed file and can downgrade an auto-updated build; to bootstrap once, use Tools > Add-ons > Install from File, then let auto-update handle the rest."
cp "$XPI_FILE" "$EXTENSIONS_DIR/thunderbird-mcp@tkasperczyk.dev.xpi"

echo "Installed! Restart Thunderbird to activate."
echo ""
echo "To configure your MCP client, add to your MCP settings:"
echo "  thunderbird-mail: node $PROJECT_DIR/mcp-bridge.cjs"
