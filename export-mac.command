#!/bin/bash

# ============================================================
#  Teams Chat Export — macOS Launcher
#  Double-click this file to export your Teams chats.
#  All data stays on this machine.
# ============================================================

# Keep window open on any error
trap 'echo ""; echo "  An error occurred. See messages above."; read -rp "  Press Enter to close..."' ERR
set -euo pipefail

cd "$(dirname "$0")"
ROOT="$(pwd)"

clear
echo ""
echo "  ╔════════════════════════════════════════════════════════════╗"
echo "  ║                                                            ║"
echo "  ║   Teams Chat Export                                        ║"
echo "  ║                                                            ║"
echo "  ║   This tool exports your Teams chats                       ║"
echo "  ║   into an HTML file you can keep.                          ║"
echo "  ║                                                            ║"
echo "  ║   Requirements:                                            ║"
echo "  ║   • Microsoft Edge must be CLOSED                          ║"
echo "  ║   • You should already be logged into Teams in Edge        ║"
echo "  ║                                                            ║"
echo "  ║   No data is sent anywhere — everything stays local.       ║"
echo "  ║                                                            ║"
echo "  ╚════════════════════════════════════════════════════════════╝"
echo ""
echo "  Press Enter to start (or Ctrl+C to cancel)..."
read -r

# ---- Step 1: Find or install Python ----
echo "[1/6] Checking Python..."

PYTHON=""
MIN_MINOR=11

find_good_python() {
    local cmd="$1"
    if command -v "$cmd" &>/dev/null; then
        local major minor
        major=$("$cmd" -c "import sys; print(sys.version_info.major)" 2>/dev/null) || return 1
        minor=$("$cmd" -c "import sys; print(sys.version_info.minor)" 2>/dev/null) || return 1
        if [ "$major" -eq 3 ] && [ "$minor" -ge "$MIN_MINOR" ]; then
            PYTHON="$cmd"
            return 0
        fi
    fi
    return 1
}

# Check local standalone Python first (faster on re-runs)
STANDALONE_PYTHON="$ROOT/.python/bin/python3"
if [ -x "$STANDALONE_PYTHON" ]; then
    PYTHON="$STANDALONE_PYTHON"
fi

# Then check system Python
if [ -z "$PYTHON" ]; then
    find_good_python python3 || \
    find_good_python python || \
    find_good_python /opt/homebrew/bin/python3 || \
    find_good_python /usr/local/bin/python3 || \
    true
fi

# Still no Python? Download standalone build
if [ -z "$PYTHON" ]; then
    echo "  Python 3.11+ not found. Downloading portable Python..."
    echo ""

    PYTHON_DIR="$ROOT/.python"
    mkdir -p "$PYTHON_DIR"

    # Detect architecture
    ARCH=$(uname -m)
    if [ "$ARCH" = "arm64" ]; then
        PYTHON_URL="https://github.com/astral-sh/python-build-standalone/releases/download/20241016/cpython-3.12.7+20241016-aarch64-apple-darwin-install_only_stripped.tar.gz"
    else
        PYTHON_URL="https://github.com/astral-sh/python-build-standalone/releases/download/20241016/cpython-3.12.7+20241016-x86_64-apple-darwin-install_only_stripped.tar.gz"
    fi

    echo "  Downloading Python 3.12 for $ARCH..."
    PYTHON_TAR="$PYTHON_DIR/python.tar.gz"

    if command -v curl &>/dev/null; then
        curl -fSL --progress-bar "$PYTHON_URL" -o "$PYTHON_TAR"
    elif command -v wget &>/dev/null; then
        wget -q --show-progress "$PYTHON_URL" -O "$PYTHON_TAR"
    else
        echo "  ERROR: Neither curl nor wget available."
        read -rp "  Press Enter to close..."
        exit 1
    fi

    if [ ! -f "$PYTHON_TAR" ]; then
        echo "  ERROR: Could not download Python. Check your internet connection."
        read -rp "  Press Enter to close..."
        exit 1
    fi

    echo "  Extracting Python..."
    tar -xzf "$PYTHON_TAR" -C "$PYTHON_DIR" --strip-components=1
    rm -f "$PYTHON_TAR"

    PYTHON="$PYTHON_DIR/bin/python3"

    if [ ! -x "$PYTHON" ]; then
        echo "  ERROR: Python extraction failed."
        read -rp "  Press Enter to close..."
        exit 1
    fi

    echo "  Portable Python installed OK."
fi

PY_VERSION=$("$PYTHON" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
echo "  OK: $PYTHON ($PY_VERSION)"

# ---- Step 2: Create venv and install dependencies ----
echo ""
echo "[2/6] Installing dependencies (first run takes a while)..."

VENV_DIR="$ROOT/.venv-export"

if [ ! -d "$VENV_DIR" ]; then
    "$PYTHON" -m venv "$VENV_DIR"
fi

# Activate venv
source "$VENV_DIR/bin/activate"

# Fix for macOS expat issue with Python 3.13+
if [ -d "/opt/homebrew/opt/expat/lib" ]; then
    export DYLD_LIBRARY_PATH="/opt/homebrew/opt/expat/lib${DYLD_LIBRARY_PATH:+:$DYLD_LIBRARY_PATH}"
fi

echo "  Installing pip..."
pip install --upgrade pip 2>&1 | tail -1
echo "  Installing playwright..."
pip install playwright 2>&1 | tail -1
echo "  Installing msteams-export..."
pip install --upgrade --force-reinstall --no-deps "$ROOT" 2>&1 | tail -1

echo "  OK: Dependencies installed."

# ---- Step 3: Install Edge for Playwright ----
echo ""
echo "[3/6] Setting up Playwright browser..."

python -m playwright install msedge 2>/dev/null || python -m playwright install chromium 2>/dev/null || true
echo "  OK."

# ---- Step 4: Login and Export ----
echo ""
echo "[4/6] Starting export..."
echo ""
echo "  A browser window will open with Teams."
echo "  If you are NOT already logged in, log in now."
echo "  Close the browser window when ready."
echo ""

PROFILE_DIR="$ROOT/.profile"
EXPORT_DIR="$ROOT/exports"

# Open interactive session for login
python -m msteams_export session-open --browser edge --profile "$PROFILE_DIR" || \
    python -m msteams_export session-open --browser auto --profile "$PROFILE_DIR"

echo ""
echo "  Starting export of all chats."
echo "  This can take 10-60 minutes depending on chat count."
echo ""

# Run the export
set +e
python -m msteams_export export all \
    --outdir "$EXPORT_DIR" \
    --browser edge \
    --profile "$PROFILE_DIR" \
    --skip-existing || \
python -m msteams_export export all \
    --outdir "$EXPORT_DIR" \
    --browser auto \
    --profile "$PROFILE_DIR" \
    --skip-existing
EXPORT_RC=$?
set -e

if [ $EXPORT_RC -ne 0 ]; then
    echo ""
    echo "  WARNING: Export completed with errors. Trying to generate archive anyway..."
fi

# ---- Step 5: Mirror image attachments ----
echo ""
echo "[5/6] Downloading images from chats (for offline viewing)..."
echo "  This can take a while depending on how many images you have."
set +e
python -m msteams_export attachments mirror "$EXPORT_DIR" \
    --browser edge \
    --profile "$PROFILE_DIR" || \
python -m msteams_export attachments mirror "$EXPORT_DIR" \
    --browser auto \
    --profile "$PROFILE_DIR"
set -e

# ---- Step 6: Generate HTML archive (folder + images) ----
echo ""
echo "[6/6] Generating HTML archive..."

ARCHIVE_DIR="$ROOT/teams-archive"
rm -rf "$ARCHIVE_DIR"
python -m msteams_export generate-html-archive "$EXPORT_DIR" --with-images --output "$ARCHIVE_DIR"

# ---- Ask user where to save the archive folder ----
echo ""
echo "  Choose where to save the archive folder..."

DEST_DIR=$(osascript <<'APPLESCRIPT' 2>/dev/null
try
    set destFolder to choose folder with prompt "Choose where to save your Teams archive folder:"
    return POSIX path of destFolder
on error errMsg number errNum
    return ""
end try
APPLESCRIPT
)

FINAL_DIR="$ARCHIVE_DIR"
if [ -n "$DEST_DIR" ]; then
    TARGET="${DEST_DIR%/}/teams-archive"
    if [ "$TARGET" != "$ARCHIVE_DIR" ]; then
        rm -rf "$TARGET"
        if cp -R "$ARCHIVE_DIR" "$TARGET" 2>/dev/null; then
            FINAL_DIR="$TARGET"
            echo "  Saved to: $FINAL_DIR"
        else
            echo "  Could not save to chosen location, keeping in current folder."
        fi
    fi
else
    echo "  Keeping archive in current folder."
fi

echo ""
echo "  ╔════════════════════════════════════════════════════════════╗"
echo "  ║                                                            ║"
echo "  ║   DONE!                                                    ║"
echo "  ║                                                            ║"
echo "  ║   Your archive has been saved.                             ║"
echo "  ║   Open it in a browser to view your chats.                 ║"
echo "  ║                                                            ║"
echo "  ╚════════════════════════════════════════════════════════════╝"
echo ""
echo "  Location: $FINAL_DIR"
echo ""

# Open the archive and reveal in Finder
open "$FINAL_DIR/index.html" 2>/dev/null || true
open -R "$FINAL_DIR/index.html" 2>/dev/null || true

read -rp "  Press Enter to close..."
