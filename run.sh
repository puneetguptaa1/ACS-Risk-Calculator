#!/usr/bin/env bash
#
# NSQIP Risk Calculator Automation -- macOS / Linux launcher
#
# This script:
#   1. Checks that Python 3 is installed
#   2. Creates a virtual environment (venv/) if it does not exist
#   3. Installs Python dependencies
#   4. Installs the Chrome browser driver for Playwright
#   5. Launches the interactive program
#
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# ── 1. Check for Python 3 ──────────────────────────────────────────
PYTHON=""
for candidate in python3 python; do
    if command -v "$candidate" &>/dev/null; then
        version=$("$candidate" --version 2>&1 | grep -oE '[0-9]+\.[0-9]+')
        major=$(echo "$version" | cut -d. -f1)
        minor=$(echo "$version" | cut -d. -f2)
        if [ "$major" -ge 3 ] && [ "$minor" -ge 9 ]; then
            PYTHON="$candidate"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo ""
    echo "ERROR: Python 3.9 or later is required but was not found."
    echo ""
    echo "Install it from: https://www.python.org/downloads/"
    echo ""
    exit 1
fi

echo ""
echo "Using $($PYTHON --version) at $(command -v $PYTHON)"

# ── 2. Create venv if needed ───────────────────────────────────────
if [ ! -d "venv" ]; then
    echo ""
    echo "Creating virtual environment..."
    "$PYTHON" -m venv venv
fi

# ── 3. Activate venv ───────────────────────────────────────────────
source venv/bin/activate

# ── 4. Install dependencies ────────────────────────────────────────
echo ""
echo "Installing dependencies (this may take a moment on first run)..."
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

# ── 5. Install Chrome driver for Playwright ────────────────────────
echo ""
echo "Checking Playwright browser..."
playwright install chrome 2>/dev/null || playwright install chromium

# ── 6. Launch ──────────────────────────────────────────────────────
echo ""
python launcher.py
