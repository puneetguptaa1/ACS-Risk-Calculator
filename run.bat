@echo off
REM
REM NSQIP Risk Calculator Automation -- Windows launcher
REM
REM This script:
REM   1. Checks that Python 3 is installed
REM   2. Creates a virtual environment (venv\) if it does not exist
REM   3. Installs Python dependencies
REM   4. Installs the Chrome browser driver for Playwright
REM   5. Launches the interactive program
REM

cd /d "%~dp0"

REM ── 1. Check for Python ──────────────────────────────────────────
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Python is not installed or not in your PATH.
    echo.
    echo Download and install Python 3.9+ from:
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANT: During installation, check "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

python --version
echo.

REM ── 2. Create venv if needed ─────────────────────────────────────
if not exist "venv\" (
    echo Creating virtual environment...
    python -m venv venv
    echo.
)

REM ── 3. Activate venv ─────────────────────────────────────────────
call venv\Scripts\activate.bat

REM ── 4. Install dependencies ──────────────────────────────────────
echo Installing dependencies (this may take a moment on first run)...
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt
echo.

REM ── 5. Install Chrome driver for Playwright ──────────────────────
echo Checking Playwright browser...
playwright install chrome 2>nul || playwright install chromium
echo.

REM ── 6. Launch ────────────────────────────────────────────────────
python launcher.py

pause
