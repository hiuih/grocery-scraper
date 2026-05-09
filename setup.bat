@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul 2>&1
title Grocery Scraper - Setup

echo.
echo ============================================================
echo   GROCERY SCRAPER  -  SETUP
echo   Sets up everything needed on a brand new computer.
echo   Keep this window open until it says DONE.
echo ============================================================
echo.

REM ─────────────────────────────────────────────────────────────
REM  STEP 1 - Check for Python
REM ─────────────────────────────────────────────────────────────
echo [Step 1 of 5]  Checking for Python...

python --version >nul 2>&1
if %errorlevel% == 0 (
    for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
    echo              Found: !PYVER!  -- good to go!
    goto :step2
)

REM Try py launcher as fallback
py --version >nul 2>&1
if %errorlevel% == 0 (
    echo              Found via py launcher -- good to go!
    set PYTHON_CMD=py
    goto :step2
)

REM Python not found -- try winget (built into Windows 10/11)
echo              Python not found. Trying to install automatically...
echo              (This may take 2-3 minutes)
echo.

winget install --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements >nul 2>&1

REM Update PATH for this session
set "PATH=%PATH%;%LOCALAPPDATA%\Programs\Python\Python312"
set "PATH=%PATH%;%LOCALAPPDATA%\Programs\Python\Python312\Scripts"
set "PATH=%PATH%;C:\Program Files\Python312"
set "PATH=%PATH%;C:\Program Files\Python312\Scripts"

python --version >nul 2>&1
if %errorlevel% == 0 (
    echo              Python installed successfully!
    goto :step2
)

REM Last resort: explain manual install
echo.
echo  *** PYTHON NOT FOUND ***
echo.
echo  Please install Python manually, then run setup.bat again:
echo.
echo    1. Open your web browser
echo    2. Go to:  https://www.python.org/downloads/
echo    3. Click the yellow "Download Python" button
echo    4. Open the downloaded file
echo    5. IMPORTANT: Tick the box "Add Python to PATH"
echo    6. Click "Install Now"  and wait for it to finish
echo    7. Close this window and double-click setup.bat again
echo.
pause
exit /b 1

:step2
echo.

REM ─────────────────────────────────────────────────────────────
REM  STEP 2 - Upgrade pip
REM ─────────────────────────────────────────────────────────────
echo [Step 2 of 5]  Updating package manager (pip)...
python -m pip install --upgrade pip -q --no-warn-script-location
echo              [OK] pip updated.
echo.

REM ─────────────────────────────────────────────────────────────
REM  STEP 3 - Install Python packages
REM ─────────────────────────────────────────────────────────────
echo [Step 3 of 5]  Installing required packages...
echo              (playwright + openpyxl -- may take 1-2 minutes)
echo.

python -m pip install playwright openpyxl -q --no-warn-script-location
if %errorlevel% neq 0 (
    echo.
    echo  [!] Package install failed. Please check your internet connection
    echo      and try running setup.bat again.
    echo.
    pause
    exit /b 1
)
echo              [OK] Packages installed.
echo.

REM ─────────────────────────────────────────────────────────────
REM  STEP 4 - Download browser for Playwright
REM ─────────────────────────────────────────────────────────────
echo [Step 4 of 5]  Downloading browser (Chromium, ~130 MB, one-time only)...
echo              This can take several minutes on a slow connection.
echo              Please be patient...
echo.

python -m playwright install chromium --with-deps
if %errorlevel% neq 0 (
    echo.
    echo  [!] Browser download failed.
    echo      Check your internet connection and try setup.bat again.
    echo.
    pause
    exit /b 1
)
echo.
echo              [OK] Browser downloaded and ready.
echo.

REM ─────────────────────────────────────────────────────────────
REM  STEP 5 - Create "Run Grocery Scraper" shortcut on Desktop
REM ─────────────────────────────────────────────────────────────
echo [Step 5 of 5]  Creating Desktop shortcut...

set "SCRIPT_DIR=%~dp0"
REM Remove trailing backslash
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

set "SCRAPER=%SCRIPT_DIR%\grocery_scraper.py"
set "SHORTCUT=%USERPROFILE%\Desktop\Run Grocery Scraper.bat"

(
    echo @echo off
    echo chcp 65001 ^>nul 2^>^&1
    echo title Grocery Scraper - Running...
    echo cd /d "%SCRIPT_DIR%"
    echo echo.
    echo echo Starting scraper - please wait, this takes 15-30 minutes.
    echo echo Results will be saved as Excel files on your Desktop.
    echo echo.
    echo python "%SCRAPER%"
    echo if %%errorlevel%% neq 0 ^(
    echo     echo.
    echo     echo Something went wrong. Please contact support.
    echo     pause
    echo ^)
) > "%SHORTCUT%"

if exist "%SHORTCUT%" (
    echo              [OK] "Run Grocery Scraper.bat" created on your Desktop.
) else (
    echo              [!] Could not create Desktop shortcut.
    echo              You can still run grocery_scraper.py directly.
)
echo.

REM ─────────────────────────────────────────────────────────────
REM  DONE
REM ─────────────────────────────────────────────────────────────
echo ============================================================
echo   SETUP COMPLETE!
echo ============================================================
echo.
echo   Everything is installed. Here is how to use the scraper:
echo.
echo   OPTION A (easiest):
echo     Double-click "Run Grocery Scraper.bat" on your Desktop
echo.
echo   OPTION B:
echo     Double-click grocery_scraper.py in this folder
echo.
echo   The scraper takes 15-30 minutes and saves two Excel files
echo   to your Desktop when it finishes:
echo     - Fresh_St_Market_Products.xlsx
echo     - Save_On_Foods_Products.xlsx
echo.
echo   You only need to run setup.bat ONCE on each computer.
echo ============================================================
echo.
pause
