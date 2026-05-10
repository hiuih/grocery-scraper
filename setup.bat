@echo off
setlocal EnableDelayedExpansion
title Grocery Scraper - Setup

echo.
echo ============================================================
echo   GROCERY SCRAPER - SETUP
echo   Sets up everything needed on a brand new computer.
echo   Keep this window open until it says DONE.
echo ============================================================
echo.

REM --- STEP 1: Check for Python ---
echo [Step 1 of 5]  Checking for Python...

python --version >nul 2>&1
if %errorlevel% == 0 (
    for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
    echo              Found: !PYVER!
    goto :step2
)

py --version >nul 2>&1
if %errorlevel% == 0 (
    echo              Found via py launcher.
    set PYTHON=py
    goto :step2
)

echo              Python not found. Installing via Windows Package Manager...
echo              This may take 2-3 minutes, please wait.
echo.

winget install --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements >nul 2>&1

set "PATH=%PATH%;%LOCALAPPDATA%\Programs\Python\Python312"
set "PATH=%PATH%;%LOCALAPPDATA%\Programs\Python\Python312\Scripts"
set "PATH=%PATH%;C:\Program Files\Python312"
set "PATH=%PATH%;C:\Program Files\Python312\Scripts"

python --version >nul 2>&1
if %errorlevel% == 0 (
    echo              Python installed OK.
    goto :step2
)

echo.
echo  *** PYTHON NOT FOUND - MANUAL INSTALL NEEDED ***
echo.
echo  1. Open your browser and go to:  https://www.python.org/downloads/
echo  2. Click the big Download Python button
echo  3. Run the downloaded installer
echo  4. IMPORTANT: Check "Add Python to PATH" before clicking Install
echo  5. Click Install Now and wait for it to finish
echo  6. Close this window, then double-click setup.bat again
echo.
pause
exit /b 1

:step2
echo              Python is ready.
echo.

REM --- STEP 2: Upgrade pip ---
echo [Step 2 of 5]  Updating pip...
python -m pip install --upgrade pip -q --no-warn-script-location
echo              Done.
echo.

REM --- STEP 3: Install packages ---
echo [Step 3 of 5]  Installing packages (playwright and openpyxl)...
echo              This may take 1-2 minutes...
echo.

python -m pip install playwright openpyxl -q --no-warn-script-location
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Package install failed.
    echo  Check your internet connection and try setup.bat again.
    echo.
    pause
    exit /b 1
)
echo              Packages installed OK.
echo.

REM --- STEP 4: Download browser ---
echo [Step 4 of 5]  Downloading browser for scraping (~130 MB, one time only)...
echo              This may take several minutes on a slow connection.
echo              Please wait...
echo.

python -m playwright install chromium --with-deps
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Browser download failed.
    echo  Check your internet connection and try setup.bat again.
    echo.
    pause
    exit /b 1
)
echo.
echo              Browser downloaded OK.
echo.

REM --- STEP 5: Create shortcut on Desktop ---
echo [Step 5 of 5]  Creating shortcut on your Desktop...

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "SCRAPER=%SCRIPT_DIR%\grocery_scraper.py"

REM Get the REAL Desktop path via PowerShell (works with OneDrive too)
for /f "usebackq delims=" %%D in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do set "DESKTOP=%%D"

if "%DESKTOP%"=="" set "DESKTOP=%USERPROFILE%\Desktop"

echo              Desktop found at: %DESKTOP%

REM Create a proper .lnk shortcut with icon using PowerShell
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; ^
   $s = $ws.CreateShortcut('%DESKTOP%\Run Grocery Scraper.lnk'); ^
   $s.TargetPath = 'python'; ^
   $s.Arguments = '"%SCRAPER%"'; ^
   $s.WorkingDirectory = '%SCRIPT_DIR%'; ^
   $s.Description = 'Run the Grocery Price Scraper'; ^
   $s.IconLocation = 'C:\Windows\System32\SHELL32.dll,23'; ^
   $s.Save()"

if exist "%DESKTOP%\Run Grocery Scraper.lnk" (
    echo              Shortcut created: "Run Grocery Scraper" on your Desktop.
) else (
    echo              Could not create .lnk shortcut. Creating .bat instead...
    (
        echo @echo off
        echo title Grocery Scraper - Running...
        echo cd /d "%SCRIPT_DIR%"
        echo echo.
        echo echo Starting scraper - please wait 15-30 minutes.
        echo echo Results will be saved as Excel files on your Desktop.
        echo echo.
        echo python "%SCRAPER%"
        echo if %%errorlevel%% neq 0 pause
    ) > "%DESKTOP%\Run Grocery Scraper.bat"
    echo              .bat shortcut created instead.
)
echo.

REM --- DONE ---
echo ============================================================
echo   SETUP COMPLETE
echo ============================================================
echo.
echo   HOW TO RUN THE SCRAPER:
echo.
echo     Double-click "Run Grocery Scraper" on your Desktop
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
