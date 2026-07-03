@echo off
:: Re-launch in persistent cmd so window stays open on crash
if "%1"=="" ( cmd /k "%~f0" run & exit /b )
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion

:: ============================================================
::  Teams Chat Export - Windows Launcher
::  Double-click this file to export your Teams chats.
::  All data stays on this machine.
:: ============================================================

title Teams Chat Export

echo.
echo  ============================================================
echo.
echo   Teams Chat Export
echo.
echo   This tool exports your Teams chats
echo   into an HTML file you can keep.
echo.
echo   Requirements:
echo   - Microsoft Edge must be CLOSED
echo   - Sign in to Teams in the window that opens
echo   - Wait until you SEE your chats before continuing
echo.
echo   No data is sent anywhere - everything stays local.
echo.
echo  ============================================================
echo.

:: Navigate to script directory
cd /d "%~dp0"

:: ---- Step 1: Find or install Python ----
echo [1/6] Checking Python...

:: Check if previously downloaded embedded Python exists
set "PYTHON_DIR=%~dp0.python"
if exist "%PYTHON_DIR%\python.exe" (
    set "PYTHON=%PYTHON_DIR%\python.exe"
    goto :python_found
)

:: Try system Python - must verify it actually works (not Windows Store alias)
py -3 --version >nul 2>&1
if %ERRORLEVEL%==0 (
    set "PYTHON=py -3"
    goto :python_found
)

python3 --version >nul 2>&1
if %ERRORLEVEL%==0 (
    set "PYTHON=python3"
    goto :python_found
)

python --version >nul 2>&1
if %ERRORLEVEL%==0 (
    set "PYTHON=python"
    goto :python_found
)

:: No Python found - download embedded Python
echo   Python not found. Downloading portable Python...
set "PYTHON_DIR=%~dp0.python"
set "PYTHON=%PYTHON_DIR%\python.exe"

if exist "%PYTHON%" goto :python_found

mkdir "%PYTHON_DIR%" 2>nul

:: Download Python Embedded (3.12)
set "PYTHON_URL=https://www.python.org/ftp/python/3.12.7/python-3.12.7-embed-amd64.zip"
set "PYTHON_ZIP=%PYTHON_DIR%\python.zip"

echo   Downloading Python 3.12...
powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_ZIP%' }"
if %ERRORLEVEL% neq 0 (
    echo   ERROR: Could not download Python. Check your internet connection.
    pause
    exit /b 1
)

echo   Extracting Python...
powershell -Command "Expand-Archive -Path '%PYTHON_ZIP%' -DestinationPath '%PYTHON_DIR%' -Force"
del "%PYTHON_ZIP%" 2>nul

:: Enable pip in embedded Python (uncomment import site in python312._pth)
set "PTH_FILE=%PYTHON_DIR%\python312._pth"
if exist "%PTH_FILE%" (
    powershell -Command "(Get-Content '%PTH_FILE%') -replace '#import site','import site' | Set-Content '%PTH_FILE%'"
)

:: Install pip
echo   Installing pip...
powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%PYTHON_DIR%\get-pip.py' }"
"%PYTHON%" "%PYTHON_DIR%\get-pip.py" --no-warn-script-location >nul 2>&1

:python_found
echo   OK: Using %PYTHON%

:: ---- Step 2: Install dependencies ----
echo.
echo [2/6] Installing dependencies (first run takes a while)...

"%PYTHON%" -m ensurepip --upgrade >nul 2>&1

echo   Installing pip and setuptools...
"%PYTHON%" -m pip install --upgrade pip setuptools wheel 2>&1
if %ERRORLEVEL% neq 0 (
    echo   ERROR: pip upgrade failed. Continuing anyway...
)

echo   Installing playwright...
"%PYTHON%" -m pip install playwright 2>&1
if %ERRORLEVEL% neq 0 (
    echo   ERROR: Could not install playwright.
    pause
    exit /b 1
)

echo   Installing msteams-export...
"%PYTHON%" -m pip install --upgrade --force-reinstall --no-deps "%~dp0." 2>&1
if %ERRORLEVEL% neq 0 (
    echo   ERROR: Could not install msteams-export.
    echo   Full path: "%~dp0."
    pause
    exit /b 1
)

echo   OK: Dependencies installed.

:: ---- Step 2b: Ensure Visual C++ Runtime is installed ----
:: greenlet (used by Playwright) needs vcruntime140_1.dll
"%PYTHON%" -c "import _greenlet" >nul 2>&1
if %ERRORLEVEL%==0 goto :vcok

echo.
echo   Visual C++ Runtime is missing (needed by Playwright).
echo   Downloading and installing Microsoft Visual C++ Redistributable...
set "VCREDIST=%PYTHON_DIR%\vc_redist.x64.exe"
powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://aka.ms/vs/17/release/vc_redist.x64.exe' -OutFile '%VCREDIST%' }"
if not exist "%VCREDIST%" (
    echo   ERROR: Could not download Visual C++ Runtime.
    echo   Please install it manually from: https://aka.ms/vs/17/release/vc_redist.x64.exe
    pause
    exit /b 1
)
echo   Installing... (may ask for admin permission)
start /wait "" "%VCREDIST%" /install /quiet /norestart
del "%VCREDIST%" 2>nul
echo   OK: Visual C++ Runtime installed.

:vcok

:: ---- Step 3: Install Edge for Playwright ----
echo.
echo [3/6] Setting up Playwright browser...

"%PYTHON%" -m playwright install msedge >nul 2>&1

:: ---- Step 4: Login and Export ----
echo.
echo [4/6] Starting export...
echo.
echo   A browser window (Edge) will open with TWO tabs.
echo.
echo   1) TAB 1 is Teams. Check it is signed in as the EXACT person whose
echo      chats you want to export. Edge often signs in as the WRONG
echo      account automatically - if so, sign out and sign back in.
echo   2) TAB 2 (Office) opens by itself. Check it too is signed in as the
echo      SAME, correct account. If it is wrong, sign out and sign back in.
echo   3) Both tabs MUST be the same, correct account - otherwise you get
echo      the wrong person's data, or no documents at all.
echo   4) When both tabs are correct, come back to THIS window and press Enter.
echo.

set "PROFILE_DIR=%~dp0.profile"
set "EXPORT_DIR=%~dp0exports"

:: Open a second tab so the same profile captures the Microsoft 365 sign-in
:: cookies needed to mirror SharePoint/OneDrive-hosted Office documents.
:: Defaults to the universal Office URL (works for any tenant, no company
:: address required); override with SHAREPOINT_URL for a specific tenant.
set "OFFICE_URL=https://powerpoint.cloud.microsoft/"
if defined SHAREPOINT_URL set "OFFICE_URL=%SHAREPOINT_URL%"
set SESSION_EXTRA=--also-url "%OFFICE_URL%"

:: Open interactive session for login
"%PYTHON%" -m msteams_export session-open --browser edge --profile "%PROFILE_DIR%" %SESSION_EXTRA%

echo.
echo   Starting export of all chats.
echo   This can take 10-60 minutes depending on chat count.
echo.

:: Run the export
"%PYTHON%" -m msteams_export export all --outdir "%EXPORT_DIR%" --browser edge --profile "%PROFILE_DIR%" --skip-existing
set "EXPORT_RC=%ERRORLEVEL%"

if not exist "%EXPORT_DIR%\index.json" (
    echo.
    echo   ERROR: No chats were exported, so there is nothing to process.
    echo.
    echo   This almost always means Teams was not signed in when the
    echo   export ran. The tool opens its OWN browser window with its own
    echo   sign-in - being signed in to your normal Edge does NOT count.
    echo.
    echo   Please run this again, and during sign-in:
    echo     1^) Sign in to Teams in the window that opens.
    echo     2^) WAIT until you can actually SEE your chats.
    echo     3^) Then come back here and press Enter.
    echo.
    pause
    exit /b 1
)

if not "%EXPORT_RC%"=="0" (
    echo.
    echo   WARNING: Export finished with some errors. Continuing with what was exported...
)

:: ---- Step 5: Mirror image attachments ----
echo.
echo [5/6] Downloading images from chats (for offline viewing)...
echo   This can take a while depending on how many images you have.
:: Optional: override the per-download pause (ms). Lower = faster; raise if throttled.
set "MIRROR_EXTRA="
if defined MIRROR_SPACING_MS set MIRROR_EXTRA=--spacing-ms %MIRROR_SPACING_MS%
"%PYTHON%" -m msteams_export attachments mirror "%EXPORT_DIR%" --browser edge --profile "%PROFILE_DIR%" %MIRROR_EXTRA%

:: ---- Step 6: Generate HTML Archive (folder + images) ----
echo.
echo [6/6] Generating HTML archive...

set "ARCHIVE_DIR=%~dp0teams-archive"
if exist "%ARCHIVE_DIR%" rmdir /S /Q "%ARCHIVE_DIR%"
"%PYTHON%" -m msteams_export generate-html-archive "%EXPORT_DIR%" --with-images --output "%ARCHIVE_DIR%"

if %ERRORLEVEL% neq 0 (
    echo   ERROR: Could not generate HTML archive.
    pause
    exit /b 1
)

:: ---- Ask user where to save the archive folder ----
echo.
echo   Choose where to save the archive folder...

set "DEST_DIR="
for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.FolderBrowserDialog; $d.Description = 'Choose where to save your Teams archive folder'; if ($d.ShowDialog() -eq 'OK') { Write-Output $d.SelectedPath }"`) do set "DEST_DIR=%%P"

set "FINAL_DIR=%ARCHIVE_DIR%"
if defined DEST_DIR set "TARGET=%DEST_DIR%\teams-archive"
if defined DEST_DIR (
    if /I not "%TARGET%"=="%ARCHIVE_DIR%" (
        if exist "%TARGET%" rmdir /S /Q "%TARGET%"
        robocopy "%ARCHIVE_DIR%" "%TARGET%" /E /NFL /NDL /NJH /NJS /NC /NS >nul
        if exist "%TARGET%\index.html" (
            set "FINAL_DIR=%TARGET%"
            echo   Saved to: %TARGET%
        ) else (
            echo   Could not save to chosen location, keeping in current folder.
        )
    ) else (
        echo   Keeping archive in current folder.
    )
) else (
    echo   Keeping archive in current folder.
)

echo.
echo  ============================================================
echo.
echo   DONE!
echo.
echo   Your archive has been saved.
echo   Open it in a browser to view your chats.
echo.
echo  ============================================================
echo.
echo   Location: %FINAL_DIR%
echo.

:: Open the archive and reveal in Explorer
start "" "%FINAL_DIR%\index.html"
explorer /select,"%FINAL_DIR%\index.html"

pause
