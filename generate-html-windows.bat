@echo off
:: Re-launch in persistent cmd so the window stays open on any error
if "%1"=="" ( cmd /k "%~f0" run & exit /b )
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion

:: ============================================================
::  Teams Chat Export - Generate HTML only
::  Builds the teams-archive\ folder from an EXISTING exports\ folder.
::  Does NOT log in and does NOT re-run the export.
:: ============================================================

title Teams Chat Export - Generate HTML

cd /d "%~dp0"

echo.
echo  ============================================================
echo.
echo   Generating teams-archive\ folder from existing exports...
echo   (No login, no re-export - uses the exports\ folder.)
echo.
echo  ============================================================
echo.

:: ---- Make sure the exports folder is present ----
if not exist "%~dp0exports\index.json" (
    echo   ERROR: Could not find exports\index.json next to this file.
    echo   Make sure the 'exports' folder is in the same place as this script.
    pause
    exit /b 1
)

:: ---- Find Python (reuse downloaded portable Python if present) ----
echo [1/2] Checking Python...

set "PYTHON_DIR=%~dp0.python"
if exist "%PYTHON_DIR%\python.exe" (
    set "PYTHON=%PYTHON_DIR%\python.exe"
    goto :python_found
)

py -3 --version >nul 2>&1
if %ERRORLEVEL%==0 ( set "PYTHON=py -3" & goto :python_found )

python --version >nul 2>&1
if %ERRORLEVEL%==0 ( set "PYTHON=python" & goto :python_found )

:: No Python found - download portable embedded Python.
:: The HTML generator only uses Python's standard library, so we do
:: NOT need pip or any extra packages here.
echo   Python not found. Downloading portable Python...
mkdir "%PYTHON_DIR%" 2>nul
set "PYTHON_URL=https://www.python.org/ftp/python/3.12.7/python-3.12.7-embed-amd64.zip"
set "PYTHON_ZIP=%PYTHON_DIR%\python.zip"
powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_ZIP%' }"
if %ERRORLEVEL% neq 0 (
    echo   ERROR: Could not download Python. Check your internet connection.
    pause
    exit /b 1
)
powershell -Command "Expand-Archive -Path '%PYTHON_ZIP%' -DestinationPath '%PYTHON_DIR%' -Force"
del "%PYTHON_ZIP%" 2>nul
set "PYTHON=%PYTHON_DIR%\python.exe"

:python_found
echo   OK: Using %PYTHON%

:: ---- Locate the HTML generator script ----
:: Its location can differ depending on how the project was downloaded,
:: so search for archive\generate.py instead of assuming one fixed path.
set "GEN_SCRIPT="
for /r "%~dp0." %%f in (generate.py) do (
    if not defined GEN_SCRIPT (
        echo %%~dpf | find /i "archive" >nul && set "GEN_SCRIPT=%%~ff"
    )
)

:: Not found locally? Download it straight from the project's GitHub repo.
:: generate.py only uses Python's standard library, so nothing else is needed.
if not defined GEN_SCRIPT (
    echo   Generator not found locally. Downloading it from GitHub...
    set "GEN_SCRIPT=%TEMP%\teams_generate.py"
    powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/audibleh/TeamsExportPortableWithImages/main/src/msteams_export/archive/generate.py' -OutFile '!GEN_SCRIPT!' }"
)

if not exist "%GEN_SCRIPT%" (
    echo   ERROR: Could not find or download the HTML generator.
    echo   Check your internet connection and try again.
    pause
    exit /b 1
)

:: ---- Generate the HTML archive ----
echo.
echo [2/2] Generating HTML archive...
echo   Using generator: %GEN_SCRIPT%

"%PYTHON%" "%GEN_SCRIPT%" "exports" "teams-archive"
if %ERRORLEVEL% neq 0 (
    echo   ERROR: Could not generate the HTML archive.
    pause
    exit /b 1
)

echo.
echo  ============================================================
echo.
echo   DONE! The teams-archive\ folder has been created.
echo   Opening it in your browser...
echo.
echo  ============================================================
echo.

start "" "%~dp0teams-archive\index.html"
pause
