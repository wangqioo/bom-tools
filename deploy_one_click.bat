@echo off
chcp 65001 >nul
setlocal EnableExtensions
title BOM Tools Deploy and Run

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"

if /i "%~1"=="/?" goto :help
if /i "%~1"=="-h" goto :help
if /i "%~1"=="--help" goto :help

if "%~1"=="" (
  set "INSTALL_DIR=%ROOT%"
) else (
  set "INSTALL_DIR=%~1"
)

if "%~2"=="" (
  set "PACKAGE="
  for /f "delims=" %%F in ('dir /b /a:-d /o:-d "%ROOT%\deploy_bundle\offline_final\bom-tools_offline_*.zip" 2^>nul') do (
    if not defined PACKAGE set "PACKAGE=%ROOT%\deploy_bundle\offline_final\%%F"
  )
  if not defined PACKAGE (
    for /f "delims=" %%F in ('dir /b /a:-d /o:-d "%ROOT%\deploy_bundle\bom-tools_offline_*.zip" 2^>nul') do (
      if not defined PACKAGE set "PACKAGE=%ROOT%\deploy_bundle\%%F"
    )
  )
  if not defined PACKAGE (
    for /f "delims=" %%F in ('dir /b /a:-d /o:-d "%ROOT%\bom-tools_offline_*.zip" 2^>nul') do (
      if not defined PACKAGE set "PACKAGE=%ROOT%\%%F"
    )
  )
) else (
  set "PACKAGE=%~2"
)

set "USERS_DB=%USERPROFILE%\Desktop\users.sqlite3"
if exist "%USERS_DB%" (
  set "USERS_DB_ARG=-UsersDbPath ""%USERS_DB%"""
) else (
  set "USERS_DB_ARG="
)

echo ========================================
echo   BOM Tools - Deploy and Run
echo ========================================
echo.
echo InstallDir: %INSTALL_DIR%
if defined PACKAGE echo Package: %PACKAGE%
if exist "%USERS_DB%" echo UsersDb: %USERS_DB%
echo.

if defined PACKAGE (
  echo Please stop the BOM Tools service before deploying.
  choice /c YN /m "Deploy this package before starting"
  if not errorlevel 2 (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%\scripts\install_offline_release.ps1" -PackagePath "%PACKAGE%" -InstallDir "%INSTALL_DIR%" %USERS_DB_ARG%
    if errorlevel 1 exit /b %errorlevel%
    echo.
    echo Deployment finished.
  )
) else (
  echo No offline package found. Starting the existing installation.
)

set "APP_ROOT=%INSTALL_DIR%"
if not exist "%APP_ROOT%\web_app2\app.py" (
  echo ERROR: Cannot find %APP_ROOT%\web_app2\app.py
  exit /b 1
)

set "PYTHON_CMD=python"
python --version >nul 2>&1
if %errorlevel% equ 0 goto :FOUND_PYTHON

py --version >nul 2>&1
if %errorlevel% equ 0 (
  set "PYTHON_CMD=py"
  goto :FOUND_PYTHON
)

echo ERROR: Python not found. Please install Python 3.10+ and add it to PATH.
pause
exit /b 1

:FOUND_PYTHON
echo.
echo Using: %PYTHON_CMD%
%PYTHON_CMD% --version

cd /d "%APP_ROOT%"
if exist venv\ (
  echo [1/3] Virtual environment found.
) else (
  echo [1/3] Creating virtual environment...
  %PYTHON_CMD% -m venv venv
  if errorlevel 1 exit /b %errorlevel%
)

call venv\Scripts\activate.bat

set "WHEELS_DIR="
if exist "%APP_ROOT%\wheels\" set "WHEELS_DIR=%APP_ROOT%\wheels"
if not defined WHEELS_DIR if exist "%APP_ROOT%\deploy_bundle\wheels\" set "WHEELS_DIR=%APP_ROOT%\deploy_bundle\wheels"
if not defined WHEELS_DIR if exist "%ROOT%\deploy_bundle\wheels\" set "WHEELS_DIR=%ROOT%\deploy_bundle\wheels"

echo [2/3] Installing dependencies...
if defined WHEELS_DIR (
  echo Offline wheels: %WHEELS_DIR%
  venv\Scripts\python.exe -m pip install --no-index --find-links "%WHEELS_DIR%" -r web_app2\requirements.txt
) else (
  echo Offline wheels not found; using pip default index.
  venv\Scripts\python.exe -m pip install -r web_app2\requirements.txt
)
if errorlevel 1 exit /b %errorlevel%

if "%PLAYWRIGHT_BROWSERS_PATH%"=="" (
  if exist "%APP_ROOT%\ms-playwright\" set "PLAYWRIGHT_BROWSERS_PATH=%APP_ROOT%\ms-playwright"
  if "%PLAYWRIGHT_BROWSERS_PATH%"=="" if exist "%APP_ROOT%\deploy_bundle\ms-playwright\" set "PLAYWRIGHT_BROWSERS_PATH=%APP_ROOT%\deploy_bundle\ms-playwright"
  if "%PLAYWRIGHT_BROWSERS_PATH%"=="" if exist "%ROOT%\deploy_bundle\ms-playwright\" set "PLAYWRIGHT_BROWSERS_PATH=%ROOT%\deploy_bundle\ms-playwright"
)

venv\Scripts\python.exe -c "from pathlib import Path; import os,sys; root=Path(os.environ.get('PLAYWRIGHT_BROWSERS_PATH') or Path.home()/'AppData/Local/ms-playwright'); sys.exit(0 if any(root.glob('chromium*')) else 1)" >nul 2>&1
if errorlevel 1 (
  echo.
  echo WARN: Playwright Chromium browser runtime was not found.
  echo PLM automation needs Chromium. Add ms-playwright or set PLAYWRIGHT_BROWSERS_PATH.
  echo.
)

echo [3/3] Starting web server...
if "%PORT%"=="" set PORT=5000
echo.
echo ========================================
echo   Open: http://localhost:%PORT%
echo   LAN:  http://SERVER_IP:%PORT%
echo   Close this window to stop the server
echo ========================================
echo.

cd /d "%APP_ROOT%\web_app2"
..\venv\Scripts\python.exe -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('waitress') else 1)" >nul 2>&1
if errorlevel 1 (
  echo waitress not installed; using Flask development server...
  ..\venv\Scripts\python.exe app.py
) else (
  echo Using waitress WSGI server...
  ..\venv\Scripts\python.exe -m waitress --host=0.0.0.0 --port=%PORT% app:app
)

echo.
echo Server has stopped.
pause
endlocal
exit /b 0

:help
echo Usage:
echo   deploy_one_click.bat
echo   deploy_one_click.bat C:\path\to\bom-tools
echo   deploy_one_click.bat C:\path\to\bom-tools C:\path\to\bom-tools_offline.zip
echo.
echo The script deploys the newest bom-tools_offline_*.zip when available,
echo preserves users.sqlite3/runtime data/offline assets, installs dependencies,
echo and starts the BOM Tools web service.
exit /b 0
