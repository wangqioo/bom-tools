@echo off
chcp 65001 >nul
setlocal EnableExtensions
title BOM Tools Offline Install and Run

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
set "APP_DIR=%ROOT%\web_app2"
set "VENV_DIR=%ROOT%\venv"
set "WHEELS_DIR=%ROOT%\wheels"

echo ========================================
echo   BOM Tools - Offline Install and Run
echo ========================================
echo.
echo Bundle: %ROOT%
echo App:    %APP_DIR%
echo.

if not exist "%APP_DIR%\app.py" (
  echo ERROR: Cannot find %APP_DIR%\app.py
  pause
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

echo ERROR: Python not found. Please install Python 3.10+ from python.org and add it to PATH.
pause
exit /b 1

:FOUND_PYTHON
echo Using: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

if exist "%VENV_DIR%\" (
  echo [1/3] Virtual environment found.
) else (
  echo [1/3] Creating virtual environment...
  %PYTHON_CMD% -m venv "%VENV_DIR%"
  if errorlevel 1 (
    echo ERROR: Failed to create virtual environment.
    pause
    exit /b %errorlevel%
  )
)

echo [2/3] Installing dependencies...
if exist "%WHEELS_DIR%\" (
  echo Offline wheels: %WHEELS_DIR%
  "%VENV_DIR%\Scripts\python.exe" -m pip install --no-index --find-links "%WHEELS_DIR%" -r "%ROOT%\requirements.txt"
) else (
  echo WARN: wheels folder not found. Trying normal pip install.
  "%VENV_DIR%\Scripts\python.exe" -m pip install -r "%ROOT%\requirements.txt"
)
if errorlevel 1 (
  echo ERROR: Dependency installation failed.
  pause
  exit /b %errorlevel%
)

if "%PLAYWRIGHT_BROWSERS_PATH%"=="" (
  if exist "%ROOT%\ms-playwright\" set "PLAYWRIGHT_BROWSERS_PATH=%ROOT%\ms-playwright"
)

"%VENV_DIR%\Scripts\python.exe" -c "from pathlib import Path; import os,sys; root=Path(os.environ.get('PLAYWRIGHT_BROWSERS_PATH') or Path.home()/'AppData/Local/ms-playwright'); sys.exit(0 if any(root.glob('chromium*')) else 1)" >nul 2>&1
if errorlevel 1 (
  echo.
  echo WARN: Playwright Chromium runtime was not found.
  echo PLM automation needs Chromium. Normal BOM tools can still run.
  echo.
)

if "%PORT%"=="" set PORT=5000

echo [3/3] Starting web server...
echo.
echo ========================================
echo   Local: http://localhost:%PORT%
echo   LAN:   http://SERVER_IP:%PORT%
echo   Close this window to stop the server
echo ========================================
echo.

cd /d "%APP_DIR%"
"%VENV_DIR%\Scripts\python.exe" -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('waitress') else 1)" >nul 2>&1
if errorlevel 1 (
  echo waitress not installed; using Flask development server...
  "%VENV_DIR%\Scripts\python.exe" app.py
) else (
  echo Using waitress WSGI server...
  "%VENV_DIR%\Scripts\python.exe" -m waitress --host=0.0.0.0 --port=%PORT% app:app
)

echo.
echo Server has stopped.
pause
endlocal
