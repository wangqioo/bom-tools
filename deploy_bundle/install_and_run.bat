@echo off
chcp 65001 >nul
title BOM Tools Web Server

echo ========================================
echo   BOM Tools - Offline Deployment
echo ========================================
echo.

:: Try to find Python (python or py launcher)
set PYTHON_CMD=python
python --version >nul 2>&1
if %errorlevel% equ 0 goto :FOUND_PYTHON

py --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=py
    goto :FOUND_PYTHON
)

echo ****************************************************
echo * ERROR: Python not found!
echo * Please install Python 3.10+ from:
echo * https://www.python.org/downloads/
echo *
echo * IMPORTANT: During installation, check
echo * "Add Python to PATH" at the bottom of the screen.
echo ****************************************************
echo.
pause
exit /b 1

:FOUND_PYTHON
echo Using: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

:: Create virtual environment (skip if exists)
if exist venv\ (
    echo [INFO] Virtual environment found, skipping creation
) else (
    echo [1/3] Creating virtual environment...
    %PYTHON_CMD% -m venv venv
    if %errorlevel% neq 0 (
        echo ****************************************************
        echo * ERROR: Failed to create virtual environment.
        echo * This can happen if Python is installed from
        echo * Microsoft Store. Please install Python from:
        echo * https://www.python.org/downloads/
        echo ****************************************************
        pause
        exit /b 1
    )
)

:: Install dependencies
echo [2/3] Installing packages (offline mode)...
call venv\Scripts\activate.bat
venv\Scripts\python.exe -m pip install --no-index --find-links wheels -r requirements.txt
if %errorlevel% neq 0 (
    echo ****************************************************
    echo * ERROR: Offline install failed.
    echo * Missing wheel files. Check that all dependencies
    echo * are present in the wheels/ folder.
    echo ****************************************************
    pause
    exit /b 1
)

:: Check Playwright browser runtime for PLM web automation
venv\Scripts\python.exe -c "from pathlib import Path; import os,sys; root=Path(os.environ.get('PLAYWRIGHT_BROWSERS_PATH') or Path.home()/'AppData/Local/ms-playwright'); sys.exit(0 if any(root.glob('chromium*')) else 1)" >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [WARN] Playwright Chromium browser runtime was not found.
    echo        PLM web automation requires Chromium.
    echo        Online install command:
    echo          venv\Scripts\python.exe -m playwright install chromium
    echo        Offline deployment must include the ms-playwright Chromium cache
    echo        or set PLAYWRIGHT_BROWSERS_PATH to a folder that contains it.
    echo.
)

:: Start server
echo [3/3] Starting web server...
if "%PORT%"=="" set PORT=5000
echo.
echo ========================================
echo   Server started! Open in browser:
echo   http://localhost:%PORT%
echo   LAN users open:
echo   http://SERVER_IP:%PORT%
echo.
echo   Close this window to stop the server
echo ========================================
echo.
cd web_app2
echo.
echo Starting web server...
echo If this fails, copy the error text above.
..\venv\Scripts\python.exe -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('waitress') else 1)" >nul 2>&1
if %errorlevel% equ 0 (
    echo Using waitress WSGI server...
    ..\venv\Scripts\python.exe -m waitress --host=0.0.0.0 --port=%PORT% app:app
) else (
    echo waitress not installed; using Flask development server...
    ..\venv\Scripts\python.exe app.py
)
echo.
echo Server has stopped.
pause
