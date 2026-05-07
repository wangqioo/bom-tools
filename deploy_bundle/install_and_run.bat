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
pip install --no-index --find-links wheels -r requirements.txt
if %errorlevel% neq 0 (
    echo ****************************************************
    echo * ERROR: Offline install failed.
    echo * Missing wheel files. Check that all dependencies
    echo * are present in the wheels/ folder.
    echo ****************************************************
    pause
    exit /b 1
)

:: Start server
echo [3/3] Starting web server...
echo.
echo ========================================
echo   Server started! Open in browser:
echo   http://localhost:5000
echo.
echo   Close this window to stop the server
echo ========================================
echo.
cd web_app
echo.
echo Starting Flask server...
echo If this fails, copy the error text above.
%PYTHON_CMD% app.py
echo.
echo Server has stopped.
pause
