@echo off
setlocal
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%PREPARE_MIGRATION_A.ps1" %*
exit /b %ERRORLEVEL%
