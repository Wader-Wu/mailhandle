@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "SUMMARY=%SCRIPT_DIR%..\tmp\mailhandle-last-start.txt"

if exist "%SUMMARY%" del /f /q "%SUMMARY%" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%start_mailhandle.ps1" >nul
set "EXITCODE=%ERRORLEVEL%"

if exist "%SUMMARY%" (
    type "%SUMMARY%"
) else (
    echo Mailhandle launcher completed, but no startup summary was written.
    echo Summary file: %SUMMARY%
)

exit /b %EXITCODE%
