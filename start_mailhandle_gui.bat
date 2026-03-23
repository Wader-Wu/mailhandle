@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "LAUNCH_SCRIPT=%ROOT_DIR%scripts\launch_mailhandle.ps1"

if not exist "%LAUNCH_SCRIPT%" (
    echo Mailhandle GUI launcher was not found.
    echo Expected: %LAUNCH_SCRIPT%
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%LAUNCH_SCRIPT%" -Mode gui
set "EXITCODE=%ERRORLEVEL%"

if not "%EXITCODE%"=="0" (
    echo.
    echo Mailhandle GUI launch failed.
    pause
)

exit /b %EXITCODE%
