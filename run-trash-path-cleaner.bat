@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "PS1=%~dp0trash-path-cleaner.ps1"
if not exist "%PS1%" (
    echo [ERROR] trash-path-cleaner.ps1 not found next to this .bat.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
set "CODE=%ERRORLEVEL%"

if not "%CODE%"=="0" (
    echo.
    echo [ERROR] Script exited with code %CODE%.
    pause
)

endlocal & exit /b %CODE%
