@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0browsers\" set "PLAYWRIGHT_BROWSERS_PATH=%~dp0browsers"
"%~dp0runtime\node.exe" "%~dp0app\commission-menu.js"
set "COMMISSION_EXIT_CODE=%errorlevel%"
pause
exit /b %COMMISSION_EXIT_CODE%
