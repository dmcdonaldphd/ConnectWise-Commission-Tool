@echo off
setlocal
cd /d "%~dp0"
"%~dp0runtime\node.exe" "%~dp0app\commission-menu.js"
set "COMMISSION_EXIT_CODE=%errorlevel%"
pause
exit /b %COMMISSION_EXIT_CODE%
