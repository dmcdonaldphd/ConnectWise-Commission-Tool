@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0PHDTech Commission Tool.exe" (
  start "" "%~dp0PHDTech Commission Tool.exe"
  exit /b 0
)
if exist "%~dp0browsers\" set "PLAYWRIGHT_BROWSERS_PATH=%~dp0browsers"
if exist "%~dp0runtime\node.exe" (
  "%~dp0runtime\node.exe" "%~dp0app\gui.js"
) else (
  node "%~dp0app\gui.js"
)
set "COMMISSION_EXIT_CODE=%errorlevel%"
if not "%COMMISSION_EXIT_CODE%"=="0" pause
exit /b %COMMISSION_EXIT_CODE%
