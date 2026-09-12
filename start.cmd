@echo off
REM Double-click this to run the AK MyToDo Hub.
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo.
  echo   Node.js is not installed. Get the current release from https://nodejs.org
  echo   then double-click this file again.
  echo.
  pause
  exit /b 1
)
node setup.mjs
pause
