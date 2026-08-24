@echo off
cd /d "%~dp0"
title ICEA LION Testcase Review
echo.
echo Starting ICEA LION Testcase Review...
echo Open: http://localhost:3100
echo Do not use Live Preview. Keep this window open. Press Ctrl+C to stop.
echo.
node scripts\sync-app-version.js
node scripts\free-port.js
if errorlevel 1 (
  echo Failed to free port 3100.
  pause
  exit /b 1
)
node server.js
echo.
echo Server stopped.
pause
