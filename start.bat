@echo off
rem One-click launcher: starts the local server (needed for PDF export) and opens the app.
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
    echo Node.js is required for PDF export. Install it from https://nodejs.org and run this again.
    pause
    exit /b 1
)
if "%PORT%"=="" set PORT=8080
start "" "http://127.0.0.1:%PORT%"
node server.js
pause
