@echo off
setlocal
title Zee AI for PowerPoint - Local Dev Server (Port 3300)
cd /d "%~dp0"

echo ============================================================
echo   Zee AI for PowerPoint - Starting Local Dev Server
echo   URL: https://localhost:3300  (or http://localhost:3300)
echo ============================================================
echo.

if not exist "node_modules" (
  echo [1/2] Installing npm dependencies...
  call npm install
  if errorlevel 1 (
    echo [ERROR] npm install failed.
    pause
    exit /b 1
  )
)

echo [2/2] Starting Vite server on port 3300...
echo NOTE: Localhost sideloads are for development only and disappear
echo when Office restarts. For daily use, run use-online.bat instead.
echo.
call npm run dev
pause
