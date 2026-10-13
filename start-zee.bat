@echo off
setlocal
title Zee AI for PowerPoint - Dev Server (Port 3300)

echo ===================================================
echo   Zee AI for PowerPoint - Starting Dev Server
echo   Port: 3300 ^| Storage Prefix: zee-ppt:
echo ===================================================
echo.

cd /d "%~dp0"

if not exist "node_modules\" (
  echo [Info] node_modules not found. Installing dependencies...
  call npm install
  if errorlevel 1 (
    echo [Error] Failed to install dependencies.
    pause
    exit /b 1
  )
)

echo [Info] Starting Vite dev server on https://localhost:3300 ...
call npm run dev

pause
