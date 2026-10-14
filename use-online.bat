@echo off
setlocal
title Zee AI for PowerPoint - Switch to Hosted Online Mode
cd /d "%~dp0"

echo ============================================================
echo   Zee AI for PowerPoint - Pointing PowerPoint to Hosted Build
echo   (Re-run this only when the manifest / ribbon buttons change)
echo ============================================================
echo.

where node >nul 2>&1
if %errorlevel% equ 0 (
  echo Refreshing manifest.prod.xml...
  call node scripts/make-manifest.mjs
)

call "%~dp0install-zee-hosted.bat"
