@echo off
setlocal
title Zee AI for PowerPoint - Local Sideload Helper (Dev Only)
cd /d "%~dp0"

set "ADDIN_ID=e74c89a2-5b14-4d93-8f61-9c2e3a4b5d78"

echo ============================================================
echo   Zee AI for PowerPoint - Sideloading Local Dev Manifest
echo   WARNING: Localhost sideloads are for development only and
echo   disappear when Office restarts. Use use-online.bat for
echo   daily use!
echo ============================================================
echo.

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /t REG_SZ /d "%~dp0manifest.xml" /f >nul
echo Registered %~dp0manifest.xml in WEF\Developer.
echo Make sure start-zee.bat is running on port 3300 before opening PowerPoint.
pause
