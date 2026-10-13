@echo off
setlocal
title Zee AI for PowerPoint - Register Hosted Add-in

echo ===================================================
echo   Zee AI for PowerPoint - Register Hosted Add-in
echo   Target: https://zeexx99.github.io/zee-ai-ppt/manifest.xml
echo ===================================================
echo.

echo [1/3] Removing old localhost registrations...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI.PPT.Dev" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI.PPT" /f >nul 2>&1

echo [2/3] Registering hosted add-in in Windows registry...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI.PPT.Hosted" /t REG_SZ /d "https://zeexx99.github.io/zee-ai-ppt/manifest.xml" /f
if errorlevel 1 (
  echo [Error] Failed to write registry key.
  pause
  exit /b 1
)

echo [3/3] Clearing Office WEF cache to ensure fresh load...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" (
  rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
  echo [Info] Office WEF cache cleared.
)

echo.
echo ===================================================
echo   SUCCESS: Zee AI for PowerPoint is now registered!
echo.
echo   Next steps:
echo   1. Close all PowerPoint windows completely.
echo   2. Reopen PowerPoint.
echo   3. Look on the Home tab for the "Zee AI" ribbon button.
echo ===================================================
pause
