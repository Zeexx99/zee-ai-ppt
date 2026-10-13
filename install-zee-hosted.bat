@echo off
setlocal
title Zee AI for PowerPoint - Register Hosted Add-in

echo ===================================================
echo   Zee AI for PowerPoint - Register Hosted Add-in
echo   Target: %~dp0manifest.prod.xml
echo ===================================================
echo.

echo [1/2] Registering ZeeAI-PPT in Windows registry...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI-PPT" /t REG_SZ /d "%~dp0manifest.prod.xml" /f
if errorlevel 1 (
  echo [Error] Failed to write registry key.
  pause
  exit /b 1
)

echo [2/2] Clearing Office WEF cache to ensure fresh load...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" (
  rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
  echo [Info] Office WEF cache cleared.
)

echo.
echo ===================================================
echo   SUCCESS: Zee AI for PowerPoint is now registered!
echo.
echo   How to open in PowerPoint:
echo   1. Close all PowerPoint windows completely.
echo   2. Reopen PowerPoint.
echo   3. Go to Insert -^> Add-ins (or My Add-ins / Get Add-ins).
echo   4. Click "SHARED FOLDER" or "DEVELOPER ADD-INS".
echo   5. Select "Zee AI for PowerPoint" and click Add.
echo   6. It will now appear on your Home tab!
echo ===================================================
pause
