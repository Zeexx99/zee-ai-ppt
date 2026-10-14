@echo off
setlocal
title Zee AI for PowerPoint - Hosted Registry Installer
cd /d "%~dp0"

echo ============================================================
echo   Zee AI for PowerPoint - Registering Hosted Add-in
echo   Manifest: %~dp0manifest.prod.xml
echo   Hosted URL: https://zeexx99.github.io/zee-ai-ppt/
echo ============================================================
echo.

set "ADDIN_ID=e74c89a2-5b14-4d93-8f61-9c2e3a4b5d78"
set "CATALOG_ID={e74c89a2-5b14-4d93-8f61-9c2e3a4b5d78}"

echo [1/3] Removing any old or duplicate registry entries...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI-PPT" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI.PPT.Hosted" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%~dp0manifest.xml" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_ID%" /f >nul 2>&1

echo [2/3] Registering ZeeAI-PPT -> %~dp0manifest.prod.xml in Windows Registry...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI-PPT" /t REG_SZ /d "%~dp0manifest.prod.xml" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\%ADDIN_ID%" /v "UseDirectDebugger" /t REG_DWORD /d 0 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\%ADDIN_ID%" /v "UseWebDebugger" /t REG_DWORD /d 0 /f >nul

echo [3/3] Clearing stale Office WEF cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
  rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
)

echo.
echo ============================================================
echo   [SUCCESS] Zee AI for PowerPoint (Hosted) is registered!
echo ============================================================
echo   How to activate in PowerPoint (one-time step):
echo   1. Close PowerPoint completely if it is open, then reopen it.
echo   2. Click "Add-ins" on the Home tab (or Insert -^> Add-ins).
echo   3. Click "More Add-ins" at the bottom of the menu.
echo   4. Click "MY ADD-INS" at the top, scroll to "Developer Add-ins",
echo      select "Zee AI for PowerPoint", and click "Add".
echo   5. The "Zee AI" button will now appear on your Home tab!
echo ============================================================
echo.
pause
