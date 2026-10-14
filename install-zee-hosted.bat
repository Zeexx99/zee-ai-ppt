@echo off
setlocal
title Zee AI for PowerPoint - Hosted Registry Installer
cd /d "%~dp0"

echo ============================================================
echo   Zee AI for PowerPoint - Registering Hosted Add-in
echo   Hosted URL: https://zeexx99.github.io/zee-ai-ppt/
echo ============================================================
echo.

set "CATALOG_DIR=%LOCALAPPDATA%\ZeeAI\PowerPoint\Catalog"
set "UNC_CATALOG_DIR=\\localhost\%LOCALAPPDATA::=$%\ZeeAI\PowerPoint\Catalog"
set "CATALOG_ID={e74c89a2-5b14-4d93-8f61-9c2e3a4b5d78}"
set "ADDIN_ID=e74c89a2-5b14-4d93-8f61-9c2e3a4b5d78"

if not exist "%CATALOG_DIR%" mkdir "%CATALOG_DIR%"

if exist "%~dp0manifest.prod.xml" (
  echo [1/4] Copying manifest.prod.xml to %CATALOG_DIR%\manifest.xml ...
  copy /Y "%~dp0manifest.prod.xml" "%CATALOG_DIR%\manifest.xml" >nul
) else (
  echo [1/4] Downloading hosted manifest from GitHub Pages...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri 'https://zeexx99.github.io/zee-ai-ppt/manifest.xml' -OutFile '%CATALOG_DIR%\manifest.xml' -UseBasicParsing"
)

echo [2/4] Removing old or conflicting localhost sideload entries...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI-PPT" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI.PPT.Hosted" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%~dp0manifest.xml" /f >nul 2>&1

echo [3/4] Registering Hosted Add-in in Windows Registry...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /t REG_SZ /d "%CATALOG_DIR%\manifest.xml" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\%ADDIN_ID%" /v "UseDirectDebugger" /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\%ADDIN_ID%" /v "UseWebDebugger" /t REG_DWORD /d 0 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\%ADDIN_ID%" /v "UseLiveReload" /t REG_DWORD /d 0 /f >nul

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_ID%" /v "Id" /t REG_SZ /d "%CATALOG_ID%" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_ID%" /v "Url" /t REG_SZ /d "%UNC_CATALOG_DIR%" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_ID%" /v "Flags" /t REG_DWORD /d 1 /f >nul

echo [4/4] Clearing stale Office WEF cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
  rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
)

echo.
echo [SUCCESS] Zee AI for PowerPoint (Hosted) is now registered!
echo   Catalog Path: %UNC_CATALOG_DIR%
echo.
echo Next steps:
echo   1. Close PowerPoint completely if it is open.
echo   2. Open PowerPoint -^> Home -^> Add-ins -^> More Add-ins.
echo   3. Under "My Add-ins" or "Shared Folder", select "Zee AI for PowerPoint" and click Add.
echo   4. "Zee AI" will now appear on your Home tab ribbon!
echo.
pause
