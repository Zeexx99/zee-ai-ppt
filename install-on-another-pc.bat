@echo off
setlocal enabledelayedexpansion
title Zee AI for PowerPoint - Standalone Installer

:MENU
cls
echo =========================================================
echo       Zee AI for PowerPoint - Standalone Installer
echo                 (Zee AI Family Add-in)
echo =========================================================
echo.
echo   [1] Install Zee AI for PowerPoint (Online Hosted)
echo   [2] Uninstall Zee AI for PowerPoint
echo   [3] Exit
echo.
echo =========================================================
set /p CHOICE="Select an option (1, 2, or 3): "

if "%CHOICE%"=="1" goto DO_INSTALL
if "%CHOICE%"=="2" goto DO_UNINSTALL
if "%CHOICE%"=="3" goto DO_EXIT

echo Invalid selection. Please enter 1, 2, or 3.
timeout /t 2 >nul
goto MENU

:DO_INSTALL
cls
echo =========================================================
echo   Installing Zee AI for PowerPoint...
echo =========================================================
echo.
echo [1/3] Adding registry entry for PowerPoint Developer Add-ins...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI.PPT.Hosted" /t REG_SZ /d "https://zeexx99.github.io/zee-ai-ppt/manifest.xml" /f >nul 2>&1
if errorlevel 1 (
  echo [Error] Could not write to registry. Please run as user.
  pause
  goto MENU
)

echo [2/3] Cleaning Office WEF cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" (
  rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
)

echo [3/3] Verifying installation...
echo.
echo ---------------------------------------------------------
echo   SUCCESS! Zee AI for PowerPoint is installed.
echo ---------------------------------------------------------
echo.
echo Instructions:
echo  1. If PowerPoint is currently open, close all windows.
echo  2. Open PowerPoint.
echo  3. Go to the Home tab and click the "Zee AI" button.
echo  4. Paste your Gemini API key in Settings (or use Team Proxy).
echo.
echo Press any key to return to menu...
pause >nul
goto MENU

:DO_UNINSTALL
cls
echo =========================================================
echo   Uninstalling Zee AI for PowerPoint...
echo =========================================================
echo.
echo Removing registry key...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "ZeeAI.PPT.Hosted" /f >nul 2>&1

echo Clearing Office cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" (
  rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
)

echo.
echo Zee AI for PowerPoint has been uninstalled.
echo Please restart PowerPoint.
echo.
echo Press any key to return to menu...
pause >nul
goto MENU

:DO_EXIT
cls
echo Thank you for using Zee AI for PowerPoint!
timeout /t 1 >nul
exit /b 0
