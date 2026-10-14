@echo off
setlocal EnableDelayedExpansion
title Zee AI for PowerPoint - Standalone Installer

set "CATALOG_DIR=%LOCALAPPDATA%\ZeeAI\PowerPoint\Catalog"
set "CATALOG_ID={e74c89a2-5b14-4d93-8f61-9c2e3a4b5d78}"
set "ADDIN_ID=e74c89a2-5b14-4d93-8f61-9c2e3a4b5d78"
set "MANIFEST_URL=https://zeexx99.github.io/zee-ai-ppt/manifest.xml"

:MENU
cls
echo ============================================================
echo        Zee AI for PowerPoint - Standalone Setup
echo        (No Node.js or Git required on this PC)
echo ============================================================
echo.
echo   [1] Install / Update Zee AI for PowerPoint
echo   [2] Uninstall Zee AI for PowerPoint
echo   [3] Exit
echo.
set /p CHOICE="Select an option (1-3): "

if "%CHOICE%"=="1" goto INSTALL
if "%CHOICE%"=="2" goto UNINSTALL
if "%CHOICE%"=="3" goto END
echo Invalid choice. Please enter 1, 2, or 3.
timeout /t 2 >nul
goto MENU

:INSTALL
echo.
echo [1/4] Creating local catalog folder: %CATALOG_DIR%
if not exist "%CATALOG_DIR%" mkdir "%CATALOG_DIR%"

echo [2/4] Fetching hosted manifest from %MANIFEST_URL% ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%MANIFEST_URL%' -OutFile '%CATALOG_DIR%\manifest.xml' -UseBasicParsing -ErrorAction Stop; Write-Host 'Downloaded latest manifest from GitHub Pages.' } catch { exit 1 }"
if errorlevel 1 (
  if exist "%~dp0manifest.prod.xml" (
    echo Using bundled manifest.prod.xml...
    copy /Y "%~dp0manifest.prod.xml" "%CATALOG_DIR%\manifest.xml" >nul
  ) else (
    echo Writing embedded production manifest...
    call :WRITE_EMBEDDED_MANIFEST "%CATALOG_DIR%\manifest.xml"
  )
)

echo [3/4] Registering Zee AI for PowerPoint in Windows Registry...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /t REG_SZ /d "%CATALOG_DIR%\manifest.xml" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_ID%" /v "Id" /t REG_SZ /d "%CATALOG_ID%" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_ID%" /v "Url" /t REG_SZ /d "%CATALOG_DIR%" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_ID%" /v "Flags" /t REG_DWORD /d 1 /f >nul

echo [4/4] Clearing old Office WEF cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
  rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
)

echo.
echo ============================================================
echo   [SUCCESS] Zee AI for PowerPoint is installed!
echo ============================================================
echo   1. Close PowerPoint completely if it is running.
echo   2. Open PowerPoint -^> Home tab -^> click "Zee AI"
echo      (Or go to Home -^> Add-ins -^> More Add-ins -^> Shared Folder -^> Zee AI).
echo   3. Open Settings (gear icon) and enter your Gemini API Key
echo      or Team Proxy Token.
echo ============================================================
echo.
pause
goto MENU

:UNINSTALL
echo.
echo [1/3] Removing registry keys...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_ID%" /f >nul 2>&1

echo [2/3] Removing local catalog folder...
if exist "%LOCALAPPDATA%\ZeeAI\PowerPoint" (
  rd /s /q "%LOCALAPPDATA%\ZeeAI\PowerPoint" >nul 2>&1
)

echo [3/3] Clearing Office WEF cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
  rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
)

echo.
echo [SUCCESS] Zee AI for PowerPoint has been uninstalled.
echo Restart PowerPoint to complete removal.
echo.
pause
goto MENU

:WRITE_EMBEDDED_MANIFEST
(
echo ^<?xml version="1.0" encoding="UTF-8" standalone="yes"?^>
echo ^<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp"^>
echo   ^<Id^>e74c89a2-5b14-4d93-8f61-9c2e3a4b5d78^</Id^>
echo   ^<Version^>1.2.0.0^</Version^>
echo   ^<ProviderName^>Zee AI^</ProviderName^>
echo   ^<DefaultLocale^>en-US^</DefaultLocale^>
echo   ^<DisplayName DefaultValue="Zee AI for PowerPoint"/^>
echo   ^<Description DefaultValue="Zee AI Assistant and Slide Agent for Microsoft PowerPoint"/^>
echo   ^<IconUrl DefaultValue="https://zeexx99.github.io/zee-ai-ppt/assets/icon-32.png"/^>
echo   ^<HighResolutionIconUrl DefaultValue="https://zeexx99.github.io/zee-ai-ppt/assets/icon-64.png"/^>
echo   ^<SupportUrl DefaultValue="https://zeexx99.github.io/zee-ai-ppt/"/^>
echo   ^<AppDomains^>
echo     ^<AppDomain^>https://zeexx99.github.io^</AppDomain^>
echo     ^<AppDomain^>https://accounts.google.com^</AppDomain^>
echo     ^<AppDomain^>https://generativelanguage.googleapis.com^</AppDomain^>
echo   ^</AppDomains^>
echo   ^<Hosts^>^<Host Name="Presentation"/^>^</Hosts^>
echo   ^<DefaultSettings^>^<SourceLocation DefaultValue="https://zeexx99.github.io/zee-ai-ppt/index.html"/^>^</DefaultSettings^>
echo   ^<Permissions^>ReadWriteDocument^</Permissions^>
echo   ^<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0"^>
echo     ^<Hosts^>^<Host xsi:type="Presentation"^>^<DesktopFormFactor^>
echo       ^<ExtensionPoint xsi:type="PrimaryCommandSurface"^>
echo         ^<OfficeTab id="TabHome"^>
echo           ^<Group id="ZeeAI.PPT.Group"^>
echo             ^<Label resid="ZeeAI.PPT.GroupLabel"/^>
echo             ^<Icon^>
echo               ^<bt:Image size="16" resid="ZeeAI.PPT.Icon16"/^>
echo               ^<bt:Image size="32" resid="ZeeAI.PPT.Icon32"/^>
echo               ^<bt:Image size="80" resid="ZeeAI.PPT.Icon80"/^>
echo             ^</Icon^>
echo             ^<Control xsi:type="Button" id="ZeeAI.PPT.OpenButton"^>
echo               ^<Label resid="ZeeAI.PPT.TaskpaneButton.Label"/^>
echo               ^<Supertip^>^<Title resid="ZeeAI.PPT.TaskpaneButton.Label"/^>^<Description resid="ZeeAI.PPT.TaskpaneButton.Tooltip"/^>^</Supertip^>
echo               ^<Icon^>
echo                 ^<bt:Image size="16" resid="ZeeAI.PPT.Icon16"/^>
echo                 ^<bt:Image size="32" resid="ZeeAI.PPT.Icon32"/^>
echo                 ^<bt:Image size="80" resid="ZeeAI.PPT.Icon80"/^>
echo               ^</Icon^>
echo               ^<Action xsi:type="ShowTaskpane"^>
echo                 ^<TaskpaneId^>ZeeAI.PPT.Taskpane^</TaskpaneId^>
echo                 ^<SourceLocation resid="ZeeAI.PPT.Taskpane.Url"/^>
echo               ^</Action^>
echo             ^</Control^>
echo           ^</Group^>
echo         ^</OfficeTab^>
echo       ^</ExtensionPoint^>
echo     ^</DesktopFormFactor^>^</Host^>^</Hosts^>
echo     ^<Resources^>
echo       ^<bt:Images^>
echo         ^<bt:Image id="ZeeAI.PPT.Icon16" DefaultValue="https://zeexx99.github.io/zee-ai-ppt/assets/icon-16.png"/^>
echo         ^<bt:Image id="ZeeAI.PPT.Icon32" DefaultValue="https://zeexx99.github.io/zee-ai-ppt/assets/icon-32.png"/^>
echo         ^<bt:Image id="ZeeAI.PPT.Icon64" DefaultValue="https://zeexx99.github.io/zee-ai-ppt/assets/icon-64.png"/^>
echo         ^<bt:Image id="ZeeAI.PPT.Icon80" DefaultValue="https://zeexx99.github.io/zee-ai-ppt/assets/icon-80.png"/^>
echo       ^</bt:Images^>
echo       ^<bt:Urls^>
echo         ^<bt:Url id="ZeeAI.PPT.Taskpane.Url" DefaultValue="https://zeexx99.github.io/zee-ai-ppt/index.html"/^>
echo       ^</bt:Urls^>
echo       ^<bt:ShortStrings^>
echo         ^<bt:String id="ZeeAI.PPT.GroupLabel" DefaultValue="Zee AI"/^>
echo         ^<bt:String id="ZeeAI.PPT.TaskpaneButton.Label" DefaultValue="Zee AI"/^>
echo       ^</bt:ShortStrings^>
echo       ^<bt:LongStrings^>
echo         ^<bt:String id="ZeeAI.PPT.TaskpaneButton.Tooltip" DefaultValue="Open Zee AI for PowerPoint"/^>
echo       ^</bt:LongStrings^>
echo     ^</Resources^>
echo   ^</VersionOverrides^>
echo ^</OfficeApp^>
) > "%~1"
exit /b 0

:END
exit /b 0
