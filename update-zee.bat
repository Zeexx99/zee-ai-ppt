@echo off
setlocal enabledelayedexpansion
title Zee AI for PowerPoint - Update and Deploy

cd /d "%~dp0"
set LOGFILE=update-log.txt

echo =================================================== > "%LOGFILE%"
echo   Zee AI for PowerPoint - Update Process >> "%LOGFILE%"
echo   Started at: %date% %time% >> "%LOGFILE%"
echo =================================================== >> "%LOGFILE%"

echo [1/4] Installing / Updating dependencies...
echo [1/4] Installing / Updating dependencies... >> "%LOGFILE%"
call npm install >> "%LOGFILE%" 2>&1
if errorlevel 1 (
  echo [ERROR] npm install failed. See %LOGFILE% for details.
  echo [ERROR] npm install failed at %date% %time% >> "%LOGFILE%"
  pause
  exit /b 1
)

echo [2/4] Running unit tests...
echo [2/4] Running unit tests... >> "%LOGFILE%"
call npm run test >> "%LOGFILE%" 2>&1
if errorlevel 1 (
  echo [ERROR] Unit tests failed! Deployment aborted. See %LOGFILE%.
  echo [ERROR] Unit tests failed at %date% %time% >> "%LOGFILE%"
  pause
  exit /b 1
)

echo [3/4] Building production package and manifests...
echo [3/4] Building production package... >> "%LOGFILE%"
call npm run build >> "%LOGFILE%" 2>&1
if errorlevel 1 (
  echo [ERROR] Build failed! See %LOGFILE%.
  echo [ERROR] Build failed at %date% %time% >> "%LOGFILE%"
  pause
  exit /b 1
)

echo [4/4] Pushing changes to GitHub...
echo [4/4] Pushing changes to GitHub... >> "%LOGFILE%"
git status --porcelain >> "%LOGFILE%" 2>&1

git add . >> "%LOGFILE%" 2>&1
set COMMIT_MSG=Update Zee AI for PowerPoint - %date% %time%
git commit -m "%COMMIT_MSG%" >> "%LOGFILE%" 2>&1
git push origin main >> "%LOGFILE%" 2>&1

if errorlevel 1 (
  echo [WARNING] Git push had an issue or no changes to push. Check %LOGFILE%.
  echo [WARNING] Git push finished with status %errorlevel% at %date% %time% >> "%LOGFILE%"
) else (
  echo [SUCCESS] Zee AI for PowerPoint updated and pushed successfully!
  echo [SUCCESS] Finished at %date% %time% >> "%LOGFILE%"
)

echo.
echo ===================================================
echo   Update completed!
echo   Log written to: %LOGFILE%
echo   Wait for the GitHub Actions green check, then
echo   restart PowerPoint to load the newest version.
echo ===================================================
pause
