@echo off
setlocal
title Zee AI for PowerPoint - Update and Deploy
cd /d "%~dp0"

set "LOGFILE=%~dp0update-log.txt"
echo ============================================================ > "%LOGFILE%"
echo   Zee AI for PowerPoint - Update Log (%DATE% %TIME%) >> "%LOGFILE%"
echo ============================================================ >> "%LOGFILE%"

echo ============================================================
echo   Zee AI for PowerPoint - Update ^& Push to GitHub Pages
echo   Logging to: %LOGFILE%
echo ============================================================
echo.

echo [1/5] Installing/verifying dependencies...
echo [1/5] npm install >> "%LOGFILE%"
call npm install >> "%LOGFILE%" 2>&1
if errorlevel 1 (
  echo [ERROR] npm install failed. Check update-log.txt for details.
  type "%LOGFILE%"
  pause
  exit /b 1
)

echo [2/5] Running unit tests...
echo [2/5] npm test >> "%LOGFILE%"
call npm test >> "%LOGFILE%" 2>&1
if errorlevel 1 (
  echo [ERROR] Unit tests failed! Check update-log.txt for details.
  type "%LOGFILE%"
  pause
  exit /b 1
)

echo [3/5] Running TypeScript check and production build...
echo [3/5] npm run build >> "%LOGFILE%"
call npm run build >> "%LOGFILE%" 2>&1
if errorlevel 1 (
  echo [ERROR] Production build failed! Check update-log.txt for details.
  type "%LOGFILE%"
  pause
  exit /b 1
)

echo [4/5] Generating hosted manifest (manifest.prod.xml)...
echo [4/5] node scripts/make-manifest.mjs >> "%LOGFILE%"
call node scripts/make-manifest.mjs >> "%LOGFILE%" 2>&1
if errorlevel 1 (
  echo [ERROR] Manifest generation failed! Check update-log.txt.
  pause
  exit /b 1
)

echo [5/5] Committing and pushing changes to GitHub (Zeexx99/zee-ai-ppt)...
echo [5/5] git add/commit/push >> "%LOGFILE%"
git add -A >> "%LOGFILE%" 2>&1
git diff --cached --quiet
if errorlevel 1 (
  git commit -m "Update Zee AI for PowerPoint (%DATE% %TIME%)" >> "%LOGFILE%" 2>&1
) else (
  echo No new file changes to commit. >> "%LOGFILE%"
)

git push origin main >> "%LOGFILE%" 2>&1
if errorlevel 1 (
  echo [WARNING] git push failed. Make sure the remote repo exists (run publish-github.bat first).
  echo Check update-log.txt for details.
  pause
  exit /b 1
)

echo.
echo [SUCCESS] Update pushed to GitHub!
echo 1. Wait 1-2 minutes for the green GitHub Actions check at:
echo    https://github.com/Zeexx99/zee-ai-ppt/actions
echo 2. Fully close PowerPoint and reopen -> Home -> Zee AI.
echo.
pause
