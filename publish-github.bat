@echo off
setlocal
title Zee AI for PowerPoint - Publish to GitHub

cd /d "%~dp0"

echo ===================================================
echo   Zee AI for PowerPoint - Publish to GitHub
echo   Repository: Zeexx99/zee-ai-ppt
echo ===================================================
echo.

if not exist ".git" (
  echo [Info] Initializing local git repository...
  git init
  git branch -M main
)

echo [Info] Setting remote origin to https://github.com/Zeexx99/zee-ai-ppt.git ...
git remote remove origin >nul 2>&1
git remote add origin https://github.com/Zeexx99/zee-ai-ppt.git

echo [Info] Running build to verify project integrity...
call npm run build
if errorlevel 1 (
  echo [Error] Build failed. Please fix any errors before publishing.
  pause
  exit /b 1
)

echo [Info] Staging files and creating initial commit...
git add .
git commit -m "Initial commit for Zee AI for PowerPoint"

echo [Info] Pushing to main branch...
git push -u origin main
if errorlevel 1 (
  echo.
  echo [Warning] Push encountered an issue.
  echo Make sure the repo 'zee-ai-ppt' exists on https://github.com/Zeexx99
  echo and that your GitHub credentials or SSH keys are configured.
  pause
  exit /b 1
)

echo.
echo ===================================================
echo   Successfully published to GitHub!
echo.
echo   IMPORTANT ONE-TIME STEP:
echo   Go to your GitHub repo settings:
echo   https://github.com/Zeexx99/zee-ai-ppt/settings/pages
echo   Under 'Build and deployment':
echo   Set Source: "GitHub Actions"
echo ===================================================
pause
