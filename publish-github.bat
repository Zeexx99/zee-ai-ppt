@echo off
setlocal
title Zee AI for PowerPoint - First-Time GitHub Publish
cd /d "%~dp0"

echo ============================================================
echo   Zee AI for PowerPoint - First-Time GitHub Setup ^& Publish
echo   Account: Zeexx99   Repo: zee-ai-ppt (public)
echo ============================================================
echo.

if not exist ".git" (
  echo Initializing Git repository...
  git init
)

git branch -M main

echo Generating icons and hosted manifest...
call node scripts/generate-icons.mjs
call node scripts/make-manifest.mjs

git add -A
git diff --cached --quiet
if errorlevel 1 (
  git commit -m "Initial release: Zee AI for PowerPoint"
)

where gh >nul 2>&1
if %errorlevel% equ 0 (
  echo Creating or linking public GitHub repo Zeexx99/zee-ai-ppt via GitHub CLI...
  gh repo view Zeexx99/zee-ai-ppt >nul 2>&1
  if errorlevel 1 (
    gh repo create Zeexx99/zee-ai-ppt --public --source=. --remote=origin --push
  ) else (
    git remote remove origin >nul 2>&1
    git remote add origin https://github.com/Zeexx99/zee-ai-ppt.git
    git push -u origin main
  )
) else (
  echo Configuring remote origin https://github.com/Zeexx99/zee-ai-ppt.git ...
  git remote remove origin >nul 2>&1
  git remote add origin https://github.com/Zeexx99/zee-ai-ppt.git
  git push -u origin main
)

echo.
echo ============================================================
echo   IMPORTANT ONE-TIME MANUAL STEP ON GITHUB:
echo   1. Open https://github.com/Zeexx99/zee-ai-ppt/settings/pages
echo   2. Under "Build and deployment" -^> "Source", select:
echo      "GitHub Actions"
echo   3. Wait 1-2 minutes for the Actions deployment to turn green.
echo   4. Run use-online.bat to register the hosted add-in in PowerPoint.
echo ============================================================
echo.
pause
