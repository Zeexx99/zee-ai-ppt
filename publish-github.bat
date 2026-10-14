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

echo Ensuring public GitHub repository Zeexx99/zee-ai-ppt exists...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$cred = \"protocol=https`nhost=github.com`n`n\" | git credential fill 2>$null; $m = ($cred | Select-String '^password=(.+)$'); if ($m) { $tok = $m.Matches.Groups[1].Value; $h = @{ Authorization = \"Bearer $tok\"; Accept = 'application/vnd.github+json'; 'X-GitHub-Api-Version' = '2022-11-28'; 'User-Agent' = 'ZeeAI-PPT' }; try { Invoke-RestMethod -Uri 'https://api.github.com/repos/Zeexx99/zee-ai-ppt' -Headers $h -Method Get -ErrorAction Stop | Out-Null; Write-Host 'Repository Zeexx99/zee-ai-ppt is ready.' } catch { Write-Host 'Creating public repository Zeexx99/zee-ai-ppt...'; $b = @{ name = 'zee-ai-ppt'; description = 'Zee AI for PowerPoint - Task-Pane Add-in'; private = $false } | ConvertTo-Json; Invoke-RestMethod -Uri 'https://api.github.com/user/repos' -Headers $h -Method Post -Body $b -ContentType 'application/json' | Out-Null; Write-Host 'Created https://github.com/Zeexx99/zee-ai-ppt' } }"

echo Configuring remote origin https://github.com/Zeexx99/zee-ai-ppt.git ...
git remote remove origin >nul 2>&1
git remote add origin https://github.com/Zeexx99/zee-ai-ppt.git
git push -u origin main
if errorlevel 1 (
  echo [ERROR] git push failed. Please check your internet connection or GitHub login.
  pause
  exit /b 1
)

echo Enabling GitHub Pages (Source: GitHub Actions)...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$cred = \"protocol=https`nhost=github.com`n`n\" | git credential fill 2>$null; $m = ($cred | Select-String '^password=(.+)$'); if ($m) { $tok = $m.Matches.Groups[1].Value; $h = @{ Authorization = \"Bearer $tok\"; Accept = 'application/vnd.github+json'; 'X-GitHub-Api-Version' = '2022-11-28'; 'User-Agent' = 'ZeeAI-PPT' }; $b = @{ build_type = 'workflow' } | ConvertTo-Json; try { Invoke-RestMethod -Uri 'https://api.github.com/repos/Zeexx99/zee-ai-ppt/pages' -Headers $h -Method Post -Body $b -ContentType 'application/json' -ErrorAction Stop | Out-Null; Write-Host 'Enabled GitHub Pages (Source: GitHub Actions)!' } catch { try { Invoke-RestMethod -Uri 'https://api.github.com/repos/Zeexx99/zee-ai-ppt/pages' -Headers $h -Method Put -Body $b -ContentType 'application/json' -ErrorAction Stop | Out-Null; Write-Host 'Updated GitHub Pages source to GitHub Actions!' } catch { Write-Host 'Note: Verify Pages source at https://github.com/Zeexx99/zee-ai-ppt/settings/pages' } } }"

echo.
echo ============================================================
echo   [SUCCESS] Published to https://github.com/Zeexx99/zee-ai-ppt
echo   1. Check deployment progress (1-2 min):
echo      https://github.com/Zeexx99/zee-ai-ppt/actions
echo   2. If needed, verify Pages is set to "GitHub Actions" at:
echo      https://github.com/Zeexx99/zee-ai-ppt/settings/pages
echo   3. Run use-online.bat to register the hosted add-in in PowerPoint.
echo ============================================================
echo.
pause
