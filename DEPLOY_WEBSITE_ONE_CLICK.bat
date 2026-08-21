@echo off
setlocal EnableExtensions
set "PYTHONUTF8=1"
title Glen Reconciliation - One Click Deploy

REM Always deploy the project folder containing this BAT file.
cd /d "%~dp0"

echo.
echo ============================================================
echo       GLEN RECONCILIATION - SAFE ONE CLICK DEPLOY
echo ============================================================
echo.

echo [1/7] Checking Git repository...
git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 goto :giterror

echo [2/7] Running local pre-flight checks...
py preflight_check.py
if errorlevel 1 goto :preflight

echo [3/7] Staging local project changes...
git add -A

git diff --cached --quiet
if %errorlevel%==0 (
    echo       No new local code changes to commit.
) else (
    echo [4/7] Creating deployment commit...
    git commit -m "Deploy reconciliation tower update"
    if errorlevel 1 goto :commiterror
)

echo [5/7] Syncing latest GitHub MAIN...
git pull --rebase origin main
if errorlevel 1 goto :pullerror

echo [6/7] Re-running pre-flight checks after sync...
py preflight_check.py
if errorlevel 1 goto :preflight

echo [7/7] Pushing MAIN to GitHub...
git push origin main
if errorlevel 1 goto :pusherror

echo.
echo ============================================================
echo DEPLOYMENT PUSH SUCCESSFUL
echo ============================================================
echo Streamlit Cloud will auto-redeploy from GitHub MAIN.
echo Your persistent Supabase data is NOT deleted by this deployment.
echo.
echo Live app:
echo https://glen-reconciliation-system-v2.streamlit.app/
echo.
start "" "https://glen-reconciliation-system-v2.streamlit.app/"
pause
exit /b 0

:giterror
echo ERROR: This folder is not a Git repository.
goto :fail
:preflight
echo ERROR: Pre-flight checks failed. Deployment stopped before push.
goto :fail
:commiterror
echo ERROR: Git commit failed.
goto :fail
:pullerror
echo ERROR: Git pull/rebase failed. Resolve the Git conflict, then run this file again.
goto :fail
:pusherror
echo ERROR: GitHub push failed. Check Git credentials/network and retry.
goto :fail
:fail
echo.
echo Nothing was intentionally deleted from Supabase.
pause
exit /b 1
