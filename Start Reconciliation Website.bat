@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Glen Reconciliation - Local Start

echo ============================================================
echo GLEN RECONCILIATION - LOCAL START
echo Folder: %CD%
echo ============================================================
echo.

echo [1/3] Project pre-flight...
py preflight_check.py
if errorlevel 1 goto :fail

echo.
echo [2/3] Database diagnostic...
py db_diagnostic.py
if errorlevel 1 (
    echo.
    echo Database is not reachable right now. The browser app will show the
    echo same safe database status instead of a raw traceback.
    echo.
)

echo [3/3] Starting Streamlit...
start "" cmd /c "timeout /t 4 >nul & start http://localhost:8501"
py -m streamlit run app.py
pause
exit /b 0

:fail
echo.
echo Project pre-flight failed. Nothing was changed in Supabase.
pause
exit /b 1
