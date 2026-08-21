@echo off
cd /d "%~dp0"
start "" cmd /c "timeout /t 4 >nul & start http://localhost:8501"
py -m streamlit run app.py
pause
