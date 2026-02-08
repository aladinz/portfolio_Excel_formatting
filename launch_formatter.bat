@echo off
REM Portfolio Formatter Streamlit App Launcher
REM Run this batch file to start the Streamlit app

cd /d "%~dp0"
.venv\Scripts\python.exe -m streamlit run portfolio_formatter_app.py
pause
