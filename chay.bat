@echo off
cd /d %~dp0
chcp 65001 >nul
if not exist venv (python -m venv venv && venv\Scripts\pip install -r requirements.txt)
venv\Scripts\python app.py
pause
