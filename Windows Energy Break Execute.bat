@echo off
REM Double-click to run from source (requires Python + pip install -r requirements.txt).
cd /d "%~dp0"
py -3 draw_prize_ui.py
if errorlevel 1 pause
