@echo off
setlocal
cd /d "%~dp0"
if "%PREHEAT_PLANNER_PORT%"=="" set "PREHEAT_PLANNER_PORT=8770"
python app.py
if errorlevel 1 pause
