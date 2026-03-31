@echo off
cd /d "%~dp0"
echo [1/2] Checking updates...
git pull origin main
echo [2/2] Running app...
py launch_gui.py
pause