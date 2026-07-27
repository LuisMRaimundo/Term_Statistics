@echo off
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONPATH=%~dp0
python textura_gui.py
if errorlevel 1 pause
endlocal
