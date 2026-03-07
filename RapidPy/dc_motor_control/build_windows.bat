@echo off
setlocal
cd /d %~dp0
python -m PyInstaller --noconfirm --clean --windowed --name RapidPyDCMotors --onefile main.py
if %errorlevel% neq 0 exit /b %errorlevel%
echo Build complete: dist\RapidPyDCMotors.exe
