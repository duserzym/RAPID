@echo off
setlocal
cd /d %~dp0
python -m PyInstaller --noconfirm --clean --windowed --name RapidPyAFTuner --onefile main.py
if %errorlevel% neq 0 exit /b %errorlevel%
echo Build complete: dist\RapidPyAFTuner.exe
