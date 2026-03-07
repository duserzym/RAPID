@echo off
setlocal
cd /d %~dp0
python tools\generate_icon.py
python -m PyInstaller --noconfirm --clean --windowed --name RapidPyUpDown --onefile --icon assets\updown_icon.ico main.py
if %errorlevel% neq 0 exit /b %errorlevel%
echo Build complete: dist\RapidPyUpDown.exe
