@echo off
setlocal
cd /d %~dp0
python -m PyInstaller --noconfirm --clean --windowed --name RapidPyChangerXY --onefile --icon assets\changer_xy_control_icon.ico --add-data assets;assets main.py
if %errorlevel% neq 0 exit /b %errorlevel%
echo Build complete: dist\RapidPyChangerXY.exe
