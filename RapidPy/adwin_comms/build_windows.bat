@echo off
setlocal
cd /d %~dp0
python tools\generate_icon.py
cd /d %~dp0..\..
python -m PyInstaller --noconfirm --clean installer\rapid_adwin_comms.spec
if %errorlevel% neq 0 exit /b %errorlevel%
echo Build complete: dist\RapidPyADWin.exe
