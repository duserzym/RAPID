@echo off
setlocal
cd /d %~dp0
python tools\generate_icon.py
cd /d %~dp0..\..
python -m PyInstaller --noconfirm --clean installer\rapid_vrm.spec
if %errorlevel% neq 0 exit /b %errorlevel%
echo Build complete: dist\RapidPyVRM.exe
