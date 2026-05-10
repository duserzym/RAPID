@echo off
setlocal
rem Run from the gaussmeter_control folder.  Changes to repo root before calling PyInstaller.
cd /d %~dp0
python tools\generate_icon.py
cd /d %~dp0..\..
python -m PyInstaller --noconfirm --clean installer\rapid_gaussmeter.spec
if %errorlevel% neq 0 exit /b %errorlevel%
echo Build complete: dist\RapidPy_Gaussmeter\
