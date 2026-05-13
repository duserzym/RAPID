@echo off
setlocal
cd /d %~dp0
python tools\generate_icon.py
python -m PyInstaller --noconfirm --clean --windowed --name RapidPyAFTuner --onefile --icon assets\af_tuner_icon.ico --add-data assets;assets main.py
if %errorlevel% neq 0 exit /b %errorlevel%
echo Build complete: dist\RapidPyAFTuner.exe
