@echo off
setlocal
cd /d %~dp0

REM Generate icon assets
echo [1/3] Generating icon assets...
python tools\generate_icon.py
if %errorlevel% neq 0 (
    echo Error: Icon generation failed
    exit /b %errorlevel%
)

REM Build with PyInstaller
echo [2/3] Building with PyInstaller...
cd /d %~dp0..\..
python -m PyInstaller --noconfirm --clean installer\rapid_com_port_mapper.spec
if %errorlevel% neq 0 (
    echo Error: PyInstaller build failed
    exit /b %errorlevel%
)

echo [3/3] Build complete!
echo The executable is available at: dist\RapidPyCOMMapper.exe
exit /b 0
