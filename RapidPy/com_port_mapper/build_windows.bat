@echo off
setlocal
set "APP_DIR=%~dp0"
set "REPO_ROOT=%~dp0..\.."
cd /d "%APP_DIR%"
set "VENV_PY=%REPO_ROOT%\.venv\Scripts\python.exe"
if exist "%VENV_PY%" (
    set "PYTHON_EXE=%VENV_PY%"
) else (
    set "PYTHON_EXE=python"
)

REM Generate icon assets
echo [1/3] Generating icon assets...
"%PYTHON_EXE%" "%APP_DIR%tools\generate_icon.py"
if %errorlevel% neq 0 (
    echo Error: Icon generation failed
    exit /b %errorlevel%
)

REM Build with PyInstaller
echo [2/3] Building with PyInstaller...
pushd "%REPO_ROOT%"
"%PYTHON_EXE%" -m PyInstaller --noconfirm --clean installer\rapid_com_port_mapper.spec
if %errorlevel% neq 0 (
    echo Error: PyInstaller build failed
    exit /b %errorlevel%
)

echo [3/3] Build complete!
popd
echo The executable is available at: %REPO_ROOT%\dist\RapidPyCOMMapper.exe
exit /b 0
