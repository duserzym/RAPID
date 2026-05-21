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
"%PYTHON_EXE%" "%APP_DIR%tools\generate_icon.py"
pushd "%REPO_ROOT%"
"%PYTHON_EXE%" -m PyInstaller --noconfirm --clean installer\rapid_af_tuner.spec
if %errorlevel% neq 0 exit /b %errorlevel%
popd
echo Build complete: %REPO_ROOT%\dist\RapidPyAFTuner.exe
