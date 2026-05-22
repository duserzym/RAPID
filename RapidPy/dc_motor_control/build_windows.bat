@echo off
setlocal
set "APP_DIR=%~dp0"
set "REPO_ROOT=%~dp0..\.."
set "VENV_PY=%REPO_ROOT%\.venv\Scripts\python.exe"
if exist "%VENV_PY%" (
	set "PYTHON_EXE=%VENV_PY%"
) else (
	set "PYTHON_EXE=python"
)
cd /d "%APP_DIR%"
"%PYTHON_EXE%" -m PyInstaller --noconfirm --clean --windowed --name RapidPyDCMotors --onefile --distpath "%REPO_ROOT%\dist" --workpath "%REPO_ROOT%\build\rapid_dc_motor_control" --specpath "%REPO_ROOT%\build\rapid_dc_motor_control" main.py
if %errorlevel% neq 0 exit /b %errorlevel%
echo Build complete: %REPO_ROOT%\dist\RapidPyDCMotors.exe
