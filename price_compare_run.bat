@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "VENV_PYTHON=%SCRIPT_DIR%myenv\Scripts\python.exe"

REM The setup script creates this environment once; launchers never recreate it.
if not exist "%VENV_PYTHON%" (
    echo Virtual environment "myenv" not found.
    echo Run "Lumens Project.bat" first.
    exit /b 1
)

pushd "%SCRIPT_DIR%"
"%VENV_PYTHON%" price_fetch.py
set EXIT_CODE=%errorlevel%
popd

endlocal & exit /b %EXIT_CODE%