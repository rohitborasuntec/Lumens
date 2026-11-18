@echo off
setlocal

set VENV_DIR=myenv

REM Check the venv exists
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo Virtual environment "%VENV_DIR%" not found.
    echo Run setup first, or create it with: python -m venv %VENV_DIR%
    exit /b 1
)

REM Activate the environment
call "%VENV_DIR%\Scripts\activate.bat"

REM Use "python" (not python3) on Windows inside a venv
python lumen_price.py

REM Capture exit code before deactivating
set EXIT_CODE=%errorlevel%

REM Deactivate the environment
call deactivate

endlocal & exit /b %EXIT_CODE%