@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on PATH.
    echo Install Python 3 and enable the PATH option.
    pause
    exit /b 1
)

if /I "%~1"=="--check" (
    python -c "import pyautogui, keyboard, PIL; print('Python environment is ready.')"
    if errorlevel 1 pause
    exit /b %errorlevel%
)

python areacap4kind1529.py
set "exit_code=%errorlevel%"

if not "%exit_code%"=="0" (
    echo.
    echo The application ended with error code %exit_code%.
    pause
)

exit /b %exit_code%
