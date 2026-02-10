@echo off
setlocal

cd /d "%~dp0"

if exist ".venv" (
    set PYTHON_CMD=.\.venv\Scripts\python.exe
) else (
    where python >nul 2>nul
    if %errorlevel% neq 0 (
        echo Error: Python is not installed or not in PATH.
        pause
        exit /b 1
    )
    set PYTHON_CMD=python
)

echo Launching Office GUI with %PYTHON_CMD%...
"%PYTHON_CMD%" office_gui.py
pause
