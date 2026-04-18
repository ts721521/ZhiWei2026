@echo off
setlocal

cd /d "%~dp0"

REM 静默启动（无控制台）：优先 pythonw.exe；失败时退回带控制台的 python.exe 方便看错误
if exist ".python\python\pythonw.exe" (
    start "" ".python\python\pythonw.exe" office_gui.py
    exit /b 0
)
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" office_gui.py
    exit /b 0
)

REM 没有 pythonw -> 用控制台版本，带 pause 方便看错误
if exist ".python\python\python.exe" (
    set PYTHON_CMD=.\.python\python\python.exe
) else if exist ".venv\Scripts\python.exe" (
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
