@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_EXE="
where py >nul 2>nul && set "PYTHON_EXE=py -3"
if not defined PYTHON_EXE (
    where python >nul 2>nul && set "PYTHON_EXE=python"
)

if not defined PYTHON_EXE (
    echo [ERROR] Python launcher not found. Install Python 3 and add it to PATH.
    pause
    exit /b 1
)

%PYTHON_EXE% -m pip install -r requirements.txt --disable-pip-version-check -q
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

start "" %PYTHON_EXE% converter.py
if errorlevel 1 (
    echo [ERROR] Failed to start converter.py.
    pause
    exit /b 1
)

exit /b 0
