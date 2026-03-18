@echo off
cd /d "%~dp0"

echo Installing required libraries...
py -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Failed to install libraries. Please check Python installation.
    pause
    exit /b
)

echo Starting CV Manager...
py src\main.py
if %errorlevel% neq 0 (
    echo Failed to start the application.
    pause
    exit /b
)