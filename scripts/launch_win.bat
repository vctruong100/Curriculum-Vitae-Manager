@echo off
cd /d "%~dp0\.."

echo Installing required libraries...
py -m pip install -r requirements.txt 2>nul
if %errorlevel% neq 0 (
    python -m pip install -r requirements.txt 2>nul
    if %errorlevel% neq 0 (
        echo Failed to install libraries. Please check Python installation.
        pause
        exit /b 1
    )
)

echo Starting CV Manager...
py src\main.py %* 2>nul
if %errorlevel% neq 0 (
    python src\main.py %*
    if %errorlevel% neq 0 (
        echo Failed to start the application.
        pause
        exit /b 1
    )
)
