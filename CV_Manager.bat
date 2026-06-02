@echo off
REM =========================================================================
REM  CV Research Experience Manager - Launcher
REM
REM  Launches the application directly from Python source.
REM  No .exe needed - eliminates Windows SmartScreen issues entirely.
REM =========================================================================

cd /d "%~dp0"

REM --- Detect Python ---
set "PY="
where py >nul 2>nul
if %errorlevel% equ 0 (
    set "PY=py"
) else (
    where python >nul 2>nul
    if %errorlevel% equ 0 (
        set "PY=python"
    ) else (
        echo ERROR: Python not found on PATH.
        echo.
        echo Please install Python 3.8+ from python.org
        echo and ensure it is added to your system PATH.
        echo.
        pause
        exit /b 1
    )
)

REM --- Launch the application ---
%PY% src\main.py

REM If we get here, the application exited (error or user closed it)
REM Exit silently - no need to pause