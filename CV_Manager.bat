@echo off
REM =========================================================================
REM  CV Research Experience Manager - Launcher
REM
REM  Auto-installs dependencies and launches the application.
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

REM --- Check / Install Dependencies ---
echo Checking dependencies...
%PY% -c "import docx, openpyxl, rapidfuzz, PIL, win32clipboard" >nul 2>nul
if %errorlevel% neq 0 (
    echo.
    echo Dependencies not found. Installing now...
    echo This may take a few minutes on first run.
    echo.
    %PY% -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo.
        echo ERROR: Failed to install dependencies.
        echo Please check your internet connection or run manually:
        echo   %PY% -m pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
    echo.
    echo Dependencies installed successfully!
    echo.
)

REM --- Launch the application ---
%PY% src\main.py

REM Check if Python execution failed
if %errorlevel% neq 0 (
    echo.
    echo ============================================
    echo ERROR: Application failed to start
    echo Exit code: %errorlevel%
    echo.
    echo Common causes:
    echo   - Missing dependencies: Run 'py -m pip install -r requirements.txt'
    echo   - Python not properly installed
    echo   - Corrupted source files
    echo.
    echo For help, check the logs or run from terminal to see full error.
    echo ============================================
    echo.
    pause
    exit /b %errorlevel%
)

REM If we get here, the application exited normally (user closed it)
REM Exit silently - no need to pause