@echo off
cd /d "%~dp0"

REM --- Build the .exe using the comprehensive build script ---
call build\build_win.bat
if %errorlevel% neq 0 (
    echo Build failed — falling back to source launch.
)

REM --- Launch the application ---
echo Starting CV Manager...
if exist CV_Manager.exe (
    start "" CV_Manager.exe
) else (
    set "PY="
    where py >nul 2>nul
    if %errorlevel% equ 0 (
        set "PY=py"
    ) else (
        set "PY=python"
    )
    %PY% src\main.py
    if %errorlevel% neq 0 (
        echo Failed to start the application.
        pause
        exit /b
    )
)