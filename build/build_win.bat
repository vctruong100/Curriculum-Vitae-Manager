@echo off
setlocal enabledelayedexpansion
REM =========================================================================
REM  Build script for CV Research Experience Manager (Windows)
REM
REM  Usage:
REM      build\build_win.bat              One-file build (default)
REM      build\build_win.bat onedir       One-folder build
REM      build\build_win.bat onefile       One-file build (explicit)
REM
REM  The script:
REM    1. Resolves py or python
REM    2. Installs dependencies + PyInstaller
REM    3. Bumps the build number (invalidates Windows icon cache)
REM    4. Generates / copies the application icon
REM    5. Cleans old build artifacts
REM    6. Runs PyInstaller with --clean
REM    7. Creates Desktop + Start Menu shortcuts with AppUserModelID
REM    8. Prints post-build guidance
REM
REM  No network requests except pip install (local PyPI mirror is fine).
REM =========================================================================

cd /d "%~dp0.."

REM --- Resolve Python launcher ---
set "PY="
where py >nul 2>nul
if %errorlevel% equ 0 (
    set "PY=py"
) else (
    where python >nul 2>nul
    if %errorlevel% equ 0 (
        set "PY=python"
    ) else (
        echo ERROR: Neither py nor python found on PATH.
        echo        Install Python 3.8+ and ensure it is on PATH.
        exit /b 1
    )
)

echo === CV Manager Build Script ===
for /f "tokens=*" %%V in ('%PY% --version 2^>^&1') do echo Python: %%V

REM --- Determine build mode ---
set "BUILD_MODE=onefile"
if /i "%~1"=="onedir" (
    set "BUILD_MODE=onedir"
)
echo Build mode: %BUILD_MODE%

REM --- Install dependencies ---
echo.
echo [1/7] Installing dependencies...
%PY% -m pip install -r requirements.txt >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    exit /b 1
)
%PY% -m pip install --upgrade pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Failed to install PyInstaller.
    exit /b 1
)

REM --- Bump build number ---
echo [2/7] Bumping build number...
%PY% build\bump_version.py
if %errorlevel% neq 0 (
    echo WARNING: Could not bump build number — continuing.
)

REM --- Generate icon ---
echo [3/7] Generating application icon...
%PY% build\generate_icon.py
if %errorlevel% neq 0 (
    echo WARNING: Icon generation failed — build will proceed without icon.
)

REM --- Clean old artifacts ---
echo [4/7] Cleaning old build artifacts...
if exist CV_Manager.exe (
    del /f CV_Manager.exe
)
if exist CV_Manager (
    rmdir /s /q CV_Manager
)

REM --- Run PyInstaller ---
echo [5/7] Running PyInstaller (--clean)...
set "CONSOLE_MODE=0"
set "BUILD_MODE=%BUILD_MODE%"
%PY% -m PyInstaller --clean --noconfirm --distpath . cv_manager.spec
if %errorlevel% neq 0 (
    echo ERROR: PyInstaller build failed.
    exit /b 1
)

REM --- Report output ---
echo [6/7] Build output:
if "%BUILD_MODE%"=="onedir" (
    if exist "CV_Manager\CV_Manager.exe" (
        echo   CV_Manager\CV_Manager.exe
    ) else (
        echo   WARNING: One-folder exe not found.
    )
) else (
    if exist "CV_Manager.exe" (
        echo   CV_Manager.exe
    ) else (
        echo   WARNING: One-file exe not found.
    )
)

REM --- Create shortcuts ---
echo [7/7] Creating shortcuts with AppUserModelID...
%PY% build\create_shortcut.py 2>nul
if %errorlevel% neq 0 (
    echo   Shortcut creation skipped (non-critical).
)

REM --- Post-build guidance ---
echo.
echo === Build complete ===
echo.
echo If the Windows taskbar still shows an old icon:
echo   1. Right-click the pinned icon and choose "Unpin from taskbar"
echo   2. Re-pin the newly built CV_Manager.exe
echo   3. The new icon and version resource will take effect immediately.
echo.

endlocal
