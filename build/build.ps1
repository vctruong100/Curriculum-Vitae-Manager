<#
.SYNOPSIS
    Build CV_Manager.exe using PyInstaller.

.DESCRIPTION
    Builds CV_Manager.exe and places it at the project root next to CV_Manager.bat.

.PARAMETER Console
    If set, builds with a visible console window (for debugging).

.EXAMPLE
    .\build\build.ps1
    .\build\build.ps1 -Console
#>

param(
    [switch]$Console
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Push-Location $ProjectRoot

try {
    Write-Host "=== CV Manager Build Script ===" -ForegroundColor Cyan
    Write-Host "Project root: $ProjectRoot"

    $python = "py"
    try {
        & $python --version 2>$null | Out-Null
    } catch {
        $python = "python"
        try {
            & $python --version 2>$null | Out-Null
        } catch {
            Write-Error "Python not found. Install Python 3.8+ and ensure it is on PATH."
            exit 1
        }
    }

    Write-Host "Python: $((& $python --version 2>&1))" -ForegroundColor Gray

    Write-Host "Installing/upgrading PyInstaller..." -ForegroundColor Yellow
    & $python -m pip install --upgrade pyinstaller 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install PyInstaller."
        exit 1
    }

    Write-Host "Installing project dependencies..." -ForegroundColor Yellow
    & $python -m pip install -r requirements.txt 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install dependencies."
        exit 1
    }

    # Generate icon from Logo.png if app.ico does not exist
    $iconPath = Join-Path $ProjectRoot "build\assets\app.ico"
    if (-not (Test-Path $iconPath)) {
        Write-Host "Generating application icon..." -ForegroundColor Yellow
        & $python (Join-Path $ProjectRoot "build\generate_icon.py")
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Icon generation failed — build will proceed without icon."
        }
    } else {
        Write-Host "Icon already exists at $iconPath" -ForegroundColor Gray
    }

    $consoleVal = "0"
    if ($Console) {
        $consoleVal = "1"
    }

    $env:CONSOLE_MODE = $consoleVal

    Write-Host "`n--- Building CV_Manager.exe ---" -ForegroundColor Green
    & $python -m PyInstaller --clean --noconfirm cv_manager.spec
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build failed."
        exit 1
    }
    $exePath = Join-Path $ProjectRoot "CV_Manager.exe"
    if (Test-Path $exePath) {
        $size = [math]::Round((Get-Item $exePath).Length / 1MB, 1)
        Write-Host "Artifact: $exePath ($size MB)" -ForegroundColor Green
    }

    Write-Host "`n=== Build complete ===" -ForegroundColor Cyan

} finally {
    Pop-Location
    Remove-Item Env:\CONSOLE_MODE -ErrorAction SilentlyContinue
}
