<#
.SYNOPSIS
    Build CV_Manager.exe using PyInstaller.

.DESCRIPTION
    Comprehensive build script that:
      1. Resolves py or python
      2. Installs dependencies + PyInstaller
      3. Bumps the build number (invalidates Windows icon cache)
      4. Generates / copies the application icon
      5. Cleans build artifacts (build/, dist/)
      6. Runs PyInstaller with --clean
      7. Creates Desktop + Start Menu shortcuts with AppUserModelID
      8. Prints post-build guidance

.PARAMETER BuildMode
    "onefile" (default) or "onedir" for one-folder build.

.PARAMETER Console
    If set, builds with a visible console window (for debugging).

.EXAMPLE
    .\build\build_win.ps1
    .\build\build_win.ps1 -BuildMode onedir
    .\build\build_win.ps1 -Console
#>

param(
    [ValidateSet("onefile", "onedir")]
    [string]$BuildMode = "onefile",

    [switch]$Console
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Push-Location $ProjectRoot

try {
    Write-Host "=== CV Manager Build Script ===" -ForegroundColor Cyan
    Write-Host "Project root : $ProjectRoot"
    Write-Host "Build mode   : $BuildMode"

    # ------------------------------------------------------------------
    # Resolve Python launcher
    # ------------------------------------------------------------------
    $python = $null
    try {
        $null = & py --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            $python = "py"
        }
    } catch { }

    if (-not $python) {
        try {
            $null = & python --version 2>&1
            if ($LASTEXITCODE -eq 0) {
                $python = "python"
            }
        } catch { }
    }

    if (-not $python) {
        Write-Error "Neither py nor python found on PATH.  Install Python 3.8+."
        exit 1
    }

    Write-Host "Python       : $((& $python --version 2>&1))" -ForegroundColor Gray

    # ------------------------------------------------------------------
    # [1/7] Install dependencies
    # ------------------------------------------------------------------
    Write-Host "`n[1/7] Installing dependencies..." -ForegroundColor Yellow
    & $python -m pip install -r requirements.txt 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install dependencies."
        exit 1
    }
    & $python -m pip install --upgrade pyinstaller 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install PyInstaller."
        exit 1
    }

    # ------------------------------------------------------------------
    # [2/7] Bump build number
    # ------------------------------------------------------------------
    Write-Host "[2/7] Bumping build number..." -ForegroundColor Yellow
    & $python (Join-Path $ProjectRoot "build\bump_version.py")
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Could not bump build number — continuing."
    }

    # ------------------------------------------------------------------
    # [3/7] Generate icon
    # ------------------------------------------------------------------
    Write-Host "[3/7] Generating application icon..." -ForegroundColor Yellow
    & $python (Join-Path $ProjectRoot "build\generate_icon.py")
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Icon generation failed — build will proceed without icon."
    }

    # ------------------------------------------------------------------
    # [4/7] Clean old artifacts
    # ------------------------------------------------------------------
    Write-Host "[4/7] Cleaning old build artifacts..." -ForegroundColor Yellow

    $exePath = Join-Path $ProjectRoot "CV_Manager.exe"
    if (Test-Path $exePath) {
        Remove-Item $exePath -Force
        Write-Host "  Removed old CV_Manager.exe" -ForegroundColor Gray
    }
    $onedirPath = Join-Path $ProjectRoot "CV_Manager"
    if (Test-Path $onedirPath) {
        Remove-Item $onedirPath -Recurse -Force
        Write-Host "  Removed CV_Manager/" -ForegroundColor Gray
    }

    # ------------------------------------------------------------------
    # [5/7] Run PyInstaller
    # ------------------------------------------------------------------
    Write-Host "[5/7] Running PyInstaller (--clean)..." -ForegroundColor Yellow

    $consoleVal = "0"
    if ($Console) {
        $consoleVal = "1"
    }

    $env:CONSOLE_MODE = $consoleVal
    $env:BUILD_MODE = $BuildMode

    & $python -m PyInstaller --clean --noconfirm --distpath . cv_manager.spec
    if ($LASTEXITCODE -ne 0) {
        Write-Error "PyInstaller build failed."
        exit 1
    }

    # ------------------------------------------------------------------
    # [6/7] Report output
    # ------------------------------------------------------------------
    Write-Host "[6/7] Build output:" -ForegroundColor Yellow

    if ($BuildMode -eq "onedir") {
        $onedirExe = Join-Path $ProjectRoot "CV_Manager\CV_Manager.exe"
        if (Test-Path $onedirExe) {
            $size = [math]::Round((Get-Item $onedirExe).Length / 1MB, 1)
            Write-Host "  $onedirExe ($size MB)" -ForegroundColor Green
        }
        else {
            Write-Warning "One-folder exe not found at $onedirExe"
        }
    }
    else {
        $exePath = Join-Path $ProjectRoot "CV_Manager.exe"
        if (Test-Path $exePath) {
            $size = [math]::Round((Get-Item $exePath).Length / 1MB, 1)
            Write-Host "  $exePath ($size MB)" -ForegroundColor Green
        }
        else {
            Write-Warning "One-file exe not found at $exePath"
        }
    }

    # ------------------------------------------------------------------
    # [7/7] Create shortcuts
    # ------------------------------------------------------------------
    Write-Host "[7/7] Creating shortcuts with AppUserModelID..." -ForegroundColor Yellow
    & $python (Join-Path $ProjectRoot "build\create_shortcut.py") 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Shortcut creation skipped (non-critical)." -ForegroundColor Gray
    }

    # ------------------------------------------------------------------
    # Post-build guidance
    # ------------------------------------------------------------------
    Write-Host "`n=== Build complete ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "If the Windows taskbar still shows an old icon:" -ForegroundColor Yellow
    Write-Host "  1. Right-click the pinned icon and choose 'Unpin from taskbar'"
    Write-Host "  2. Re-pin the newly built CV_Manager.exe"
    Write-Host "  3. The new icon and version resource will take effect immediately."
    Write-Host ""

} finally {
    Pop-Location
    Remove-Item Env:\CONSOLE_MODE -ErrorAction SilentlyContinue
    Remove-Item Env:\BUILD_MODE -ErrorAction SilentlyContinue
}
