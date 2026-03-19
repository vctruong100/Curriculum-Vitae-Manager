<#
.SYNOPSIS
    Build CV_Manager.exe using PyInstaller.

.DESCRIPTION
    Delegates to the comprehensive build_win.ps1 script.
    Kept for backward compatibility.

.PARAMETER Console
    If set, builds with a visible console window (for debugging).

.PARAMETER BuildMode
    "onefile" (default) or "onedir" for one-folder build.

.EXAMPLE
    .\build\build.ps1
    .\build\build.ps1 -Console
    .\build\build.ps1 -BuildMode onedir
#>

param(
    [ValidateSet("onefile", "onedir")]
    [string]$BuildMode = "onefile",

    [switch]$Console
)

$buildScript = Join-Path $PSScriptRoot "build_win.ps1"

$params = @{
    BuildMode = $BuildMode
}
if ($Console) {
    $params["Console"] = $true
}

& $buildScript @params
