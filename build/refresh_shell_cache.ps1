<#
.SYNOPSIS
    Print instructions to refresh the Windows Explorer icon cache.

.DESCRIPTION
    Windows caches .exe icons aggressively.  After a rebuild the taskbar or
    Explorer may still show the previous icon.  This script prints the
    recommended manual steps.  It does NOT restart Explorer automatically
    because that disrupts the user's open windows.

    Run this only if the icon does not update after a clean rebuild.

.EXAMPLE
    .\build\refresh_shell_cache.ps1
#>

Write-Host ""
Write-Host "=== Windows Icon Cache Refresh Guide ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "If the .exe or taskbar still shows an old icon after a rebuild,"
Write-Host "try the following steps in order:"
Write-Host ""
Write-Host "  1. Unpin the old shortcut from the taskbar." -ForegroundColor Yellow
Write-Host "  2. Delete the old CV_Manager.exe (the build script does this)."
Write-Host "  3. Rebuild using  build\build_win.bat  or  build\build_win.ps1"
Write-Host "  4. Re-pin the newly built exe to the taskbar."
Write-Host ""
Write-Host "If the icon STILL does not update:" -ForegroundColor Yellow
Write-Host ""
Write-Host '  a. Open an elevated Command Prompt and run:'
Write-Host '       ie4uinit.exe -show' -ForegroundColor White
Write-Host ""
Write-Host '  b. Or clear the icon cache manually:'
Write-Host '       taskkill /IM explorer.exe /F' -ForegroundColor White
Write-Host '       del /A /Q "%localappdata%\IconCache.db"' -ForegroundColor White
Write-Host '       del /A /F /Q "%localappdata%\Microsoft\Windows\Explorer\iconcache*"' -ForegroundColor White
Write-Host '       start explorer.exe' -ForegroundColor White
Write-Host ""
Write-Host "  NOTE: Step (b) will briefly close all Explorer windows." -ForegroundColor Red
Write-Host ""
Write-Host "The build script embeds a new version resource on every build,"
Write-Host "which should make step (b) unnecessary in most cases."
Write-Host ""
