' =========================================================================
'  CV Research Experience Manager - Silent Launcher
'
'  Launches the application without showing a terminal window.
'  Double-click this file for a completely silent launch.
' =========================================================================

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this script is located
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Change to the script directory
objShell.CurrentDirectory = strScriptDir

' Detect Python (try 'py' first, then 'python')
strPython = "py"
On Error Resume Next
objShell.Run "cmd /c where py >nul 2>nul", 0, True
If Err.Number <> 0 Then
    strPython = "python"
End If
On Error GoTo 0

' Launch the application (0 = hidden window, False = don't wait)
objShell.Run strPython & " src\main.py", 0, False

' Clean up
Set objShell = Nothing
Set objFSO = Nothing
