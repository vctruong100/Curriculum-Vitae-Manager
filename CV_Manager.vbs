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
Dim strPython, intPythonFound
strPython = ""
intPythonFound = 0

' Check for 'py' launcher
If objFSO.FileExists("C:\Windows\py.exe") Then
    strPython = "py"
    intPythonFound = 1
Else
    ' Try 'python' command
    Dim objExec, strOutput
    On Error Resume Next
    Set objExec = objShell.Exec("cmd /c where python 2>nul")
    If Err.Number = 0 Then
        strOutput = objExec.StdOut.ReadAll()
        If Trim(strOutput) <> "" Then
            strPython = "python"
            intPythonFound = 1
        End If
    End If
    On Error GoTo 0
End If

If intPythonFound = 0 Then
    MsgBox "ERROR: Python not found!" & vbCrLf & vbCrLf & _
           "Please install Python 3.8+ from python.org" & vbCrLf & _
           "and ensure it is added to your system PATH.", _
           vbCritical, "CV Manager - Python Not Found"
    WScript.Quit 1
End If

' Check if main.py exists
If Not objFSO.FileExists(strScriptDir & "\src\main.py") Then
    MsgBox "ERROR: Source file not found!" & vbCrLf & vbCrLf & _
           "Could not find: src\main.py" & vbCrLf & vbCrLf & _
           "Please ensure all application files are extracted properly.", _
           vbCritical, "CV Manager - File Not Found"
    WScript.Quit 1
End If

' Test Python can run and check dependencies
Dim intTestResult
intTestResult = 0

On Error Resume Next
Dim objTestExec
Set objTestExec = objShell.Exec(strPython & " -c ""import docx, openpyxl, rapidfuzz, PIL, win32clipboard"" 2>nul")
If Err.Number <> 0 Then
    intTestResult = 1
Else
    ' Wait for process to complete
    Do While objTestExec.Status = 0
        WScript.Sleep 100
    Loop
    If objTestExec.ExitCode <> 0 Then
        intTestResult = 1
    End If
End If
On Error GoTo 0

' Install dependencies if missing
If intTestResult <> 0 Then
    ' Show message that we're installing
    MsgBox "First-time setup: Installing required dependencies..." & vbCrLf & vbCrLf & _
           "This may take a few minutes. Click OK to continue.", _
           vbInformation, "CV Manager - First Run Setup"
    
    ' Run pip install in visible window so user sees progress (1 = visible, True = wait)
    Dim intInstallResult
    intInstallResult = objShell.Run("cmd /c """ & strPython & """ -m pip install -r """ & strScriptDir & "\requirements.txt"""", 1, True)
    
    If intInstallResult <> 0 Then
        MsgBox "ERROR: Failed to install dependencies." & vbCrLf & vbCrLf & _
               "Please check your internet connection and run manually:" & vbCrLf & _
               "  " & strPython & " -m pip install -r requirements.txt", _
               vbCritical, "CV Manager - Installation Failed"
        WScript.Quit 1
    End If
    
    ' Verify installation worked
    On Error Resume Next
    Dim objVerifyExec
    Set objVerifyExec = objShell.Exec(strPython & " -c ""import docx, openpyxl, rapidfuzz, PIL, win32clipboard"" 2>nul")
    If Err.Number <> 0 Then
        MsgBox "ERROR: Dependencies still missing after installation." & vbCrLf & vbCrLf & _
               "Please try installing manually:" & vbCrLf & _
               "  " & strPython & " -m pip install -r requirements.txt", _
               vbCritical, "CV Manager - Installation Failed"
        WScript.Quit 1
    End If
    Do While objVerifyExec.Status = 0
        WScript.Sleep 100
    Loop
    If objVerifyExec.ExitCode <> 0 Then
        MsgBox "ERROR: Dependencies still missing after installation." & vbCrLf & vbCrLf & _
               "Please try installing manually:" & vbCrLf & _
               "  " & strPython & " -m pip install -r requirements.txt", _
               vbCritical, "CV Manager - Installation Failed"
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    MsgBox "Dependencies installed successfully!" & vbCrLf & vbCrLf & _
           "Launching CV Manager...", _
           vbInformation, "CV Manager - Setup Complete"
End If

' Launch the application (0 = hidden window, False = don't wait)
objShell.Run strPython & " src\main.py", 0, False

' Clean up
Set objShell = Nothing
Set objFSO = Nothing
