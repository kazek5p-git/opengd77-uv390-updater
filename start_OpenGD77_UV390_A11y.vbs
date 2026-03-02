Option Explicit

Dim shell, fso, scriptDir, psCmd, logDir, logFile, ts
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
logDir = scriptDir & "\logs"
If Not fso.FolderExists(logDir) Then
    fso.CreateFolder(logDir)
End If
ts = Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2)
logFile = logDir & "\OpenGD77_A11y_launcher_" & ts & ".log"

Sub LogLine(msg)
    On Error Resume Next
    Dim fh
    Set fh = fso.OpenTextFile(logFile, 8, True, 0)
    fh.WriteLine "[" & Now & "] " & msg
    fh.Close
    On Error GoTo 0
End Sub

psCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptDir & "\OpenGD77_UV390_A11y.ps1"""
LogLine "Launch command: " & psCmd
shell.Run psCmd, 1, False
LogLine "Launch command sent to shell."
