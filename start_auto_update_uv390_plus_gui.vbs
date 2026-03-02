Option Explicit

Dim shell, fso, scriptDir, pythonwPath, cmd
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonwPath = "C:\Python314\pythonw.exe"

If Not fso.FileExists(pythonwPath) Then
    pythonwPath = "pythonw.exe"
End If

cmd = """" & pythonwPath & """ """ & scriptDir & "\opengd77_auto_update_uv390_plus_gui.py"""
shell.Run cmd, 0, False
