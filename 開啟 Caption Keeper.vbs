Option Explicit

Dim shell, fso, folder, pythonw, script
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = shell.ExpandEnvironmentStrings("%LocalAppData%") & "\Programs\Python\Python312\pythonw.exe"
script = folder & "\caption_keeper.py"

If Not fso.FileExists(pythonw) Then
  MsgBox "Python 3.12 was not found. Please reinstall Caption Keeper.", vbCritical, "Caption Keeper"
  WScript.Quit 1
End If

shell.Run Chr(34) & pythonw & Chr(34) & " " & Chr(34) & script & Chr(34), 1, False

