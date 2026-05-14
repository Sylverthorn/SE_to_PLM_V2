' SE_to_PLM Launcher
' Automates venv creation and launches the GUI

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the root folder where this script is located
strRootPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strVenvPath = strRootPath & "\venv"
strPythonExe = strVenvPath & "\Scripts\pythonw.exe"

' 1. Check if venv exists
If Not objFSO.FolderExists(strVenvPath) Then
    ' Create venv
    objShell.Run "cmd /c python -m venv " & strVenvPath, 0, True
    objShell.Run strVenvPath & "\Scripts\pip.exe install -r " & strRootPath & "\requirements.txt", 0, True
End If

' 2. Run the application
' We target the app/main.py as the entry point
' Note: Running as module (-m) requires the root directory to be in PYTHONPATH
' or the script to be run from the root.
strCommand = "cmd /c cd /d """ & strRootPath & """ && """ & strPythonExe & """ -m SE_to_PLM.app.main"
objShell.Run strCommand, 0, False
