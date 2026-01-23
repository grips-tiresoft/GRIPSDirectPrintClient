Set objFSO = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
psScript = scriptDir & "\Print-GRDPFile.ps1"
sh.Run "powershell.exe -STA -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File """ & psScript & """ """ & WScript.Arguments(0) & """", 0, False