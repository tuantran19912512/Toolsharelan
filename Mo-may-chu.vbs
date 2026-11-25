Option Explicit
Dim files, shell, root, arguments
Set files = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Shell.Application")
root = files.GetParentFolderName(WScript.ScriptFullName)
arguments = "-NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File " & Chr(34) & files.BuildPath(root, "bin\May-chu.ps1") & Chr(34)
shell.ShellExecute "powershell.exe", arguments, root, "runas", 0
