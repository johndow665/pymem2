Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "python.exe q.py", 0, True
Set WshShell = Nothing
