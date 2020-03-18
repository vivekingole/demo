Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "%windir%\notepad.exe"
WshShell.AppActivate "Notepad"
Dim x
x = 1   
While x < 5
Wscript.sleep 290000
WshShell.SendKeys "."
Wend