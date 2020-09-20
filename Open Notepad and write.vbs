'Open notepad and write contents to it

set wshell=WScript.CreateObject("Wscript.Shell")
wshell.run "notepad.exe"
Wshell.AppActivate "Notepad"
Wscript.sleep 1500  'Needs a sleep increment right after notepad is opened
wshell.sendkeys "sumit good luck"
