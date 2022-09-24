Set wshShell =wscript.CreateObject("WScript.Shell")
do
wscript.sleep 100
wshshell.sendkeys "{l}"
wscript.sleep 300
wshshell.sendkeys "{ENTER}"
loop     
