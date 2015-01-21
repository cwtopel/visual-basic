set x = createobject("wscript.shell")

x.run "notepad.exe"
wscript.sleep 200
x.sendkeys "Now we're open."
x.sendkeys "{enter}"
x.sendkeys "Line 2"

x.sendkeys "%fs"