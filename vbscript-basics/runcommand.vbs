Option Explicit
Dim obj, a, b, c

Set obj = CreateObject("wscript.shell")
'obj.run "C:\Users\chris\Documents\GitHub\visual-basic\vbscript-basics"

a = msgbox("Open a Program?", vbYesNoCancel+vbQuestion+vbSystemModal)

if a = vbYes then
	obj.run "mspaint.exe"
	b = msgbox("Open a Folder?", vbYesNo+vbQuestion+vbSystemModal)
else
	b = msgbox("Open a Folder?", vbYesNo+vbQuestion+vbSystemModal)
end if

if b = vbYes then
	obj.run "C:\Users\chris\Documents\GitHub\visual-basic\vbscript-basics"
	c = msgbox("Open a File?", vbYesNo+vbQuestion+vbSystemModal)
else
	c = msgbox("Open a File?", vbYesNo+vbQuestion+vbSystemModal)
end if

if c = vbYes then
	obj.run "C:\Users\chris\Documents\GitHub\visual-basic\vbscript-basics\vbscript_msgbox_chart_numbers.jpg"
else
	wscript.quit
end if