Option Explicit
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("C:\Users\Jeremy\Desktop\sinking.jpg") then
fso.CopyFile "C:\Users\Jeremy\Desktop\sinking.jpg" , "C:\Users\Jeremy\Downloads\Test\"
Else
wscript.echo "doesn't exist"
End If