Option Explicit
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists("C:\Users\chris\Documents\GitHub\visual-basic\vbscript-basics") Then
wscript.echo "folder exists"
ElseIf Not fso.FolderExists("C:\Users\chris\Documents\GitHub\visual-basic\vbscript-basics") Then
wscript.echo "folder doesn't exist"
End If

If fso.FileExists("C:\Users\chris\Documents\GitHub\visual-basic\vbscript-basics\folderexists.vbs") Then
wscript.echo "file exists"
Else
wscript.echo "file doesn't exist"
End If