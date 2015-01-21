Option Explicit
Dim ie, x

Set x = CreateObject("wscript.shell")
Set ie = CreateObject("InternetExplorer.Application")

Sub WaitForLoad
Do while ie.Busy
wscript.sleep 200
Loop
End Sub

ie.Navigate "https://www.facebook.com/"
ie.Toolbar=0
ie.StatusBar=0
ie.Height=560
ie.Width=1000
ie.Top=0
ie.Left=0
ie.Resizable=0
ie.Visible=1

Call WaitForLoad
x.sendkeys "cow"
x.sendkeys "{tab}"
x.sendkeys "pass"
x.sendkeys "{enter}"