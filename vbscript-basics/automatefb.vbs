Option Explicit
Dim ie, ipf

Set ie = CreateObject("InternetExplorer.Application")

On Error Resume Next

Sub WaitForLoad
Do While IE.Busy
WScript.Sleep 500
Loop
End Sub

Sub Find(x)
Set ipf = ie.Document.All.Item(x)
End Sub

ie.Left = 0
ie.Top = 0
ie.Toolbar = 0
ie.StatusBar = 0
ie.Height = 120
ie.Width = 1020
ie.Resizable = 0

ie.Navigate "https://www.facebook.com/"

Call WaitForLoad

ie.Visible = True

Call Find("email")
ipf.Value = "USERNAME OR EMAIL"
Call Find("pass")
ipf.Value = "PASSWORD"
Call Find("login_form")
ipf.Submit

Call WaitForLoad
ie.Height = 700