set Shell = CreateObject("Wscript.shell")
set oNet = CreateObject("Wscript.Network")

DomainName = oNet.UserDomain

Set Domain = GetObject("WinNT://" & DomainName)

For Each ADSIObject In Domain
If ADSIObject.Class = "Computer" Then
Shell.Run "C:\tools\psexec -d \\" & ADSIObject.Name & " gpupdate /force"
End If
Next
