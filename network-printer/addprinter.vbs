' addprinter.vbs - Windows Logon Script
Set objNetwork = CreateObject("WScript.Network")
objNetwork.AddWindowsPrinterConnection "\\PRINTSERVER\PRINTER"
