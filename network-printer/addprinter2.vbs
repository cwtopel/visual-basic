Set WshNetwork = WScript.CreateObject("WScript.Network")
PrinterPath = "\\printserv\DefaultPrinter"
PrinterDriver = "HP OfficeJet 6000"
WshNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
