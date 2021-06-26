Attribute VB_Name = "SetLabelPrinter"
Public Sub Set_Printer(sPrinterName As String)
Attribute Set_Printer.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim networks() As Variant
    Dim network As Variant
    
    'Brute force it to find printer from pool of networks.
    On Error Resume Next
    
    networks = Array("Ne04", "Ne06", "Ne08", "Ne09", "Ne10")
    
    For Each network In networks
        Application.ActivePrinter = sPrinterName & " on " & network & ":"
    Next network
        
    ' MsgBox Application.ActivePrinter
    SetupProp
End Sub
