Attribute VB_Name = "GetSheets"
Function GetSheetNames(wb As Workbook) As Variant
    Dim iNumSheets As Long
    Dim aSheetNames()
    
    iNumSheets = wb.Sheets.Count
    ReDim aSheetNames(iNumSheets)
    
    For i = 1 To iNumSheets
        aSheetNames(i) = wb.Sheets(i).Name
    Next i
    
    GetSheetNames = aSheetNames
End Function
