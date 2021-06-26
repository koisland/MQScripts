Attribute VB_Name = "GetSheetName"
Public Function SheetName(rCellRef As Range) As String
    ' rCellRef because if using Application.ActiveSheet, will apply to all other sheets using same function.
    SheetName = rCellRef.Worksheet.Name
End Function
