Attribute VB_Name = "AddCol"
Sub AddColumns(iBeforeCol As Long, sText As String, iStartRow As Long, iEndRow As Long)
    ' Add column at selected column index and input desired string
    
    On Error GoTo CleanUp
    
    Dim ws As Worksheet
    Dim sAreaToFill As String
    
    For Each ws In ActiveWorkbook.Sheets
        ' Insert column
        ws.Columns(iBeforeCol).Insert
        
        ' Get address of region and set text value.
        sAreaToFill = Range(Cells(iStartRow, iBeforeCol), Cells(iEndRow, iBeforeCol)).Address
        ws.Range(sAreaToFill).Value = sText
    Next ws

Exit Sub

CleanUp:
    'Print error and delete added column.
    Debug.Print Err.Description
    ws.Columns(iBeforeCol).Delete
    
End Sub
