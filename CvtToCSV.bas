Attribute VB_Name = "CvtToCSV"
' Adapted from https://www.exceltip.com/cells-ranges-rows-and-columns-in-vba/consolidatemerge-multiple-worksheets-into-one-master-sheet-using-vba.html

Sub ConvertToCSV(rHeaderRow As Range)
    Dim ws As Worksheet
    Dim wbMaster As Workbook
    Dim wbMerged As Workbook
    Dim wbMergedPath As String
    
    Set wbMaster = ActiveWorkbook
    Set wbMerged = Workbooks.Add
    
    Dim iStartRow, iStartCol As Long
    Dim iLastRow, iLastCol As Long
    Dim rSectionToMove As Range
    Dim iLastRowNewWb As Long
    
    Adjust_Notif (False)
    wbMergedPath = ThisWorkbook.Path & "\" & "CSV_" & Left(wbMaster.Name, Len(wbMaster.Name) - 5)
    If Dir(wbMergedPath) <> "" Then
        MsgBox "File already exists."
        Exit Sub
    End If
    
    ' Copy and paste header row.
    rHeaderRow.Copy
    With wbMerged.Sheets(1).Range("A1")
        .PasteSpecial xlPasteValues
    End With
    
    For Each ws In wbMaster.Sheets
        ' Skip key and template
        If ws.Name <> "Key" And ws.Name <> "Template" Then
            iStartRow = rHeaderRow.Row
            iStartCol = rHeaderRow.Column
            
            ws.Activate
            ' Get last row for coord (row, col) and offset by 1 row.
            iLastRow = Cells(iStartRow, iStartCol).End(xlDown).Offset(1, 0).Row
            iLastCol = Cells(iStartRow, iStartCol).End(xlToRight).Offset(0, 1).Column
            
            ' +1 to row to omit header
            Set rSectionToMove = Range(Cells(iStartRow + 1, iStartCol), Cells(iLastRow, iLastCol))
            
            rSectionToMove.Copy
            
            ' Start from new wb at A col, then get cell at end of sheet. Move up from bottom until first cell value. Get row and add 1 to offset
            ' Start from bottom to avoid hitting gaps which would paste at wrong position.
            iLastRowNewWb = wbMerged.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
            With wbMerged.Sheets(1).Range("A" & iLastRowNewWb)
                ' Paste only values. Formulas will result in error.
                .PasteSpecial xlPasteValues
            End With
        End If
    Next ws
    
    With wbMerged
        .SaveAs Filename:=wbMergedPath, FileFormat:=6
        .Close
    End With
    
    MsgBox "Merged csv created at " & wbMergedPath
    Adjust_Notif (True)
End Sub



