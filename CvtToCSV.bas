Attribute VB_Name = "CvtToCSV"
' Adapted from https://www.exceltip.com/cells-ranges-rows-and-columns-in-vba/consolidatemerge-multiple-worksheets-into-one-master-sheet-using-vba.html

Sub ConvertToCSV()
    Dim ws As Worksheet
    Dim wbMaster As Workbook
    Dim wbMerged As Workbook
    Dim wbMergedPath As String
    Dim rSelectHeaders As Range
    Dim iStartRow, iStartCol As Long
    Dim iLastRow, iLastCol As Long
    Dim rSectionToMove As Range
    Dim iLastRowNewWb As Long
    
    ' On Error GoTo CanceledPrompt
    
    ' Ask for header row and copy it.
    Set rSelectHeaders = Application.InputBox("Select headers.", Title:="Headers", Type:=8)
    rSelectHeaders.Copy
    
    ' Set master and merged workbooks.
    Set wbMaster = ActiveWorkbook
    Set wbMerged = Workbooks.Add
    
    Adjust_Notif (False)
    
    ' Set filename and check if file already exists.
    wbMergedPath = ThisWorkbook.Path & "\" & "CSV_" & Left(wbMaster.Name, Len(wbMaster.Name) - 5)
    If Dir(wbMergedPath) <> "" Then
        MsgBox "File already exists."
        Exit Sub
    End If
    
    ' Paste header rows into merged wb.
    With wbMerged.Sheets(1).Range("A1")
        .PasteSpecial xlPasteValues
    End With
    
    For Each ws In wbMaster.Sheets
        ' Skip key and template
        If ws.Name <> "Key" And ws.Name <> "Template" Then
            iStartRow = rSelectHeaders.Row
            iStartCol = rSelectHeaders.Column
            
            ws.Activate
            ' Get last row for coord (row, col) and offset by 1 row.
            ' Add one to last row because using RDI to get end of number of rows.
            iLastRow = Cells(iStartRow, iStartCol + 1).End(xlDown).Offset(1, 0).Row
            iLastCol = Cells(iStartRow, iStartCol).End(xlToRight).Offset(0, 1).Column
            
            ' +1 to row to omit header
            Set rSectionToMove = Range(Cells(iStartRow + 1, iStartCol), Cells(iLastRow, iLastCol))
            
            Debug.Print rSectionToMove.Address
            rSectionToMove.Copy
            
            ' Start from new wb at A col, then get cell at end of sheet. Move up from bottom until first cell value. Get row.
            ' Start from bottom to avoid hitting gaps which would paste at wrong position.
            iLastRowNewWb = wbMerged.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
            With wbMerged.Sheets(1).Range("A" & iLastRowNewWb)
                ' Paste only values. Formulas will result in error.
                .PasteSpecial xlPasteValues
            End With
        End If
    Next ws
    
    ' Save as CSV (6)
    With wbMerged
        .SaveAs Filename:=wbMergedPath, FileFormat:=6
        .Close
    End With
    
    MsgBox "Merged csv created at " & wbMergedPath
    Adjust_Notif (True)
Exit Sub

CanceledPrompt:
    ' Raised since Nothing set instead of Range when prompt canceled. Classic VBA.
    If Err.Number = 424 Then
        Exit Sub
    Else
        MsgBox Err.Number & vbNewLine & Err.Description
    End If
End Sub

