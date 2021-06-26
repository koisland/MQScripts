Attribute VB_Name = "AddDay"
Sub AddNewDay()
    Dim sCurrentDate As String
    sCurrentDate = Month(Now) & "." & Day(Now) & "." & Right(Year(Now), 2)
    
    ' If date isn't a sheet...
    If CheckSheetExists(sCurrentDate) = False Then
        ' Copy template and rename to current date.
        ActiveWorkbook.Worksheets("Template").Copy After:=ActiveWorkbook.Worksheets(Sheets.Count)
        With Worksheets(Sheets.Count)
            .Name = sCurrentDate
        End With
    End If
End Sub

