Attribute VB_Name = "HideCol"
Sub HideColumns(iColumnNumber As Long)
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ' Filter function won't work with partial matches. No way to check array for items w/o iteration. Vba ur a joke.
        If ws.Name <> "Key" Then
            ws.Range("A" & CStr(iColumnNumber)).EntireColumn.Hidden = True
        End If
    Next ws

End Sub
  
