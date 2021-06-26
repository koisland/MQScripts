Attribute VB_Name = "GenerateCopy"
Sub Generate_Copy()
    ' Boolean if an item has been selected in listbox.
    Dim bItemSelected As Boolean
    Dim iListNum As Integer
    ' Index of item selected on listbox. Zero indexed.
    Dim iNumSelected As Integer

    Dim wbOld As Workbook
    Dim wbNew As Workbook
    Set wbOld = ActiveWorkbook
    
    Dim objIncLstBox As Object
    Set objIncLstBox = wbOld.Sheets("Generate Copy").ListBox1
    Dim sNewFilename As String
    
    ' Create new filename with current date only if doesn't exist.
    sNewFilename = Left(wbOld.Name, Len(wbOld.Name) - 5) & " (" & Replace(Date, "/", ".") & ").xlsx"
    If Dir(ThisWorkbook.Path & "\" & sNewFilename) <> "" Then
        MsgBox "File already exists."
        Exit Sub
    End If
    
    ' Ignore notifications so script runs in background.
    Adjust_Notif (False)
    
    bItemSelected = False
    iNumSelected = 0
    With objIncLstBox
        For iListNum = 0 To .ListCount - 1
            If .Selected(iListNum) = True Then
                ' Do not like this one bit! Must be better, more elegant way.
                ' ItemsSelected property for Access not Excel. :/
                If bItemSelected = False Then
                    Set wbNew = Workbooks.Add
                    bItemSelected = True
                End If
                iNumSelected = iNumSelected + 1
                wbOld.Sheets(.List(iListNum)).Copy wbNew.Sheets(iNumSelected)
            End If
        Next
    End With
    
    If bItemSelected = True Then
        ' Delete Sheet1 from new workbook
        ' Save in current wb path and close.
        With wbNew
            .Sheets("Sheet1").Delete
            .SaveAs Filename:=ThisWorkbook.Path & "\" & sNewFilename, FileFormat:=51
            .Close
        End With
        MsgBox "File generated." & vbNewLine & "@ " & ThisWorkbook.Path & "\" & sNewFilename
    End If
    
    ' Reenable notifications.
    Adjust_Notif (True)
End Sub
