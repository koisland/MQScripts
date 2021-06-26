Attribute VB_Name = "RegexVLookup"
Function Regex_VLookup(sRegexPattern As String, rRange As Range, iReturnCol As Long, Optional sDelimiter As String = ", ") As String
    Dim objRegex As Object
    Set objRegex = New RegExp
    
    Dim sResult As String
    sResult = ""
    
    ' Set regex pattern
    objRegex.Pattern = sRegexPattern
    For Each rw In rRange.Rows
        If objRegex.Test(rw.Cells(1, 1).Value) Then
            If sResult = "" Then
                sResult = sResult & CStr(rw.Cells(1, iReturnCol).Value)
            Else
                sResult = sResult & sDelimiter & CStr(rw.Cells(1, iReturnCol).Value)
            End If
        End If
    Next rw
    If sResult = "" Then
        Regex_VLookup = "None"
    Else
        Regex_VLookup = sResult
    End If
    
End Function
