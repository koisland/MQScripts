Attribute VB_Name = "ChkSheetExists"
Function CheckSheetExists(sSheetName As String) As Boolean
    Dim aWkshtNames As Variant
    Dim iWkshtMatches As Long
    
    aWkshtNames = GetSheetNames(Application.ActiveWorkbook)
    ' Will include partial matches unfortuantely.
    aRes = Filter(SourceArray:=aWkshtNames, Match:=sSheetName, Include:=False)
    ' VBA being awful. length of array.
    iWkshtMatches = UBound(aRes) - LBound(aRes) + 1
    
    If iWkshtMatches = 0 Then
        CheckSheetExists = False
    Else
        CheckSheetExists = True
    End If
End Function
