Attribute VB_Name = "FindLastRow"
Public Function Adj_Range_Last_Col(rColRange As Range, rColToCheck As Range) As Range
    ' rColRange should be the whole column ex. Range("A:A")
    Dim iLastRow As Long
    Dim rAdjRange As Range
    
    iLastRow = rColToCheck.End(xlDown).Offset(1, 0).Row

    Set rAdjRange = rColRange.Resize(iLastRow, rColRange.Columns.Count)
    Set Adj_Range_Last_Col = rAdjRange.Offset(1, 0)
    
End Function
