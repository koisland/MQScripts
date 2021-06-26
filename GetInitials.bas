Attribute VB_Name = "GetInitials"
Public Function Initials(ByVal rev As Integer) As String
    Dim FullName As String
    Dim SplitName() As String
    Dim FirstLetters(2) As String
    Dim length As Integer
    Dim index As Integer
    
    FullName = Application.UserName
    ' Three initial example.
    ' FullName = "G, M C"
    ' Giving me different result than Excel version used in the lab
    
    'Outputs array. Zero index based.
    SplitName = Split(FullName, " ")
    
    length = UBound(SplitName) - LBound(SplitName)
    For index = 0 To length
        If length = 1 Then
            ' Two character initials
            ' Ex. Oshima, Keith -> OK
            FirstLetters(index) = Left(SplitName(index), 1)
        ElseIf length > 1 And index = 0 Or index = length Then
            ' Longer character initials. Ignore middle initial and take first and last.
            ' Ex. Morales, Paul John -> JM
            FirstLetters(index) = Left(SplitName(index), 1)
        End If
    Next index
    
    'Reverse initials with this excel version.
    If rev = 1 Then
        Initials = StrReverse(Join(FirstLetters, ""))
    Else
        Initials = Join(FirstLetters, "")
    End If
    
End Function
