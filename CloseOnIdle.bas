Attribute VB_Name = "CloseOnIdle"
'Must be in global scope to get replaced with wb open or selection change
Dim iStartTime As Long
Dim iRunTime As Variant

Sub Start_Timer()
    ' Setup start time.
    iStartTime = Timer
    iRunTime = Now + TimeValue("00:10:00")
    ' When time is reached, run Close_Save
    Application.OnTime EarliestTime:=iRunTime, Procedure:="Close_Save"
End Sub

Sub Stop_Timer()
    Application.OnTime EarliestTime:=iRunTime, Procedure:="Close_Save", Schedule:=False
End Sub

Sub Close_Save()
    If Timer - iStartTime > 590 Then
        ThisWorkbook.Close SaveChanges:=True
    End If
End Sub
