Attribute VB_Name = "AdjustNotifications"
Sub Adjust_Notif(bSetting As Boolean)
    With Application
        .ScreenUpdating = bSetting
        .DisplayAlerts = bSetting
        .EnableEvents = bSetting
    End With
End Sub
