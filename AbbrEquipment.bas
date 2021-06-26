Attribute VB_Name = "AbbrEquipment"
Public Function Abbr_Equipment(iEquipNum As Range) As String
    ' Returns abbreviated equipment name along with number given
    ' Equipment name = worksheet name.
    ' If running from vba, activate wksht first.
    Dim sSheetName
    
    'Need to get cell ref range because ActiveSheet.Name will apply result across sheets.
    sSheetName = iEquipNum.Worksheet.Name
    
    With Application
        Select Case sSheetName
            Case "Incubators"
                Abbr_Equipment = "INC #" & CStr(iEquipNum.Value)
            Case "Refrigerators"
                Abbr_Equipment = "REF #" & CStr(iEquipNum.Value)
            Case "Freezers"
                Abbr_Equipment = "FRZ #" & CStr(iEquipNum.Value)
            Case "Waterbaths"
                Abbr_Equipment = "WTB #" & CStr(iEquipNum.Value)
            Case "Balances"
                Abbr_Equipment = "BLC #" & CStr(iEquipNum.Value)
            Case "Hotplates"
                Abbr_Equipment = "HP #" & CStr(iEquipNum.Value)
            Case "Vortexers"
                Abbr_Equipment = "VTX #" & CStr(iEquipNum.Value)
            Case "Heating Blocks"
                Abbr_Equipment = "HB #" & CStr(iEquipNum.Value)
            Case Else
                Abbr_Equipment = "None"
        End Select
    End With

End Function
