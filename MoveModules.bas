Attribute VB_Name = "MoveModules"
' Subroutine for copying vba module from one wb to another
' Answer from Sean Hare on StackOverflow: https://stackoverflow.com/questions/40956465/vba-to-copy-module-from-one-excel-workbook-to-another-workbook

Sub CopyModules(wbSource As Workbook, wbTarget As Workbook)
   Dim vbcompSource As VBComponent, vbcompTarget As VBComponent
   Dim sText As String, nType As Long
   For Each vbcompSource In wbSource.VBProject.VBComponents
      nType = vbcompSource.Type
      If nType < 100 Then  '100=vbext_ct_Document -- the only module type we would not want to copy
         Set vbcompTarget = wbTarget.VBProject.VBComponents.Add(nType)
         sText = vbcompSource.CodeModule.Lines(1, vbcompSource.CodeModule.CountOfLines)
         vbcompTarget.CodeModule.AddFromString (sText)
         vbcompTarget.Name = vbcompSource.Name
      End If
   Next vbcompSource
End Sub

