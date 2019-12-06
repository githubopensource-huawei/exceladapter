Attribute VB_Name = "pubModule"
Public Function GetSheetsPass() As String
    GetSheetsPass = "XCT100"
End Function

Public Sub UnprotectWorkSheet(curSheet As Worksheet)
    On Error Resume Next
    curSheet.Unprotect (GetSheetsPass)
    Application.ScreenUpdating = True
End Sub

Public Sub ProtectWorkSheet(curSheet As Worksheet)
    On Error Resume Next
    curSheet.Protect Password:=GetSheetsPass, AllowFormattingCells:=True, AllowFormattingColumns:=True
    Application.ScreenUpdating = True
End Sub

Public Sub ProtectWorkBook()
    On Error Resume Next
    'ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
    'ThisWorkbook.Save
End Sub

Public Sub UnprotectWorkBook()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=GetSheetsPass
    'ThisWorkbook.Save
End Sub

Sub SetSysOption()
    With ActiveWindow
        If .DisplayZeros Then Exit Sub
        .DisplayGridlines = False
        .DisplayZeros = True
    End With
End Sub

