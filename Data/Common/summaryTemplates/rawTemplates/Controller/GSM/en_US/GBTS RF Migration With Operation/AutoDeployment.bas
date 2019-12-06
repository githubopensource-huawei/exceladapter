Attribute VB_Name = "AutoDeployment"
'即插即用数据特殊处理
'用以设置颜色
Const SolidColorIdx = 16
Const SolidPattern = xlGray16
Const NullPattern = xlNone

Public Sub AutoDeploySheetChange(ByVal sheet As Object, ByVal Target As range)
    Dim connTypeCol As Long
    Dim authenticationTypeCol As Long
    connTypeCol = 8
    authenticationTypeCol = 9
    
    If (Target.Interior.colorIndex = SolidColorIdx) And Target.value <> "" Then
        Target.value = ""
        MsgBox getResByKey("NoInput")
        Exit Sub
    End If
    
    If Target.row > 2 And Target.column = connTypeCol And Target.value = getResByKey("commConn") Then
                sheet.Cells(Target.row, authenticationTypeCol).Interior.colorIndex = SolidColorIdx
                sheet.Cells(Target.row, authenticationTypeCol).Interior.Pattern = SolidPattern
                sheet.Cells(Target.row, authenticationTypeCol).value = ""
                sheet.Cells(Target.row, authenticationTypeCol).Validation.ShowInput = False
    ElseIf Target.row > 2 And Target.column = connTypeCol And Target.value = getResByKey("sslConn") Then
                sheet.Cells(Target.row, authenticationTypeCol).Interior.colorIndex = NullPattern
                sheet.Cells(Target.row, authenticationTypeCol).Interior.Pattern = NullPattern
                sheet.Cells(Target.row, authenticationTypeCol).Validation.ShowInput = True
    End If
    
End Sub

