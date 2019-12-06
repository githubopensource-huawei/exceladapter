Attribute VB_Name = "templateMgr"
Public Const CONST_LAN_EN = 1
Public Const CONST_LAN_ZH = 2


Public Sub CustomizeXls(iLan As Integer)
    If CONST_LAN_ZH = iLan Then
        TemplateCMForm.Caption = "定制模板"
        TemplateCMForm.ToolFrame.Caption = "汇总"
    ElseIf CONST_LAN_EN = iLan Then
        TemplateCMForm.Caption = "Customize template"
        TemplateCMForm.ToolFrame.Caption = "Summary"
    End If
    
    TemplateCMForm.cbATDM.value = ThisWorkbook.Sheets("category").Cells(1, 2).value
    TemplateCMForm.cbAIP.value = ThisWorkbook.Sheets("category").Cells(2, 2).value
    TemplateCMForm.cbATDMIP.value = ThisWorkbook.Sheets("category").Cells(3, 2).value
    TemplateCMForm.cbATERTDM.value = ThisWorkbook.Sheets("category").Cells(4, 2).value
    TemplateCMForm.cbATERIP.value = ThisWorkbook.Sheets("category").Cells(5, 2).value
    TemplateCMForm.cbGbFR.value = ThisWorkbook.Sheets("category").Cells(6, 2).value
    TemplateCMForm.cbGbIP.value = ThisWorkbook.Sheets("category").Cells(7, 2).value

    TemplateCMForm.Show vbModeless
End Sub
