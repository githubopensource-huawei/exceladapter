Attribute VB_Name = "TransPortSub"
'「Base Station Transport Data」页记录起始行
Private Const constRecordRow = 2
Private Const productTypeAttr As String = "PRODUCTTYPE"
Private Const productTypeMoc As String = "Node"
Private Const siteTemplateAttr As String = "SiteTemplateName"
Private Const siteTemplateMoc As String = "Node"
Private Const GRadioAttr As String = "RadioTemplateName"
Private Const GRadioMoc As String = "GbtsFunction"
Private Const URadioAttr As String = "RadioTemplateName"
Private Const URadioMoc As String = "NodeBFunction"
Private Const LRadioAttr As String = "RadioTemplateName"
Private Const LRadioMoc As String = "eNodeBFunction"

Private Const gsmRadioType As String = "GSM Radio Template"
Private Const umtsRadioType As String = "UMTS Radio Template"
Private Const lteRadioType As String = "LTE Radio Template"




'定义「*Site Type」→「Cabinet Type」→「*Site Template」的联动事件
Public Sub TransPortSheetChange(ByVal sh As Object, ByVal Target As Range)
    Dim m_Str As String
    '「Base Station Transport Data」页「*Site Type」所在列
    constTypeCol = getColNum(sh.name, constRecordRow, productTypeAttr, productTypeMoc)
    
    '「Base Station Transport Data」页「*Site Template」所在列
    constPattenCol = getColNum(sh.name, constRecordRow, siteTemplateAttr, siteTemplateMoc)
    
    '「*Site Type」值变更时，重新设置「Site Template」的侯选值并清除当前值。
     If Target.column = constTypeCol And Target.row > constRecordRow And Target.count = 1 Then
        If constTypeCol <> -1 And constPattenCol <> -1 Then
            '获取「*Site Template」列侯选值
            m_Str_Template = Get_Template_Related(Target.value)
            If m_Str_Template <> "" Then
                With Target.Offset(0, constPattenCol - constTypeCol).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Str_Template
                End With
                If Not Target.Offset(0, constPattenCol - constTypeCol).Validation.value Then
                    Target.Offset(0, constPattenCol - constTypeCol).value = ""
                End If
            Else
                With Target.Offset(0, constPattenCol - constTypeCol).Validation
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .inputTitle = ""
                    .ErrorTitle = ""
                    .inputMessage = ""
                    .ErrorMessage = ""
                    .IMEMode = xlIMEModeNoControl
                    .ShowInput = True
                    .ShowError = True
                End With
                    Target.Offset(0, constPattenCol - constTypeCol).value = ""
            End If
        End If
         
    End If

End Sub

'从「MappingSiteTemplate」页获取「*Site Template」列侯选值
Public Function Get_Template_Related(SiteType As String) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    Dim neType As String
    
    neType = getNeType()
    
    m_start = 0
    m_Str = ""
    For m_rowNum = 2 To MappingSiteTemplate.Range("a1048576").End(xlUp).row
        If SiteType = MappingSiteTemplate.Cells(m_rowNum, 1).value _
        And neType = MappingSiteTemplate.Cells(m_rowNum, 5).value Then
            If m_Str = "" Then
                m_Str = MappingSiteTemplate.Cells(m_rowNum, 4).value
            Else
                m_Str = m_Str & "," & MappingSiteTemplate.Cells(m_rowNum, 4).value
            End If
        End If
    Next
    
    'If m_start = 0 Then
        'm_Str = ""
    'Else
        'm_Str = "=INDIRECT(""MappingSiteTemplate!D" & CStr(m_start) & ":D" & CStr(m_end) & """)"
    'End If
    
    Get_Template_Related = m_Str

End Function

'从指定sheet页的指定行，查找指定列，返回列号
Private Function Get_Col(sheetName As String, RecordRow As Long, ColValue As String) As Long
    Dim m_ColNum As Long

    Get_Col = -1
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For m_ColNum = 1 To ws.Range("XFD2").End(xlToLeft).column
        If ColValue = ws.Cells(RecordRow, m_ColNum).value Then
            Get_Col = m_ColNum
            Exit For
        End If
    Next
End Function






