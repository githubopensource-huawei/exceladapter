Attribute VB_Name = "TransPortSub"
'��Base Station Transport Data��ҳ��¼��ʼ��
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
Private Const LRESRadioMoc As String = "eNodeBEqmFunction"
Private Const MRadioAttr As String = "RadioTemplateName"
Private Const MRadioMoc As String = "NBBSFunction"

Private Const gsmRadioType As String = "GSM Radio Template"
Private Const umtsRadioType As String = "UMTS Radio Template"
Private Const lteRadioType As String = "LTE Radio Template"
Private Const nbiotRadioType As String = "NB-IoT Radio Template"




'���塸*Site Type������Cabinet Type������*Site Template���������¼�
Public Sub TransPortSheetChange(ByVal sh As Object, ByVal target As range)
    Dim m_Str As String
    '��Base Station Transport Data��ҳ��*Site Type��������
    constTypeCol = getColNum(sh.name, constRecordRow, productTypeAttr, productTypeMoc)
    
    '��Base Station Transport Data��ҳ��*Site Template��������
    constPattenCol = getColNum(sh.name, constRecordRow, siteTemplateAttr, siteTemplateMoc)
    
    '��*Site Type��ֵ���ʱ���������á�Site Template���ĺ�ѡֵ�������ǰֵ��
     If target.column = constTypeCol And target.row > constRecordRow And target.count = 1 Then
        If constTypeCol <> -1 And constPattenCol <> -1 Then
            '��ȡ��*Site Template���к�ѡֵ
            m_Str_Template = Get_Template_Related(target.value, sh, target.Offset(0, constPattenCol - constTypeCol))
            If m_Str_Template <> "" Then
                With target.Offset(0, constPattenCol - constTypeCol).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Str_Template
                End With
                If Not target.Offset(0, constPattenCol - constTypeCol).Validation.value Then
                    target.Offset(0, constPattenCol - constTypeCol).value = ""
                End If
            Else
                With target.Offset(0, constPattenCol - constTypeCol).Validation
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
                    target.Offset(0, constPattenCol - constTypeCol).value = ""
            End If
        End If
         
    End If

End Sub
    
'�������á�*Site Type���к͡�*Site Template���������б���¼�
Public Sub TransPortSheetSelectionChange(ByVal sh As Object, ByVal target As range)
    Dim m_Str_Cabinet As String
    Dim m_Str_Template As String
    Debug.Print Now
    '��Base Station Transport Data��ҳ��*Site Type��������
    constTypeCol = getColNum(sh.name, constRecordRow, productTypeAttr, productTypeMoc)
    
    '��Base Station Transport Data��ҳ��*Site Template��������
    constPattenCol = getColNum(sh.name, constRecordRow, siteTemplateAttr, siteTemplateMoc)

    '��Base Station Transport Data��ҳ��*GBTS Radio Template��������
    constGsmRadionPattenCol = getColNum(sh.name, constRecordRow, GRadioAttr, GRadioMoc)
    
    '��Base Station Transport Data��ҳ��*Umts Radio Template��������
    constUmtsRadionPattenCol = getColNum(sh.name, constRecordRow, URadioAttr, URadioMoc)
    
    '��Base Station Transport Data��ҳ��*LTE Radio Template��������
    constLteRadionPattenCol = getColNum(sh.name, constRecordRow, LRadioAttr, LRadioMoc)
    
    If constLteRadionPattenCol < 1 Then
        constLteRadionPattenCol = getColNum(sh.name, constRecordRow, LRadioAttr, LRESRadioMoc)
    End If
    
     '��Base Station Transport Data��ҳ��*NBIOT Radio Template��������
    constNbiotRadionPattenCol = getColNum(sh.name, constRecordRow, MRadioAttr, MRadioMoc)
    Debug.Print Now
    '��ȡ��*Site Type���к�ѡֵ�����趨Ϊ������
    If target.column = constTypeCol And target.count = 1 And target.row > constRecordRow Then
        '��ȡ��*Site Type���к�ѡֵ
        m_Str_Cabinet = GetSiteType(sh, target)
        If m_Str_Cabinet <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Str_Cabinet
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            With target.Validation
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
                target.value = ""
        End If
    '��ȡ��*Site Template���к�ѡֵ�����趨Ϊ�����б�
    ElseIf target.column = constPattenCol And target.count = 1 And target.row > constRecordRow Then
        If constTypeCol <> -1 Then
            If Cells(target.row, constTypeCol).value <> "" Then
                '�ӡ�MappingSiteTemplate��ҳ��ȡ��*Site Template���к�ѡֵ
                m_Str_Template = Get_Template_Related(target.Offset(0, constTypeCol - constPattenCol).value, sh, target)
                If m_Str_Template <> "" Then
                    With target.Validation
                           .Delete
                           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Str_Template
                    End With
                    If Not target.Validation.value Then
                        target.value = ""
                    End If
                Else
                    With target.Validation
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
                        target.value = ""
                End If
            Else
                With target.Validation
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
                    target.value = ""
            End If
        End If
    ElseIf target.column = constGsmRadionPattenCol And target.count = 1 And target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(getResByKey(gsmRadioType), sh, target)
        If m_Str_Radio <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Str_Radio
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            With target.Validation
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
                target.value = ""
        End If
    ElseIf target.column = constUmtsRadionPattenCol And target.count = 1 And target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(getResByKey(umtsRadioType), sh, target)
        If m_Str_Radio <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Str_Radio
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            With target.Validation
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
                target.value = ""
        End If
    ElseIf target.column = constLteRadionPattenCol And target.count = 1 And target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(getResByKey(lteRadioType), sh, target)
        If m_Str_Radio <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Str_Radio
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            With target.Validation
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
                target.value = ""
        End If
     ElseIf target.column = constNbiotRadionPattenCol And target.count = 1 And target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(getResByKey(nbiotRadioType), sh, target)
        If m_Str_Radio <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Str_Radio
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            With target.Validation
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
                target.value = ""
        End If
    End If
End Sub

'�ӡ�MappingSiteTypeCabinetType��ҳ��ȡ��*Site Type���к�ѡֵ
Function GetSiteType(sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim neType As String
    
    neType = getNeType()
    m_Str = ""
    For m_rowNum = 2 To Worksheets("ProductType").range("a1048576").End(xlUp).row
            If Worksheets("ProductType").Cells(m_rowNum, 2) = neType Then
                If m_Str = "" Then
                     m_Str = Worksheets("ProductType").Cells(m_rowNum, 1).value
                Else
                    m_Str = m_Str & "," & Worksheets("ProductType").Cells(m_rowNum, 1).value
                End If
            End If
    Next
    If Len(m_Str) > 255 Then
        Dim groupName As String
        Dim columnName As String
        Dim valideDef As CValideDef
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        Set valideDef = initDefaultDataSub.getInnerValideDef(sheet.name + "," + groupName + "," + columnName)
        If valideDef Is Nothing Then
            Set valideDef = addInnerValideDef(sheet.name, groupName, columnName, m_Str)
        Else
            Call modiflyInnerValideDef(sheet.name, groupName, columnName, m_Str, valideDef)
        End If
        m_Str = valideDef.getValidedef
    End If
    GetSiteType = m_Str
End Function

Function getRadioTemplate(RadioType As String, sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_RowNum_Inner As Long
    Dim flag As Boolean
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    Dim neType As String
    neType = getNeType()
    m_Str = ""
      For m_rowNum = 2 To Worksheets("MappingRadioTemplate").range("a1048576").End(xlUp).row
        If (RadioType = Worksheets("MappingRadioTemplate").Cells(m_rowNum, 2).value Or Len(Trim(Worksheets("MappingRadioTemplate").Cells(m_rowNum, 2).value)) = 0) _
        And neType = Worksheets("MappingRadioTemplate").Cells(m_rowNum, 3).value Then
            If m_Str = "" Then
                m_Str = Worksheets("MappingRadioTemplate").Cells(m_rowNum, 1).value
            ElseIf VBA.Trim(Worksheets("MappingRadioTemplate").Cells(m_rowNum, 1).value) <> "" Then
                m_Str = m_Str & "," & Worksheets("MappingRadioTemplate").Cells(m_rowNum, 1).value
            End If
        End If
    Next
    
    If Len(m_Str) > 255 Then
        Dim groupName As String
        Dim columnName As String
        Dim valideDef As CValideDef
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        Set valideDef = initDefaultDataSub.getInnerValideDef(sheet.name + "," + groupName + "," + columnName)
        If valideDef Is Nothing Then
            Set valideDef = addInnerValideDef(sheet.name, groupName, columnName, m_Str)
        Else
            Call modiflyInnerValideDef(sheet.name, groupName, columnName, m_Str, valideDef)
        End If
        m_Str = valideDef.getValidedef
    End If
    getRadioTemplate = m_Str
End Function


'�ӡ�MappingSiteTemplate��ҳ��ȡ��*Site Template���к�ѡֵ
Public Function Get_Template_Related(SiteType As String, sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    Dim neType As String
    
    neType = getNeType()
    
    m_start = 0
    m_Str = ""
    For m_rowNum = 2 To Worksheets("MappingSiteTemplate").range("a1048576").End(xlUp).row
        If SiteType = Worksheets("MappingSiteTemplate").Cells(m_rowNum, 1).value _
        And neType = Worksheets("MappingSiteTemplate").Cells(m_rowNum, 5).value Then
            If m_Str = "" Then
                m_Str = Worksheets("MappingSiteTemplate").Cells(m_rowNum, 4).value
            ElseIf VBA.Trim(Worksheets("MappingSiteTemplate").Cells(m_rowNum, 4).value) <> "" Then
                m_Str = m_Str & "," & Worksheets("MappingSiteTemplate").Cells(m_rowNum, 4).value
            End If
        End If
    Next
    
    If Len(m_Str) > 255 Then
        Dim groupName As String
        Dim columnName As String
        Dim valideDef As CValideDef
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        Set valideDef = initDefaultDataSub.getInnerValideDef(sheet.name + "," + groupName + "," + columnName)
        If valideDef Is Nothing Then
            Set valideDef = addInnerValideDef(sheet.name, groupName, columnName, m_Str)
        Else
            Call modiflyInnerValideDef(sheet.name, groupName, columnName, m_Str, valideDef)
        End If
        m_Str = valideDef.getValidedef
    End If
    
    Get_Template_Related = m_Str

End Function

'��ָ��sheetҳ��ָ���У�����ָ���У������к�
Private Function Get_Col(sheetName As String, recordRow As Long, ColValue As String) As Long
    Dim m_colNum As Long

    Get_Col = -1
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        If ColValue = ws.Cells(recordRow, m_colNum).value Then
            Get_Col = m_colNum
            Exit For
        End If
    Next
End Function










