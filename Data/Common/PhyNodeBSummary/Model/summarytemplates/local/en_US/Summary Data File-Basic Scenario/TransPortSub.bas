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

Private Const NRadioMoc As String = "gNodeBFunction"
Private Const NRadioAttr As String = "RadioTemplateName"

Private Const gsmRadioType As String = "GSM Radio Template"
Private Const umtsRadioType As String = "UMTS Radio Template"
Private Const lteRadioType As String = "LTE Radio Template"
Private Const lresRadioType As String = "LRES Radio Template"
Private Const nbiotRadioType As String = "NB-IoT Radio Template"

Private Const nrRadioType As String = "NR Radio Template"


'���塸*Site Type������Cabinet Type������*Site Template���������¼�
Public Sub TransPortSheetChange(ByVal sh As Object, ByVal Target As range)
    Dim m_Str As String
    '��Base Station Transport Data��ҳ��*Site Type��������
    constTypeCol = getColNum(sh.name, constRecordRow, productTypeAttr, productTypeMoc)
    
    '��Base Station Transport Data��ҳ��*Site Template��������
    constPattenCol = getColNum(sh.name, constRecordRow, siteTemplateAttr, siteTemplateMoc)
    
    '��*Site Type��ֵ���ʱ���������á�Site Template���ĺ�ѡֵ�������ǰֵ��
     If Target.column = constTypeCol And Target.row > constRecordRow And Target.count = 1 Then
        If constTypeCol <> -1 And constPattenCol <> -1 Then
            '��ȡ��*Site Template���к�ѡֵ
            m_Str_Template = Get_Template_Related(Target.value, sh, Target.Offset(0, constPattenCol - constTypeCol))
            If m_Str_Template <> "" Then
                With Target.Offset(0, constPattenCol - constTypeCol).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Template
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
                    .inputmessage = ""
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
    
'�������á�*Site Type���к͡�*Site Template���������б���¼�
Public Sub TransPortSheetSelectionChange(ByVal sh As Object, ByVal Target As range)
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
    
    constNRRadionPattenCol = getColNum(sh.name, constRecordRow, NRadioAttr, NRadioMoc)
    
    Debug.Print Now
    '��ȡ��*Site Type���к�ѡֵ�����趨Ϊ������
    If Target.column = constTypeCol And Target.count = 1 And Target.row > constRecordRow Then
        '��ȡ��*Site Type���к�ѡֵ
        m_Str_Cabinet = GetSiteType(sh, Target)
        If m_Str_Cabinet <> "" Then
            With Target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Cabinet
            End With
            If Not Target.Validation.value Then
                Target.value = ""
            End If
        Else
            With Target.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .inputTitle = ""
                .ErrorTitle = ""
                .inputmessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
                Target.value = ""
        End If
    '��ȡ��*Site Template���к�ѡֵ�����趨Ϊ�����б�
    ElseIf Target.column = constPattenCol And Target.count = 1 And Target.row > constRecordRow Then
        If constTypeCol <> -1 Then
            If Cells(Target.row, constTypeCol).value <> "" Then
                '�ӡ�MappingSiteTemplate��ҳ��ȡ��*Site Template���к�ѡֵ
                m_Str_Template = Get_Template_Related(Target.Offset(0, constTypeCol - constPattenCol).value, sh, Target)
                If m_Str_Template <> "" Then
                    With Target.Validation
                           .Delete
                           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Template
                    End With
                    If Not Target.Validation.value Then
                        Target.value = ""
                    End If
                Else
                    With Target.Validation
                        .Delete
                        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .inputTitle = ""
                        .ErrorTitle = ""
                        .inputmessage = ""
                        .ErrorMessage = ""
                        .IMEMode = xlIMEModeNoControl
                        .ShowInput = True
                        .ShowError = True
                    End With
                        Target.value = ""
                End If
            Else
                With Target.Validation
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .inputTitle = ""
                    .ErrorTitle = ""
                    .inputmessage = ""
                    .ErrorMessage = ""
                    .IMEMode = xlIMEModeNoControl
                    .ShowInput = True
                    .ShowError = True
                End With
                    Target.value = ""
            End If
        End If
    ElseIf Target.column = constGsmRadionPattenCol And Target.count = 1 And Target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(getResByKey(gsmRadioType), sh, Target)
        If m_Str_Radio <> "" Then
            With Target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Radio
            End With
            If Not Target.Validation.value Then
                Target.value = ""
            End If
        Else
            With Target.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .inputTitle = ""
                .ErrorTitle = ""
                .inputmessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
                Target.value = ""
        End If
    ElseIf Target.column = constUmtsRadionPattenCol And Target.count = 1 And Target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(getResByKey(umtsRadioType), sh, Target)
        If m_Str_Radio <> "" Then
            With Target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Radio
            End With
            If Not Target.Validation.value Then
                Target.value = ""
            End If
        Else
            With Target.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .inputTitle = ""
                .ErrorTitle = ""
                .inputmessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
                Target.value = ""
        End If
    ElseIf Target.column = constLteRadionPattenCol And Target.count = 1 And Target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(getResByKey(lteRadioType), sh, Target)
        If Len(m_Str_Radio) = 0 Then
            m_Str_Radio = getRadioTemplate(getResByKey(lresRadioType), sh, Target)
        End If
        If m_Str_Radio <> "" Then
            With Target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Radio
            End With
            If Not Target.Validation.value Then
                Target.value = ""
            End If
        Else
            With Target.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .inputTitle = ""
                .ErrorTitle = ""
                .inputmessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
                Target.value = ""
        End If
     ElseIf Target.column = constNbiotRadionPattenCol And Target.count = 1 And Target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(getResByKey(nbiotRadioType), sh, Target)
        If m_Str_Radio <> "" Then
            With Target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Radio
            End With
            If Not Target.Validation.value Then
                Target.value = ""
            End If
        Else
            With Target.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .inputTitle = ""
                .ErrorTitle = ""
                .inputmessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
                Target.value = ""
        End If
     ElseIf Target.column = constNRRadionPattenCol And Target.count = 1 And Target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(getResByKey(nrRadioType), sh, Target)
        If m_Str_Radio <> "" Then
            With Target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Radio
            End With
            If Not Target.Validation.value Then
                Target.value = ""
            End If
        Else
            With Target.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .inputTitle = ""
                .ErrorTitle = ""
                .inputmessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
                Target.value = ""
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
    For m_rowNum = 2 To Worksheets("ProductType").range("a65536").End(xlUp).row
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
      For m_rowNum = 2 To Worksheets("MappingRadioTemplate").range("a65536").End(xlUp).row
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
    For m_rowNum = 2 To Worksheets("MappingSiteTemplate").range("a65536").End(xlUp).row
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
    For m_colNum = 1 To ws.range("IV2").End(xlToLeft).column
        If ColValue = ws.Cells(recordRow, m_colNum).value Then
            Get_Col = m_colNum
            Exit For
        End If
    Next
End Function

'delete illegal sites in transport data sheet defined in config file from GUI Java
Public Sub deleteIllegalSites()
    Dim illegalSites() As String
    Dim msgInfo As String
    Dim siteArray() As String
    Dim tmpArray() As String
    Call changeAlerts(False)
    
    msgInfo = readUTF8File(ThisWorkbook.Path + "\Parameter.ini")
    
    tmpArray = Split(msgInfo, "=")
    If tmpArray(0) <> "NeedDelSites" Then
        Exit Sub
    End If
    
    illegalSites = Split(tmpArray(1), ",")
    If UBound(illegalSites) - LBound(illegalSites) = 0 Then
        Exit Sub
    End If
    
    Application.EnableEvents = False
    
    Call deleteIllegalSites_i(illegalSites)
    
    Call changeAlerts(True)
End Sub


Public Sub deleteIllegalSites_i(ByRef illegalSites() As String)
    Dim transportSheet As Worksheet
    Set transportSheet = ThisWorkbook.Worksheets(GetMainSheetName())
    
    Dim rowIdx As Integer
    Dim idx As Integer
    Dim siteName As String
    Const dataBeginRow As Integer = 4
    
    For idx = 0 To UBound(illegalSites)
        For rowIdx = transportSheet.range("a65535").End(xlUp).row To dataBeginRow
            If CStr(transportSheet.Cells(rowIdx, 1).value) = illegalSites(idx) Then
                transportSheet.rows(rowIdx).Delete
                Exit For
            End If
        Next
    Next

End Sub


