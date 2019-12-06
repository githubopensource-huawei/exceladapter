Attribute VB_Name = "TransPortSub"
'��Base Station Transport Data��ҳ��¼��ʼ��
Private Const constRecordRow = 2
Private Const productTypeAttr As String = "PRODUCTTYPE"
Private Const productTypeMoc As String = "Node"
Private Const siteTemplateAttr As String = "SiteTemplateName"
Private Const siteTemplateMoc As String = "Node"
Private Const LRadioAttr As String = "RadioTemplateName"
Private Const LRadioMoc As String = "eNodeBFunction"
Private Const LRESRadioMoc As String = "eNodeBEqmFunction"
Private Const lteRadioType As String = "LTE Radio Template"

Private Function getFddTddTypeString(ByRef mappingSiteTemplateSheet As Worksheet, ByRef SiteType As String) As String
    Dim maxRowNumber As Long, rowIndex As Long
    maxRowNumber = mappingSiteTemplateSheet.range("A65535").End(xlUp).row
    
    Dim fddTddTypeCol As New Collection
    Dim fddTddValue As Variant
    '���������SiteType�õ�֧�ֵ�FddTdd��������
    For rowIndex = 2 To maxRowNumber
        With mappingSiteTemplateSheet
            If .range("A" & rowIndex).value = SiteType Then
                fddTddValue = .range("C" & rowIndex).value
                If fddTddValue <> "" And Not Contains(fddTddTypeCol, CStr(fddTddValue)) Then
                    fddTddTypeCol.Add item:=fddTddValue, key:=fddTddValue
                End If
            End If
        End With
    Next rowIndex
    
    '�õ�fddTddTypeString����Ч���ַ���
    Dim fddTddTypeString As String
    For Each fddTddValue In fddTddTypeCol
        If fddTddTypeString = "" Then
            fddTddTypeString = fddTddValue
        Else
            fddTddTypeString = fddTddTypeString & "," & fddTddValue
        End If
    Next fddTddValue
    
    getFddTddTypeString = fddTddTypeString
End Function



'���塸*Site Type������Cabinet Type������*Site Template���������¼�
Public Sub TransPortSheetChange(ByVal sh As Object, ByVal target As range)
    Dim m_Str As String
    '��Base Station Transport Data��ҳ��*Site Type��������
    constTypeCol = getColNum(sh.name, constRecordRow, productTypeAttr, productTypeMoc)
    
    
    '��Base Station Transport Data��ҳ��*FDD/TDD Mode��������
    Dim baseStationVersion As String
    baseStationVersion = UCase(getBaseStationVersion)
    constFddTddCol = Get_Col(sh.name, constRecordRow, getResByKey("FDD/TDD Mode"))
    
    If constFddTddCol < 1 Then
    constFddTddCol = Get_Col(sh.name, constRecordRow, getResByKey("Mode"))
    End If
    '��Base Station Transport Data��ҳ��Cabinet Type��������
    constVersionCol = Get_Col(sh.name, constRecordRow, getResByKey("Cabinet Type"))
    
    '��Base Station Transport Data��ҳ��*Site Template��������
    constPattenCol = getColNum(sh.name, constRecordRow, siteTemplateAttr, siteTemplateMoc)
    
    '��*Site Type��ֵ���ʱ���������á�Site Template���ĺ�ѡֵ�������ǰֵ��
     If target.column = constTypeCol And target.row > constRecordRow And target.count = 1 Then
        If constVersionCol <> -1 Then
            '��ȡ��Cabinet Type���к�ѡֵ
            m_Str = Get_Site_Cabinet_Related(target.value, sh, target.Offset(0, constVersionCol - constTypeCol))
            If m_Str <> "" Then
                With target.Offset(0, constVersionCol - constTypeCol).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str
                End With
                If Not target.Offset(0, constPattenCol - constTypeCol).Validation.value Then
                    target.Offset(0, constVersionCol - constTypeCol).value = ""
                End If
            Else
                With target.Offset(0, constVersionCol - constTypeCol).Validation
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
                    target.Offset(0, constVersionCol - constTypeCol).value = ""
            End If
        End If
        If constFddTddCol <> -1 Then
            'DBS3900_LTE֧��FDD/TDD/FDDTDD������վ��ֻ֧��FDD�����˵���Ǿɵģ�����3205E��Сվ�Ѿ��޷�֧�֣���Ҫ����SiteType������Ӧ��FDD/TDDģʽ
            
            '����վ�͵õ�֧�ֵ�FddTdd�����ַ���
            Dim fddTddTypeString As String
            Dim mappingSiteTemplateSht As Worksheet
            Set mappingSiteTemplateSht = ThisWorkbook.Worksheets("MappingSiteTemplate")
            fddTddTypeString = getFddTddTypeString(mappingSiteTemplateSht, target.value)
            
'            If Target.value = "DBS3900_LTE" Then
'                With Target.Offset(0, constFDDTDDCol - constTypeCol).Validation
'                    .Delete
'                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="FDD,TDD,FDDTDD"
'                End With
'                If Not Target.Offset(0, constFDDTDDCol - constTypeCol).Validation.value Then
'                    Target.Offset(0, constFDDTDDCol - constTypeCol).value = ""
'                End If
            If target.value = "" Then
                With target.Offset(0, constFddTddCol - constTypeCol).Validation
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
                    target.Offset(0, constFddTddCol - constTypeCol).value = ""
            Else
                With target.Offset(0, constFddTddCol - constTypeCol).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=fddTddTypeString
                End With
                If Not target.Offset(0, constFddTddCol - constTypeCol).Validation.value Then
                    target.Offset(0, constFddTddCol - constTypeCol).value = ""
                End If
            End If
        End If
'��Cabinet Type��ֵ���ʱ���������á�*Site Template���ĺ�ѡֵ�������ǰֵ��
    ElseIf target.column = constVersionCol And target.row > constRecordRow And target.count = 1 Then
        If constTypeCol <> -1 And constFddTddCol <> -1 And constPattenCol <> -1 Then
            '��ȡ��*Site Template���к�ѡֵ
            m_Str_Template = Get_Template_Related(target.Offset(0, constTypeCol - constVersionCol).value, target.Offset(0, constFddTddCol - constVersionCol).value, target.value, sh, target.Offset(0, constPattenCol - constVersionCol))
            If m_Str_Template <> "" Then
                With target.Offset(0, constPattenCol - constVersionCol).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Template
                End With
                If Not target.Offset(0, constPattenCol - constVersionCol).Validation.value Then
                    target.Offset(0, constPattenCol - constVersionCol).value = ""
                End If
            Else
                With target.Offset(0, constPattenCol - constVersionCol).Validation
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
                    target.Offset(0, constPattenCol - constVersionCol).value = ""
            End If
        End If
    '��*FDD/TDD Mode��ֵ���ʱ���������á�*Site Template���͡�Radio Template���ĺ�ѡֵ�������ǰֵ��
    ElseIf target.column = constFddTddCol And target.row > constRecordRow And target.count = 1 Then
        If constTypeCol <> -1 And constVersionCol <> -1 And constPattenCol <> -1 Then
            '��ȡ��*Site Template���к�ѡֵ
            m_Str_Template = Get_Template_Related(target.Offset(0, constTypeCol - constFddTddCol).value, target.value, target.Offset(0, constVersionCol - constFddTddCol).value, sh, target.Offset(0, constPattenCol - constFddTddCol))
            If m_Str_Template <> "" Then
                With target.Offset(0, constPattenCol - constFddTddCol).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Template
                End With
                If Not target.Offset(0, constPattenCol - constFddTddCol).Validation.value Then
                    target.Offset(0, constPattenCol - constFddTddCol).value = ""
                End If
            Else
                With target.Offset(0, constPattenCol - constFddTddCol).Validation
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
                    target.Offset(0, constPattenCol - constFddTddCol).value = ""
            End If
        End If
        
        If constRadioCol <> -1 Then
             '��ȡ��*Radio Template���к�ѡֵ
            m_Str_RadioTemplate = getRadioTemplate(target.value, sh, target.Offset(0, constRadioCol - constFddTddCol))
            If m_Str_RadioTemplate <> "" Then
                With target.Offset(0, constRadioCol - constFddTddCol).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_RadioTemplate
                End With
                If Not target.Offset(0, constRadioCol - constFddTddCol).Validation.value Then
                    target.Offset(0, constRadioCol - constFddTddCol).value = ""
                End If
            Else
                With target.Offset(0, constRadioCol - constFddTddCol).Validation
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
                    target.Offset(0, constRadioCol - constFddTddCol).value = ""
            End If
        End If
        
    End If
End Sub

Private Function targetHasFormula1(ByRef target As range) As Boolean
    On Error GoTo ErrorHandler
    targetHasFormula1 = True
    If target.Validation Is Nothing Then 'û����Ч�ԣ���û��formula1
        targetHasFormula1 = False
        Exit Function
    End If
    
    Dim formula1 As String
    formula1 = target.Validation.formula1 '�����formula1����ֵ�ɹ������û�У���ֵ��������ErrorHandler
    If formula1 = "" Then targetHasFormula1 = False '����ǿգ���Ҳû��formula1
    Exit Function
ErrorHandler:
    targetHasFormula1 = False
End Function

Private Sub setRangeValidation(ByRef certainRange As range, ByRef validationString As String)
    On Error Resume Next

    With certainRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=validationString
    End With
    
    If (Not certainRange.Validation.value) And certainRange.value <> "" Then
        certainRange.value = ""
    End If
End Sub

Private Sub setCabinetTypeValidation(ByRef cabinetTypeRange As range, ByRef productTypeRange As range, ByRef transportSheet As Worksheet)
    Dim cabinetTypeString As String
    cabinetTypeString = Get_Site_Cabinet_Related(productTypeRange.value, transportSheet, cabinetTypeRange)
    Call setRangeValidation(cabinetTypeRange, cabinetTypeString)
End Sub

Private Sub setFddTddValidation(ByRef fddTddTypeRange As range, ByRef productTypeRange As range)
    Dim fddTddTypeString As String
    Dim mappingSiteTemplateSht As Worksheet
    Set mappingSiteTemplateSht = ThisWorkbook.Worksheets("MappingSiteTemplate")
    fddTddTypeString = getFddTddTypeString(mappingSiteTemplateSht, productTypeRange.value)
    Call setRangeValidation(fddTddTypeRange, fddTddTypeString)
End Sub
    
'�������á�*Site Type���к͡�*Site Template���������б���¼�
Public Sub TransPortSheetSelectionChange(ByVal sh As Object, ByVal target As range)
    Dim m_Str_Cabinet As String
    Dim m_Str_Template As String
    Dim constProductTypeCol As Long, constPatternCol As Long, constFddTddCol As Long
    Dim constCabinetTypeCol As Long, constLteRadioPatternCol As Long
    Dim productTypeRange As range
    '��Base Station Transport Data��ҳ��*Site Type��������
    constProductTypeCol = getColNum(sh.name, constRecordRow, productTypeAttr, productTypeMoc)
    
    '��Base Station Transport Data��ҳ��*Site Template��������
    constPattenCol = getColNum(sh.name, constRecordRow, siteTemplateAttr, siteTemplateMoc)
    
        '��Base Station Transport Data��ҳ��*FDD/TDD Mode��������
    Dim baseStationVersion As String
    baseStationVersion = UCase(getBaseStationVersion)
    constFddTddCol = Get_Col(sh.name, constRecordRow, getResByKey("FDD/TDD Mode"))
    If constFddTddCol < 1 Then
    constFddTddCol = Get_Col(sh.name, constRecordRow, getResByKey("Mode"))
    End If
    '��Base Station Transport Data��ҳ��Cabinet Type��������
    constCabinetTypeCol = Get_Col(sh.name, constRecordRow, getResByKey("Cabinet Type"))
    
    '��Base Station Transport Data��ҳ��*LTE Radio Template��������
    constLteRadioPattenCol = getColNum(sh.name, constRecordRow, LRadioAttr, LRadioMoc)
    
      If constLteRadioPattenCol < 1 Then
        constLteRadioPattenCol = getColNum(sh.name, constRecordRow, LRadioAttr, LRESRadioMoc)
    End If
    
    
    '��ȡ��*Site Type���к�ѡֵ�����趨Ϊ������
    If target.column = constProductTypeCol And target.count = 1 And target.row > constRecordRow Then
        '��ȡ��*Site Type���к�ѡֵ
        m_Str_Cabinet = GetSiteType(sh, target)
        If m_Str_Cabinet <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Cabinet
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
    '���塸Cabinet Type����������ֵ����Ҫ�������ı��û�д���Site Typeֵ���¼�����û�������������
    ElseIf target.column = constCabinetTypeCol And target.count = 1 And target.row > constRecordRow Then
        Set productTypeRange = target.Offset(0, constProductTypeCol - constCabinetTypeCol)
        Call setCabinetTypeValidation(target, productTypeRange, sh)
    '���塸*FDD/TDD Mode����������ֵ����Ҫ�������ı��û�д���Site Typeֵ���¼�����û�������������
    ElseIf target.column = constFddTddCol And target.count = 1 And target.row > constRecordRow Then
        Set productTypeRange = target.Offset(0, constProductTypeCol - constFddTddCol)
        Call setFddTddValidation(target, productTypeRange)
    '��ȡ��*Site Template���к�ѡֵ�����趨Ϊ�����б�
    ElseIf target.column = constPattenCol And target.count = 1 And target.row > constRecordRow Then
        If constProductTypeCol <> -1 Then
            If Cells(target.row, constProductTypeCol).value <> "" Then
                '�ӡ�MappingSiteTemplate��ҳ��ȡ��*Site Template���к�ѡֵ
                If getNeType() <> "USU" Then
                    m_Str_Template = Get_Template_Related(target.Offset(0, constProductTypeCol - constPattenCol).value, target.Offset(0, constFddTddCol - constPattenCol).value, target.Offset(0, constCabinetTypeCol - constPattenCol).value, sh, target)
                Else
                    m_Str_Template = Get_Template_Related(target.Offset(0, constProductTypeCol - constPattenCol).value, "", "", sh, target)
                End If
                If m_Str_Template <> "" Then
                    With target.Validation
                           .Delete
                           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Template
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
    ElseIf target.column = constLteRadioPattenCol And target.count = 1 And target.row > constRecordRow Then
        '��ȡ��*Radio Template���к�ѡֵ
        m_Str_Radio = getRadioTemplate(Cells(target.row, constFddTddCol).value, sh, target)
        If m_Str_Radio <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str_Radio
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

'�ӡ�MappingSiteTypeCabinetType��ҳ��ȡ��Cabinet Type���к�ѡֵ
Function Get_Site_Cabinet_Related(SiteType As String, sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    
    m_Str = ""
    For m_rowNum = 2 To Worksheets("Mapping SiteType_CabinetType").range("a65536").End(xlUp).row
        If SiteType = Worksheets("Mapping SiteType_CabinetType").Cells(m_rowNum, 1).value Then
            If m_Str = "" Then
                m_Str = Worksheets("Mapping SiteType_CabinetType").Cells(m_rowNum, 2).value
            Else
                m_Str = m_Str & "," & Worksheets("Mapping SiteType_CabinetType").Cells(m_rowNum, 2).value
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
    
    Get_Site_Cabinet_Related = m_Str
End Function


'�ӡ�ProductType��ҳ��ȡ��*Site Type���к�ѡֵ
Function GetSiteType(sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_RowNum_Inner As Long
    Dim flag As Boolean
    Dim m_Str As String
    
    Dim siteTypeCol As New Collection
    Dim SiteType As String
    m_Str = ""
    For m_rowNum = 2 To Worksheets("Mapping SiteType_CabinetType").range("a65536").End(xlUp).row
        SiteType = Worksheets("Mapping SiteType_CabinetType").range("A" & m_rowNum).value
        If Not Contains(siteTypeCol, SiteType) And SiteType <> "" Then
            siteTypeCol.Add item:=SiteType, key:=SiteType
            m_Str = m_Str & SiteType & ","
        End If
    Next m_rowNum
    
    Call eraseLastChar(m_Str) 'ȥ�����һ��,
    
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

Function getRadioTemplate(fddTdd As String, sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_RowNum_Inner As Long
    Dim flag As Boolean
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    
    m_Str = ""
    If fddTdd = "TDD Only" Then
    fddTdd = "TDD"
    End If
      For m_rowNum = 2 To Worksheets("MappingRadioTemplate").range("a65536").End(xlUp).row
        If (fddTdd = Worksheets("MappingRadioTemplate").Cells(m_rowNum, 1).value Or fddTdd = "") Then
            If m_Str = "" Then
                m_Str = Worksheets("MappingRadioTemplate").Cells(m_rowNum, 3).value
            ElseIf VBA.Trim(Worksheets("MappingRadioTemplate").Cells(m_rowNum, 3).value) <> "" Then
                m_Str = m_Str & "," & Worksheets("MappingRadioTemplate").Cells(m_rowNum, 3).value
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
Public Function Get_Template_Related(SiteType As String, fddTdd As String, CabinetType As String, sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    
    m_Str = ""
    For m_rowNum = 2 To Worksheets("MappingSiteTemplate").range("a65536").End(xlUp).row
        If (SiteType = Worksheets("MappingSiteTemplate").Cells(m_rowNum, 1).value Or SiteType = "") _
        And (CabinetType = Worksheets("MappingSiteTemplate").Cells(m_rowNum, 2).value Or CabinetType = "") _
        And (fddTdd = Worksheets("MappingSiteTemplate").Cells(m_rowNum, 3).value Or fddTdd = "") Then
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
    For m_colNum = 1 To Worksheets(sheetName).range("IV2").End(xlToLeft).column
        If ColValue = Worksheets(sheetName).Cells(recordRow, m_colNum).value Then
            Get_Col = m_colNum
            Exit For
        End If
    Next
End Function




