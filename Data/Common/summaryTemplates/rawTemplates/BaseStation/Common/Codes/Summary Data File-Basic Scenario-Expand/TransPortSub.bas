Attribute VB_Name = "TransPortSub"
Option Explicit

'「Base Station Transport Data」页记录起始行
Private Const constRecordRow = 2
Private Const productTypeAttr As String = "PRODUCTTYPE"
Private Const productTypeMoc As String = "NODE"
Private Const siteTemplateAttr As String = "SITETEMPLATENAME"
Private Const siteTemplateMoc As String = "NODE"
Private Const GRadioAttr As String = "RADIOTEMPLATENAME"
Private Const GRadioMoc As String = "GBTSFUNCTION"
Private Const URadioAttr As String = "RADIOTEMPLATENAME"
Private Const URadioMoc As String = "NODEBFUNCTION"
Private Const LRadioAttr As String = "RADIOTEMPLATENAME"
Private Const LRadioMoc As String = "ENODEBFUNCTION"
'Private Const LRESRadioMoc As String = "ENODEBEQMFUNCTION"
Private Const MRadioAttr As String = "RADIOTEMPLATENAME"
Private Const MRadioMoc As String = "NBBSFUNCTION"

Private Const NRadioMoc As String = "GNODEBFUNCTION"
Private Const NRadioAttr As String = "RADIOTEMPLATENAME"

Private Const gsmRadioType As String = "GSM RADIO TEMPLATE"
Private Const umtsRadioType As String = "UMTS RADIO TEMPLATE"
Private Const lteRadioType As String = "LTE RADIO TEMPLATE"
'Private Const lresRadioType As String = "LRES RADIO TEMPLATE"
Private Const nbiotRadioType As String = "NB-IOT RADIO TEMPLATE"

Private Const nrRadioType As String = "NR RADIO TEMPLATE"


'定义「*Site Type」→「Cabinet Type」→「*Site Template」的联动事件
Public Sub TransPortSheetChange(ByVal sh As Object, ByVal target As range)
    Dim m_Str As String
    Dim productTypeCol As Integer, siteTemplateCol As Integer

    productTypeCol = getColNum(sh.name, constRecordRow, productTypeAttr, productTypeMoc)
    
    siteTemplateCol = getColNum(sh.name, constRecordRow, siteTemplateAttr, siteTemplateMoc)
    
    '「*Site Type」值变更时，重新设置「Site Template」的侯选值并清除当前值。
    If target.column = productTypeCol And target.row > constRecordRow And target.count = 1 Then
        If productTypeCol <> -1 And siteTemplateCol <> -1 Then
            Dim siteTemplateListValue As String
            siteTemplateListValue = Get_Template_Related(target.value, sh, target.Offset(0, siteTemplateCol - productTypeCol))
            Call setValidationEnum(target.Offset(0, siteTemplateCol - productTypeCol), siteTemplateListValue)
        End If
    End If
End Sub
    
'定义设置「*Site Type」列和「*Site Template」列下拉列表的事件
Public Sub transportSheetSelectionChange(ByVal sh As Object, ByVal target As range)
    On Error GoTo ErrorHandler
    If target.count > 1 Or target.row <= constRecordRow Then Exit Sub
    
    Dim sht As Worksheet
    Set sht = sh
    
    Dim mappingDef As CMappingDef
    Dim grpName As String, colName As String, mocName As String, attrName As String
    
    grpName = get_GroupName(sht.name, target.column)
    colName = get_ColumnName(sht.name, target.column)
    Set mappingDef = getMappingDefine(sht.name, grpName, colName)
    
    If mappingDef Is Nothing Then Exit Sub
    
    mocName = UCase(mappingDef.mocName)
    attrName = UCase(mappingDef.attributeName)
    
    If mocName = "" Or attrName = "" Then
        Call iubTransportSheetSelectionChange(sht, target)
        Exit Sub
    End If
    
    Dim enumValueList As String
    If mocName = productTypeMoc And attrName = productTypeAttr Then
        enumValueList = GetSiteType(sht, target)
        Call setValidationEnum(target, enumValueList)
    ElseIf mocName = siteTemplateMoc And attrName = siteTemplateAttr Then
        Dim productTypeCol As Integer, productTypeValue As String
        productTypeCol = getColNum(sht.name, constRecordRow, productTypeAttr, productTypeMoc)
        productTypeValue = target.Offset(0, productTypeCol - target.column).value
        If productTypeValue = "" Then
            target.Validation.Delete
            target.value = ""
            Exit Sub
        Else
            enumValueList = Get_Template_Related(productTypeValue, sht, target)
            Call setValidationEnum(target, enumValueList)
        End If
    ElseIf mocName = GRadioMoc And attrName = GRadioAttr Then
        enumValueList = getRadioTemplate(getResByKey(gsmRadioType), sht, target)
        Call setValidationEnum(target, enumValueList)
    ElseIf mocName = URadioMoc And attrName = URadioAttr Then
        enumValueList = getRadioTemplate(getResByKey(umtsRadioType), sht, target)
        Call setValidationEnum(target, enumValueList)
    ElseIf mocName = LRadioMoc And attrName = LRadioAttr Then
        enumValueList = getRadioTemplate(getResByKey(lteRadioType), sht, target)
        Call setValidationEnum(target, enumValueList)
    ElseIf mocName = MRadioMoc And attrName = MRadioAttr Then
        enumValueList = getRadioTemplate(getResByKey(nbiotRadioType), sht, target)
        Call setValidationEnum(target, enumValueList)
    ElseIf mocName = NRadioMoc And attrName = NRadioAttr Then
        enumValueList = getRadioTemplate(getResByKey(nrRadioType), sht, target)
        Call setValidationEnum(target, enumValueList)
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "some exception in transportSheetSelectionChange, " & Err.Description
End Sub

'从「MappingSiteTypeCabinetType」页获取「*Site Type」列侯选值
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

Function getRadioTemplate(radioType As String, sheet As Worksheet, cellRange As range) As String
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
        If (radioType = Worksheets("MappingRadioTemplate").Cells(m_rowNum, 2).value Or Len(Trim(Worksheets("MappingRadioTemplate").Cells(m_rowNum, 2).value)) = 0) _
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


'从「MappingSiteTemplate」页获取「*Site Template」列侯选值
Public Function Get_Template_Related(siteType As String, sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    Dim neType As String
    
    neType = getNeType()
    
    m_start = 0
    m_Str = ""
    For m_rowNum = 2 To Worksheets("MappingSiteTemplate").range("a65536").End(xlUp).row
        If (siteType = Worksheets("MappingSiteTemplate").Cells(m_rowNum, 1).value Or Worksheets("MappingSiteTemplate").Cells(m_rowNum, 1).value = "") _
        And neType = Worksheets("MappingSiteTemplate").Cells(m_rowNum, 5).value Then
            If m_Str = "" Then
                m_Str = Worksheets("MappingSiteTemplate").Cells(m_rowNum, 4).value
            ElseIf VBA.Trim(Worksheets("MappingSiteTemplate").Cells(m_rowNum, 4).value) <> "" And InStr(m_Str, VBA.Trim(Worksheets("MappingSiteTemplate").Cells(m_rowNum, 4).value)) <= 0 Then
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

'从指定sheet页的指定行，查找指定列，返回列号
Private Function Get_Col(sheetName As String, recordRow As Long, ColValue As String) As Long
    Get_Col = -1

    Dim targetRange As range
    Set targetRange = ThisWorkbook.Worksheets(sheetName).rows(recordRow).Find(ColValue, lookat:=xlWhole, LookIn:=xlValues)
    If Not targetRange Is Nothing Then Get_Col = targetRange.column
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

Public Sub setValidationEnum(target As range, validationText As String)
    On Error GoTo ErrorHandler
    If validationText <> "" Then
        With target.Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=validationText
        End With
        If Not target.Validation.value Then
            target.value = ""
        End If
    Else
        target.Validation.Delete
        target.value = ""
    End If
        
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setValidationEnum, " & Err.Description
End Sub

Public Sub iubTransportSheetSelectionChange(sheet As Worksheet, target As range)
    On Error GoTo ErrorHandler
    Dim address As String
    Dim addressArray() As String
    Dim iRange As range
    Dim iSheet As Worksheet
    
    Dim mocGroupName As String
    Dim mocColumnName As String
    Dim mocSheetName As String
    
    Dim controlDef As CControlDef

    If sheet.Cells(3, target.column).value = "" Then
        With target.Validation
            .InCellDropdown = False
            .Delete
        End With
        Exit Sub
    End If
    addressArray = Split(sheet.Cells(3, target.column).value, ",")
    address = addressArray(0)
    For Each iSheet In ThisWorkbook.Sheets
        If iSheet.Tab.colorIndex = BluePrintSheetColor Then
            Set iRange = iSheet.range(address)
            Exit For
        End If
    Next iSheet
    Call getGroupNameShNameAndAttrName(iSheet, iRange, mocGroupName, mocSheetName, mocColumnName)
    Set controlDef = getControlDefine(mocSheetName, mocGroupName, mocColumnName)
    
    Dim enumListValue As String
    enumListValue = controlDef.lstValue
    
    If Len(enumListValue) > 255 Then
       enumListValue = getIndirectValidateListValue(mocSheetName, mocGroupName, mocColumnName)
    End If
    
    If Not controlDef Is Nothing Then
        If UCase(controlDef.dataType) <> "ENUM" Then Exit Sub
        Call setValidationEnum(target, enumListValue)
    Else
        With target.Validation
            .InCellDropdown = False
            .Delete
        End With
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in iubTransportSheetSelectionChange, " & Err.Description
End Sub



