Attribute VB_Name = "TransPortSub"
Option Explicit

'「Base Station Transport Data」页记录起始行
Private Const constRecordRow = 2
Private Const productTypeAttr As String = "PRODUCTTYPE"
Private Const productTypeMoc As String = "NODE"
Private Const siteTemplateAttr As String = "SITETEMPLATENAME"
Private Const siteTemplateMoc As String = "NODE"
Private Const LRadioAttr As String = "RADIOTEMPLATENAME"
Private Const LRadioMoc As String = "ENODEBFUNCTION"
'Private Const LRESRadioMoc As String = "ENODEBEQMFUNCTION"
Private Const lteRadioType As String = "LTE RADIO TEMPLATE"

Private Function getFddTddTypeString(ByRef SiteType As String) As String
    Dim mappingSiteTemplateSht As Worksheet
    Set mappingSiteTemplateSht = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    Dim maxRowNumber As Long, rowIndex As Long
    maxRowNumber = mappingSiteTemplateSht.range("A65535").End(xlUp).row
    
    Dim fddTddTypeCol As New Collection
    Dim fddTddValue As Variant
    '根据输入的SiteType得到支持的FddTdd类型容器
    For rowIndex = 2 To maxRowNumber
        With mappingSiteTemplateSht
            If .range("A" & rowIndex).value = SiteType Then
                fddTddValue = .range("C" & rowIndex).value
                If fddTddValue <> "" And Not Contains(fddTddTypeCol, CStr(fddTddValue)) Then
                    fddTddTypeCol.Add Item:=fddTddValue, key:=fddTddValue
                End If
            End If
        End With
    Next rowIndex
    
    getFddTddTypeString = collectionJoin(fddTddTypeCol)
End Function



'定义「*Site Type」→「Cabinet Type」→「*Site Template」的联动事件
Public Sub transportSheetChange(ByVal sh As Object, ByVal target As range)
    On Error GoTo ErrorHandler
    If target.row <= constRecordRow Or target.count > 1 Then Exit Sub

    Dim productTypeCol As Integer, fddTddCol As Integer, cabinetTypeCol As Integer, siteTemplateCol As Integer, radioTemplateCol As Integer
    productTypeCol = getColNum(sh.name, constRecordRow, productTypeAttr, productTypeMoc)

    fddTddCol = Get_Col(sh.name, constRecordRow, getResByKey("FDD/TDD Mode"))
    If fddTddCol < 1 Then
        fddTddCol = Get_Col(sh.name, constRecordRow, getResByKey("Mode"))
    End If

    cabinetTypeCol = Get_Col(sh.name, constRecordRow, getResByKey("Cabinet Type"))
    
    siteTemplateCol = getColNum(sh.name, constRecordRow, siteTemplateAttr, siteTemplateMoc)
    radioTemplateCol = getColNum(sh.name, constRecordRow, LRadioAttr, LRadioMoc)
    
    Dim enumValueList As String
    If target.column = productTypeCol Then  '「*Site Type」值变更时，重新设置「Site Template」的侯选值并清除当前值。
        If cabinetTypeCol <> -1 Then
            enumValueList = Get_Site_Cabinet_Related(target.value, sh, target.Offset(0, cabinetTypeCol - productTypeCol))
            Call setValidationEnum(target.Offset(0, cabinetTypeCol - productTypeCol), enumValueList)
        End If
        If fddTddCol <> -1 Then
            'DBS3900_LTE支持FDD/TDD/FDDTDD，其它站型只支持FDD：这段说明是旧的，对于3205E的小站已经无法支持，需要根据SiteType生成相应的FDD/TDD模式
            Dim fddTddTypeString As String
            fddTddTypeString = getFddTddTypeString(target.value)

            If target.value = "" Then
                target.Offset(0, fddTddCol - productTypeCol).Validation.Delete
                target.Offset(0, fddTddCol - productTypeCol).value = ""
            Else
                With target.Offset(0, fddTddCol - productTypeCol).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=fddTddTypeString
                End With
                If Not target.Offset(0, fddTddCol - productTypeCol).Validation.value Then
                    target.Offset(0, fddTddCol - productTypeCol).value = ""
                End If
            End If
        End If
    ElseIf target.column = cabinetTypeCol Then    '「Cabinet Type」值变更时，重新设置「*Site Template」的侯选值并清除当前值。
        If productTypeCol <> -1 And fddTddCol <> -1 And siteTemplateCol <> -1 Then
            enumValueList = Get_Template_Related(target.Offset(0, productTypeCol - cabinetTypeCol).value, target.Offset(0, fddTddCol - cabinetTypeCol).value, target.value, sh, target.Offset(0, siteTemplateCol - cabinetTypeCol))
            Call setValidationEnum(target.Offset(0, siteTemplateCol - cabinetTypeCol), enumValueList)
        End If
    ElseIf target.column = fddTddCol Then    '「*FDD/TDD Mode」值变更时，重新设置「*Site Template」和「Radio Template」的侯选值并清除当前值。
        If productTypeCol <> -1 And cabinetTypeCol <> -1 And siteTemplateCol <> -1 Then
            enumValueList = Get_Template_Related(target.Offset(0, productTypeCol - fddTddCol).value, target.value, target.Offset(0, cabinetTypeCol - fddTddCol).value, sh, target.Offset(0, siteTemplateCol - fddTddCol))
            Call setValidationEnum(target.Offset(0, siteTemplateCol - fddTddCol), enumValueList)
        End If
        
        If radioTemplateCol <> -1 Then
            enumValueList = getRadioTemplate(target.value, sh, target.Offset(0, radioTemplateCol - fddTddCol))
            Call setValidationEnum(target.Offset(0, radioTemplateCol - fddTddCol), enumValueList)
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in transportSheetChange, " & Err.Description
End Sub

Private Function targetHasFormula1(ByRef target As range) As Boolean
    On Error GoTo ErrorHandler
    targetHasFormula1 = True
    If target.Validation Is Nothing Then '没有有效性，则没有formula1
        targetHasFormula1 = False
        Exit Function
    End If
    
    Dim formula1 As String
    formula1 = target.Validation.formula1 '如果有formula1，则赋值成功，如果没有，则赋值出错，进入ErrorHandler
    If formula1 = "" Then targetHasFormula1 = False '如果是空，则也没有formula1
    Exit Function
ErrorHandler:
    targetHasFormula1 = False
End Function
    
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
    
    Dim productTypeCol As Integer, productTypeValue As String
    productTypeCol = getColNum(sht.name, constRecordRow, productTypeAttr, productTypeMoc)
    productTypeValue = sht.Cells(target.row, productTypeCol).value
    
    Dim enumValueList As String
    If mocName = productTypeMoc And attrName = productTypeAttr Then
        enumValueList = GetSiteType(sht, target)
        Call setValidationEnum(target, enumValueList)
    ElseIf mocName = siteTemplateMoc And attrName = siteTemplateAttr Then
        If productTypeValue = "" Then
            target.Validation.Delete
            target.value = ""
            Exit Sub
        Else
            Dim fddTddCol As Integer, cabinetTypeCol As Integer
            Dim fddTddValue As String, cabinetTypeValue As String
            fddTddCol = Get_Col(sht.name, constRecordRow, getResByKey("FDD/TDD Mode"))
            If fddTddCol = -1 Then fddTddCol = Get_Col(sht.name, constRecordRow, getResByKey("Mode"))
            cabinetTypeCol = Get_Col(sht.name, constRecordRow, getResByKey("Cabinet Type"))
            If fddTddCol <> -1 Then fddTddValue = sht.Cells(target.row, fddTddCol)
            If cabinetTypeCol <> -1 Then cabinetTypeValue = sht.Cells(target.row, cabinetTypeCol)
            
            enumValueList = Get_Template_Related(productTypeValue, fddTddValue, cabinetTypeValue, sht, target)
            Call setValidationEnum(target, enumValueList)
        End If
    ElseIf colName = getResByKey("Cabinet Type") Then
        enumValueList = Get_Site_Cabinet_Related(productTypeValue, sht, target)
        Call setValidationEnum(target, enumValueList)
    ElseIf colName = getResByKey("FDD/TDD Mode") Or colName = getResByKey("Mode") Then
        enumValueList = getFddTddTypeString(productTypeValue)
        Call setValidationEnum(target, enumValueList)
    ElseIf mocName = LRadioMoc And attrName = LRadioAttr Then
        fddTddCol = Get_Col(sht.name, constRecordRow, getResByKey("FDD/TDD Mode"))
        If fddTddCol = -1 Then fddTddCol = Get_Col(sht.name, constRecordRow, getResByKey("Mode"))
        If fddTddCol <> -1 Then fddTddValue = sht.Cells(target.row, fddTddCol)
        enumValueList = getRadioTemplate(fddTddValue, sht, target)
        Call setValidationEnum(target, enumValueList)
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in transportSheetSelectionChange, " & Err.Description
End Sub

'从「MappingSiteTypeCabinetType」页获取「Cabinet Type」列侯选值
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


'从「ProductType」页获取「*Site Type」列侯选值
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
            siteTypeCol.Add Item:=SiteType, key:=SiteType
            m_Str = m_Str & SiteType & ","
        End If
    Next m_rowNum
    
    Call eraseLastChar(m_Str) '去掉最后一个,
    
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


'从「MappingSiteTemplate」页获取「*Site Template」列侯选值
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
    Dim groupName As String
    Dim columnName As String
    Dim mappingDef As CMappingDef
    Dim address As String
    Dim addressArray() As String
    Dim iRange As range
    Dim iSheet As Worksheet
    
    Dim mocGroupName As String
    Dim mocColumnName As String
    Dim mocSheetName As String
    
    Dim controlDef As CControlDef
    
    If isIubStyleWorkBook() Then
        groupName = get_GroupName(sheet.name, target.column)
        columnName = get_ColumnName(sheet.name, target.column)
        Set mappingDef = getMappingDefine(sheet.name, groupName, columnName)
        If Not mappingDef Is Nothing Then
           If (mappingDef.mocName <> "" And mappingDef.attributeName <> "") Or _
           mappingDef.columnName = getResByKey("Cabinet Type") Or mappingDef.columnName = getResByKey("FDD/TDD Mode") Or mappingDef.columnName = getResByKey("Mode") Then
                Exit Sub
           End If
        End If
        
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
        
        Dim validLen As Long
        validLen = Len(controlDef.lstValue)
        Dim longValidatastr As String
        If validLen > 255 Then
            longValidatastr = getLongValidatastr(mocSheetName, mocGroupName, mocColumnName)
        End If
        
        
        If Not controlDef Is Nothing Then
            If UCase(controlDef.dataType) = "ENUM" And controlDef.lstValue <> "" And longValidatastr = "" Then
                With target.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=controlDef.lstValue
                End With
            ElseIf UCase(controlDef.dataType) = "ENUM" And controlDef.lstValue <> "" And longValidatastr <> "" Then
                With target.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=longValidatastr
                End With
            Else
                With target.Validation
                    .InCellDropdown = False
                    .Delete
                End With
            End If
        Else
            With target.Validation
                .InCellDropdown = False
                .Delete
            End With
        End If
    End If
End Sub




