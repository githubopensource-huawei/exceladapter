Attribute VB_Name = "TransPortSub"
Option Explicit

'「Base Station Transport Data」页记录起始行
Private Const listShtTitleRow = 2
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
Private Const DSARadioMoc As String = "DSAFUNCTION"
Private Const DSARadioAttr As String = "RADIOTEMPLATENAME"

Private Const gsmRadioType As String = "GSM RADIO TEMPLATE"
Private Const umtsRadioType As String = "UMTS RADIO TEMPLATE"
Private Const lteRadioType As String = "LTE RADIO TEMPLATE"
'Private Const lresRadioType As String = "LRES RADIO TEMPLATE"
Private Const nbiotRadioType As String = "NB-IOT RADIO TEMPLATE"
Private Const nrRadioType As String = "NR RADIO TEMPLATE"
Private Const dsaRadioType As String = "DSA RADIO TEMPLATE"


'定义「*Site Type」→「Cabinet Type」→「*Site Template」的联动事件
Public Sub TransPortSheetChange(ByVal sh As Object, ByVal target As range)
    Dim m_Str As String
    Dim productTypeCol As Integer, siteTemplateCol As Integer

    productTypeCol = getColNum(sh.name, listShtTitleRow, productTypeAttr, productTypeMoc)
    
    siteTemplateCol = getColNum(sh.name, listShtTitleRow, siteTemplateAttr, siteTemplateMoc)
    
    '「*Site Type」值变更时，重新设置「Site Template」的侯选值并清除当前值。
    If target.column = productTypeCol And target.row > listShtTitleRow And target.count = 1 Then
        If productTypeCol <> -1 And siteTemplateCol <> -1 Then
            Dim siteTemplateListValue As String
            siteTemplateListValue = getTemplateListValue(target.value, sh, target.Offset(0, siteTemplateCol - productTypeCol))
            Call setValidationEnum(target.Offset(0, siteTemplateCol - productTypeCol), siteTemplateListValue)
        End If
    End If
End Sub
    
'定义设置「*Site Type」列和「*Site Template」列下拉列表的事件
Public Sub transportSheetSelectionChange(ByVal sh As Object, ByVal target As range)
    On Error GoTo ErrorHandler
    If target.count > 1 Or target.row <= listShtTitleRow Then Exit Sub
    
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
        enumValueList = getSiteType(sht, target)
        Call setValidationEnum(target, enumValueList)
    ElseIf mocName = siteTemplateMoc And attrName = siteTemplateAttr Then
        Dim productTypeCol As Integer, productTypeValue As String
        productTypeCol = getColNum(sht.name, listShtTitleRow, productTypeAttr, productTypeMoc)
        productTypeValue = target.Offset(0, productTypeCol - target.column).value
        If productTypeValue = "" Then
            target.Validation.Delete
            target.value = ""
            Exit Sub
        Else
            enumValueList = getTemplateListValue(productTypeValue, sht, target)
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
    ElseIf mocName = DSARadioMoc And attrName = DSARadioAttr Then
        enumValueList = getRadioTemplate(getResByKey(dsaRadioType), sh, target)
        Call setValidationEnum(target, enumValueList)
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "some exception in transportSheetSelectionChange, " & Err.Description
End Sub

'从「MappingSiteTypeCabinetType」页获取「*Site Type」列侯选值
Function getSiteType(sheet As Worksheet, cellRange As range) As String
On Error GoTo ErrorHandler
    Dim neType As String
    neType = getNeType()

    Dim productTypeTemplate As Worksheet
    Set productTypeTemplate = Worksheets("ProductType")
    
    Dim productTypes As New Collection
    Dim ProductType As String
    Dim rowIdx As Integer
    With productTypeTemplate
        For rowIdx = 2 To getUsedRowCount(productTypeTemplate, 1)
            If .Cells(rowIdx, 2).value = neType Then
                ProductType = .Cells(rowIdx, 1).value
                If ProductType <> "" And Not Contains(productTypes, ProductType) Then
                    productTypes.Add Item:=ProductType, key:=ProductType
                End If
            End If
        Next
    End With
    
    getSiteType = collectionJoin(productTypes)
    
    If Len(getSiteType) > 255 Then getSiteType = getIndirectListValue(sheet, cellRange.column, getSiteType)
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getSiteType, " & Err.Description
End Function

Function getRadioTemplate(radioType As String, sheet As Worksheet, cellRange As range) As String
On Error GoTo ErrorHandler
    Dim neType As String
    neType = getNeType()
    
    Dim mapRadioTemplate As Worksheet
    Set mapRadioTemplate = Worksheets("MappingRadioTemplate")
    
    Dim radioTemplates As New Collection
    Dim radioTemplate As String
    Dim rowIdx As Integer
    With mapRadioTemplate
        For rowIdx = 2 To getUsedRowCount(mapRadioTemplate, 1)
            If (.Cells(rowIdx, 2).value = radioType Or Trim(.Cells(rowIdx, 2).value) = "") And .Cells(rowIdx, 3).value = neType Then
                radioTemplate = .Cells(rowIdx, 1).value
                If radioTemplate <> "" And Not Contains(radioTemplates, radioTemplate) Then
                    radioTemplates.Add Item:=radioTemplate, key:=radioTemplate
                End If
            End If
        Next
    End With
    
    getRadioTemplate = collectionJoin(radioTemplates)
    
    If Len(getRadioTemplate) > 255 Then getRadioTemplate = getIndirectListValue(sheet, cellRange.column, getRadioTemplate)
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getRadioTemplate, " & Err.Description
End Function


'从「MappingSiteTemplate」页获取「*Site Template」列侯选值
Private Function getTemplateListValue(siteType As String, sheet As Worksheet, cellRange As range) As String
On Error GoTo ErrorHandler
    Dim neType As String
    neType = getNeType()
    
    Dim mapSiteTemplate As Worksheet
    Set mapSiteTemplate = Worksheets("MappingSiteTemplate")
    
    Dim siteTemplates As New Collection
    Dim siteTemplate As String
    Dim rowIdx As Integer
    With mapSiteTemplate
        For rowIdx = 2 To getUsedRowCount(mapSiteTemplate, 1)
            If (.Cells(rowIdx, 1).value = siteType Or Trim(.Cells(rowIdx, 1).value = "")) And .Cells(rowIdx, 5).value = neType Then
                siteTemplate = .Cells(rowIdx, 4).value
                If siteTemplate <> "" And Not Contains(siteTemplates, siteTemplate) Then
                    siteTemplates.Add Item:=siteTemplate, key:=siteTemplate
                End If
            End If
        Next
    End With
    
    getTemplateListValue = collectionJoin(siteTemplates)
    
    If Len(getTemplateListValue) > 255 Then getTemplateListValue = getIndirectListValue(sheet, cellRange.column, getTemplateListValue)

    Exit Function
ErrorHandler:
    Debug.Print "some exception in getTemplateListValue, " & Err.Description
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



