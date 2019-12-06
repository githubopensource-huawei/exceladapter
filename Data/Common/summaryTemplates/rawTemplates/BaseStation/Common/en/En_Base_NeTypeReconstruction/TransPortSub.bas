Attribute VB_Name = "TransPortSub"
Option Explicit

'「Base Station Transport Data」页记录起始行
Private Const listShtTitleRow = 2
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
Private Const DSARadioMoc As String = "DsaFunction"
Private Const DSARadioAttr As String = "RadioTemplateName"

Private Const gsmRadioType As String = "GSM Radio Template"
Private Const umtsRadioType As String = "UMTS Radio Template"
Private Const lteRadioType As String = "LTE Radio Template"
Private Const lresRadioType As String = "LRES Radio Template"
Private Const nbiotRadioType As String = "NB-IoT Radio Template"
Private Const nrRadioType As String = "NR Radio Template"
Private Const dsaRadioType As String = "DSA Radio Template"


'定义「*Site Type」→「Cabinet Type」→「*Site Template」的联动事件
Public Sub TransPortSheetChange(ByVal sh As Object, ByVal target As range)
On Error GoTo ErrorHandler
    If target.row <= listShtTitleRow Or target.count > 1 Then Exit Sub
    
    Dim productTypeColNum As Integer
    productTypeColNum = getColNum(sh.name, listShtTitleRow, productTypeAttr, productTypeMoc) '「Base Station Transport Data」页「*Site Type」所在列
    
    Dim siteTemplateColNum As Integer
    siteTemplateColNum = getColNum(sh.name, listShtTitleRow, siteTemplateAttr, siteTemplateMoc) '「Base Station Transport Data」页「*Site Template」所在列
    
    If target.column = productTypeColNum Then
        If productTypeColNum <> -1 And siteTemplateColNum <> -1 Then
            Dim siteTemplateListValue As String
            siteTemplateListValue = getTemplateListValue(target.value, sh, target.Offset(0, siteTemplateColNum - productTypeColNum))
            If siteTemplateListValue <> "" Then
                With target.Offset(0, siteTemplateColNum - productTypeColNum).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=siteTemplateListValue
                End With
                If Not target.Offset(0, siteTemplateColNum - productTypeColNum).Validation.value Then
                    target.Offset(0, siteTemplateColNum - productTypeColNum).value = ""
                End If
            ElseIf Not special() Then
                Call clearValidation(target.Offset(0, siteTemplateColNum - productTypeColNum))
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in TransPortSheetChange, " & Err.Description
End Sub
    
'定义设置「*Site Type」列和「*Site Template」列下拉列表的事件
Public Sub TransPortSheetSelectionChange(ByVal sh As Object, ByVal target As range)
On Error GoTo ErrorHandler
    If target.count > 1 Or target.row <= listShtTitleRow Then Exit Sub
    
    Dim productTypeColNum As Long, siteTemplateColNum As Long, gsmRadioTemplateColNum As Long, umtsRadioTemplateColNum As Long
    Dim lteRadioTemplateColNum As Long, nbiotRadioTemplateColNum As Long, nrRadioTemplateColNum As Long, dsaRadioTemplateColNum As Long
    
    productTypeColNum = getColNum(sh.name, listShtTitleRow, productTypeAttr, productTypeMoc) '「Base Station Transport Data」页「*Site Type」所在列
    siteTemplateColNum = getColNum(sh.name, listShtTitleRow, siteTemplateAttr, siteTemplateMoc) '「Base Station Transport Data」页「*Site Template」所在列
    gsmRadioTemplateColNum = getColNum(sh.name, listShtTitleRow, GRadioAttr, GRadioMoc) '「Base Station Transport Data」页「*GBTS Radio Template」所在列
    umtsRadioTemplateColNum = getColNum(sh.name, listShtTitleRow, URadioAttr, URadioMoc) '「Base Station Transport Data」页「*Umts Radio Template」所在列
    lteRadioTemplateColNum = getColNum(sh.name, listShtTitleRow, LRadioAttr, LRadioMoc) '「Base Station Transport Data」页「*LTE Radio Template」所在列
    If lteRadioTemplateColNum < 1 Then
        lteRadioTemplateColNum = getColNum(sh.name, listShtTitleRow, LRadioAttr, LRESRadioMoc)
    End If
    
    nbiotRadioTemplateColNum = getColNum(sh.name, listShtTitleRow, MRadioAttr, MRadioMoc) '「Base Station Transport Data」页「*NBIOT Radio Template」所在列
    nrRadioTemplateColNum = getColNum(sh.name, listShtTitleRow, NRadioAttr, NRadioMoc) '「Base Station Transport Data」页「*NR Radio Template」所在列
    dsaRadioTemplateColNum = getColNum(sh.name, listShtTitleRow, DSARadioAttr, DSARadioMoc) '「Base Station Transport Data」页「*DSA Radio Template」所在列
    
    Dim listValue As String
    If target.column = productTypeColNum Then
        listValue = getProductTypeListValue(sh, target) '获取「*Site Type」列侯选值
        If listValue <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listValue
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        ElseIf Not special() Then
            Call clearValidation(target)
        End If
    ElseIf target.column = siteTemplateColNum And productTypeColNum <> -1 Then '获取「*Site Template」列侯选值，并设定为下拉列表
        If Cells(target.row, productTypeColNum).value <> "" Then
            listValue = getTemplateListValue(target.Offset(0, productTypeColNum - siteTemplateColNum).value, sh, target) '从「MappingSiteTemplate」页获取「*Site Template」列侯选值
            If listValue <> "" Then
                With target.Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listValue
                End With
                If Not target.Validation.value Then
                    target.value = ""
                End If
            Else
                Call clearValidation(target)
            End If
        Else
            Call clearValidation(target)
        End If
    ElseIf target.column = gsmRadioTemplateColNum Then
        listValue = getRadioTemplate(getResByKey(gsmRadioType), sh, target) '获取「*Radio Template」列侯选值
        If listValue <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listValue
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            Call clearValidation(target)
        End If
    ElseIf target.column = umtsRadioTemplateColNum Then
        listValue = getRadioTemplate(getResByKey(umtsRadioType), sh, target)
        If listValue <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listValue
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            Call clearValidation(target)
        End If
    ElseIf target.column = lteRadioTemplateColNum Then
        listValue = getRadioTemplate(getResByKey(lteRadioType), sh, target)
        If Len(listValue) = 0 Then
            listValue = getRadioTemplate(getResByKey(lresRadioType), sh, target)
        End If
        If listValue <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listValue
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            Call clearValidation(target)
        End If
     ElseIf target.column = nbiotRadioTemplateColNum Then
        listValue = getRadioTemplate(getResByKey(nbiotRadioType), sh, target)
        If listValue <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listValue
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            Call clearValidation(target)
        End If
    ElseIf target.column = nrRadioTemplateColNum Then
        listValue = getRadioTemplate(getResByKey(nrRadioType), sh, target)
        If listValue <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listValue
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            Call clearValidation(target)
        End If
    ElseIf target.column = dsaRadioTemplateColNum Then
        listValue = getRadioTemplate(getResByKey(dsaRadioType), sh, target)
        If listValue <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listValue
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            Call clearValidation(target)
        End If
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in TransPortSheetSelectionChange, " & Err.Description
End Sub

Function getProductTypeListValue(sheet As Worksheet, cellRange As range) As String
On Error GoTo ErrorHandler
    Dim neType As String
    neType = getNeType()
    
    Dim productTypeSht As Worksheet
    Set productTypeSht = Worksheets("ProductType")
    
    Dim productTypes As New Collection
    Dim ProductType As String
    Dim rowIdx As Integer
    With productTypeSht
        For rowIdx = 2 To getUsedRowCount(productTypeSht, 1)
            If .Cells(rowIdx, 2) = neType Then
                ProductType = .Cells(rowIdx, 1)
                If ProductType <> "" And Not Contains(productTypes, ProductType) Then
                    productTypes.Add item:=ProductType, key:=ProductType
                End If
            End If
        Next
    End With
    
    getProductTypeListValue = collectionJoin(productTypes)
    
    If Len(getProductTypeListValue) > 255 Then getProductTypeListValue = getIndirectListValue(sheet, cellRange.column, getProductTypeListValue)
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getProductTypeListValue, " & Err.Description
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
                    radioTemplates.Add item:=radioTemplate, key:=radioTemplate
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
Public Function getTemplateListValue(siteType As String, sheet As Worksheet, cellRange As range) As String
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
            If (.Cells(rowIdx, 1).value = siteType Or Trim(.Cells(rowIdx, 1).value) = "") And .Cells(rowIdx, 5).value = neType Then
                siteTemplate = .Cells(rowIdx, 4).value
                If siteTemplate <> "" And Not Contains(siteTemplates, siteTemplate) Then
                    siteTemplates.Add item:=siteTemplate, key:=siteTemplate
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









