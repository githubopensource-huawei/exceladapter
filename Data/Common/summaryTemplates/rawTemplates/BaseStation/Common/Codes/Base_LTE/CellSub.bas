Attribute VB_Name = "CellSub"
Option Explicit

Private Const listShtTitleRow = 2

Public Function isCellSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("LTE Cell") Or sheetName = getResByKey("DCell") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function

'定义「DlBandWidth」→「SubframeAssignment」→「FddTddInd」→「AntennaMode」 →「Cell Template」的联动
Public Sub Cell_Worksheet_Change(ByVal sh As Worksheet, ByVal target As range)
On Error GoTo ErrorHandler
    If target.row <= listShtTitleRow Or target.count > 1 Then Exit Sub
    
    Dim saColNum As Long, fddTddColNum As Long, txrxModeColNum As Long, nbiotFlagColNum As Long, bandwidthColNum As Long, cellTempColNum As Long
    
    saColNum = getColNum(sh.name, listShtTitleRow, "SubframeAssignment", "Cell") '「LTE Cell」页「SubframeAssignment」所在列
    fddTddColNum = getColNum(sh.name, listShtTitleRow, "FddTddInd", "Cell") '「LTE Cell」页「*FddTddInd」所在列
    txrxModeColNum = getColNum(sh.name, listShtTitleRow, "TxRxMode", "Cell") '「LTE Cell」页「*TxRxMode」所在列
    bandwidthColNum = getColNum(sh.name, listShtTitleRow, "DlBandWidth", "Cell") '「LTE Cell」页「*DlBandwidth」所在列
    cellTempColNum = getColNum(sh.name, listShtTitleRow, "CellTemplateName", "Cell") '「LTE Cell」页「*Cell Template」所在列
    nbiotFlagColNum = getColNum(sh.name, listShtTitleRow, "NbCellFlag", "Cell") '「LTE Cell」页「*NB-IoT TA Flag」所在列,有些网元模型中没有这个属性
    
    If bandwidthColNum = -1 Or saColNum = -1 Or fddTddColNum = -1 Or txrxModeColNum = -1 Or cellTempColNum = -1 Then
        Exit Sub
    End If
    
    Dim curCol As Long
    curCol = target.column
    If curCol <> bandwidthColNum And curCol <> saColNum And curCol <> fddTddColNum And curCol <> txrxModeColNum And (nbiotFlagColNum = -1 Or curCol <> nbiotFlagColNum) Then
        Exit Sub
    End If
    
    Dim saValue As String, fddTddValue As String, txrxModeValue As String, nbiotFlagValue As String, bandwidthValue As String
    
    With sh
        saValue = .Cells(target.row, saColNum)
        fddTddValue = .Cells(target.row, fddTddColNum)
        txrxModeValue = .Cells(target.row, txrxModeColNum)
        bandwidthValue = .Cells(target.row, bandwidthColNum)
        nbiotFlagValue = "False"
        If nbiotFlagColNum <> -1 Then nbiotFlagValue = .Cells(target.row, nbiotFlagColNum)
    End With
    
    Dim listValue As String
    listValue = getListValue(bandwidthValue, txrxModeValue, fddTddValue, saValue, nbiotFlagValue, sh, cellTempColNum)
        
    If listValue <> "" Then
        With target.Offset(0, cellTempColNum - curCol).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=listValue
        End With
        If Not target.Offset(0, cellTempColNum - curCol).Validation.value Then
            target.Offset(0, cellTempColNum - curCol).value = ""
        End If
    Else
        Call clearValidation(target.Offset(0, cellTempColNum - bandwidthColNum))
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in Cell_Worksheet_Change," & Err.Description
End Sub

'定义设置「Cell Template」列下拉列表的事件
Public Sub Cell_Worksheet_SelectionChange(ByVal sh As Worksheet, ByVal target As range)
On Error GoTo ErrorHandler
    If target.count > 1 Or target.row <= listShtTitleRow Then Exit Sub
    
    Dim bandwidthColNum As Long, saColNum As Long, fddTddColNum As Long, txrxModeColNum As Long, nbiotFlagColNum As Long, cellTempColNum As Long
    
    saColNum = getColNum(sh.name, listShtTitleRow, "SubframeAssignment", "Cell") '「LTE Cell」页「SubframeAssignment」所在列
    fddTddColNum = getColNum(sh.name, listShtTitleRow, "FddTddInd", "Cell") '「LTE Cell」页「*FddTddInd」所在列
    txrxModeColNum = getColNum(sh.name, listShtTitleRow, "TxRxMode", "Cell") '「LTE Cell」页「*TxRxMode」所在列
    cellTempColNum = getColNum(sh.name, listShtTitleRow, "CellTemplateName", "Cell") '「LTE Cell」页「*Cell Template」所在列
    bandwidthColNum = getColNum(sh.name, listShtTitleRow, "DlBandWidth", "Cell") '「LTE Cell」页「*DlBandwidth」所在列
    nbiotFlagColNum = getColNum(sh.name, listShtTitleRow, "NbCellFlag", "Cell") '「LTE Cell」页「*NB-IoT TA Flag」所在列,有些网元模型中没有这个属性
    
    If target.column <> cellTempColNum Then Exit Sub
    
    Dim bandwidthValue As String, saValue As String, fddTddValue As String, txrxModeValue As String, nbiotValue As String
    
    If bandwidthColNum <> -1 Then bandwidthValue = Cells(target.row, bandwidthColNum).value
    If txrxModeColNum <> -1 Then txrxModeValue = Cells(target.row, txrxModeColNum).value
    If fddTddColNum <> -1 Then fddTddValue = Cells(target.row, fddTddColNum).value
    If saColNum <> -1 Then saValue = Cells(target.row, saColNum).value
    
    nbiotValue = "false"
    If nbiotFlagColNum <> -1 Then nbiotValue = Cells(target.row, nbiotFlagColNum).value
    
    Dim listValue As String
    listValue = getListValue(bandwidthValue, txrxModeValue, fddTddValue, saValue, nbiotValue, sh, target.column)
    If listValue <> "" Then
        With target.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=listValue
        End With
        If Not target.Validation.value Then
            target.value = ""
        End If
    Else
        Call clearValidation(target)
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in Cell_Worksheet_SelectionChange, " & Err.Description
End Sub

'从「MappingCellTemplate」页获取「Cell Template」列侯选值
Private Function getListValue(dlBandwidth As String, TxRxMode As String, fddTdd As String, SA As String, _
    nbioTFlag As String, ByVal sheet As Worksheet, ByVal cellTempColNum As Long) As String
On Error GoTo ErrorHandler
    Dim dlBandwidthValue As String
    dlBandwidthValue = getDlBandwidthValue(dlBandwidth)
       
    Dim fddTddValue As String
    fddTddValue = getFddTddValue(fddTdd)

    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = Worksheets("MappingCellTemplate")
    
    Dim bandwidthCol As Long, txRxModeCol As Long, fddTddCol As Long, saCol As Long, cellPatternCol As Long
    saCol = 4 'attrNameColNumInSpecialDef(mapCellTemplate, "SA")
    fddTddCol = 3 'attrNameColNumInSpecialDef(mapCellTemplate, "FDD/TDD")
    txRxModeCol = 2 'attrNameColNumInSpecialDef(mapCellTemplate, "TxRxMode")
    bandwidthCol = 1 ' attrNameColNumInSpecialDef(mapCellTemplate, "Bandwidth")
    cellPatternCol = 5 'attrNameColNumInSpecialDef(mapCellTemplate, "CellPattern")
    
    '这个新增的容器用来去重复用的，控制器和基站小区模板解耦可能会导致小区模板页有重复模板
    Dim cellTemplates As New Collection
    Dim cellTemplate As String
    Dim rowIdx As Long
    Dim listValue As String
    With mapCellTemplate
        For rowIdx = 2 To getUsedRowCount(mapCellTemplate, 1)
            If templateConditionMatch(mapCellTemplate, rowIdx, saCol, fddTddCol, txRxModeCol, _
                bandwidthCol, cellPatternCol, SA, nbioTFlag, dlBandwidthValue, fddTddValue) Then
                    cellTemplate = .Cells(rowIdx, cellPatternCol).value
                    If Not Contains(cellTemplates, cellTemplate) And cellTemplate <> "" Then
                        cellTemplates.Add item:=cellTemplate, key:=cellTemplate
                    End If
            End If
        Next
    End With
    
    listValue = collectionJoin(cellTemplates)
    
    If Len(listValue) > 255 Then
        Dim groupName As String
        Dim columnName As String
        Dim valideDef As CValideDef
        Call getGrpAndColName(sheet, cellTempColNum, groupName, columnName)
        Set valideDef = initDefaultDataSub.getInnerValideDef(sheet.name + "," + groupName + "," + columnName)
        If valideDef Is Nothing Then
            Set valideDef = addInnerValideDef(sheet.name, groupName, columnName, listValue)
        Else
            Call modiflyInnerValideDef(sheet.name, groupName, columnName, listValue, valideDef)
        End If
        listValue = valideDef.getValidedef
    End If
    getListValue = listValue
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getListValue, " & Err.Description
End Function

Private Function templateConditionMatch(mapCellTemplate As Worksheet, _
    rowIdx As Long, _
    saCol As Long, _
    fddTddCol As Long, _
    txRxModeCol As Long, _
    bandwidthCol As Long, _
    cellPatternCol As Long, _
    saValue As String, _
    nbioTFlag As String, _
    dlBandwidthValue As String, _
    fddTddValue As String) As Boolean
On Error GoTo ErrorHandler
    templateConditionMatch = False
    
    With mapCellTemplate
        If UCase(nbioTFlag) = "TRUE" And UCase(.Cells(rowIdx, fddTddCol).value) <> "NB-IOT" Then Exit Function
        
        If nbioTFlag = "" Or UCase(nbioTFlag) = "FALSE" Then
            If dlBandwidthValue <> "" And .Cells(rowIdx, bandwidthCol).value <> "" And dlBandwidthValue <> .Cells(rowIdx, bandwidthCol).value Then Exit Function
            
            If fddTddValue <> "" And .Cells(rowIdx, fddTddCol).value <> "" And .Cells(rowIdx, fddTddCol).value <> fddTddValue Then Exit Function
            
            If saValue <> "" And .Cells(rowIdx, saCol).value <> saValue Then Exit Function
        End If
    End With
    
    templateConditionMatch = True
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in templateConditionMatch, " & Err.Description
End Function

Private Sub clearValidation(target As range)
On Error GoTo ErrorHandler
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
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in clearValidation, " & Err.Description
End Sub

Private Function getDlBandwidthValue(dlBandwidth As String) As String
    Dim dlBandwidthValue As String
    Select Case dlBandwidth
        Case "CELL_BW_N6"
            dlBandwidthValue = "1.4M"
        Case "CELL_BW_N15"
            dlBandwidthValue = "3M"
        Case "CELL_BW_N25"
            dlBandwidthValue = "5M"
        Case "CELL_BW_N50"
            dlBandwidthValue = "10M"
        Case "CELL_BW_N75"
            dlBandwidthValue = "15M"
        Case "CELL_BW_N100"
            dlBandwidthValue = "20M"
        Case Else
            dlBandwidthValue = ""
    End Select
    getDlBandwidthValue = dlBandwidthValue
End Function

Private Function getFddTddValue(fddTdd As String) As String
    Dim fddTddValue As String
    Select Case fddTdd
        Case "CELL_TDD"
            fddTddValue = "TDD"
        Case "CELL_FDD"
            fddTddValue = "FDD"
        Case "CELL_NB-IoT"
            fddTddValue = "NB-IoT"
        Case Else
            fddTddValue = ""
    End Select
    getFddTddValue = fddTddValue
End Function

Private Sub getGrpAndColName(sht As Worksheet, ByVal colNum As Long, grpName As String, colName As String)
    Dim col As Long
    With sht
        colName = .Cells(listShtTitleRow, colNum).value
        For col = colNum To 1 Step -1
            If .Cells(1, col).value <> "" Then
                grpName = .Cells(1, col).value
                Exit For
            End If
        Next
    End With
End Sub

Private Function collectionJoin(coll As Collection, Optional delimiter As String = ",") As String
On Error GoTo ErrorHandler
    collectionJoin = ""
    If coll.count = 0 Then Exit Function
    
    Dim deli As String
    deli = ""
    
    Dim item
    For Each item In coll
        collectionJoin = collectionJoin & deli & CStr(item)
        deli = delimiter
    Next
    Exit Function
ErrorHandler:
    Debug.Print "some exception in collectionJoin, " & Err.Description
End Function



