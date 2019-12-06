Attribute VB_Name = "CellSub"
Option Explicit

'「eNodeB Radio Data」页记录起始行
Private Const listShtTitleRow = 2
Private Const GCellMocName As String = "GLoCell"
Private Const attrName As String = "CellTemplateName"
Private Const GCellType As String = "GSM Local Cell"

Private Const logicCellMocName As String = "GCELL"
Private Const logicAttrName As String = "TemplateName"
Private Const logicCellType As String = "GSM Logic Cell"

Private Const UCellMocName As String = "ULOCELL"
Private Const UAttrName As String = "CellTemplateName"
Private Const UCellType As String = "UMTS Local Cell"

Private Const logicUCellMocName As String = "CELL"
Private Const logicUAttrName As String = "TemplateName"
Private Const logicUCellType As String = "UMTS Logic Cell"

Private Const LCellMocName As String = "Cell"
Private Const LAttrName As String = "CellTemplateName"
Private Const LCellType As String = "LTE Cell"

Private Const MCellMocName As String = "MCell"
Private Const MAttrName As String = "CellTemplateName"
Private Const MCellType As String = "NB-IoT Cell"

Private Const RCellMocName As String = "RFALoCell"
Private Const RAttrName As String = "CellTemplateName"
Private Const RCellType As String = "RFA Cell"

Private Const NRDuCellMocName As String = "NRDUCell"
Private Const NRDuCellAttrName As String = "CellTemplateName"
Private Const NRDuCellCellType As String = "NR DU Cell"

Private Const NRCellMocName As String = "NRCell"
Private Const NRAttrName As String = "CellTemplateName"
Private Const NRCellType As String = "NR Cell"

Private Const DSACellMocName As String = "DCell"
Private Const DSAAttrName As String = "CellTemplateName"
Private Const DSACellType As String = "DCell"


Public Function isCellSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("GSM Logic Cell") Or sheetName = getResByKey("UMTS Logic Cell") Or _
        sheetName = getResByKey("RFA Cell") Or sheetName = getResByKey("LTE Cell") Or sheetName = getResByKey("NB-IoT Cell") Or _
        sheetName = getResByKey("NR Cell") Or sheetName = getResByKey("NR DU Cell") Or sheetName = getResByKey("DCell") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function
 
Public Function isTransportSheet(sheetName As String) As Boolean
    If isCellSheet(sheetName) Or sheetName = getResByKey("GTRXGROUP") Or sheetName = getResByKey("NRDUCellTrp") Or sheetName = getResByKey("NRDUCellCoverage") Then
        isTransportSheet = True
        Exit Function
    End If
    isTransportSheet = False
End Function

Public Function isGsmCellSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("GSM Logic Cell") Then
        isGsmCellSheet = True
        Exit Function
    End If
    isGsmCellSheet = False
End Function

Public Function isLteCellSheet(ByRef sheetName As String) As Boolean
    If sheetName = getResByKey("LTE Cell") Then
        isLteCellSheet = True
    Else
        isLteCellSheet = False
    End If
End Function

'定义设置「Cell Template」列下拉列表的事件
Public Sub CellSelectionChange(ByVal sheet As Worksheet, ByVal target As range)
    If target.count > 1 Or target.row <= listShtTitleRow Then Exit Sub
    
    Call initCellTemplate(sheet, target, attrName, GCellMocName, getResByKey(GCellType)) 'init 物理GCell
    
    Call initCellTemplate(sheet, target, logicAttrName, logicCellMocName, getResByKey(logicCellType)) 'init 逻辑GCell
    
    Call initCellTemplate(sheet, target, UAttrName, UCellMocName, getResByKey(UCellType)) 'init 物理UCell
    
    Call initCellTemplate(sheet, target, logicUAttrName, logicUCellMocName, getResByKey(logicUCellType)) 'init 逻辑UCell
    
    Call initCellTemplate(sheet, target, LAttrName, LCellMocName, getResByKey(LCellType))
    
    Call initCellTemplate(sheet, target, MAttrName, MCellMocName, getResByKey(MCellType))
    
    Call initCellTemplate(sheet, target, RAttrName, RCellMocName, getResByKey(RCellType))
    
    Call initCellTemplate(sheet, target, NRDuCellAttrName, NRDuCellMocName, getResByKey(NRDuCellCellType))
    
    Call initCellTemplate(sheet, target, NRAttrName, NRCellMocName, getResByKey(NRCellType))
    
    Call initCellTemplate(sheet, target, DSAAttrName, DSACellMocName, getResByKey(DSACellType))
End Sub




Private Sub initCellTemplate(sheet As Worksheet, ByVal target As range, myAttrName As String, myCellMocName As String, cellType As String)
On Error GoTo ErrorHandler
    Dim cellTemplateColNum As Long
    cellTemplateColNum = getColNum(sheet.name, listShtTitleRow, myAttrName, myCellMocName) '「物理Cell Template」所在列
    
    If cellTemplateColNum < 0 Or target.column <> cellTemplateColNum Then Exit Sub

    If cellType = getResByKey(LCellType) Then
        Call initLteCellTemplate(sheet, target, cellType)
        Exit Sub
    ElseIf cellType = getResByKey(NRCellType) Or cellType = getResByKey(NRDuCellCellType) Then
        Call initNrCellTemplate(sheet, target, myCellMocName, cellType)
        Exit Sub
    End If
    
    Dim cellTemplateListValue As String
    cellTemplateListValue = getCellTemplateListValue(cellType, sheet, target) '获取「CellTemplate」列侯选值
    If cellTemplateListValue <> "" Then
        With target.Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=cellTemplateListValue
        End With
        If Not target.Validation.value Then
            target.value = ""
        End If
    Else
        Call clearValidation(target)
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in initCellTemplate, " & Err.Description
End Sub

Private Sub initNrCellTemplate(sh As Worksheet, target As range, myCellMocName As String, cellType As String)
On Error GoTo ErrorHandler
    Dim cellMocName As String
    cellMocName = myCellMocName
    
    Dim bandwidthColNum As Long, saColNum As Long, fddTddColNum As Long, txrxModeColNum As Long, nbIotFlagColNum As Long
    saColNum = -1
    txrxModeColNum = -1
    nbIotFlagColNum = -1
    fddTddColNum = getColNum(sh.name, listShtTitleRow, "DuplexMode", cellMocName) '「NR Cell」页「*FddTddInd」所在列
    bandwidthColNum = getColNum(sh.name, listShtTitleRow, "DlBandwidth", cellMocName) '「NR Cell」页「*DlBandwidth」所在列
    
    Dim bandwidthValue As String, saValue As String, fddtddValue As String, txrxModeValue As String, nbIotCellFlag As String
    
    If bandwidthColNum <> -1 Then bandwidthValue = Cells(target.row, bandwidthColNum).value
    If txrxModeColNum <> -1 Then txrxModeValue = Cells(target.row, txrxModeColNum).value
    If fddTddColNum <> -1 Then fddtddValue = Cells(target.row, fddTddColNum).value
    If saColNum <> -1 Then saValue = Cells(target.row, saColNum).value
     
    nbIotCellFlag = "FALSE"
    If nbIotFlagColNum <> -1 Then nbIotCellFlag = Cells(target.row, nbIotFlagColNum).value
    
    Dim cellTemplateListValue As String
    cellTemplateListValue = getCellTemplateListValueEx(bandwidthValue, fddtddValue, saValue, nbIotCellFlag, sh, target, cellType)
    If cellTemplateListValue <> "" Then
        With target.Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=cellTemplateListValue
        End With
        If Not target.Validation.value Then
                target.value = ""
        End If
    Else
        Call clearValidation(target)
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in initNrCellTemplate, " & Err.Description
End Sub

Private Sub initLteCellTemplate(sh As Worksheet, target As range, cellType As String)
On Error GoTo ErrorHandler
    Dim bandwidthColNum As Long, saColNum As Long, fddTddColNum As Long, txrxModeColNum As Long, nbIotFlagColNum As Long
    
    saColNum = getColNum(sh.name, listShtTitleRow, "SubframeAssignment", "Cell")
    fddTddColNum = getColNum(sh.name, listShtTitleRow, "FddTddInd", "Cell")
    txrxModeColNum = getColNum(sh.name, listShtTitleRow, "TxRxMode", "Cell")
    bandwidthColNum = getColNum(sh.name, listShtTitleRow, "DlBandWidth", "Cell")
    
    nbIotFlagColNum = -1
    If getNBIOTFlag = True Then nbIotFlagColNum = getColNum(sh.name, listShtTitleRow, "NbCellFlag", "Cell")
    
    Dim bandwidthValue As String, saValue As String, fddtddValue As String, txrxModeValue As String, nbIotCellFlag As String
    
    If bandwidthColNum <> -1 Then bandwidthValue = Cells(target.row, bandwidthColNum).value
    If txrxModeColNum <> -1 Then txrxModeValue = Cells(target.row, txrxModeColNum).value
    If fddTddColNum <> -1 Then fddtddValue = Cells(target.row, fddTddColNum).value
    If saColNum <> -1 Then saValue = Cells(target.row, saColNum).value

    nbIotCellFlag = "FALSE"
    If nbIotFlagColNum <> -1 Then nbIotCellFlag = Cells(target.row, nbIotFlagColNum).value
    
    Dim cellTemplateListValue As String
    cellTemplateListValue = getCellTemplateListValueEx(bandwidthValue, fddtddValue, saValue, nbIotCellFlag, sh, target, cellType)
    If cellTemplateListValue <> "" Then
        With target.Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=cellTemplateListValue
        End With
        If Not target.Validation.value Then
                target.value = ""
        End If
    Else
        Call clearValidation(target)
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in initLteCellTemplate, " & Err.Description
End Sub

'从「MappingCellTemplate」页获取「Cell Template」列侯选值
Private Function getCellTemplateListValue(myType As String, sheet As Worksheet, cellRange As range) As String
On Error GoTo ErrorHandler
    Dim neType As String
    neType = getNeType()
    
    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = Worksheets("MappingCellTemplate")
    
    Dim cellTemplates As New Collection
    Dim cellTemplate As String
    
    Dim rowIdx As Integer
    With mapCellTemplate
        For rowIdx = 2 To getUsedRowCount(mapCellTemplate, 1)
            cellTemplate = .Cells(rowIdx, 1).value
            If (myType = .Cells(rowIdx, 2).value Or Trim(.Cells(rowIdx, 2).value) = "") _
                And neType = .Cells(rowIdx, 3).value And Not Contains(cellTemplates, cellTemplate) And cellTemplate <> "" Then
                    cellTemplates.Add item:=cellTemplate, key:=cellTemplate
            End If
        Next
    End With
    
    getCellTemplateListValue = collectionJoin(cellTemplates)

    If Len(getCellTemplateListValue) > 255 Then getCellTemplateListValue = getIndirectListValue(sheet, cellRange.column, getCellTemplateListValue)
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getCellTemplateListValue, " & Err.Description
End Function

Private Function getCellTemplateListValueEx(dlBandwidth As String, fddTdd As String, sa As String, nbiotFlag As String, sheet As Worksheet, cellRange As range, ByRef cellType As String) As String
On Error GoTo ErrorHandler
    Dim dlBandwidthValue As String
    Dim fddtddValue As String

    If cellType = getResByKey(LCellType) Then
        dlBandwidthValue = getLTEDlBandwidthValue(dlBandwidth)
        fddtddValue = getLTEFddTddValue(fddTdd)
    ElseIf cellType = getResByKey(NRCellType) Or cellType = getResByKey(NRDuCellCellType) Then
        dlBandwidthValue = getNRDlBandwidthValue(dlBandwidth)
        fddtddValue = getNRFddTddValue(fddTdd)
    End If

    Dim neType As String
    neType = getNeType()

    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = Worksheets("MappingCellTemplate")

    Dim cellTemplates As New Collection
    Dim cellTemplate As String

    Dim rowIdx As Long
    With mapCellTemplate
        For rowIdx = 2 To getUsedRowCount(mapCellTemplate, 1)
            If templateConditionMatch(mapCellTemplate, rowIdx, sa, nbiotFlag, dlBandwidthValue, fddtddValue, cellType, neType) Then
                cellTemplate = .Cells(rowIdx, 1).value
                If Not Contains(cellTemplates, cellTemplate) And cellTemplate <> "" Then
                    cellTemplates.Add item:=cellTemplate, key:=cellTemplate
                End If
            End If
        Next
    End With

    getCellTemplateListValueEx = collectionJoin(cellTemplates)

    If Len(getCellTemplateListValueEx) > 255 Then getCellTemplateListValueEx = getIndirectListValue(sheet, cellRange.column, getCellTemplateListValueEx)

    Exit Function
ErrorHandler:
    Debug.Print "some exception in getCellTemplateListValueEx, " & Err.Description
End Function

Private Function templateConditionMatch(mapCellTemplate As Worksheet, _
    ByVal rowIdx As Long, _
    saValue As String, _
    nbiotFlag As String, _
    dlBandwidthValue As String, _
    fddtddValue As String, _
    cellType As String, _
    neType As String) As Boolean
On Error GoTo ErrorHandler
    templateConditionMatch = False
    
    With mapCellTemplate
        If UCase(nbiotFlag) = "TRUE" And UCase(.Cells(rowIdx, 6).value) <> "NB-IOT" Then Exit Function
        
        If UCase(nbiotFlag) = "" Or UCase(nbiotFlag) = "FALSE" Then
            If dlBandwidthValue <> "" And .Cells(rowIdx, 4).value <> "" And dlBandwidthValue <> .Cells(rowIdx, 4).value Then Exit Function
            
            If fddtddValue <> "" And .Cells(rowIdx, 6).value <> "" And fddtddValue <> .Cells(rowIdx, 6).value Then Exit Function
            
            If saValue <> "" And .Cells(rowIdx, 7).value <> "" And saValue <> .Cells(rowIdx, 7).value Then Exit Function
            
            If .Cells(rowIdx, 2).value <> "" And cellType <> .Cells(rowIdx, 2).value Then Exit Function
            
            If neType <> .Cells(rowIdx, 3) Then Exit Function
        End If
    End With
    
    templateConditionMatch = True
    Exit Function
ErrorHandler:
    Debug.Print "some exception in templateConditionMatch, " & Err.Description
End Function

Private Function getNRDlBandwidthValue(dlBandwidth As String) As String
    Select Case dlBandwidth
        Case "CELL_BW_N10"
            getNRDlBandwidthValue = "10M"
        Case "CELL_BW_N15"
            getNRDlBandwidthValue = "15M"
        Case "CELL_BW_N20"
            getNRDlBandwidthValue = "20M"
        Case "CELL_BW_N40"
            getNRDlBandwidthValue = "40M"
        Case "CELL_BW_N60"
            getNRDlBandwidthValue = "60M"
        Case "CELL_BW_N80"
            getNRDlBandwidthValue = "80M"
        Case "CELL_BW_N100"
            getNRDlBandwidthValue = "100M"
        Case "CELL_BW_N200"
            getNRDlBandwidthValue = "200M"
        Case "CELL_BW_10M"
            getNRDlBandwidthValue = "10M"
        Case "CELL_BW_15M"
            getNRDlBandwidthValue = "15M"
        Case "CELL_BW_20M"
            getNRDlBandwidthValue = "20M"
        Case "CELL_BW_40M"
            getNRDlBandwidthValue = "40M"
        Case "CELL_BW_60M"
            getNRDlBandwidthValue = "60M"
        Case "CELL_BW_80M"
            getNRDlBandwidthValue = "80M"
        Case "CELL_BW_100M"
            getNRDlBandwidthValue = "100M"
        Case "CELL_BW_200M"
            getNRDlBandwidthValue = "200M"
        Case "CELL_BW_30M"
            getNRDlBandwidthValue = "30M"
        Case "CELL_BW_50M"
            getNRDlBandwidthValue = "50M"
        Case "CELL_BW_70M"
            getNRDlBandwidthValue = "70M"
        Case "CELL_BW_90M"
            getNRDlBandwidthValue = "90M"
        Case Else
            getNRDlBandwidthValue = ""
    End Select
End Function

Private Function getNRFddTddValue(fddTdd As String) As String
    Select Case fddTdd
        Case "CELL_TDD"
            getNRFddTddValue = "TDD"
        Case "CELL_FDD"
            getNRFddTddValue = "FDD"
        Case "CELL_SUL"
             getNRFddTddValue = "SUL"
        Case Else
            getNRFddTddValue = ""
    End Select
End Function

Private Function getLTEDlBandwidthValue(dlBandwidth As String) As String
    Select Case dlBandwidth
        Case "CELL_BW_N6"
            getLTEDlBandwidthValue = "1.4M"
        Case "CELL_BW_N15"
            getLTEDlBandwidthValue = "3M"
        Case "CELL_BW_N25"
            getLTEDlBandwidthValue = "5M"
        Case "CELL_BW_N50"
            getLTEDlBandwidthValue = "10M"
        Case "CELL_BW_N75"
            getLTEDlBandwidthValue = "15M"
        Case "CELL_BW_N100"
            getLTEDlBandwidthValue = "20M"
        Case Else
            getLTEDlBandwidthValue = ""
    End Select
End Function

Private Function getLTEFddTddValue(fddTdd As String) As String
    Select Case fddTdd
        Case "CELL_TDD"
            getLTEFddTddValue = "TDD"
        Case "CELL_FDD"
            getLTEFddTddValue = "FDD"
        Case Else
            getLTEFddTddValue = ""
    End Select
End Function
