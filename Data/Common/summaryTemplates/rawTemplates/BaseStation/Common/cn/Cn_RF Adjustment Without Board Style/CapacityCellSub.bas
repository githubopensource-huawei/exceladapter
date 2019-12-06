Attribute VB_Name = "CapacityCellSub"
Option Explicit

Private Const CELL_BAR_NAME = "SectorEqmBar"
Private Const DELFREQ_BAR_NAME = "DeleteTrxBar"
Private Const BATCH_DELFREQ_BAR_NAME = "BatchDeleteFreqBar"
Private Const FINISH_BAR_NAME = "AdjustFinishBar"
Private Const CANCEL_BAR_NAME = "AdjustCancelBar"
Private Const Col_Width = 12
Private SITE_NAME As String
Private WRITESUCCESS As Boolean
Public InAdjustAntnPort As Boolean
Private CELL_SHEET_NAME As String
Public CELL_TYPE As Long '0代表是GSM小区,1是UMTS小区,2是LTE小区,对于GSM小区,需要根据小区单独判断类型Cell_CellType_Map.
Public IsCloudRAN_DU As Boolean

Private ROW_COUNT As Long
Private cellValueStr As String
Private trxValueStr As String
Private boardValueStr As String
Private antnValueStr As String
Private modelValueStr As String
Private sectoreqmValueStr As String
Private valueMap As Collection
Private errCollect As Collection
Private dupCollect As Collection

Private Cell_Row_Map As CMap  'GSM Cell and Row Mapping
Private Cell_CellType_Map As CMap  'GSM Cell Type
Private Cell_TrxListInfo As CMap
Private GCell_TrxListInfo As CMap  'key:BTSNAME#CELLID
Private Cell_Index As Long
Private Trx_Index As Long
Private Sector_Index As Long
Private Board_Index As Long
Private Ante_Index As Long
Private Model_Index As Long
Private BaseEqm_Index As Long
Private SectorEqm_Index As Long
Private SectorEqmGrpId_Index As Long
Private CellBeamMode_Index As Long

Private BTSName_Index As Long
Private selectBtsNameCol As Collection '用户选择基站名称列表
Private btsNameRowCountMap As CMap 'key:BTSNAME#CELLNAME,value:RowCount，基站名称和临时页签行数的映射
Private tempShtStartRow As Long '临时页签boardNo起始行记录
Private btsNameBrdNoMap As CMap 'key:BTSNAME,value:boardNo.，基站名称和单板列表的映射


Sub popUpTempSheetCannotChangeMsgbox()
    Call MsgBox(getResByKey("TempSheetCannotChange"), vbInformation, getResByKey("Warning"))
    ThisWorkbook.Worksheets(getResByKey("Temp Sheet")).Select
End Sub

Sub popUpTempAdjustSheetCannotChangeMsgbox()
    Call MsgBox(getResByKey("TempSheetCannotChange"), vbInformation, getResByKey("Warning"))
    ThisWorkbook.Worksheets(getResByKey("Temp_Adjust_Sheet")).Select
End Sub

Sub popUpTempSheetCannotSaveMsgbox()
    Call MsgBox(getResByKey("TempSheetCannotSave"), vbInformation, getResByKey("Warning"))
End Sub

Sub changeRangeColor(sheet As Worksheet, cRange As range)
    Dim rStr As String
    Dim rowNum As Long
    rStr = cRange.address(False, False)
    If Contains(errCollect, rStr) Then
        cRange.Interior.colorIndex = -4142
        Exit Sub
    End If
    
    rowNum = cRange.row
    Dim rowSet As Variant
    Dim tVal As Variant
    Dim sVal As Variant
    Dim lineStr As String
    Dim maxColLen As Long
    maxColLen = getColMaxLength(sheet)
    For Each tVal In dupCollect
        If InStr(1, tVal, CStr(rowNum)) <> 0 Then
            rowSet = Split(tVal, ",")
            If UBound(rowSet) = 1 Then
                For Each sVal In rowSet
                    rowNum = CLng(sVal)
                    sheet.range(Cells(rowNum, Cell_Index), Cells(rowNum, maxColLen)).Interior.colorIndex = -4142
                Next
            Else
                sheet.range(Cells(rowNum, Cell_Index), Cells(rowNum, maxColLen)).Interior.colorIndex = -4142
                dupCollect.Remove (tVal)
                For Each sVal In rowSet
                    If sVal <> str(rowNum) Then
                        If lineStr = "" Then
                            lineStr = sVal
                        Else
                            lineStr = lineStr + "," + sVal
                        End If
                    End If
                Next
                dupCollect.Add Item:=lineStr, key:=lineStr
            End If
            Exit Sub
        End If
    Next
End Sub

Private Function getColMaxLength(sheet As Worksheet) As Long
    getColMaxLength = sheet.range("XFD1").End(xlToLeft).column
End Function

'?ó??é?3y?μμ?oí?úé?3y?μμ?°′?￥
Public Sub createDelFreqBar()
    Dim delChooseBar As CommandBar
    Dim delFreqStyle As CommandBarButton
    Dim BatchdelFreqStyle As CommandBarButton
    Dim BatchdelChooseBar As CommandBar
    
    'Call deleteCellBar
    
    Dim actSheetName As String
    actSheetName = ThisWorkbook.ActiveSheet.name
    CELL_SHEET_NAME = actSheetName
    CELL_TYPE = cellSheetType(actSheetName)
      
      If CELL_TYPE = 0 Then
            Set delChooseBar = Application.CommandBars.Add(DELFREQ_BAR_NAME, msoBarTop)
            Dim delbarDescLbl As String
            delbarDescLbl = "DeleteFreq"
            With delChooseBar
               .Protection = msoBarNoResize
               .Visible = True
               Set delFreqStyle = .Controls.Add(Type:=msoControlButton)
               With delFreqStyle
                    .Style = msoButtonIconAndCaption
                    .Caption = getResByKey(delbarDescLbl)
                    .TooltipText = getResByKey(delbarDescLbl)
                    .OnAction = "deleteFrequency"
                    .FaceId = 186
                    .Enabled = True
                End With
              End With
                        Set BatchdelChooseBar = Application.CommandBars.Add(BATCH_DELFREQ_BAR_NAME, msoBarTop)
                        Dim BatchdelbarDescLbl As String
                        BatchdelbarDescLbl = "BatchDeleteFreq"
                        With BatchdelChooseBar
                                        .Protection = msoBarNoResize
                                        .Visible = True
                                        Set BatchdelFreqStyle = .Controls.Add(Type:=msoControlButton)
                                        With BatchdelFreqStyle
                                                .Style = msoButtonIconAndCaption
                                                .Caption = getResByKey(BatchdelbarDescLbl)
                                                .TooltipText = getResByKey(BatchdelbarDescLbl)
                                                .OnAction = "BatchdeleteFrequency"
                                                .FaceId = 186
                                                .Enabled = True
                                        End With
                        End With
          
        End If
End Sub

Sub createCellBar()
    Dim baseStationChooseBar As CommandBar
    Dim baseStationStyle As CommandBarButton
    
    Call deleteCellBar
    Set baseStationChooseBar = Application.CommandBars.Add(CELL_BAR_NAME, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("AdjustCellAntnPort")
            .TooltipText = getResByKey("AdjustCellAntnPort")
            .OnAction = "baseStationChoose"
            .FaceId = 186
            .Enabled = True
        End With
      End With
End Sub

Private Sub baseStationChoose()
    Dim actSheetName As String
    
    actSheetName = ThisWorkbook.ActiveSheet.name
    
    CELL_SHEET_NAME = actSheetName
    CELL_TYPE = cellSheetType(actSheetName)
    If hasEuDuLocalCellRes Then IsCloudRAN_DU = True
    
    MuliBtsFilterForm.Show
End Sub


Private Sub deleteFrequency()
    On Error GoTo ErrorHandler
    
    DeleteFreqForm.Show
    Exit Sub
ErrorHandler:
End Sub

Private Sub BatchdeleteFrequency()
    On Error GoTo ErrorHandler
    
    BatchDeleteFreqForm.Show
    Exit Sub
ErrorHandler:
End Sub

Sub deleteCellBar()
    If existToolBar(CELL_BAR_NAME) Then
        Application.CommandBars(CELL_BAR_NAME).Delete
    End If
    If existToolBar(DELFREQ_BAR_NAME) Then
        Application.CommandBars(DELFREQ_BAR_NAME).Delete
    End If
        
        If existToolBar(BATCH_DELFREQ_BAR_NAME) Then
        Application.CommandBars(BATCH_DELFREQ_BAR_NAME).Delete
    End If
End Sub
 
 Sub initTempSheetControl(ByRef flag As Boolean)
    On Error Resume Next
    Dim k As Long
    Dim controlId As Long
    For k = 1 To Application.CommandBars("Ply").Controls.count
        controlId = Application.CommandBars("Ply").Controls(k).ID
        Application.CommandBars("Ply").FindControl(ID:=controlId).Enabled = flag
    Next
    With Application.CommandBars("Column")
        .FindControl(ID:=3183).Enabled = flag
        .FindControl(ID:=297).Enabled = flag
        .FindControl(ID:=294).Enabled = flag
    End With
End Sub

Sub createTempBar()
    Dim finishBar As CommandBar
    Dim finishStyle As CommandBarButton
    Dim cancelStyle As CommandBarButton
    If ThisWorkbook.ActiveSheet.name <> getResByKey("Temp Sheet") Then
        Exit Sub
    End If
    
    Set finishBar = Application.CommandBars.Add(FINISH_BAR_NAME, msoBarTop)
    With finishBar
        .Protection = msoBarNoResize
        .Visible = True
        Set finishStyle = .Controls.Add(Type:=msoControlButton)
        With finishStyle
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Finish")
            .TooltipText = getResByKey("Finish")
            .OnAction = "writeBackData"
            .FaceId = 186
            .Enabled = True
        End With
        Set cancelStyle = .Controls.Add(Type:=msoControlButton)
        With cancelStyle
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Cancel")
            .TooltipText = getResByKey("Cancel")
            .OnAction = "deleteTempSheet"
            .FaceId = 186
            .Enabled = True
        End With
      End With
      
End Sub

Sub deleteTempBar()
    If existToolBar(FINISH_BAR_NAME) Then
        Application.CommandBars(FINISH_BAR_NAME).Delete
    End If
    If existToolBar(CANCEL_BAR_NAME) Then
        Application.CommandBars(CANCEL_BAR_NAME).Delete
    End If
End Sub


Private Function existToolBar(ByRef barName As String) As Boolean
    On Error GoTo ErrorHandler
    existToolBar = True
    Dim bar As CommandBar
    Set bar = CommandBars(barName)
    Exit Function
ErrorHandler:
    existToolBar = False
End Function


Sub deleteFreqAssoEqm(cellRow As Long, freqIndex As Long)
On Error GoTo ErrorHandler

    If CELL_TYPE <> 0 Then
        Exit Sub
    End If
    
    If cellRow < 3 Then
        GoTo ErrorHandler
    End If
    
    Dim cellsheet As Worksheet
    Set cellsheet = ThisWorkbook.ActiveSheet
    
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim antnIndex As Long
    Dim sectorIndex As Long
    Dim antnStr As String
    Dim sectorStr As String
    
    '获取天线端口所在列
    Call getAntenneMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antnIndex = findColumnByName(cellsheet, columnName, 2)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    sectorIndex = findColumnByName(cellsheet, columnName, 2)
    
    antnStr = cellsheet.Cells(cellRow, antnIndex).value
    sectorStr = cellsheet.Cells(cellRow, sectorIndex).value
    
    Dim antnArray() As String
    Dim sectorArray() As String
    
    antnArray = Split(antnStr, ";")
    sectorArray = Split(sectorStr, ";")
    
    If UBound(antnArray) <> UBound(sectorArray) Then
        GoTo ErrorHandler
    End If
    
    If UBound(antnArray) < freqIndex Then
        GoTo ErrorHandler
    End If
    
    Dim index As Long
    Dim valueStr As String
    valueStr = "-1"
    For index = LBound(antnArray) To UBound(antnArray)
        If index <> freqIndex Then
            If valueStr = "-1" Then
                valueStr = antnArray(index)
            Else
                valueStr = valueStr + ";" + antnArray(index)
            End If
        End If
    Next
    
    If valueStr = "-1" Then
        valueStr = ""
    End If
    cellsheet.Cells(cellRow, antnIndex).value = valueStr
    
    valueStr = "-1"
    For index = LBound(sectorArray) To UBound(sectorArray)
        If index <> freqIndex Then
            If valueStr = "-1" Then
                valueStr = sectorArray(index)
            Else
                valueStr = valueStr + ";" + sectorArray(index)
            End If
        End If
    Next
    
    If valueStr = "-1" Then
        valueStr = ""
    End If
    cellsheet.Cells(cellRow, sectorIndex).value = valueStr
    
ErrorHandler:
End Sub

Sub AddSectorEqm(ByRef selectedMocCol As Collection, ByRef CellSheetName As String)
    On Error GoTo ErrorHandler
    Set valueMap = New Collection
    'SITE_NAME = siteName
    Set btsNameBrdNoMap = New CMap
    Set selectBtsNameCol = selectedMocCol
    WRITESUCCESS = False
    Call judgeGNormalCell
    
    Dim chkPassed As Boolean
    chkPassed = True

    ROW_COUNT = calculateRow(chkPassed)
    If Not chkPassed Then
        Exit Sub
    End If
    
    ThisWorkbook.Worksheets.Add after:=ThisWorkbook.ActiveSheet
    ThisWorkbook.ActiveSheet.name = getResByKey("Temp Sheet")

    Call createTempBar
    Call initMenuStatus(ThisWorkbook.ActiveSheet)
    InAdjustAntnPort = True
    
    Call insertBtsNameColumn
    
    If Not insertCellIdColumn() Then
        Exit Sub
    End If
    If Not insertSectorIdColumn() Then
        Exit Sub
    End If
    
    Dim temBtsName As Variant
    tempShtStartRow = 2
    '只刷RXUBoardColumn的列头
    If Not insertRxuBoardColumn() Then
        Exit Sub
    End If
    
    If CELL_TYPE = 0 Then
        Call insertTrxColumn
    End If
    
    Call insertAntenneColumn
    Call insertAnteModelColumn
    If CELL_TYPE = 2 Then
        If Not insertBaseEqmColumn() Then
            Exit Sub
        End If
        If Not insertSectorEqmGrpIdColumn() Then
        End If
        If Not insertCellBeamModeColumn() Then
        End If
    End If
    If CELL_TYPE = 1 Then
        If Not insertSectoreqmColumn() Then
        End If
    End If
    
    Call writeData
    
    '只需调用refreshBtsRxuBoardMap，当前防止有改漏Thisworkbook的处理，现在还是会插入list
    'Call refreshBtsRxuBoardMap
'    Call insertRxuBoardList4AllSheet
    
    Call AdjustSheetStyle
    WRITESUCCESS = True
    Exit Sub
ErrorHandler:
    WRITESUCCESS = False
End Sub


Private Sub judgeGNormalCell()
    On Error GoTo ErrorHandler
    Dim mocName As String
    Dim attrName As String
    Dim cellTypeIndex As Long
    Dim cellTypeVal As String
    
    Dim rowNum As Long
    Dim maxRow As Long
    Dim cellsheet As Worksheet
    Dim cellId As String
    Dim temBtsName As Variant
    Dim mapKey As String
    
    Set Cell_Row_Map = New CMap
    Set Cell_CellType_Map = New CMap
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Call getCellMocNameAndAttrName(mocName, attrName)
    '判断是否为EuCellSectorEqm或EuPrbSectorEqm页签，如果是则mocName重新赋值
    If CELL_SHEET_NAME = "EUCELLSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUCELLSECTOREQM") Then
        mocName = "EuCellSectorEqm"
        attrName = "VLOCALCELLID"
    ElseIf CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
        mocName = "EuPrbSectorEqm"
        attrName = "VLOCALCELLID"
    End If
    cellTypeIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    maxRow = cellsheet.range("a1048576").End(xlUp).row
    
    For Each temBtsName In selectBtsNameCol
        For rowNum = 3 To maxRow
            If cellsheet.Cells(rowNum, 1).value = CStr(temBtsName) Then
                cellId = cellsheet.Cells(rowNum, cellTypeIndex).value
                mapKey = CStr(temBtsName) + "*" + cellId
                If Cell_Row_Map.hasKey(mapKey) = False Then
                    Call Cell_Row_Map.SetAt(mapKey, rowNum)
                End If
                
                If Cell_CellType_Map.hasKey(mapKey) = False Then
                    Call Cell_CellType_Map.SetAt(mapKey, CELL_TYPE)
                End If
            End If
        Next
    Next
    
    mocName = "GLOCELL"
    attrName = "LOCELLTYPE"
    cellTypeIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    Dim tVal As Variant
    For Each tVal In Cell_CellType_Map.KeyCollection
        If Cell_Row_Map.hasKey(tVal) Then
            rowNum = CLng(Cell_Row_Map.GetAt(tVal))
            cellTypeVal = cellsheet.Cells(rowNum, cellTypeIndex).value
            If cellTypeVal = "MULTISITE_CELL" Then
                Call Cell_CellType_Map.RemoveKey(tVal)
                Call Cell_CellType_Map.SetAt(tVal, 4) '当前是共小区,CellType为4
            End If
        End If
    Next
    Exit Sub
ErrorHandler:
End Sub

Private Sub AdjustSheetStyle()
    Dim tmpsheet As Worksheet
    Dim sheetRange As range
    Dim maxColLen As Long
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    
    '设置临时页签单元格字体格式
    With tmpsheet.Cells.Font
        .name = "Arial"
        .Size = 10
    End With
    
    '设置临时页签单元格格式
    With tmpsheet.Cells
        .RowHeight = 14
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Columns.AutoFit
    End With
    
    '设置临时页签边框（预留空行）
    maxColLen = getColMaxLength(tmpsheet)
    With tmpsheet.range(Cells(1, 1), Cells(ROW_COUNT + 3, maxColLen))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders.LineStyle = xlContinuous
    End With
    
    '刷新批注，自适应调整批注框大小
    Dim maxColumnNumber As Long
    maxColumnNumber = tmpsheet.range("XFD1").End(xlToLeft).column
    Call refreshComment(tmpsheet.range(tmpsheet.Cells(1, 1), tmpsheet.Cells(1, maxColumnNumber)), True)
End Sub

Private Sub writeData()
    Dim cell As CAntennes
    Dim row As Variant
    Dim index As Long
    Dim sflag As Boolean
    Dim boardStr As String
    Dim antnStr As String
    Dim rsModel As String
    Dim tempSheet As Worksheet
    Dim antenneCollection As Collection
    Dim rangeStr As String
    Dim rowVal As Long
    sflag = True
    Set tempSheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    index = 2
    For Each cell In valueMap
    
        If cell.antennes = "" Or cell.sectorIds = "" Then
            GoTo NextLoop
        End If
        
        Set antenneCollection = cell.getAntenneCollection(sflag, rowVal)
        If sflag = False Then
            Call deleteTempSheet
            ThisWorkbook.Worksheets(CELL_SHEET_NAME).rows(rowVal).Select
            Exit Sub
        End If
        For Each row In antenneCollection
            
            boardStr = row(3)
            antnStr = row(4)
            rsModel = row(5)
            rangeStr = row(6)
            If checkBoardBasedValid(boardStr) = False Or (Trim(antnStr) <> "" And checkDataValid(antnValueStr, antnStr) = False) _
            Or (Trim(rsModel) <> "" And checkDataValid(modelValueStr, rsModel) = False) Then
                Call MsgBox(getResByKey("wrongDataInput") & rangeStr, vbInformation, getResByKey("Warning"))
                Call deleteTempSheet
                ThisWorkbook.Worksheets(CELL_SHEET_NAME).range(rangeStr).Select
                Exit Sub
            End If
            tempSheet.Cells(index, BTSName_Index).value = row(0)
            tempSheet.Cells(index, Cell_Index).value = row(1)
            If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
                tempSheet.Cells(index, Trx_Index).value = row(7)
            End If
            tempSheet.Cells(index, Sector_Index).value = row(2)
            tempSheet.Cells(index, Board_Index).value = boardStr
            tempSheet.Cells(index, Ante_Index).value = antnStr
            tempSheet.Cells(index, Model_Index).value = rsModel
            If CELL_TYPE = 2 Then
                tempSheet.Cells(index, BaseEqm_Index).value = row(8)
                If row(10) <> "" Then tempSheet.Cells(index, SectorEqmGrpId_Index).value = row(10)
                If row(11) <> "" Then tempSheet.Cells(index, CellBeamMode_Index).value = row(11)
            End If
            If CELL_TYPE = 1 And row(9) <> "" Then
                tempSheet.Cells(index, SectorEqm_Index).value = row(9)
            End If
            index = index + 1
        Next
NextLoop:
    Next cell
End Sub

Private Function checkBoardBasedValid(tVal As String) As Boolean
        If tVal = "" Then
            checkBoardBasedValid = True
            Exit Function
        End If
        Dim tCount As Long
        Dim strArray() As String
        strArray = Split(tVal, "_")
        tCount = UBound(strArray)
        If tCount = 3 Then
            checkBoardBasedValid = True
        Else
            checkBoardBasedValid = False
        End If
End Function


Private Function checkDataValid(tValueStr As String, tVal As String, Optional ignoreEmptyStr As Boolean = False) As Boolean
        checkDataValid = True
        If ignoreEmptyStr = True And Len(tVal) = 0 Then
            Exit Function
        End If
        
        Dim index As Long
        Dim strArray() As String
        strArray = Split(tValueStr, ",")
        For index = LBound(strArray) To UBound(strArray)
              If strArray(index) = tVal Then
                checkDataValid = True
                Exit Function
                End If
        Next
        checkDataValid = False
End Function

Private Sub insertBtsNameColumn()
    
    Dim cellsheet As Worksheet
    Dim btsNameColIndex As Long
    Dim mocName As String
    Dim attrName As String
    
    Call getBaseStationMocNameAndAttrName(mocName, attrName)
    btsNameColIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    BTSName_Index = 1
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    cellsheet.Cells(2, btsNameColIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, BTSName_Index)
End Sub


Private Function insertCellIdColumn() As Boolean
        Dim myAttrName As String
        Dim myCellMocName As String
        Dim constCellTempCol As Long
        Dim multiRruCellModeIndex As Long
        Dim mocName As String
        Dim attrName As String
    
        Call getCellMocNameAndAttrName(myCellMocName, myAttrName)
        '判断是否为EuCellSectorEqm或EuPrbSectorEqm页签，如果是则mocName重新赋值
        If CELL_SHEET_NAME = "EUCELLSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUCELLSECTOREQM") Then
            myCellMocName = "EuCellSectorEqm"
            myAttrName = "VLOCALCELLID"
        ElseIf CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
            myCellMocName = "EuPrbSectorEqm"
            myAttrName = "VLOCALCELLID"
        End If
        constCellTempCol = getColNum(CELL_SHEET_NAME, 2, myAttrName, myCellMocName)
        '???RRU??????
        Call getCellMultiRruCellModeMocAndAttrName(myCellMocName, myAttrName)
        multiRruCellModeIndex = getColNum(CELL_SHEET_NAME, 2, myAttrName, myCellMocName)
        
        Dim cellsheet As Worksheet
        insertCellIdColumn = True
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Cell_Index = 2
        cellsheet.Cells(2, constCellTempCol).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Cell_Index)
        Dim cellsStr As String
        cellsStr = ""
        
        Dim btsNameColIndex As Long
        Dim temBtsName As String
        Dim temCellId As String
        
        Call getBaseStationMocNameAndAttrName(mocName, attrName)
        btsNameColIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
        
        Dim index As Long
        For index = 3 To cellsheet.range("a1048576").End(xlUp).row
            temBtsName = cellsheet.Cells(index, btsNameColIndex).value
            If cellsheet.Cells(index, 1).value = temBtsName And cellsheet.Cells(index, 2).value <> "RMV" Then
                If multiRruCellModeIndex > 0 Then
                    If cellsheet.Cells(index, multiRruCellModeIndex).value = "MPRU_AGGREGATION" Or cellsheet.Cells(index, multiRruCellModeIndex).value = "MIXED_MULTIRRU_CELL" Then
                        GoTo NextLoop
                    End If
                End If
                temCellId = cellsheet.Cells(index, constCellTempCol).value
                If cellsStr <> "" Then
                     cellsStr = cellsStr + "," + temCellId
                Else
                    cellsStr = temCellId
                End If
            End If
NextLoop:
        Next

'        Dim cellRang As range
'        Set cellRang = ThisWorkbook.ActiveSheet.range("A2:A" + CStr(2 + ROW_COUNT + 3))
        If cellsStr <> "" Then
                cellValueStr = cellsStr
'                With cellRang.Validation
'                   .delete
'                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=cellsStr
'                End With
'                ThisWorkbook.ActiveSheet.Columns(Cell_Index).EntireColumn.AutoFit
        End If
        insertCellIdColumn = True
End Function

Public Sub insertTrxColumn()
        Dim myAttrName As String
        Dim myCellMocName As String
        Dim constCellTempCol As Long
        Dim trxlistStr As String
        Dim cellId As String
        Dim constTrxCol As Long
        Dim constBTrxCol As Long
        Dim alltrxlistStr As String
        
        Dim cellsheet As Worksheet
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Set GCell_TrxListInfo = New CMap
        
        Call getCellMocNameAndAttrName(myCellMocName, myAttrName)
        constCellTempCol = getColNum(CELL_SHEET_NAME, 2, myAttrName, myCellMocName)
        constTrxCol = getColNum(CELL_SHEET_NAME, 2, "NONBCCHFREQLIST", "TRXINFO")
        constBTrxCol = getColNum(CELL_SHEET_NAME, 2, "BCCHFREQ", "TRXINFO")
        
        Dim index As Long
        Dim keyStr As String
        Dim temBtsName As Variant
        For Each temBtsName In selectBtsNameCol
            For index = 2 To cellsheet.range("a1048576").End(xlUp).row
                If cellsheet.Cells(index, 1).value = temBtsName And cellsheet.Cells(index, 2).value <> "RMV" Then
                    cellId = cellsheet.Cells(index, constCellTempCol).value
                    trxlistStr = cellsheet.Cells(index, constTrxCol).value + "," + cellsheet.Cells(index, constBTrxCol).value
                    
                    keyStr = CStr(temBtsName) + "*" + cellId
                    If GCell_TrxListInfo.hasKey(keyStr) = False Then
                        Call GCell_TrxListInfo.SetAt(keyStr, trxlistStr)
                    End If
                    
                    If trxValueStr <> "" Then
                         trxValueStr = trxValueStr + "," + trxlistStr
                    Else
                        trxValueStr = trxlistStr
                    End If
                    
                End If
            Next
        Next
        Trx_Index = 3
        ThisWorkbook.ActiveSheet.Cells(1, Board_Index).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Trx_Index)
        ThisWorkbook.ActiveSheet.Cells(1, Trx_Index).value = getResByKey("Frequency")
        Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, Trx_Index), getResByKey("FREQ"))

End Sub


Private Function insertSectorIdColumn() As Boolean
        Dim mocName As String
        Dim attrName As String
        Dim sectorColumnName As String
        Dim columnIndex As Long
        Dim cellsheet As Worksheet
        
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Call getSectorMocNameAndAttr(mocName, attrName)
        If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
            mocName = "EuPrbSectorEqm"
        End If
        sectorColumnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex < 0 Then
            insertSectorIdColumn = False
            Exit Function
        End If
        If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
            Sector_Index = 4
        Else
            Sector_Index = 3
        End If
        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Sector_Index)
        Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, Sector_Index), getResByKey("SECTOR_SECTORID"))
        ThisWorkbook.ActiveSheet.Columns(Sector_Index).EntireColumn.AutoFit
        insertSectorIdColumn = True
End Function

Private Function insertRxuBoardColumn() As Boolean
    Dim temBtsName As Variant
    Dim brdStyleSheetName As String
    Dim brdStyleSheet As Worksheet
    Dim grpCollection As Collection
    Dim brdGrp
    Dim startRow As Long
    Dim endRow As Long
    Dim boardNoIndex As Long
    Dim charStr As String
  
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        Board_Index = 5
        charStr = "E"
    Else
        Board_Index = 4
        charStr = "D"
    End If

    insertRxuBoardColumn = False
    
    '找任意一个boardStyle，取RXUBoardColumn的信息
    For Each temBtsName In selectBtsNameCol
               
        brdStyleSheetName = findBoardStyleSheetByBtsName(CStr(temBtsName))
        If brdStyleSheetName = "" Then
            Call MsgBox(getResByKey("NoBoradStyle"), vbInformation, getResByKey("Warning"))
            Call deleteTempSheet
            insertRxuBoardColumn = False
            Exit Function
        End If
        
        Set brdStyleSheet = ThisWorkbook.Worksheets(brdStyleSheetName)
        
        Set grpCollection = findBrdGroups
        For Each brdGrp In grpCollection
            Call getGroupStartAndEndRowByGroupName(brdStyleSheet, CStr(brdGrp), startRow, endRow)
            If startRow <> -1 Then
                boardNoIndex = getboradNoColumNumber(brdStyleSheet, startRow + 1, CStr(brdGrp))
                brdStyleSheet.Cells(startRow + 1, boardNoIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Board_Index)
                Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, Board_Index), getResByKey("RxuBoard"))
                insertRxuBoardColumn = True
                Exit Function
            End If
        Next
    Next
    
End Function

Public Sub insertRxuBoardList(ByRef ws As Worksheet, ByRef cellRange As range)
    If cellRange.rows.count <> 1 And cellRange.Columns.count <> 1 Then Exit Sub
    If cellRange.column <> Board_Index Then Exit Sub
    
    Dim btsName As String
    btsName = ws.Cells(cellRange.row, 1)
    If btsName = "" Then
        cellRange.Validation.Delete
        Exit Sub
    End If
    
    Call insertRxuBoardListByBtsName(btsName, cellRange)
End Sub

Private Sub insertRxuBoardList4AllSheet()
    '先刷新所有基站的BoardNoList
    Call refreshBtsRxuBoardMap
    
    Dim ws As Worksheet
    Dim maxRowNumber As Long
    Dim maxColNumber As Long
    
    Set ws = ThisWorkbook.ActiveSheet
    maxRowNumber = ws.UsedRange.rows.count
    maxColNumber = ws.UsedRange.Columns.count
           
    '取临时页签所有数据
    Dim DataRange As Variant
    DataRange = ws.range("A1:" + getColStr(maxColNumber) + CStr(maxRowNumber)).value '一次从Excel单元格中读取所有的值,将其放入数组

    '默认临时页签从第2行开始写数据
    Dim btsName As String
    Dim rowIndex As Long
    For rowIndex = 2 To maxRowNumber
        btsName = DataRange(rowIndex, 1)
        Call insertRxuBoardListByBtsName(btsName, ws.Cells(rowIndex, Board_Index))
    Next rowIndex
    
End Sub


Private Sub insertRxuBoardListByBtsName(ByRef btsName As String, ByRef cellRang As range)
    Dim brdStr As String
    If btsNameBrdNoMap.hasKey(btsName) Then
        brdStr = btsNameBrdNoMap.GetAt(btsName)
    Else
        brdStr = findBoardNosStrByBtsName(btsName)
        Call btsNameBrdNoMap.SetAt(btsName, brdStr)
    End If
    
    If Len(brdStr) = 0 Then Exit Sub
            
    '插入List
    With cellRang.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=brdStr
    End With

End Sub


Public Function refreshBtsRxuBoardMap() As String
    
    Dim grpCollection As Collection
    Set grpCollection = findBrdGroups
        
    Dim brdStyleSheetName As String
    Dim brdStyleSheet As Worksheet
    
    Dim btsName As String
    Dim brdStr As String
    
    Dim btsVal As Variant
    For Each btsVal In selectBtsNameCol
        btsName = CStr(btsVal)
        brdStyleSheetName = findBoardStyleSheetByBtsName(btsName)
        If Len(brdStyleSheetName) <> 0 Then
            Set brdStyleSheet = ThisWorkbook.Worksheets(brdStyleSheetName)
            brdStr = findBoardNosStrByBtsName_i(brdStyleSheet, grpCollection, btsName)
            Call btsNameBrdNoMap.SetAt(btsName, brdStr)
        End If
    Next
    
End Function


Private Function findBoardNosStrByBtsName(ByRef btsName As String) As String
    findBoardNosStrByBtsName = ""
    
    Dim brdStyleSheetName As String
    Dim brdStyleSheet As Worksheet
    
    brdStyleSheetName = findBoardStyleSheetByBtsName(btsName)
    If brdStyleSheetName = "" Then
        Call MsgBox(getResByKey("NoBoradStyle"), vbInformation, getResByKey("Warning"))
        'Call deleteTempSheet
        Exit Function
    End If
    
    Set brdStyleSheet = ThisWorkbook.Worksheets(brdStyleSheetName)
    
    Dim grpCollection As Collection
    Set grpCollection = findBrdGroups

    findBoardNosStrByBtsName = findBoardNosStrByBtsName_i(brdStyleSheet, grpCollection, btsName)
    
End Function


Public Function findBoardNosStrByBtsName_i(ByRef brdStyleSheet As Worksheet, ByRef grpCollection As Collection, ByRef btsName As String) As String
    Dim startRow As Long
    Dim endRow As Long
    Dim index As Long
    Dim boardNoIndex As Long
    Dim brdGrp
    Dim brdStr

    findBoardNosStrByBtsName_i = ""
    brdStr = ""
    '先判断是否为场景列功能，是场景列功能，初始化BoardStyle和BaseStation页签数据，并且判断场景化isCustomMatchRow
    If isCustomSheet Then
        '初始化数据
        Set baseStationData = New CBaseStationData
        Set boardStyleData = New CBoardStyleData
        Call baseStationData.init
        Call boardStyleData.init
        
        For Each brdGrp In grpCollection
            Call getGroupStartAndEndRowByGroupName(brdStyleSheet, CStr(brdGrp), startRow, endRow)
            If startRow = -1 Then GoTo NextLoopCustom
            
            boardNoIndex = getboradNoColumNumber(brdStyleSheet, startRow + 1, CStr(brdGrp))
            
            For index = startRow + 2 To endRow
                If isCustomMatchRow(btsName, brdStyleSheet, CStr(brdGrp), index) Then
                    If brdStr = "" Then
                        brdStr = brdStyleSheet.Cells(index, boardNoIndex).value
                    Else
                        brdStr = brdStr + "," + brdStyleSheet.Cells(index, boardNoIndex).value
                    End If
                End If
            Next
NextLoopCustom:
        Next brdGrp
    Else
    '如果不是场景列功能，则按一般流程处理
        For Each brdGrp In grpCollection
            Call getGroupStartAndEndRowByGroupName(brdStyleSheet, CStr(brdGrp), startRow, endRow)
            If startRow = -1 Then GoTo NextLoop
            
            boardNoIndex = getboradNoColumNumber(brdStyleSheet, startRow + 1, CStr(brdGrp))
            
            For index = startRow + 2 To endRow
                If brdStr = "" Then
                    brdStr = brdStyleSheet.Cells(index, boardNoIndex).value
                Else
                    brdStr = brdStr + "," + brdStyleSheet.Cells(index, boardNoIndex).value
                End If
            Next
NextLoop:
        Next brdGrp
    End If
    
    findBoardNosStrByBtsName_i = brdStr
    
End Function

Public Function getboradNoColumNumber(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef groupName As String) As Long
    Dim colIndex As Long
    Dim colnumName As String
    For colIndex = 1 To ws.range("XFD" & CStr(rowNumber)).End(xlToLeft).column
        colnumName = ws.Cells(rowNumber, colIndex).value
        If isBoardNoColum(groupName, colnumName) Then
            getboradNoColumNumber = colIndex
            Exit Function
        End If
    Next
    getboradNoColumNumber = 1
End Function
Private Function isBoardNoColum(ByRef groupName As String, ByRef columnName As String) As Boolean
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    isBoardNoColum = False
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mappingdefgroupName = groupName And columnName = mappingdefcolumnName And attributeName = "BoardNo" Then
            isBoardNoColum = True
            Exit For
        End If
    Next
End Function

'Private Sub getStartAndEndRowNum(brdSheetName As String, ByRef groupName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
'    Dim sh As Worksheet
'    Set sh = ThisWorkbook.Worksheets(brdSheetName)
'    groupNameStartRowNumber = getGroupNameStartRowNumber(sh, groupName)
'    If groupNameStartRowNumber = -1 Then
'        groupNameEndRowNumber = -1
'    Else
'        groupNameEndRowNumber = groupNameStartRowNumber + getCurrentRegionRowsCount(sh, groupNameStartRowNumber) - 1
'    End If
'End Sub

Private Function getCurrentRegionRowsCount(ByRef ws As Worksheet, ByRef startRowNumber As Long) As Long
    Dim rowNumber As Long
    Dim rowscount As Long
    rowscount = 1
    For rowNumber = startRowNumber + 1 To 2000
        If rowIsBlank(ws, rowNumber) = True Then
            Exit For
        Else
            rowscount = rowscount + 1
        End If
    Next rowNumber
    getCurrentRegionRowsCount = rowscount
End Function

Private Function findBrdGroups() As Collection
    Dim relationSheet As Worksheet
    Dim rowCount As Long
    Dim index As Long
    Dim grpCollection As New Collection
    Set relationSheet = ThisWorkbook.Worksheets("RELATION DEF")
    rowCount = relationSheet.range("a1048576").End(xlUp).row
    For index = 2 To rowCount
        If UCase(relationSheet.Cells(index, 6).value) = "RRU" Or UCase(relationSheet.Cells(index, 6).value) = "RFU" Or _
            UCase(relationSheet.Cells(index, 6).value) = "AARU" Then
                If Not Contains(grpCollection, UCase(relationSheet.Cells(index, 6).value)) Then
                    grpCollection.Add Item:=relationSheet.Cells(index, 2).value, key:=UCase(relationSheet.Cells(index, 6).value)
                End If
        End If
    Next
    Set findBrdGroups = grpCollection
End Function

Private Function findBoardStyleSheet(btsIndex As Long) As String
    Dim groupName As String
    Dim columnName As String
    Dim mainSheet As Worksheet
    Dim mainSheetName As String
    mainSheetName = GetMainSheetName
    Set mainSheet = ThisWorkbook.Worksheets(mainSheetName)
    
    Dim siteIndex As Long
    Dim brdStyleIndex As Long
    'site index
    groupName = ThisWorkbook.Worksheets(CELL_SHEET_NAME).Cells(1, 1).value
    columnName = ThisWorkbook.Worksheets(CELL_SHEET_NAME).Cells(2, 1).value
    siteIndex = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    'brd index
    Call findBrdStyleGrpNameAndColName(mainSheetName, groupName, columnName)
    brdStyleIndex = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    
    Dim row As Long
    For row = 2 To mainSheet.range("a1048576").End(xlUp).row
         If mainSheet.Cells(row, siteIndex).value = SITE_NAME Then
                findBoardStyleSheet = mainSheet.Cells(row, brdStyleIndex).value
                btsIndex = row
                Exit Function
         End If
    Next
    findBoardStyleSheet = ""
End Function


Private Function findBoardStyleSheetByBtsName(ByRef btsName As String) As String
    Dim groupName As String
    Dim columnName As String
    Dim mainSheet As Worksheet
    Dim mainSheetName As String
    mainSheetName = GetMainSheetName
    Set mainSheet = ThisWorkbook.Worksheets(mainSheetName)
    
    Dim siteIndex As Long
    Dim brdStyleIndex As Long
    'site index
    groupName = ThisWorkbook.Worksheets(CELL_SHEET_NAME).Cells(1, 1).value
    columnName = ThisWorkbook.Worksheets(CELL_SHEET_NAME).Cells(2, 1).value
    siteIndex = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    'brd index
    Call findBrdStyleGrpNameAndColName(mainSheetName, groupName, columnName)
    brdStyleIndex = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    
    Dim row As Long
    For row = 2 To mainSheet.range("a1048576").End(xlUp).row
         If mainSheet.Cells(row, siteIndex).value = btsName Then
                findBoardStyleSheetByBtsName = mainSheet.Cells(row, brdStyleIndex).value
                Exit Function
         End If
    Next
    findBoardStyleSheetByBtsName = ""
End Function

Private Function findColNumByGrpNameAndColName(sh As Worksheet, groupName As String, columnName As String)
    Dim m_colNum As Long
    For m_colNum = 1 To sh.range("XFD2").End(xlToLeft).column
        If get_GroupName(sh.name, m_colNum) = groupName Then
            If GetDesStr(columnName) = GetDesStr(sh.Cells(2, m_colNum).value) Then
                findColNumByGrpNameAndColName = m_colNum
                Exit For
            End If
        End If
    Next
End Function

Private Sub findBrdStyleGrpNameAndColName(sheetName As String, groupName As String, columanName As String)
    Dim relationSheet As Worksheet
    Dim rowCount As Long
    Dim index As Long
    
    Set relationSheet = ThisWorkbook.Worksheets("RELATION DEF")
    rowCount = relationSheet.range("a1048576").End(xlUp).row
    For index = 2 To rowCount
        If relationSheet.Cells(index, 1).value = sheetName And LCase(relationSheet.Cells(index, 4).value) = "true" And _
            LCase(relationSheet.Cells(index, 5).value) = "false" Then
                groupName = relationSheet.Cells(index, 2).value
                columanName = relationSheet.Cells(index, 3).value
                Exit Sub
            Exit Sub
        End If
    Next
End Sub

Private Function calculateRow(ByRef chkPassed As Boolean) As Long
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim rowCount As Long
    Dim columnIndex As Long
    Dim sectorIndex As Long
    Dim baseEqmIndex As Long
    Dim sectorEqmIndex As Long
    Dim sectorEqmGrpIdIndex As Long
    Dim cellBeamModeIndex As Long
    Dim multiRruCellModeIndex As Long
    Dim cellsheet As Worksheet
    Dim index As Long
    Dim antenneIndex As Long
    Dim sectorArray As Validation
    
    Set Cell_TrxListInfo = New CMap
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Set btsNameRowCountMap = New CMap
    
    '基站名称默认在第一列
    Dim btsNameColIndex As Long
    Call getBaseStationMocNameAndAttrName(mocName, attrName)
    btsNameColIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    '找小区ID所在的列
    Dim constCellTempCol As Long
    Call getCellMocNameAndAttrName(mocName, attrName)
    '判断是否为EuCellSectorEqm或EuPrbSectorEqm页签，如果是则mocName重新赋值
    If CELL_SHEET_NAME = "EUCELLSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUCELLSECTOREQM") Then
        mocName = "EuCellSectorEqm"
        attrName = "VLOCALCELLID"
    ElseIf CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
        mocName = "EuPrbSectorEqm"
        attrName = "VLOCALCELLID"
    End If
    constCellTempCol = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    '获取天线端口所在列
    Call getAntenneMocNameAndAttr(mocName, attrName)
    If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
        mocName = "EuPrbSectorEqm"
    End If
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    columnIndex = findColumnByName(cellsheet, columnName, 2)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
            mocName = "EuPrbSectorEqm"
        End If
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    sectorIndex = findColumnByName(cellsheet, columnName, 2)
    
    If CELL_TYPE = 2 Then
        Call getBaseEqmMocNameAndAttrName(mocName, attrName)
        If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
            mocName = "EuPrbSectorEqm"
                        attrName = "BasebandEqmId"
        End If
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        baseEqmIndex = findColumnByName(cellsheet, columnName, 2)
        
        Call getSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        sectorEqmGrpIdIndex = findColumnByName(cellsheet, columnName, 2)
        
        Call getCellbeamModeMocNameAndAttrName(mocName, attrName)
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        cellBeamModeIndex = findColumnByName(cellsheet, columnName, 2)
        
        '获取LTE小区页签多RRU共小区模式列
        Call getCellMultiRruCellModeMocAndAttrName(mocName, attrName)
        multiRruCellModeIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    End If
    
    If CELL_TYPE = 1 Then
        Call getSectoreqmMocNameAndAttrName(mocName, attrName)
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        sectorEqmIndex = findColumnByName(cellsheet, columnName, 2)
        
        '获取U小区页签多RRU共小区模式列
        Call getCellMultiRruCellModeMocAndAttrName(mocName, attrName)
        multiRruCellModeIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    End If
    
    rowCount = 0
    If columnIndex <= 0 Then
        calculateRow = rowCount
        Exit Function
    End If
    Dim antennes As CAntennes
    Dim temBtsName As String
    
    For index = 3 To cellsheet.range("a1048576").End(xlUp).row
        temBtsName = cellsheet.Cells(index, btsNameColIndex).value
        If existInCollection(temBtsName, selectBtsNameCol) And cellsheet.Cells(index, constCellTempCol).value <> "" And cellsheet.Cells(index, 2).value <> "RMV" Then
            '此处加判断，如果是LTE或UMTS小区，且多RRU共小区模式为MPRU_AGGREGATION，则不转入临时页签
            If (CELL_TYPE = 1 Or CELL_TYPE = 2) And multiRruCellModeIndex > 0 Then
                If cellsheet.Cells(index, multiRruCellModeIndex).value = "MPRU_AGGREGATION" Or cellsheet.Cells(index, multiRruCellModeIndex).value = "MIXED_MULTIRRU_CELL" Then
                    GoTo NextLoop
                End If
            End If
            Set antennes = New CAntennes
            Dim tem_key As String
            
            antennes.btsName = cellsheet.Cells(index, btsNameColIndex).value
            antennes.cellId = cellsheet.Cells(index, constCellTempCol).value
            
            'key采用BTSNAME#CELLNAME的形式
            Dim tempBtsName As String
            Dim tempCellId As String
            tempBtsName = cellsheet.Cells(index, btsNameColIndex).value
            tempCellId = cellsheet.Cells(index, constCellTempCol).value
            tem_key = tempBtsName + "*" + tempCellId
            
            If (CELL_TYPE = 0) Then
                antennes.trxId = TrxInfoMgr.getFreqLstStrByGCell(CELL_SHEET_NAME, temBtsName, antennes.cellId)
                If antennes.trxId = "" Then
                    chkPassed = False
                    Exit Function
                End If
                If Cell_TrxListInfo.hasKey(tem_key) = False Then
                    Call Cell_TrxListInfo.SetAt(tem_key, antennes.trxId)
                End If
            Else
                antennes.trxId = ""
            End If
            antennes.antennes = cellsheet.Cells(index, columnIndex).value
            antennes.sectorIds = cellsheet.Cells(index, sectorIndex).value
            If CELL_TYPE = 2 Then
                antennes.baseEqmIds = cellsheet.Cells(index, baseEqmIndex).value
                If sectorEqmGrpIdIndex > 0 Then antennes.sectorEqmGrpIds = cellsheet.Cells(index, sectorEqmGrpIdIndex).value
                If cellBeamModeIndex > 0 Then antennes.cellbeamModes = cellsheet.Cells(index, cellBeamModeIndex).value
            End If
            If CELL_TYPE = 1 And sectorEqmIndex > 0 Then
                antennes.sectorEqmIds = cellsheet.Cells(index, sectorEqmIndex).value
            End If
            antennes.ranges = cellsheet.Cells(index, columnIndex).address(False, False)
            antennes.row = index
            
            
            '获取每个“基站#小区”对应的行数，用于确定单板编号下拉列表范围
            Dim antennesCol As Collection
            Dim sflag As Boolean
            Dim rowVal As Long
            
           Set antennesCol = antennes.getAntenneCollection(sflag, rowVal)
            
            '当sflag为false证明获取antennesCol失败，直接删除临时页签，退出
            If sflag = False Then
                Call deleteTempSheet
                ThisWorkbook.Worksheets(CELL_SHEET_NAME).rows(rowVal).Select
                Exit Function
            End If
            
            If Not btsNameRowCountMap.hasKey(CStr(tem_key)) Then
                Call btsNameRowCountMap.SetAt(CStr(tem_key), antennesCol.count)
                rowCount = rowCount + antennesCol.count
            End If

            If Not Contains(valueMap, tem_key) Then
                valueMap.Add Item:=antennes, key:=tem_key
            End If
        End If
NextLoop:
    Next
    
    calculateRow = rowCount
End Function

Public Function findColumnByName(sh As Worksheet, columnName As String, row As Long) As String
    Dim columnCount As Long
    Dim index As Long
    columnCount = sh.range("XFD" + CStr(row)).End(xlToLeft).column
    For index = 1 To columnCount
           If sh.Cells(row, index).value = columnName Then
                 findColumnByName = index
                 Exit Function
           End If
    Next
    findColumnByName = -1
End Function

Public Function findColumnFromRelationDef(sheetName As String, mocName As String, attrName As String) As String
    Dim relationSheet As Worksheet
    Dim rowCount As Long
    Dim index As Long
    
    Set relationSheet = ThisWorkbook.Worksheets("RELATION DEF")
    rowCount = relationSheet.range("a1048576").End(xlUp).row
    For index = 2 To rowCount
        If relationSheet.Cells(index, 1).value = sheetName And relationSheet.Cells(index, 6).value = mocName And _
            relationSheet.Cells(index, 7).value = attrName Then
            findColumnFromRelationDef = relationSheet.Cells(index, 3).value
            Exit Function
        End If
    Next
    findColumnFromRelationDef = ""
End Function

Public Sub getCellMultiRruCellModeMocAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 1 Then
        mocName = "ULOCELL"
        attrName = "LOCELLTYPE"
    ElseIf CELL_TYPE = 2 Then
        mocName = "Cell"
        attrName = "MultiRruCellMode"
        
        If IsCloudRAN_DU Then
            mocName = "EuDuLocalCellRes"
        End If
    End If
End Sub

Private Sub getSectorMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        mocName = "GTRXGROUPSECTOREQM"
        attrName = "SECTORID"
    ElseIf CELL_TYPE = 1 Then
        mocName = "ULOCELLSECTOREQM"
        attrName = "SECTORID"
    ElseIf CELL_TYPE = 2 Then
        mocName = "eUCellSectorEqm"
        attrName = "SECTORID"
        
        If IsCloudRAN_DU Then
            mocName = "EuDuCellSectorEqm"
        End If
    End If
End Sub

Private Sub getAntenneMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        mocName = "GTRXGROUPSECTOREQM"
        attrName = "SECTORANTENNA"
    ElseIf CELL_TYPE = 1 Then
        mocName = "ULOCELLSECTOREQM"
        attrName = "SECTORANTENNA"
    ElseIf CELL_TYPE = 2 Then
        mocName = "eUCellSectorEqm"
        attrName = "SECTORANTENNA"
        
        If IsCloudRAN_DU Then
            mocName = "EuDuCellSectorEqm"
        End If
    End If
End Sub


Public Function getBaseStationMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If getNeType() = "MRAT" Then
        If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
            attrName = "GBTSFUNCTIONNAME"
            mocName = "GBTSFUNCTION"
        ElseIf CELL_TYPE = 1 Then
            attrName = "NODEBFUNCTIONNAME"
            mocName = "NODEBFUNCTION"
        ElseIf CELL_TYPE = 2 Then
            attrName = "eNodeBFunctionName"
            mocName = "eNodeBFunction"
            
            If IsCloudRAN_DU Then
                attrName = "eNodeBEqmFunctionName"
                mocName = "eNodeBEqmFunction"
            End If
        End If
    Else
        attrName = "NENAME"
        mocName = "NE"
    End If
End Function

Private Function getCellMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        attrName = "GLoCellID"
        mocName = "GLoCell"
    ElseIf CELL_TYPE = 1 Then
        attrName = "ULOCELLID"
        mocName = "ULOCELL"
    ElseIf CELL_TYPE = 2 Then
        attrName = "LocalCellId"
        mocName = "Cell"
        
        If IsCloudRAN_DU Then
            attrName = "EuDuLocalCellId"
            mocName = "EuDuLocalCellRes"
        End If
    End If
End Function
Public Function getCellbeamModeMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 2 Then
        attrName = "CellBeamMode"
        mocName = "eUCellSectorEqm"
        
        If IsCloudRAN_DU Then
            mocName = "EuDuCellSectorEqm"
        End If
    End If
End Function
Public Function getSectoreqmGrpIdMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 2 Then
        attrName = "SectorEqmCombineGrpId"
        mocName = "eUCellSectorEqm"
        
        If IsCloudRAN_DU Then
            mocName = "EuDuCellSectorEqm"
        End If
    End If
End Function
Private Function getBaseEqmMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 2 Then
        attrName = "BaseBandEqmId"
        mocName = "eUCellSectorEqm"
        
        If IsCloudRAN_DU Then
            mocName = "EuDuCellSectorEqm"
        End If
    End If
End Function
Public Function getSectoreqmMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 1 Then
        attrName = "SECTOREQMPROPERTY"
        mocName = "ULOCELLSECTOREQM"
    End If
End Function

Public Function cellSheetType(sheetName As String) As Long
    If sheetName = "GSM Cell" Or sheetName = getResByKey("GSMCellSheetName") Then
        cellSheetType = 0
    ElseIf sheetName = "UMTS Cell" Or sheetName = getResByKey("UMTSCellSheetName") Or sheetName = getResByKey("ULOCELLSECEQMGRP") Or sheetName = "ULOCELLSECEQMGRP" Then
        cellSheetType = 1
    ElseIf sheetName = "LTE Cell" Or sheetName = getResByKey("LTECellSheetName") Or sheetName = getResByKey("EUSECTOREQMGROUP") Or sheetName = "EUSECTOREQMGROUP" _
    Or sheetName = getResByKey("EUCELLSECTOREQM") Or sheetName = "EUCELLSECTOREQM" Or sheetName = getResByKey("EUPRBSECTOREQM") Or sheetName = "EUPRBSECTOREQM" _
    Or sheetName = getResByKey("EUPRBSECTOREQMGROUP") Or sheetName = "EUPRBSECTOREQMGROUP" Then
        cellSheetType = 2
    Else
        cellSheetType = -1
    End If
End Function


Private Sub insertAntenneColumn()
    Dim referencedString As String
    Dim tmpsheet As Worksheet
    Dim rowNum As Long
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim antnIndex As Long
    Dim cellsheet As Worksheet
    Dim charStr As String
    
    referencedString = "R0A,R0B,R0C,R0D,R0E,R0F,R0G,R0H,R1A,R1B,R1C,R1D,R2A,R2B,R2C,R2D,R3A,R3B,R3C,R3D,R4A,R4B,R4C,R4D,R5A,R5B,R5C,R5D,R6A,R6B,R6C,R6D,R7A,R7B,R7C,R7D,NULL"
    
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Call getAntenneMocNameAndAttr(mocName, attrName)
    If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
        mocName = "EuPrbSectorEqm"
    End If
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antnIndex = findColumnByName(cellsheet, columnName, 2)
    
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        Ante_Index = 6
        charStr = "F"
    Else
        Ante_Index = 5
        charStr = "E"
    End If
    
    cellsheet.Cells(2, antnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Ante_Index)
    Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, Ante_Index), getResByKey("Antenne"))
    
    Dim antnRang As range
    Set antnRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
    If referencedString <> "" Then
        antnValueStr = referencedString
        With antnRang.Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
        End With
        ThisWorkbook.ActiveSheet.Columns(Ante_Index).EntireColumn.AutoFit
    End If
End Sub

Public Sub insertAnteModelColumn()
    Dim charStr As String
    Dim referencedString As String
     
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        Model_Index = 7
        charStr = "G"
    Else
        Model_Index = 6
        charStr = "F"
    End If
    referencedString = "RXTX_MODE,RX_MODE,TX_MODE,NULL"
    
    ThisWorkbook.ActiveSheet.Cells(1, Board_Index).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Model_Index)
    ThisWorkbook.ActiveSheet.Cells(1, Model_Index).value = getResByKey("anteModel")
    Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, Model_Index), getResByKey("AnteModelInfo"))
    
    Dim cellRang As range
    modelValueStr = referencedString
    Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
        With cellRang.Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
        End With
    ThisWorkbook.ActiveSheet.Columns(Model_Index).EntireColumn.AutoFit
End Sub


Private Function insertBaseEqmColumn() As Boolean
        Dim mocName As String
        Dim attrName As String
        Dim sectorColumnName As String
        Dim columnIndex As Long
        Dim cellsheet As Worksheet
        
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Call getBaseEqmMocNameAndAttrName(mocName, attrName)
        
        If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
            mocName = "EuPrbSectorEqm"
                        attrName = "BasebandEqmId"
        End If
        
        sectorColumnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex < 0 Then
            insertBaseEqmColumn = False
            Exit Function
        End If
        BaseEqm_Index = 7
        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, BaseEqm_Index)
        Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, BaseEqm_Index), getResByKey("LTEBASEBANDEQMID"))
        ThisWorkbook.ActiveSheet.Columns(BaseEqm_Index).EntireColumn.AutoFit
        insertBaseEqmColumn = True
End Function
Private Function insertSectorEqmGrpIdColumn() As Boolean
        Dim mocName As String
        Dim attrName As String
        Dim sectorColumnName As String
        Dim columnIndex As Long
        Dim cellsheet As Worksheet
        Dim referencedString As String
        Dim charStr As String
        referencedString = "255,0"
        
        charStr = "H"
        
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Call getSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
        sectorColumnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex < 0 Then
            insertSectorEqmGrpIdColumn = False
            SectorEqmGrpId_Index = 0
            Exit Function
        End If
        SectorEqmGrpId_Index = 8
        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, SectorEqmGrpId_Index)
        
        Dim resName As String
        resName = "SECTOREQMCOMBINEGRPID"
        If IsCloudRAN_DU Then resName = "SECTOREQMCOMBINEGRPID_DU"
        Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, SectorEqmGrpId_Index), getResByKey(resName))
        
        Dim cellRang As range
        sectoreqmValueStr = referencedString
        Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
        With cellRang.Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
        End With
        
        ThisWorkbook.ActiveSheet.Columns(SectorEqmGrpId_Index).EntireColumn.AutoFit
        
        insertSectorEqmGrpIdColumn = True
End Function

Private Function insertCellBeamModeColumn() As Boolean
        Dim mocName As String
        Dim attrName As String
        Dim sectorColumnName As String
        Dim columnIndex As Long
        Dim cellsheet As Worksheet
        Dim referencedString As String
        Dim charStr As String
        referencedString = "NORMAL,ADVANCED_BEAMFORMING,MASSIVE_MIMO_Ph1,MASSIVE_MIMO_Ph2"
        
        charStr = "I"
        
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Call getCellbeamModeMocNameAndAttrName(mocName, attrName)
        sectorColumnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex < 0 Then
            insertCellBeamModeColumn = False
            CellBeamMode_Index = 0
            Exit Function
        End If
        CellBeamMode_Index = 9
        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, CellBeamMode_Index)
        
        Dim resName As String
        resName = "CELLBEAMMODE"
        If IsCloudRAN_DU Then resName = "CELLBEAMMODE_DU"
        Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, CellBeamMode_Index), getResByKey(resName))
        
        Dim cellRang As range
        Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
        With cellRang.Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
        End With
        
        ThisWorkbook.ActiveSheet.Columns(CellBeamMode_Index).EntireColumn.AutoFit
        
        insertCellBeamModeColumn = True
End Function



Private Function insertSectoreqmColumn() As Boolean
        Dim mocName As String
        Dim attrName As String
        Dim sectorColumnName As String
        Dim columnIndex As Long
        Dim cellsheet As Worksheet
        Dim referencedString As String
        Dim charStr As String
        referencedString = "NORMAL,ASSISTANCE"
        
        charStr = "G"
        
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Call getSectoreqmMocNameAndAttrName(mocName, attrName)
        sectorColumnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex < 0 Then
            insertSectoreqmColumn = False
            SectorEqm_Index = 0
            Exit Function
        End If
        SectorEqm_Index = 7
        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, SectorEqm_Index)
        Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, SectorEqm_Index), getResByKey("UMTSSECTOREQMPROPERTY"))
        
        Dim cellRang As range
        sectoreqmValueStr = referencedString
        Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
        With cellRang.Validation
           .Delete
           .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
        End With
        
        ThisWorkbook.ActiveSheet.Columns(SectorEqm_Index).EntireColumn.AutoFit
        
        insertSectoreqmColumn = True
End Function

Private Function getAnteRSModelValue(rsModel As String, changeT As Long) As String
    If changeT = 0 Then
        If rsModel = "RX And TX" Then
            getAnteRSModelValue = "RXTX"
        Else
            getAnteRSModelValue = rsModel
        End If
    Else
         If rsModel = "RXTX" Then
            getAnteRSModelValue = "RX And TX"
        Else
            getAnteRSModelValue = rsModel
        End If
    End If
End Function

Private Sub writeBackData()
On Error GoTo ErrorHandler
    'DTS2019050703315，回写数据前清空粘贴板
    Application.CutCopyMode = False
    
    Dim cellInfoMap As CMapValueObject
    Dim error As Boolean
    Set errCollect = New Collection
    If WRITESUCCESS = True Then
        error = checkUserData()
        If error = False Then
            Exit Sub
        End If
        If CELL_TYPE = 0 Then
            Call writeGSMCellData
        ElseIf CELL_TYPE = 2 Then
            Call writeLTECellData
        ElseIf CELL_TYPE = 1 And SectorEqm_Index <> 0 Then
            Call writeUMTSCellData
        Else
            Set cellInfoMap = genCellInfoMap()
            Call sortMapByKey(cellInfoMap, error)
            Call writeCellData(cellInfoMap)
        End If
    End If
    WRITESUCCESS = False
    Call deleteTempSheet
    Exit Sub
ErrorHandler:
    Call deleteTempSheet
End Sub

Private Sub writeGSMCellData()
    Dim tmpsheet As Worksheet
    Dim cellsheet As Worksheet
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim antnIndex As Long
    Dim secIndex As Long
    
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Call getAntenneMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antnIndex = findColumnByName(cellsheet, columnName, 2)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    secIndex = findColumnByName(cellsheet, columnName, 2)
    
    Dim sVal As Variant
    Dim trxlistStr As String
    For Each sVal In Cell_TrxListInfo.KeyCollection
        trxlistStr = Cell_TrxListInfo.GetAt(sVal)
        Call writeGSMOneCell(tmpsheet, cellsheet, CStr(sVal), trxlistStr, antnIndex, secIndex)
    Next
End Sub


Private Function writeGSMOneCell(tmpsheet As Worksheet, cellsheet As Worksheet, temKey As String, trxlistStr As String, antnIndex As Long, secIndex As Long) As Boolean
    Dim trxArray() As String
    Dim index As Long
    Dim rowNum As Long
    Dim maxRow As Long
    Dim trxId As String
    Dim sector As String
    Dim antenna As String
    Dim secStr As String
    Dim antaStr As String
    Dim boardAnte As String
    Dim sectorId As String
    Dim temKeyArr() As String
    
    Dim cellInRow As String
    
    writeGSMOneCell = True
    If Cell_Row_Map.hasKey(temKey) = False Then
        writeGSMOneCell = False
        Exit Function
    End If
    cellInRow = Cell_Row_Map.GetAt(temKey)
     
    secStr = "-1"
    antaStr = ""
    trxArray = Split(trxlistStr, ",")
    temKeyArr = Split(temKey, "*")
    If UBound(temKeyArr) < 0 Then Exit Function
    
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For index = LBound(trxArray) To UBound(trxArray)
        sector = "-1"
        For rowNum = 2 To maxRow
            trxId = tmpsheet.Cells(rowNum, Trx_Index).value
            sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
            '当基站名称和小区名称匹配到，则写入一条数据
            If tmpsheet.Cells(rowNum, BTSName_Index).value = temKeyArr(0) And tmpsheet.Cells(rowNum, Cell_Index).value = temKeyArr(1) Then
                trxId = tmpsheet.Cells(rowNum, Trx_Index).value
                sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
                If trxId = trxArray(index) Then
                    boardAnte = tmpsheet.Cells(rowNum, Board_Index).value + "_" + tmpsheet.Cells(rowNum, Ante_Index).value + ":" + tmpsheet.Cells(rowNum, Model_Index).value
                    If sector = "-1" Then
                        sector = sectorId
                        antenna = boardAnte
                    Else
                        sector = sector + "," + sectorId
                        antenna = antenna + "," + boardAnte
                    End If
                End If
            End If
        Next
        
        If sector = "-1" Then
            If secStr = "-1" Then
                secStr = ""
                antaStr = ""
            ElseIf index < UBound(trxArray) Then
                secStr = secStr + ";"
                antaStr = antaStr + ";"
            End If
        Else
            If secStr = "-1" Then
                secStr = sector
                antaStr = antenna
            Else
                secStr = secStr + ";" + sector
                antaStr = antaStr + ";" + antenna
            End If
        End If
    Next
    
    cellsheet.Cells(cellInRow, secIndex).value = secStr
    cellsheet.Cells(cellInRow, antnIndex).value = antaStr
    
End Function

Private Sub writeLTECellData()
    Dim tmpsheet As Worksheet
    Dim cellsheet As Worksheet
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim antnIndex As Long
    Dim secIndex As Long
    Dim baseEqmIndex As Long
    Dim sectorEqmGrpIdIndex As Long
    Dim cellBeamModeIndex As Long
    
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    
    Call getAntenneMocNameAndAttr(mocName, attrName)
    If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
        mocName = "EuPrbSectorEqm"
    End If
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antnIndex = findColumnByName(cellsheet, columnName, 2)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
        mocName = "EuPrbSectorEqm"
    End If
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    secIndex = findColumnByName(cellsheet, columnName, 2)
    
    If CELL_TYPE = 2 Then
        Call getBaseEqmMocNameAndAttrName(mocName, attrName)
        If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
            mocName = "EuPrbSectorEqm"
                        attrName = "BasebandEqmId"
        End If
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        baseEqmIndex = findColumnByName(cellsheet, columnName, 2)
        
        Call getSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        sectorEqmGrpIdIndex = findColumnByName(cellsheet, columnName, 2)
        
        Call getCellbeamModeMocNameAndAttrName(mocName, attrName)
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        cellBeamModeIndex = findColumnByName(cellsheet, columnName, 2)
        
    End If
    
    Dim sVal As Variant
    For Each sVal In Cell_Row_Map.KeyCollection
        Call writeLTEOneCell(tmpsheet, cellsheet, CStr(sVal), secIndex, antnIndex, baseEqmIndex, sectorEqmGrpIdIndex, cellBeamModeIndex)
    Next
End Sub

Private Sub writeUMTSCellData()
    Dim tmpsheet As Worksheet
    Dim cellsheet As Worksheet
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim antnIndex As Long
    Dim secIndex As Long
    Dim sectorEqmIndex As Long
    
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    
    Call getAntenneMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antnIndex = findColumnByName(cellsheet, columnName, 2)
    
    Call getSectorMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    secIndex = findColumnByName(cellsheet, columnName, 2)
    
    If CELL_TYPE = 1 Then
        Call getSectoreqmMocNameAndAttrName(mocName, attrName)
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        sectorEqmIndex = findColumnByName(cellsheet, columnName, 2)
    End If
    
    Dim sVal As Variant
    For Each sVal In Cell_Row_Map.KeyCollection
        Call writeUMTSOneCell(tmpsheet, cellsheet, CStr(sVal), secIndex, antnIndex, sectorEqmIndex)
    Next
End Sub

Private Sub writeUMTSOneCell(tmpsheet As Worksheet, cellsheet As Worksheet, temKey As String, _
                                                    secIndex As Long, antnIndex As Long, sectorEqmIndex As Long)
    Dim maxRow As Long
    Dim rowNum As Long
    Dim sectorId As String
    Dim baseEqmId As String
    Dim boardAnte As String
    Dim sectorIds As String
    Dim sectoreqmPro As String
    Dim temKeyArr() As String
    
    temKeyArr = Split(temKey, "*")
    If UBound(temKeyArr) < 0 Then Exit Sub
    
    sectorIds = ""
    
    Dim cellInRow As String
    cellInRow = Cell_Row_Map.GetAt(temKey)
    
    Dim sectoreqmProSectorMap As CMapValueObject
    Set sectoreqmProSectorMap = New CMapValueObject
    
    Dim sectoreqmProColl As Collection
    
    Dim sectorMap As CMap
    Set sectorMap = New CMap
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
         If tmpsheet.Cells(rowNum, BTSName_Index).value = temKeyArr(0) And tmpsheet.Cells(rowNum, Cell_Index).value = temKeyArr(1) Then
            sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
            sectoreqmPro = tmpsheet.Cells(rowNum, SectorEqm_Index).value
            If sectorMap.hasKey(sectorId) = False Then
                Call sectorMap.SetAt(sectorId, sectorId)
            End If
            If sectoreqmProSectorMap.hasKey(sectorId) = False Then
                Set sectoreqmProColl = New Collection
                sectoreqmProColl.Add Item:=sectoreqmPro, key:=sectoreqmPro
                Call sectoreqmProSectorMap.SetAt(sectorId, sectoreqmProColl)
            End If
            Set sectoreqmProColl = sectoreqmProSectorMap.GetAt(sectorId)
            If sectoreqmProSectorMap.hasKey(sectorId) And Contains(sectoreqmProColl, sectoreqmPro) = False Then
                sectoreqmProSectorMap.RemoveKey (sectorId)
                sectoreqmProColl.Add Item:=sectoreqmPro, key:=sectoreqmPro
                Call sectoreqmProSectorMap.SetAt(sectorId, sectoreqmProColl)
            End If
        End If
    Next
    
    Dim sVal As Variant
    For Each sVal In sectorMap.KeyCollection
        If sectorIds = "" Then
            sectorIds = CStr(sVal)
        Else
            sectorIds = sectorIds + "," + CStr(sVal)
        End If
    Next
    
    getSortedStr (sectorIds)
    
    Dim sectorArray() As String
    Dim baseEqmIdStr As String
    baseEqmIdStr = ""
    
    Dim antenna As String
    Dim antaStr As String
    antaStr = ""
    
    sectorArray = Split(sectorIds, ",")
    Dim sectorIdStr As String
    Dim index As Long
    sectorIdStr = ""
    For index = LBound(sectorArray) To UBound(sectorArray)
        baseEqmId = ""
        Set sectoreqmProColl = sectoreqmProSectorMap.GetAt(sectorArray(index))
        Dim aVal As Variant
        For Each aVal In sectoreqmProColl
            antenna = ""
            If baseEqmId = "" Then
                baseEqmId = CStr(aVal)
            Else
                baseEqmId = baseEqmId + "," + CStr(aVal)
            End If
            If sectorIdStr = "" Then
                sectorIdStr = sectorArray(index)
            Else
                sectorIdStr = sectorIdStr + "," + sectorArray(index)
            End If
            
            For rowNum = 2 To maxRow
                If tmpsheet.Cells(rowNum, BTSName_Index).value = temKeyArr(0) And tmpsheet.Cells(rowNum, Cell_Index).value = temKeyArr(1) Then
                    sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
                    sectoreqmPro = tmpsheet.Cells(rowNum, SectorEqm_Index).value
                    If sectorId = sectorArray(index) And sectoreqmPro = CStr(aVal) Then
                        boardAnte = tmpsheet.Cells(rowNum, Board_Index).value + "_" + tmpsheet.Cells(rowNum, Ante_Index).value + ":" + tmpsheet.Cells(rowNum, Model_Index).value
                        If antenna = "" Then
                            antenna = boardAnte
                        Else
                            antenna = antenna + "," + boardAnte
                        End If
                    End If
                 End If
            
            Next
        
            If antaStr = "" Then
                antaStr = antenna
            Else
                antaStr = antaStr + ";" + antenna
            End If
        Next
        

        If baseEqmIdStr = "" Then
            baseEqmIdStr = baseEqmId
        Else
            baseEqmIdStr = baseEqmIdStr + "," + baseEqmId
        End If
    Next
    
    cellsheet.Cells(cellInRow, secIndex).value = sectorIdStr
    cellsheet.Cells(cellInRow, antnIndex).value = antaStr
    If CELL_TYPE = 1 Then
        cellsheet.Cells(cellInRow, sectorEqmIndex).value = baseEqmIdStr
    End If
    
End Sub

Private Sub writeLTEOneCell(tmpsheet As Worksheet, cellsheet As Worksheet, temKey As String, _
                                                    secIndex As Long, antnIndex As Long, baseEqmIndex As Long, sectorEqmGrpIdIndex As Long, cellBeamModeIndex As Long)
    Dim maxRow As Long
    Dim rowNum As Long
    Dim sectorId As String
    Dim baseEqmId As String
    Dim boardAnte As String
    Dim sectorIds As String
    Dim sectorEqmGrpId As String
    Dim cellbeamMode As String
    Dim temKeyArr() As String
    
    temKeyArr = Split(temKey, "*")
    If UBound(temKeyArr) < 0 Then Exit Sub
    
    sectorIds = ""
    
    Dim cellInRow As String
    cellInRow = Cell_Row_Map.GetAt(temKey)
    
    Dim sectorMap As CMap
    Set sectorMap = New CMap
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
         If tmpsheet.Cells(rowNum, BTSName_Index).value = temKeyArr(0) And tmpsheet.Cells(rowNum, Cell_Index).value = temKeyArr(1) Then
            sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
            If sectorMap.hasKey(sectorId) = False Then
                Call sectorMap.SetAt(sectorId, sectorId)
            End If
        End If
    Next
    
    Dim sVal As Variant
    For Each sVal In sectorMap.KeyCollection
        If sectorIds = "" Then
            sectorIds = CStr(sVal)
        Else
            sectorIds = sectorIds + "," + CStr(sVal)
        End If
    Next
    
    getSortedStr (sectorIds)
    
    Dim sectorArray() As String
    Dim baseEqmIdStr As String
    baseEqmIdStr = ""
    
    Dim sectorEqmGrpIdStr As String
    sectorEqmGrpIdStr = ""
    
    Dim cellbeamModeStr As String
    cellbeamModeStr = ""
    
    Dim antenna As String
    Dim antaStr As String
    antaStr = ""
    
    sectorArray = Split(sectorIds, ",")
    
    Dim index As Long
    For index = LBound(sectorArray) To UBound(sectorArray)
        antenna = ""
        For rowNum = 2 To maxRow
            If tmpsheet.Cells(rowNum, BTSName_Index).value = temKeyArr(0) And tmpsheet.Cells(rowNum, Cell_Index).value = temKeyArr(1) Then
                sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
                If sectorId = sectorArray(index) Then
                    If CELL_TYPE = 2 Then
                        baseEqmId = tmpsheet.Cells(rowNum, BaseEqm_Index).value
                        If SectorEqmGrpId_Index <> 0 Then sectorEqmGrpId = tmpsheet.Cells(rowNum, SectorEqmGrpId_Index).value
                        If CellBeamMode_Index <> 0 Then cellbeamMode = tmpsheet.Cells(rowNum, CellBeamMode_Index).value
                    End If
                
                    boardAnte = tmpsheet.Cells(rowNum, Board_Index).value + "_" + tmpsheet.Cells(rowNum, Ante_Index).value + ":" + tmpsheet.Cells(rowNum, Model_Index).value
                    If antenna = "" Then
                        antenna = boardAnte
                    Else
                        antenna = antenna + "," + boardAnte
                    End If
                End If
            End If
        Next
        
        If antaStr = "" Then
            antaStr = antenna
        Else
            antaStr = antaStr + ";" + antenna
        End If
        
        If baseEqmIdStr = "" Then
            baseEqmIdStr = baseEqmId
        Else
            baseEqmIdStr = baseEqmIdStr + "," + baseEqmId
        End If
        
        If cellbeamModeStr = "" Then
            cellbeamModeStr = cellbeamMode
        Else
            cellbeamModeStr = cellbeamModeStr + "," + cellbeamMode
        End If
        
        If sectorEqmGrpIdStr = "" Then
            sectorEqmGrpIdStr = sectorEqmGrpId
        Else
            sectorEqmGrpIdStr = sectorEqmGrpIdStr + "," + sectorEqmGrpId
        End If
    Next
    
    cellsheet.Cells(cellInRow, secIndex).value = sectorIds
    cellsheet.Cells(cellInRow, antnIndex).value = antaStr
    If CELL_TYPE = 2 Then
        cellsheet.Cells(cellInRow, baseEqmIndex).value = baseEqmIdStr
        If sectorEqmGrpIdIndex <> -1 Then cellsheet.Cells(cellInRow, sectorEqmGrpIdIndex).value = sectorEqmGrpIdStr
        If cellBeamModeIndex <> -1 Then cellsheet.Cells(cellInRow, cellBeamModeIndex).value = cellbeamModeStr
    End If
    
End Sub

Private Sub deleteTempSheet()
    Dim tmpsheet As Worksheet
    Dim cellsheet As Worksheet
    InAdjustAntnPort = False
    If CELL_SHEET_NAME <> "" Then
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        cellsheet.Activate
    End If
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    Call changeAlerts(False)
    tmpsheet.Delete
    Call changeAlerts(True)
End Sub


Private Function genCellInfoMap() As CMapValueObject
    Dim tmpsheet As Worksheet
    Dim maxRow As Long
    Dim rowNum As Long
    Dim cellInfoMap As CMapValueObject
    Dim tmpMap As CMap
    Dim cellId As String
    Dim sectorId As String
    Dim board As String
    Dim antn As String
    Dim model As String
    Dim boardInfo As String
    Dim tVal As CMap
    Dim sVal As String
    Dim isExist As Boolean
    Dim celldes As String
    Dim sectordes As String
    Dim antndes As String
    Dim btsName As String
    Dim temKey As String
    
    Set cellInfoMap = New CMapValueObject
    
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
        btsName = tmpsheet.Cells(rowNum, BTSName_Index)
        cellId = tmpsheet.Cells(rowNum, Cell_Index).value
        sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
        board = tmpsheet.Cells(rowNum, Board_Index).value
        antn = tmpsheet.Cells(rowNum, Ante_Index).value
        model = tmpsheet.Cells(rowNum, Model_Index).value
        boardInfo = board & "_" & antn & ":" & model
        
        temKey = btsName + "*" + cellId
        
        isExist = cellInfoMap.hasKey(temKey)
        If isExist Then
            Set tVal = cellInfoMap.GetAt(temKey)
            isExist = tVal.hasKey(sectorId)
            If isExist Then
                sVal = tVal.GetAt(sectorId)
                boardInfo = sVal + "," + boardInfo
                tVal.RemoveKey (sectorId)
            End If
            Call tVal.SetAt(sectorId, boardInfo)
        Else
            Set tmpMap = New CMap
            Call tmpMap.SetAt(sectorId, boardInfo)
            Call cellInfoMap.SetAt(temKey, tmpMap)
        End If
    Next
    Set genCellInfoMap = cellInfoMap
End Function

Private Function checkUserData() As Boolean
    Dim tmpsheet As Worksheet
    Dim errRangeCol As Collection
    Dim maxRow As Long
    Dim rowNum As Long
    Dim cellStr As String
    Dim sectorStr As String
    Dim boardStr As String
    Dim trxStr As String
    Dim antnStr As String
    Dim modelStr As String
    Dim baseEqmStr As String
    Dim sectoreqmproStr As String
    Dim btsNameStr As String
    Dim beamModeStr As String
    Dim sectorEqmGrpIdStr As String
    
    Dim errflag As Boolean
    Dim lineStr As String
    Dim tVal As Variant
    Dim eRange As range
    Dim dupCol As Collection
    Dim keyStr As String
    Dim umtSkeyStr As String
    Dim cellTrxMap As CMapValueObject  '校验GSM小区的频点必须都要配置
    'Dim trxSectorMap As New CMapValueObject '校验GSM共小区情况下每个频点配置的扇区都要一致
    
    Dim sectorBaseEqmMap As CMapValueObject
    Dim sectorBaseEqm As String
    Set sectorBaseEqmMap = New CMapValueObject
    
    Dim sectorEqmproMap As CMapValueObject
    Dim sectoreqmPro As String
    Set sectorEqmproMap = New CMapValueObject
    
    Dim umtSsectorEqmproMap As CMap
    Dim umtSsectorEqmpro As String
    Set umtSsectorEqmproMap = New CMap
    
    Dim btsNameSectorEqmProMap As CMap
    Set btsNameSectorEqmProMap = New CMap
    
    Dim btsNameBeamModeMap As CMap
    Set btsNameBeamModeMap = New CMap
    
    Dim btsNameCombineGrpIdMap As CMap
    Set btsNameCombineGrpIdMap = New CMap
    
    Dim lteSectorEqmGrpIdMap As CMap
    Set lteSectorEqmGrpIdMap = New CMap
    
    Dim lteCellBeamModeMap As CMap
    Set lteCellBeamModeMap = New CMap
    
    Set cellTrxMap = New CMapValueObject
    checkUserData = True
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    Set errRangeCol = New Collection
    Set dupCol = New Collection
    
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
        btsNameStr = tmpsheet.Cells(rowNum, BTSName_Index).value
        cellStr = tmpsheet.Cells(rowNum, Cell_Index).value
        sectorStr = tmpsheet.Cells(rowNum, Sector_Index).value
        boardStr = tmpsheet.Cells(rowNum, Board_Index).value
        antnStr = tmpsheet.Cells(rowNum, Ante_Index).value

        modelStr = tmpsheet.Cells(rowNum, Model_Index).value
        If CELL_TYPE = 2 Then
            baseEqmStr = tmpsheet.Cells(rowNum, BaseEqm_Index).value
            If SectorEqmGrpId_Index <> 0 Then
                sectorEqmGrpIdStr = tmpsheet.Cells(rowNum, SectorEqmGrpId_Index).value
            End If
            If CellBeamMode_Index <> 0 Then
                beamModeStr = tmpsheet.Cells(rowNum, CellBeamMode_Index).value
            End If
        End If
        If CELL_TYPE = 0 Then
            trxStr = tmpsheet.Cells(rowNum, Trx_Index).value
        End If
        If CELL_TYPE = 1 And SectorEqm_Index <> 0 Then
            sectoreqmproStr = tmpsheet.Cells(rowNum, SectorEqm_Index).value
        End If
        
        boardValueStr = ""
        If btsNameBrdNoMap.hasKey(btsNameStr) Then boardValueStr = btsNameBrdNoMap.GetAt(btsNameStr)
        
        errflag = checkInputData(tmpsheet, rowNum, btsNameStr, cellStr, sectorStr, boardStr, antnStr, modelStr, baseEqmStr, trxStr, sectoreqmproStr, errRangeCol)
        If errflag Then
            If CELL_TYPE = 0 Then
                trxStr = tmpsheet.Cells(rowNum, Trx_Index).value
                Dim aMap As CMapValueObject
                Set aMap = New CMapValueObject
                Dim col As Collection
                Set col = New Collection
                Dim temCellTrxKey As String
                temCellTrxKey = btsNameStr + "*" + cellStr
                If cellTrxMap.hasKey(temCellTrxKey) Then
                     Dim needAdd As Boolean
                     needAdd = True
                     Set aMap = cellTrxMap.GetAt(temCellTrxKey)
                     If aMap.hasKey(sectorStr) Then
                        Set col = aMap.GetAt(sectorStr)
                        If Contains(col, trxStr) Then
                            needAdd = False
                        End If
                    End If
                    If needAdd = True Then
                        col.Add Item:=trxStr, key:=trxStr
                        Call aMap.SetAt(sectorStr, col)
                        cellTrxMap.RemoveKey (temCellTrxKey)
                        Call cellTrxMap.SetAt(temCellTrxKey, aMap)
                    End If
                Else
                    col.Add Item:=trxStr, key:=trxStr
                    Call aMap.SetAt(sectorStr, col)
                    Call cellTrxMap.SetAt(temCellTrxKey, aMap)
                End If
            
                keyStr = btsNameStr + "_" + cellStr + "_" + sectorStr + "_" + trxStr + "_" + boardStr + "_" + antnStr + "_" + modelStr
            ElseIf CELL_TYPE = 2 Then
                keyStr = btsNameStr + "_" + cellStr + "_" + sectorStr + "_" + boardStr + "_" + antnStr + "_" + modelStr + "_" + baseEqmStr
                
                sectorBaseEqm = btsNameStr + "_" + cellStr + "_" + sectorStr
                Dim bMap As CMap
                Set bMap = New CMap
                If sectorBaseEqmMap.hasKey(sectorBaseEqm) Then
                    Set bMap = sectorBaseEqmMap.GetAt(sectorBaseEqm)
                    sectorBaseEqmMap.RemoveKey (sectorBaseEqm)
                End If
                Call bMap.SetAt(rowNum, baseEqmStr)
                Call sectorBaseEqmMap.SetAt(sectorBaseEqm, bMap)
                   
                If btsNameBeamModeMap.hasKey(btsNameStr) Then
                    If (Len(btsNameBeamModeMap.GetAt(btsNameStr)) = 0 And Len(beamModeStr) <> 0) _
                        Or (Len(btsNameBeamModeMap.GetAt(btsNameStr)) <> 0 And Len(beamModeStr) = 0) Then
                        Call MsgBox(Replace(getResByKey("CellBeamModeNotComplete"), "0%", btsNameStr), vbInformation, getResByKey("Warning"))
                        checkUserData = False
                        Exit Function
                    End If
                Else
                    Call btsNameBeamModeMap.SetAt(btsNameStr, beamModeStr)
                End If
                
                If btsNameCombineGrpIdMap.hasKey(btsNameStr) Then
                    If (Len(btsNameCombineGrpIdMap.GetAt(btsNameStr)) = 0 And Len(sectorEqmGrpIdStr) <> 0) _
                        Or (Len(btsNameCombineGrpIdMap.GetAt(btsNameStr)) <> 0 And Len(sectorEqmGrpIdStr) = 0) Then
                        Call MsgBox(Replace(getResByKey("SectorEqmCombineGrpIdNotComplete"), "0%", btsNameStr), vbInformation, getResByKey("Warning"))
                        checkUserData = False
                        Exit Function
                    End If
                Else
                    Call btsNameCombineGrpIdMap.SetAt(btsNameStr, sectorEqmGrpIdStr)
                End If
                
                '校验 扇区设备合并组标识 和 小区波束模式
                If SectorEqmGrpId_Index <> 0 Then
                    'sectorEqmGrpIdStr = tmpsheet.Cells(rowNum, SectorEqmGrpId_Index).value
                    If lteSectorEqmGrpIdMap.hasKey(sectorBaseEqm) Then
                        Dim lteSecEqmGrpId As String
                        lteSecEqmGrpId = lteSectorEqmGrpIdMap.GetAt(sectorBaseEqm)
                        If lteSecEqmGrpId <> sectorEqmGrpIdStr Then
                            Call MsgBox(getResByKey("lteSectorEqmGrpIdNotConsis"), vbInformation, getResByKey("Warning"))
                            tmpsheet.range(Cells(rowNum - 1, SectorEqmGrpId_Index), Cells(rowNum, SectorEqmGrpId_Index)).Interior.colorIndex = 3
                            checkUserData = False
                            Exit Function
                        End If
                    Else
                        Call lteSectorEqmGrpIdMap.SetAt(sectorBaseEqm, sectorEqmGrpIdStr)
                    End If
                End If
                
                If CellBeamMode_Index <> 0 Then
                    'beamModeStr = tmpsheet.Cells(rowNum, CellBeamMode_Index).value
                    If lteCellBeamModeMap.hasKey(sectorBaseEqm) Then
                        Dim lteCellBeamModeStr As String
                        lteCellBeamModeStr = lteCellBeamModeMap.GetAt(sectorBaseEqm)
                        If lteCellBeamModeStr <> beamModeStr Then
                            Call MsgBox(getResByKey("lteCellBeamModeNotConsis"), vbInformation, getResByKey("Warning"))
                            tmpsheet.range(Cells(rowNum - 1, CellBeamMode_Index), Cells(rowNum, CellBeamMode_Index)).Interior.colorIndex = 3
                            checkUserData = False
                            Exit Function
                        End If
                    Else
                        Call lteCellBeamModeMap.SetAt(sectorBaseEqm, beamModeStr)
                    End If
                End If
                
            ElseIf CELL_TYPE = 1 And SectorEqm_Index <> 0 Then
                keyStr = btsNameStr + "_" + cellStr + "_" + sectorStr + "_" + boardStr + "_" + antnStr + "_" + modelStr + "_" + sectoreqmproStr
                
                sectoreqmPro = btsNameStr + "_" + cellStr + "_" + sectorStr
                Dim sbMap As CMap
                Set sbMap = New CMap
                If sectorEqmproMap.hasKey(sectoreqmPro) Then
                    Set sbMap = sectorEqmproMap.GetAt(sectoreqmPro)
                    sectorEqmproMap.RemoveKey (sectoreqmPro)
                End If
                Call sbMap.SetAt(rowNum, sectoreqmproStr)
                Call sectorEqmproMap.SetAt(sectoreqmPro, sbMap)
                
                umtSkeyStr = btsNameStr + "_" + cellStr + "_" + boardStr
                
                If umtSsectorEqmproMap.hasKey(umtSkeyStr) Then
                    umtSsectorEqmpro = umtSsectorEqmproMap.GetAt(umtSkeyStr)
                    If umtSsectorEqmpro <> sectoreqmproStr Then
                        Call MsgBox(getResByKey("sectorEqmNotConsis"), vbInformation, getResByKey("Warning"))
                        tmpsheet.range(Cells(rowNum - 1, SectorEqm_Index), Cells(rowNum, SectorEqm_Index)).Interior.colorIndex = 3
                        checkUserData = False
                        Exit Function
                    End If
                End If
                Call umtSsectorEqmproMap.SetAt(umtSkeyStr, sectoreqmproStr)
                
                If btsNameSectorEqmProMap.hasKey(btsNameStr) Then
                    If (Len(btsNameSectorEqmProMap.GetAt(btsNameStr)) = 0 And Len(sectoreqmproStr) <> 0) _
                        Or (Len(btsNameSectorEqmProMap.GetAt(btsNameStr)) <> 0 And Len(sectoreqmproStr) = 0) Then
                        Call MsgBox(Replace(getResByKey("SectorEqmProDataNotComplete"), "0%", btsNameStr), vbInformation, getResByKey("Warning"))
                        checkUserData = False
                        Exit Function
                    End If
                Else
                    Call btsNameSectorEqmProMap.SetAt(btsNameStr, sectoreqmproStr)
                End If
                
            Else
                keyStr = btsNameStr + "_" + cellStr + "_" + sectorStr + "_" + boardStr + "_" + antnStr + "_" + modelStr
            End If
            If Contains(dupCol, keyStr) Then
                lineStr = dupCol(keyStr) + "," + str(rowNum)
                dupCol.Remove (keyStr)
            Else
                lineStr = str(rowNum)
            End If
            dupCol.Add Item:=lineStr, key:=keyStr
        End If
    Next
    
    If errRangeCol.count() <> 0 Then
        Call MsgBox(getResByKey("recordError"), vbInformation, getResByKey("Warning"))
        For Each tVal In errRangeCol
            Set eRange = tmpsheet.range(tVal)
            eRange.Interior.colorIndex = 3
        Next
        Set errCollect = errRangeCol
        checkUserData = False
        Exit Function
    End If
    
    Set dupCollect = New Collection
    
    If CheckDataDuplicate(tmpsheet, dupCol) = False Then
        checkUserData = False
        Exit Function
    End If
    
    If CELL_TYPE = 2 Then
        If CheckLteCellValid(tmpsheet, sectorBaseEqmMap) = False Then
            checkUserData = False
            Exit Function
        End If
    End If
    
    If CELL_TYPE = 0 Then
        If CheckGsmCellValid(cellTrxMap) = False Then
            checkUserData = False
            Exit Function
        End If
    End If

End Function

Private Function CheckLteCellValid(tmpsheet As Worksheet, sectorBaseEqmMap As CMapValueObject) As Boolean
        Dim aVal As Variant
        Dim tMap As CMap
        Dim rowErrStr As String
        Dim errReportStr As String
        CheckLteCellValid = True
        
        For Each aVal In sectorBaseEqmMap.KeyCollection
            Set tMap = sectorBaseEqmMap.GetAt(aVal)
            Dim rowList As String
            rowList = ""
            If existDiffData(tMap, rowList) = True Then
                dupCollect.Add Item:=rowList, key:=rowList
                If rowErrStr = "" Then
                    rowErrStr = rowList
                    errReportStr = rowList
                Else
                    rowErrStr = rowErrStr + "," + rowList
                    errReportStr = errReportStr + ";" + rowList
                End If
            End If
        Next
        
        Dim lineSet As Variant
        Dim maxColLen As Long
        Dim tVal As Variant
        Dim rowNum  As Long
        maxColLen = getColMaxLength(tmpsheet)
        If rowErrStr <> "" Then
            Call MsgBox(getResByKey("recordNotConsis") + errReportStr, vbInformation, getResByKey("Warning"))
            lineSet = Split(rowErrStr, ",")
            For Each tVal In lineSet
                rowNum = CLng(tVal)
                tmpsheet.range(Cells(rowNum, Cell_Index), Cells(rowNum, maxColLen)).Interior.colorIndex = 3
            Next
            CheckLteCellValid = False
            Exit Function
        End If
        
End Function

Private Function CheckUmtsCellValid(tmpsheet As Worksheet, sectorEqmproMap As CMapValueObject) As Boolean
        Dim aVal As Variant
        Dim tMap As CMap
        Dim rowErrStr As String
        Dim errReportStr As String
        CheckUmtsCellValid = True
        
        For Each aVal In sectorEqmproMap.KeyCollection
            Set tMap = sectorEqmproMap.GetAt(aVal)
            Dim rowList As String
            rowList = ""
            If existDiffData(tMap, rowList) = True Then
                dupCollect.Add Item:=rowList, key:=rowList
                If rowErrStr = "" Then
                    rowErrStr = rowList
                    errReportStr = rowList
                Else
                    rowErrStr = rowErrStr + "," + rowList
                    errReportStr = errReportStr + ";" + rowList
                End If
            End If
        Next
        
        Dim lineSet As Variant
        Dim maxColLen As Long
        Dim tVal As Variant
        Dim rowNum  As Long
        maxColLen = getColMaxLength(tmpsheet)
        If rowErrStr <> "" Then
            Call MsgBox(getResByKey("recordNotConsis") + errReportStr, vbInformation, getResByKey("Warning"))
            lineSet = Split(rowErrStr, ",")
            For Each tVal In lineSet
                rowNum = CLng(tVal)
                tmpsheet.range(Cells(rowNum, Cell_Index), Cells(rowNum, maxColLen)).Interior.colorIndex = 3
            Next
            CheckUmtsCellValid = False
            Exit Function
        End If
        
End Function


Private Function CheckDataDuplicate(tmpsheet As Worksheet, dupCol As Collection) As Boolean
    CheckDataDuplicate = True
    
    Dim lineSet As Variant
    Dim selectStr As String
    Dim dupStr As String
    Dim rowNum As Long
    Dim tVal As Variant
    
    For Each tVal In dupCol
        If InStr(1, tVal, ",") <> 0 Then
            dupCollect.Add Item:=tVal, key:=tVal
            If dupStr = "" Then
                dupStr = tVal
                selectStr = tVal
            Else
                dupStr = dupStr + ";" + tVal
                selectStr = selectStr + "," + tVal
            End If
        End If
    Next
    
    Dim maxColLen As Long
    maxColLen = getColMaxLength(tmpsheet)
    If dupStr <> "" Then
        Call MsgBox(getResByKey("recordDuplicate") + dupStr, vbInformation, getResByKey("Warning"))
        lineSet = Split(selectStr, ",")
        For Each tVal In lineSet
            rowNum = CLng(tVal)
            tmpsheet.range(Cells(rowNum, Cell_Index), Cells(rowNum, maxColLen)).Interior.colorIndex = 3
        Next
        CheckDataDuplicate = False
        Exit Function
    End If
    
End Function
Private Function CheckGsmCellValid(cellTrxMap As CMapValueObject) As Boolean
        Dim sVal As Variant
        Dim cellId As String
        Dim cellListStr As String
        Dim mulCellErrorListStr As String
        cellListStr = ""
        mulCellErrorListStr = ""
        
        CheckGsmCellValid = True
        
        For Each sVal In cellTrxMap.KeyCollection
            If IsAllTrxSetToCell(sVal, cellTrxMap) = False Then
                If cellListStr = "" Then
                    cellListStr = CStr(sVal)
                Else
                    cellListStr = cellListStr + "," + CStr(sVal)
                End If
            End If
            If IsAllTrxSetToSector(sVal, cellTrxMap) = False Then
                If mulCellErrorListStr = "" Then
                    mulCellErrorListStr = CStr(sVal)
                Else
                    mulCellErrorListStr = mulCellErrorListStr + "," + CStr(sVal)
                End If
            End If
        Next
        
        If cellListStr <> "" Then
            getSortedStr (cellListStr)
            Call MsgBox(getResByKey("trxNeedSet") + cellListStr, vbInformation, getResByKey("Warning"))
            CheckGsmCellValid = False
            Exit Function
        End If
        If mulCellErrorListStr <> "" Then
            getSortedStr (mulCellErrorListStr)
            Call MsgBox(getResByKey("sectorNeedSame") + mulCellErrorListStr, vbInformation, getResByKey("Warning"))
            CheckGsmCellValid = False
            Exit Function
        End If
End Function

Private Function existDiffData(aMap As CMap, rowList As String) As Boolean
    existDiffData = False
    Dim keyVal As Variant
    Dim tVal As Variant
    
    Dim dataCollect As Collection
    Set dataCollect = New Collection
    For Each keyVal In aMap.KeyCollection
        tVal = aMap.GetAt(keyVal)
        If Not Contains(dataCollect, CStr(tVal)) Then
            dataCollect.Add Item:=CStr(tVal), key:=CStr(tVal)
        End If
        If rowList = "" Then
            rowList = CStr(keyVal)
        Else
            rowList = rowList + "," + CStr(keyVal)
        End If
    Next
    
    If dataCollect.count() <> 1 Then
        existDiffData = True
        Exit Function
    End If
End Function
Private Function IsAllTrxSetToCell(tVal As Variant, cellTrxMap As CMapValueObject) As Boolean
    Dim aMap As CMapValueObject
    Dim col As New Collection
    Dim sVal As Variant
    Dim tmpVal As Variant
    Dim trxlistStr As String
    Dim trxArray() As String
    Dim index As Long
    Dim trxStr As String
    
    IsAllTrxSetToCell = True
    Set aMap = cellTrxMap.GetAt(tVal)

    For Each sVal In aMap.KeyCollection
        Dim tmpCol As New Collection
        Set tmpCol = aMap.GetAt(sVal)
        For Each tmpVal In tmpCol
            If Not Contains(col, CStr(tmpVal)) Then
                col.Add Item:=CStr(tmpVal), key:=CStr(tmpVal)
            End If
        Next
    Next

    If Cell_TrxListInfo.hasKey(tVal) Then
        trxlistStr = Cell_TrxListInfo.GetAt(tVal)
        trxArray = Split(trxlistStr, ",")
        For index = LBound(trxArray) To UBound(trxArray)
            If Not Contains(col, trxArray(index)) Then
                IsAllTrxSetToCell = False
                Exit Function
            End If
        Next
    End If
    
End Function

Private Function IsAllTrxSetToSector(tVal As Variant, cellTrxMap As CMapValueObject) As Boolean
    Dim aMap As CMapValueObject
    Dim sVal As Variant
    Dim tmpVal As Variant
    Dim trxlistStr As String
    Dim trxArray() As String
    Dim index As Long
    Dim trxStr As String

    IsAllTrxSetToSector = True
    If Cell_CellType_Map.hasKey(tVal) Then
        If CLng(Cell_CellType_Map.GetAt(tVal)) = 0 Then
            Exit Function
        End If
    End If

    Set aMap = cellTrxMap.GetAt(tVal)

    If Cell_TrxListInfo.hasKey(tVal) Then
        trxlistStr = Cell_TrxListInfo.GetAt(tVal)
        trxArray = Split(trxlistStr, ",")
    End If

    For Each sVal In aMap.KeyCollection
        Dim tmpCol As New Collection
        Set tmpCol = aMap.GetAt(sVal)
        For index = LBound(trxArray) To UBound(trxArray)
            If Not Contains(tmpCol, trxArray(index)) Then
                IsAllTrxSetToSector = False
                Exit Function
            End If
        Next
    Next
End Function

Private Function IsGSMMulCellVaild(tVal As Variant, trxSectorMap As CMapValueObject) As Boolean
    Dim col As Collection
    Dim sVal As Variant
    Dim val As Variant
    
    Set col = trxSectorMap.GetAt(tVal)
    For Each sVal In trxSectorMap.KeyCollection
        If tVal <> sVal Then
            Dim kcol As Collection
            Set kcol = trxSectorMap.GetAt(sVal)
            For Each val In col
                If Not Contains(kcol, CStr(val)) Then
                    IsGSMMulCellVaild = False
                    Exit Function
                End If
            Next
        End If
    Next
    IsGSMMulCellVaild = True
End Function

Private Function existInCollection(strValue As Variant, strCollection As Collection) As Boolean
    Dim sItem As Variant
    If Trim(CStr(strValue)) = "" Then
        existInCollection = True
        Exit Function
    End If
    For Each sItem In strCollection
        If sItem = strValue Then
            existInCollection = True
            Exit Function
        End If
    Next
    existInCollection = False
End Function


Private Function checkInputData(sheet As Worksheet, lineNo As Long, btsName As String, cell As String, _
        sector As String, board As String, antn As String, model As String, baseEqm As String, trx As String, sectoreqmproStr As String, errRangeCol As Collection) As Boolean
        checkInputData = True
        Dim rangeStr As String
        
        If btsName = "" Or existInCollection(btsName, selectBtsNameCol) = False Then
            rangeStr = sheet.Cells(lineNo, BTSName_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
        
        If checkDataValid(cellValueStr, cell) = False Then
            rangeStr = sheet.Cells(lineNo, Cell_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
    
        If Trim(sector) = "" Or isAInteger(sector) = False Then
            rangeStr = sheet.Cells(lineNo, Sector_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
        
        If checkDataValid(boardValueStr, board) = False And boardValueStr <> "" Then
            rangeStr = sheet.Cells(lineNo, Board_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
            
        If checkDataValid(antnValueStr, antn) = False Then
            rangeStr = sheet.Cells(lineNo, Ante_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
        
        If checkDataValid(modelValueStr, model) = False Then
            rangeStr = sheet.Cells(lineNo, Model_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
        
        If CELL_TYPE = 2 Then
            If Trim(sector) = "" Or isAInteger(baseEqm) = False Then
                rangeStr = sheet.Cells(lineNo, BaseEqm_Index).address(False, False)
                errRangeCol.Add Item:=rangeStr, key:=rangeStr
                checkInputData = False
            End If
        End If
        
        If CELL_TYPE = 0 Then
            If checkDataValid(trxValueStr, trx) = False Then
                rangeStr = sheet.Cells(lineNo, Trx_Index).address(False, False)
                errRangeCol.Add Item:=rangeStr, key:=rangeStr
                checkInputData = False
            End If
        End If
                
        If CELL_TYPE = 1 And SectorEqm_Index <> 0 Then
            If checkDataValid(sectoreqmValueStr, sectoreqmproStr, True) = False Then
                rangeStr = sheet.Cells(lineNo, SectorEqm_Index).address(False, False)
                errRangeCol.Add Item:=rangeStr, key:=rangeStr
                checkInputData = False
            End If
        End If
        
        
End Function
Public Sub AdjustAnntDataCheck(ByVal sheet As Object, ByVal target As range)
    Dim loccellIdCol As Long
    Dim sectoIdCol As Long
    Dim basebandeqmIdCol As Long
    
    loccellIdCol = findCertainValColumnNumber(sheet, 1, getResByKey("locellId"))
    sectoIdCol = findCertainValColumnNumber(sheet, 1, getResByKey("sectorId"))
    basebandeqmIdCol = findCertainValColumnNumber(sheet, 1, getResByKey("basebandeqmId"))
    
    Dim loccellIdValue As String
    Dim sectoIdValue As String
    Dim basebandeqmIdValue As String
    Dim nResponse As String
    

    Dim columnRange As range, cellRange As range
    
    For Each columnRange In target.Columns
        If columnRange.column = loccellIdCol Then
            For Each cellRange In columnRange
                loccellIdValue = Trim(cellRange.value)
                If target.row > 1 And loccellIdValue <> "" Then
                    If (CStr(loccellIdValue) < 0 Or CStr(loccellIdValue) > 268435455) Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~268435455]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                    End If
                End If
            Next cellRange
        ElseIf columnRange.column = sectoIdCol Then
            For Each cellRange In columnRange
                sectoIdValue = Trim(cellRange.value)
                If target.row > 1 And sectoIdValue <> "" Then
                    If CStr(sectoIdValue) < 0 Or CStr(sectoIdValue) > 1048576 Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~1048576]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                    End If
                End If
            Next cellRange
            
         ElseIf columnRange.column = basebandeqmIdCol Then
            For Each cellRange In columnRange
                basebandeqmIdValue = Trim(cellRange.value)
                If basebandeqmIdValue <> "" And target.row > 1 Then
                    If CStr(basebandeqmIdValue) < 0 Or (CStr(basebandeqmIdValue) > 23 And CStr(basebandeqmIdValue) <> 255) Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~23],[255]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                    End If
                End If
            Next cellRange
            
        End If
    Next columnRange
End Sub


Private Function isAInteger(ByRef tVal As String) As Boolean
    On Error GoTo ErrorHandler
    Dim k As Long
    If InStr(tVal, ".") <> 0 Then GoTo ErrorHandler
    k = CLng(tVal)
    If k <= 1048576 And k >= 0 Then
        isAInteger = True
    Else
        isAInteger = False
    End If
    Exit Function
ErrorHandler:
    isAInteger = False
End Function

Private Sub writeCellData(ByRef cellInfoMap As CMapValueObject)
    Dim cellsheet As Worksheet
    Dim maxRow As Long
    Dim rowNum As Long
    Dim baseStationName As String
    Dim cellId As String
    Dim keyVal As Variant
    Dim tmpVal As Variant
    Dim sectorStr As String
    Dim boradStr As String
    Dim tVal As CMap
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim antnIndex As Long
    Dim sectorIndex As Long
    Dim constCellTempCol As Long
    Dim tmpsheet As Worksheet
    
    Dim btsNameColIndex As Long
    
    Call getBaseStationMocNameAndAttrName(mocName, attrName)
    btsNameColIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    '找小区ID所在的列
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Call getCellMocNameAndAttrName(mocName, attrName)
    '判断是否为EuCellSectorEqm或EuPrbSectorEqm页签，如果是则mocName重新赋值
    If CELL_SHEET_NAME = "EUCELLSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUCELLSECTOREQM") Then
        mocName = "EuCellSectorEqm"
        attrName = "VLOCALCELLID"
    ElseIf CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
        mocName = "EuPrbSectorEqm"
        attrName = "VLOCALCELLID"
    End If
    constCellTempCol = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
        
    '获取天线端口所在列
    Call getAntenneMocNameAndAttr(mocName, attrName)
    If CELL_SHEET_NAME = "EUPRBSECTOREQM" Or CELL_SHEET_NAME = getResByKey("EUPRBSECTOREQM") Then
        mocName = "EuPrbSectorEqm"
    End If
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antnIndex = findColumnByName(cellsheet, columnName, 2)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    sectorIndex = findColumnByName(cellsheet, columnName, 2)
    
    maxRow = cellsheet.range("a1048576").End(xlUp).row
    
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    Dim sVal As Variant
    For Each sVal In Cell_Row_Map.KeyCollection
        Dim temKeyArr() As String

        temKeyArr = Split(CStr(sVal), "*")
        If UBound(temKeyArr) < 0 Then Exit Sub
        
        For rowNum = 3 To maxRow
'                baseStationName = cellsheet.Cells(rowNum, 1).value
                cellId = cellsheet.Cells(rowNum, constCellTempCol).value
                sectorStr = ""
                boradStr = ""
    '            If baseStationName = SITE_NAME Then
                If cellsheet.Cells(rowNum, btsNameColIndex).value = temKeyArr(0) And cellsheet.Cells(rowNum, constCellTempCol).value = temKeyArr(1) Then
'                    For Each keyVal In cellInfoMap.KeyCollection
'                        If cellId = keyVal Then
                            Set tVal = cellInfoMap.GetAt(CStr(sVal))
                            For Each tmpVal In tVal.KeyCollection
                                If sectorStr = "" Then
                                    sectorStr = tmpVal
                                    boradStr = tVal.GetAt(tmpVal)
                                Else
                                    sectorStr = sectorStr & "," & tmpVal
                                    boradStr = boradStr & ";" & tVal.GetAt(tmpVal)
                                End If
                            Next
'                        End If
'                   Next
                    cellsheet.Cells(rowNum, sectorIndex).value = sectorStr
                    cellsheet.Cells(rowNum, antnIndex).value = boradStr
                End If
        Next

    Next
End Sub


Private Sub sortMapByKey(ByRef mapObject As CMapValueObject, error As Boolean)
    On Error GoTo ErrorHandler
    Dim tmpValueMap As CMapValueObject
    Dim tCount As Long
    Dim keyVal As Variant
    Dim tmpVal As Variant
    Dim tVal As CMap
    Dim tmpMap As CMap
    Dim keyArray() As Variant
    Dim index As Long
    Dim antnStr As String
    Dim secStr As String
    Set tmpValueMap = New CMapValueObject
    error = False
    For Each keyVal In mapObject.KeyCollection
        Set tVal = mapObject.GetAt(keyVal)
        tCount = tVal.KeyCollection.count
        ReDim keyArray(tCount - 1)
        index = 0
        For Each tmpVal In tVal.KeyCollection
            keyArray(index) = CLng(tmpVal)
            index = index + 1
        Next
        Call QuickSort(keyArray())
        Set tmpMap = New CMap
        For index = LBound(keyArray) To UBound(keyArray)
            secStr = CStr(keyArray(index))
            If tVal.hasKey(secStr) = False Then Exit Sub
            antnStr = tVal.GetAt(secStr)
            Call getSortedStr(antnStr)
            Call tmpMap.SetAt(secStr, antnStr)
        Next
        Call tmpValueMap.SetAt(keyVal, tmpMap)
    Next
    Set mapObject = tmpValueMap
ErrorHandler:
    error = True
End Sub

Private Sub getSortedStr(ByRef infoStr As String)
    Dim strArray() As Variant
    Dim tmpArray As Variant
    Dim index As Long
    Dim tCount As Long
    Dim tmpStr As String
    
    If infoStr = "" Then Exit Sub
    
    tmpArray = Split(infoStr, ",")
    tCount = UBound(tmpArray) - LBound(tmpArray)
    ReDim strArray(tCount) As Variant
    For index = LBound(tmpArray) To UBound(tmpArray)
        strArray(index) = tmpArray(index)
    Next
    Call QuickSort(strArray())
    tmpStr = ""
    
    For index = LBound(strArray) To UBound(strArray)
        If tmpStr = "" Then
            tmpStr = strArray(index)
        Else
            tmpStr = tmpStr & "," & strArray(index)
        End If
    Next
    infoStr = tmpStr
End Sub

Private Sub changeAlerts(ByRef flag As Boolean)
    Application.EnableEvents = flag
    Application.DisplayAlerts = flag
    Application.ScreenUpdating = flag
End Sub

Public Sub QuickSort(ByRef lngArray() As Variant)
    Dim iLBound As Long
    Dim iUBound As Long
    Dim iTemp As Variant
    Dim iOuter As Long
    Dim iMax As Long
    
    iLBound = LBound(lngArray)
    iUBound = UBound(lngArray)

    If (iUBound - iLBound) Then
        For iOuter = iLBound To iUBound
            If lngArray(iOuter) > lngArray(iMax) Then iMax = iOuter
        Next iOuter

        iTemp = lngArray(iMax)
        lngArray(iMax) = lngArray(iUBound)
        lngArray(iUBound) = iTemp

        Call InnerQuickSort(lngArray, iLBound, iUBound)
    End If
End Sub


Private Sub InnerQuickSort(ByRef lngArray() As Variant, ByVal iLeftEnd As Long, ByVal iRightEnd As Long)
    Dim iLeftCur As Long
    Dim iRightCur As Long
    Dim iPivot As Variant
    Dim iTemp As Variant

    If iLeftEnd >= iRightEnd Then Exit Sub

    iLeftCur = iLeftEnd
    iRightCur = iRightEnd + 1
    iPivot = lngArray(iLeftEnd)
    
    Do
        Do
            iLeftCur = iLeftCur + 1
        Loop While lngArray(iLeftCur) < iPivot

        Do
            iRightCur = iRightCur - 1
        Loop While lngArray(iRightCur) > iPivot
        
        If iLeftCur >= iRightCur Then Exit Do
        
        iTemp = lngArray(iLeftCur)
        lngArray(iLeftCur) = lngArray(iRightCur)
        lngArray(iRightCur) = iTemp
    Loop

    lngArray(iLeftEnd) = lngArray(iRightCur)
    lngArray(iRightCur) = iPivot
    Call InnerQuickSort(lngArray, iLeftEnd, iRightCur - 1)
    Call InnerQuickSort(lngArray, iRightCur + 1, iRightEnd)
End Sub

Public Sub AdjustAntennPortSheet(ByVal sheet As Worksheet, ByVal target As range)
    Dim columnNumber As Long
    Dim rowNumber As Long
    Dim columName As String
    Dim keyStr As String
    Dim sVal As Variant
    Dim trxlistStr As String

    If target.rows.count <> 1 Or target.Columns.count <> 1 Then Exit Sub
    If CELL_TYPE <> 0 Then Exit Sub
  
    rowNumber = target.row
    columnNumber = target.column
    columName = sheet.Cells(1, columnNumber).value

    If columName = getResByKey("Frequency") Then
        keyStr = CStr(sheet.Cells(rowNumber, 1).value) + "*" + CStr(sheet.Cells(rowNumber, 2).value)
        
        If GCell_TrxListInfo.hasKey(keyStr) Then
            trxlistStr = GCell_TrxListInfo.GetAt(keyStr)
            With target.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=trxlistStr
            End With
        Else
            target.Validation.Delete
        End If
    End If
End Sub


Public Function getNeNamebyFunctionName(ByRef functionNeMap As CMap) As String
    Dim neNameVar As Variant
    Dim functionNameStr As String
    Dim functionNameStrArry() As String
    Dim index As Long
    getNeNamebyFunctionName = ""
    For Each neNameVar In functionNeMap.KeyCollection
        functionNameStr = functionNeMap.GetAt(neNameVar)
        If InStr(functionNameStr, ",") <> 0 Then
            functionNameStrArry = Split(functionNameStr, ",")
            For index = LBound(functionNameStrArry) To UBound(functionNameStrArry)
                If functionNameStrArry(index) = SITE_NAME Then
                    getNeNamebyFunctionName = CStr(neNameVar)
                    Exit Function
                End If
            Next
        Else
            If SITE_NAME = CStr(neNameVar) Then
                getNeNamebyFunctionName = SITE_NAME
                Exit Function
            End If
        End If
    Next
End Function

Private Function isCustomRxuBoardNo(ByRef boardStyleSheetName As String, ByRef groupName As String, ByRef rowNumber As Long) As Boolean
    Dim tempboardStyleData As CBoardStyleData
    If tempboardStyleData Is Nothing Then Set tempboardStyleData = New CBoardStyleData
    Call tempboardStyleData.init
    
    isCustomRxuBoardNo = True
    
    Dim baseStationData As CBaseStationData
    Set baseStationData = New CBaseStationData
    Dim boardStyleNeMap As CMap
    Dim neName As String
    Dim baseNeCustomInfoMap As CMapValueObject
    Dim customColValue As CMap
    Dim functionNeMap As CMap
    
    Call baseStationData.init
    Set functionNeMap = baseStationData.functionNeMap
    neName = getNeNamebyFunctionName(functionNeMap)
    Set baseNeCustomInfoMap = baseStationData.baseNeCustomInfoMap
    
    If Not baseNeCustomInfoMap.hasKey(neName) Then Exit Function
    Set customColValue = baseNeCustomInfoMap.GetAt(neName)
    
    Dim keyStr As String
    Dim boardStyleCustomMocInfoMap As CMap
    Dim mocName As String
    Dim customInfo As String
    Dim customInfoArry() As String
    Dim boardStylecustomColletter As String
    Dim boardStylecustomColName As String
    Dim boardStyleSheet As Worksheet
    Dim boardstylecellvalue As String
    
    Set boardStyleSheet = ThisWorkbook.Worksheets(boardStyleSheetName)
    
    mocName = tempboardStyleData.getMocNameByGroupName(groupName)
    Set boardStyleCustomMocInfoMap = tempboardStyleData.getBoardStyleCustomMocInfoMap
    
    If mocName <> "" Then
        keyStr = mocName + "," + groupName
        If boardStyleCustomMocInfoMap.hasKey(keyStr) Then
            customInfo = boardStyleCustomMocInfoMap.GetAt(keyStr)
            If customInfo <> "" Then
                customInfoArry = Split(customInfo, ",")
                boardStylecustomColName = customInfoArry(1)
                boardStylecustomColletter = customInfoArry(2)
                boardstylecellvalue = boardStyleSheet.range(boardStylecustomColletter & rowNumber).value
                isCustomRxuBoardNo = isExistSce(customColValue, boardstylecellvalue, boardStylecustomColName)
            End If
        End If

    End If

End Function

Private Function isExistSce(ByRef customColValue As CMap, ByRef scenValue As String, ByRef boardStylecustomColName As String) As Boolean
    Dim scenValueArry() As String
    Dim baseCustomColValue As String
    Dim index As Long
    
    isExistSce = True
    If Not customColValue.hasKey(boardStylecustomColName) Then Exit Function
    baseCustomColValue = customColValue.GetAt(boardStylecustomColName)
    If scenValue = "" And baseCustomColValue = "" Then Exit Function
    If scenValue = "" And baseCustomColValue <> "" Then
        isExistSce = False
        Exit Function
    End If
    scenValueArry = Split(scenValue, ",")

    For index = LBound(scenValueArry) To UBound(scenValueArry)
     If scenValueArry(index) = baseCustomColValue Then
        isExistSce = True
        Exit Function
     End If
    Next
    isExistSce = False
End Function

'刷新批注内容
Public Sub refreshCommentText(target As range, commentText As String)
    On Error GoTo ErrorHandler
    
    If target.comment Is Nothing Then
        With target
            With .addComment
                .Visible = False
                .text commentText
            End With
            With .comment.Shape
            .TextFrame.AutoSize = True
            .TextFrame.Characters.Font.Bold = True
            End With
        End With
    Else
        target.comment.text text:=commentText
    End If
    Exit Sub
     
ErrorHandler:
    Debug.Print "some exception in refreshCommentText, " & Err.Description
End Sub
