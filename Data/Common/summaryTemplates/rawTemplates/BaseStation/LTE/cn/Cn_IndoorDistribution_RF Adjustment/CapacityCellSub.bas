Attribute VB_Name = "CapacityCellSub"
Option Explicit

Private Const CELL_BAR_NAME = "SectorEqmBar"
Private Const DELFREQ_BAR_NAME = "DeleteTrxBar"
Private Const FINISH_BAR_NAME = "AdjustFinishBar"
Private Const CANCEL_BAR_NAME = "AdjustCancelBar"
Private Const Col_Width = 12
Private SITE_NAME As String
Private WRITESUCCESS As Boolean
Public InAdjustAntnPort As Boolean
Private CELL_SHEET_NAME As String
Public CELL_TYPE As Long '0代表是GSM小区，1是UMTS小区，2是LTE小区，对于GSM小区，需要根据小区单独判断类型Cell_CellType_Map
Private ROW_COUNT As Long
Private cellValueStr As String
'Private trxValueStr As String
Private boardValueStr As String
Private antnValueStr As String
Private modelValueStr As String
Private valueMap As Collection
Private errCollect As Collection
Private dupCollect As Collection

Private Cell_Row_Map As CMap  'GSM Cell and Row Mapping
Private Cell_CellType_Map As CMap  'GSM Cell Type
Private Cell_TrxListInfo As CMap
Private Cell_Index As Long
Private Trx_Index As Long
Private Sector_Index As Long
Private Board_Index As Long
Private Ante_Index As Long
Private Model_Index As Long
Private BaseEqm_Index As Long

Sub popUpTempSheetCannotChangeMsgbox()
    Call MsgBox(getResByKey("TempSheetCannotChange"), vbInformation, getResByKey("Warning"))
    ThisWorkbook.Worksheets(getResByKey("Temp Sheet")).Select
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

Sub createCellBar()
    Dim baseStationChooseBar As CommandBar
    Dim baseStationStyle As CommandBarButton
    Dim delChooseBar As CommandBar
    Dim delFreqStyle As CommandBarButton
    
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
      
      Dim actSheetName As String
      actSheetName = ThisWorkbook.ActiveSheet.name
      CELL_SHEET_NAME = actSheetName
      CELL_TYPE = cellSheetType(actSheetName)
      
      If CELL_TYPE = 0 Then
            Set delChooseBar = Application.CommandBars.Add(DELFREQ_BAR_NAME, msoBarBottom)
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
        End If
End Sub

Private Sub baseStationChoose()
    Dim actSheetName As String
    
    actSheetName = ThisWorkbook.ActiveSheet.name
    
    CELL_SHEET_NAME = actSheetName
    CELL_TYPE = cellSheetType(actSheetName)
    
    BaseStationForm.Show
End Sub


Private Sub deleteFrequency()
    On Error GoTo ErrorHandler
    
    DeleteFreqForm.Show
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

Sub AddSectorEqm(siteName As String)
    On Error GoTo ErrorHandler
    Set valueMap = New Collection
    SITE_NAME = siteName
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
    If Not insertCellIdColumn() Then
        Exit Sub
    End If
    If Not insertSectorIdColumn() Then
        Exit Sub
    End If
    If insertRxuBoardColumn() = False Then
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
    End If
    
    Call writeData
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
    
    Set Cell_Row_Map = New CMap
    Set Cell_CellType_Map = New CMap
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Call getCellMocNameAndAttrName(mocName, attrName)
    cellTypeIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    maxRow = cellsheet.range("a1048576").End(xlUp).row
    For rowNum = 3 To maxRow
        If cellsheet.Cells(rowNum, 1).value = SITE_NAME Then
            cellId = cellsheet.Cells(rowNum, cellTypeIndex).value
            If Cell_Row_Map.hasKey(cellId) = False Then
                Call Cell_Row_Map.SetAt(cellId, rowNum)
            End If
            
            If Cell_CellType_Map.hasKey(cellId) = False Then
                Call Cell_CellType_Map.SetAt(cellId, CELL_TYPE)
            End If
        End If
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
    With tmpsheet.Cells.Font
        .name = "Arial"
        .Size = 10
    End With
    maxColLen = getColMaxLength(tmpsheet)
    With tmpsheet.range(Cells(2, Cell_Index), Cells(2 + ROW_COUNT + 3, maxColLen))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub writeData()
    Dim Cell As CAntennes
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
    For Each Cell In valueMap
        Set antenneCollection = Cell.getAntenneCollection(sflag, rowVal)
        If sflag = False Then
            Call deleteTempSheet
            ThisWorkbook.Worksheets(CELL_SHEET_NAME).rows(rowVal).Select
            Exit Sub
        End If
        For Each row In antenneCollection
            boardStr = row(2)
            antnStr = row(3)
            rsModel = row(4)
            rangeStr = row(5)
            If checkBoardBasedValid(boardStr) = False Or (Trim(antnStr) <> "" And checkDataValid(antnValueStr, antnStr) = False) _
            Or (Trim(rsModel) <> "" And checkDataValid(modelValueStr, rsModel) = False) Then
                Call MsgBox(getResByKey("wrongDataInput") & rangeStr, vbInformation, getResByKey("Warning"))
                Call deleteTempSheet
                ThisWorkbook.Worksheets(CELL_SHEET_NAME).range(rangeStr).Select
                Exit Sub
            End If
            tempSheet.Cells(index, Cell_Index).value = row(0)
            If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
                tempSheet.Cells(index, Trx_Index).value = row(6)
            End If
            tempSheet.Cells(index, Sector_Index).value = row(1)
            tempSheet.Cells(index, Board_Index).value = boardStr
            tempSheet.Cells(index, Ante_Index).value = antnStr
            tempSheet.Cells(index, Model_Index).value = rsModel
            If CELL_TYPE = 2 Then
                tempSheet.Cells(index, BaseEqm_Index).value = row(7)
            End If
            index = index + 1
        Next
    Next
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


Private Function checkDataValid(tValueStr As String, tVal As String) As Boolean
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


Private Function insertCellIdColumn() As Boolean
        Dim myAttrName As String
        Dim myCellMocName As String
        Dim constCellTempCol As Long
        Call getCellMocNameAndAttrName(myCellMocName, myAttrName)
        constCellTempCol = getColNum(CELL_SHEET_NAME, 2, myAttrName, myCellMocName)
        
        Dim cellsheet As Worksheet
        insertCellIdColumn = True
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Cell_Index = 1
        cellsheet.Cells(2, constCellTempCol).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Cell_Index)
        Dim cellsStr As String
        cellsStr = ""

        Dim index As Long
        For index = 2 To cellsheet.range("a1048576").End(xlUp).row
            If cellsheet.Cells(index, 1).value = SITE_NAME Then
                If cellsStr <> "" Then
                     cellsStr = cellsStr + "," + cellsheet.Cells(index, constCellTempCol).value
                Else
                    cellsStr = cellsheet.Cells(index, constCellTempCol).value
                End If
            End If
        Next

        Dim cellRang As range
        Set cellRang = ThisWorkbook.ActiveSheet.range("A2:A" + CStr(2 + ROW_COUNT + 3))
        If cellsStr <> "" Then
                cellValueStr = cellsStr
                With cellRang.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=cellsStr
                End With
                ThisWorkbook.ActiveSheet.columns(Cell_Index).ColumnWidth = Col_Width
        End If
        insertCellIdColumn = True
End Function

Public Sub insertTrxColumn()
        Trx_Index = 2
        ThisWorkbook.ActiveSheet.Cells(1, Board_Index).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Trx_Index)
        ThisWorkbook.ActiveSheet.Cells(1, Trx_Index).value = getResByKey("Frequency")

        ThisWorkbook.ActiveSheet.columns(Trx_Index).ColumnWidth = Col_Width
End Sub


Private Function insertSectorIdColumn() As Boolean
        Dim mocName As String
        Dim attrName As String
        Dim sectorColumnName As String
        Dim columnIndex As Long
        Dim cellsheet As Worksheet
        
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Call getSectorMocNameAndAttr(mocName, attrName)
        sectorColumnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex < 0 Then
            insertSectorIdColumn = False
            Exit Function
        End If
        If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
            Sector_Index = 3
        Else
            Sector_Index = 2
        End If
        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Sector_Index)
        ThisWorkbook.ActiveSheet.columns(Sector_Index).ColumnWidth = Col_Width
        insertSectorIdColumn = True
End Function

Private Function insertRxuBoardColumn() As Boolean
        insertRxuBoardColumn = True
        Dim brdStyleSheetName As String
        Dim grpCollection As Collection
        Dim brdStr As String
        Dim brdGrp
        Dim startRow As Long
        Dim endRow As Long
        Dim index As Long
        Dim btsIndex As Long
        Dim charStr As String
        Dim mainSheetName As String
        Dim mainSheet As Worksheet
        btsIndex = -1
        
        brdStyleSheetName = findBoardStyleSheet(btsIndex)
        If brdStyleSheetName = "" Then
                Call MsgBox(getResByKey("NoBoradStyle"), vbInformation, getResByKey("Warning"))
                Call deleteTempSheet
                mainSheetName = GetMainSheetName
                Set mainSheet = ThisWorkbook.Worksheets(mainSheetName)
                mainSheet.Select
                If btsIndex <> -1 Then
                    mainSheet.rows(btsIndex).Select
                End If
                insertRxuBoardColumn = False
                Exit Function
        End If
        
        brdStr = ""
        Set grpCollection = findBrdGroups
        Dim boardStyleSheet As Worksheet
        
        Set boardStyleSheet = ThisWorkbook.Worksheets(brdStyleSheetName)
        
        If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
            Board_Index = 4
            charStr = "D"
        Else
            Board_Index = 3
            charStr = "C"
        End If
        
        For Each brdGrp In grpCollection
            Call getStartAndEndRowNum(brdStyleSheetName, CStr(brdGrp), startRow, endRow)
            boardStyleSheet.Cells(startRow + 1, 1).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Board_Index)
            For index = startRow + 2 To endRow
                If brdStr = "" Then
                    brdStr = boardStyleSheet.Cells(index, 1).value
                Else
                    brdStr = brdStr + "," + boardStyleSheet.Cells(index, 1).value
                End If
            Next
        Next
        
        Dim cellRang As range
        Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
        If brdStr <> "" Then
                boardValueStr = brdStr
                With cellRang.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=brdStr
                End With
                ThisWorkbook.ActiveSheet.columns(Board_Index).ColumnWidth = Col_Width
        End If
End Function

Private Sub getStartAndEndRowNum(brdSheetName As String, ByRef groupName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(brdSheetName)
    groupNameStartRowNumber = getGroupNameStartRowNumber(sh, groupName)
    groupNameEndRowNumber = groupNameStartRowNumber + getCurrentRegionRowsCount(sh, groupNameStartRowNumber) - 1
End Sub

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
    Dim cellsheet As Worksheet
    Dim index As Long
    Dim antenneIndex As Long
    Dim sectorArray As Validation
    
    Set Cell_TrxListInfo = New CMap
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    '找小区ID所在的列
    Dim constCellTempCol As Long
    Call getCellMocNameAndAttrName(mocName, attrName)
    constCellTempCol = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    '获取天线端口所在列
    Call getAntenneMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    columnIndex = findColumnByName(cellsheet, columnName, 2)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    sectorIndex = findColumnByName(cellsheet, columnName, 2)
    
    If CELL_TYPE = 2 Then
        Call getBaseEqmMocNameAndAttrName(mocName, attrName)
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        baseEqmIndex = findColumnByName(cellsheet, columnName, 2)
    End If
    
    rowCount = 0
    If columnIndex <= 0 Then
        calculateRow = rowCount
        Exit Function
    End If
    Dim antennes As CAntennes
  
    For index = 3 To cellsheet.range("a1048576").End(xlUp).row
        If cellsheet.Cells(index, 1).value = SITE_NAME And cellsheet.Cells(index, constCellTempCol).value <> "" Then
            Set antennes = New CAntennes
            antennes.cellId = cellsheet.Cells(index, constCellTempCol).value
            If (CELL_TYPE = 0) Then
                antennes.trxId = TrxInfoMgr.getFreqLstStrByGCell(CELL_SHEET_NAME, SITE_NAME, antennes.cellId)
                If antennes.trxId = "" Then
                    chkPassed = False
                    Exit Function
                End If
                If Cell_TrxListInfo.hasKey(antennes.cellId) = False Then
                    Call Cell_TrxListInfo.SetAt(antennes.cellId, antennes.trxId)
                End If
            Else
                antennes.trxId = ""
            End If
            antennes.antennes = cellsheet.Cells(index, columnIndex).value
            antennes.sectorIds = cellsheet.Cells(index, sectorIndex).value
            If CELL_TYPE = 2 Then
                antennes.baseEqmIds = cellsheet.Cells(index, baseEqmIndex).value
            End If
            antennes.ranges = cellsheet.Cells(index, columnIndex).address(False, False)
            antennes.row = index
            rowCount = rowCount + antennes.rowCount
            If Not Contains(valueMap, cellsheet.Cells(index, constCellTempCol).value) Then
                valueMap.Add Item:=antennes, key:=cellsheet.Cells(index, constCellTempCol).value
            End If
        End If
    Next
    
    calculateRow = rowCount
End Function

Private Function findColumnByName(sh As Worksheet, columnName As String, row As Long) As String
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
    End If
End Sub

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
    End If
End Function

Private Function getBaseEqmMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 2 Then
        attrName = "BaseBandEqmId"
        mocName = "eUCellSectorEqm"
    End If
End Function

Private Function cellSheetType(sheetName As String) As Long
    If sheetName = "GSM Cell" Or sheetName = "GSM小区" Then
        cellSheetType = 0
    ElseIf sheetName = "UMTS Cell" Or sheetName = "UMTS小区" Then
        cellSheetType = 1
    ElseIf sheetName = "LTE Cell" Or sheetName = "LTE小区" Then
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
    
    referencedString = "R0A,R0B,R0C,R0D,R0E,R0F,R0G,R0H,R1A,R1B,R1C,R1D,R2A,R2B,R2C,R2D,R3A,R3B,R3C,R3D,R4A,R4B,R4C,R4D,R5A,R5B,R5C,R5D,R6A,R6B,R6C,R6D,R7A,R7B,R7C,R7D"
    
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Call getAntenneMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antnIndex = findColumnByName(cellsheet, columnName, 2)
    
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        Ante_Index = 5
        charStr = "E"
    Else
        Ante_Index = 4
        charStr = "D"
    End If
    
    cellsheet.Cells(2, antnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Ante_Index)
    
    Dim antnRang As range
    Set antnRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
    If referencedString <> "" Then
            antnValueStr = referencedString
            With antnRang.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=referencedString
            End With
            ThisWorkbook.ActiveSheet.columns(Ante_Index).ColumnWidth = Col_Width
    End If
End Sub

Public Sub insertAnteModelColumn()
         Dim charStr As String
         Dim referencedString As String
         
        If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
            Model_Index = 6
            charStr = "F"
        Else
            Model_Index = 5
            charStr = "E"
        End If
        referencedString = "RXTX_MODE,RX_MODE,TX_MODE"
        
        ThisWorkbook.ActiveSheet.Cells(1, Board_Index).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Model_Index)
        ThisWorkbook.ActiveSheet.Cells(1, Model_Index).value = getResByKey("anteModel")

        Dim cellRang As range
        modelValueStr = referencedString
        Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
                With cellRang.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=referencedString
                End With
        ThisWorkbook.ActiveSheet.columns(Model_Index).ColumnWidth = Col_Width
End Sub


Private Function insertBaseEqmColumn() As Boolean
        Dim mocName As String
        Dim attrName As String
        Dim sectorColumnName As String
        Dim columnIndex As Long
        Dim cellsheet As Worksheet
        
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        Call getBaseEqmMocNameAndAttrName(mocName, attrName)
        sectorColumnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex < 0 Then
            insertBaseEqmColumn = False
            Exit Function
        End If
        BaseEqm_Index = 6
        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, BaseEqm_Index)
        ThisWorkbook.ActiveSheet.columns(BaseEqm_Index).ColumnWidth = Col_Width
        insertBaseEqmColumn = True
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
    Dim trxListStr As String
    For Each sVal In Cell_TrxListInfo.KeyCollection
        trxListStr = Cell_TrxListInfo.GetAt(sVal)
        Call writeGSMOneCell(tmpsheet, cellsheet, CStr(sVal), trxListStr, antnIndex, secIndex)
    Next
End Sub


Private Function writeGSMOneCell(tmpsheet As Worksheet, cellsheet As Worksheet, cellId As String, trxListStr As String, antnIndex As Long, secIndex As Long) As Boolean
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
    
    Dim cellInRow As String
    
    writeGSMOneCell = True
    If Cell_Row_Map.hasKey(cellId) = False Then
        writeGSMOneCell = False
        Exit Function
    End If
    cellInRow = Cell_Row_Map.GetAt(cellId)
     
    secStr = "-1"
    antaStr = ""
    trxArray = Split(trxListStr, ",")
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For index = LBound(trxArray) To UBound(trxArray)
        sector = "-1"
        For rowNum = 2 To maxRow
            trxId = tmpsheet.Cells(rowNum, Trx_Index).value
            sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
            If tmpsheet.Cells(rowNum, Cell_Index).value = cellId Then
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
    
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    
    Call getAntenneMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antnIndex = findColumnByName(cellsheet, columnName, 2)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    secIndex = findColumnByName(cellsheet, columnName, 2)
    
    If CELL_TYPE = 2 Then
        Call getBaseEqmMocNameAndAttrName(mocName, attrName)
        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
        baseEqmIndex = findColumnByName(cellsheet, columnName, 2)
    End If
    
    Dim sVal As Variant
    For Each sVal In Cell_Row_Map.KeyCollection
        Call writeLTEOneCell(tmpsheet, cellsheet, CStr(sVal), secIndex, antnIndex, baseEqmIndex)
    Next
End Sub

Private Sub writeLTEOneCell(tmpsheet As Worksheet, cellsheet As Worksheet, cellId As String, _
                                                    secIndex As Long, antnIndex As Long, baseEqmIndex As Long)
    Dim maxRow As Long
    Dim rowNum As Long
    Dim sectorId As String
    Dim baseEqmId As String
    Dim boardAnte As String
    Dim sectorIds As String
    sectorIds = ""
    
    Dim cellInRow As String
    cellInRow = Cell_Row_Map.GetAt(cellId)
    
    Dim sectorMap As CMap
    Set sectorMap = New CMap
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
         If tmpsheet.Cells(rowNum, Cell_Index).value = cellId Then
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
    
    Dim antenna As String
    Dim antaStr As String
    antaStr = ""
    
    sectorArray = Split(sectorIds, ",")
    
    Dim index As Long
    For index = LBound(sectorArray) To UBound(sectorArray)
        antenna = ""
        For rowNum = 2 To maxRow
            If tmpsheet.Cells(rowNum, Cell_Index).value = cellId Then
                sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
                If sectorId = sectorArray(index) Then
                    If CELL_TYPE = 2 Then
                        baseEqmId = tmpsheet.Cells(rowNum, BaseEqm_Index).value
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
    Next
    
    cellsheet.Cells(cellInRow, secIndex).value = sectorIds
    cellsheet.Cells(cellInRow, antnIndex).value = antaStr
    If CELL_TYPE = 2 Then
        cellsheet.Cells(cellInRow, baseEqmIndex).value = baseEqmIdStr
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
    
    Set cellInfoMap = New CMapValueObject
    
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
        cellId = tmpsheet.Cells(rowNum, Cell_Index).value
        sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
        board = tmpsheet.Cells(rowNum, Board_Index).value
        antn = tmpsheet.Cells(rowNum, Ante_Index).value
        model = tmpsheet.Cells(rowNum, Model_Index).value
        boardInfo = board & "_" & antn & ":" & model
        
        isExist = cellInfoMap.hasKey(cellId)
        If isExist Then
            Set tVal = cellInfoMap.GetAt(cellId)
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
            Call cellInfoMap.SetAt(cellId, tmpMap)
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
    Dim errflag As Boolean
    Dim lineStr As String
    Dim tVal As Variant
    Dim eRange As range
    Dim dupCol As Collection
    Dim keyStr As String
    Dim cellTrxMap As CMapValueObject  '校验GSM小区的频点必须都要配置
    'Dim trxSectorMap As New CMapValueObject '校验GSM共小区情况下每个频点配置的扇区都要一致
    
    Dim sectorBaseEqmMap As CMapValueObject
    Dim sectorBaseEqm As String
    Set sectorBaseEqmMap = New CMapValueObject
    
    Set cellTrxMap = New CMapValueObject
    checkUserData = True
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    Set errRangeCol = New Collection
    Set dupCol = New Collection
    
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
        cellStr = tmpsheet.Cells(rowNum, Cell_Index).value
        sectorStr = tmpsheet.Cells(rowNum, Sector_Index).value
        boardStr = tmpsheet.Cells(rowNum, Board_Index).value
        antnStr = tmpsheet.Cells(rowNum, Ante_Index).value

        modelStr = tmpsheet.Cells(rowNum, Model_Index).value
        If CELL_TYPE = 2 Then
            baseEqmStr = tmpsheet.Cells(rowNum, BaseEqm_Index).value
        End If
        errflag = checkInputData(tmpsheet, rowNum, cellStr, sectorStr, boardStr, antnStr, modelStr, baseEqmStr, errRangeCol)
        If errflag Then
            If CELL_TYPE = 0 Then
                trxStr = tmpsheet.Cells(rowNum, Trx_Index).value
                Dim aMap As CMapValueObject
                Set aMap = New CMapValueObject
                Dim col As Collection
                Set col = New Collection
                If cellTrxMap.hasKey(cellStr) Then
                     Dim needAdd As Boolean
                     needAdd = True
                     Set aMap = cellTrxMap.GetAt(cellStr)
                     If aMap.hasKey(sectorStr) Then
                        Set col = aMap.GetAt(sectorStr)
                        If Contains(col, trxStr) Then
                            needAdd = False
                        End If
                    End If
                    If needAdd = True Then
                        col.Add Item:=trxStr, key:=trxStr
                        Call aMap.SetAt(sectorStr, col)
                        cellTrxMap.RemoveKey (cellStr)
                        Call cellTrxMap.SetAt(cellStr, aMap)
                    End If
                Else
                    col.Add Item:=trxStr, key:=trxStr
                    Call aMap.SetAt(sectorStr, col)
                    Call cellTrxMap.SetAt(cellStr, aMap)
                End If
            
                keyStr = cellStr + "_" + sectorStr + "_" + trxStr + "_" + boardStr + "_" + antnStr + "_" + modelStr
            ElseIf CELL_TYPE = 2 Then
                keyStr = cellStr + "_" + sectorStr + "_" + boardStr + "_" + antnStr + "_" + modelStr + "_" + baseEqmStr
                
                sectorBaseEqm = cellStr + "_" + sectorStr
                Dim bMap As CMap
                Set bMap = New CMap
                If sectorBaseEqmMap.hasKey(sectorBaseEqm) Then
                    Set bMap = sectorBaseEqmMap.GetAt(sectorBaseEqm)
                    sectorBaseEqmMap.RemoveKey (sectorBaseEqm)
                End If
                Call bMap.SetAt(rowNum, baseEqmStr)
                Call sectorBaseEqmMap.SetAt(sectorBaseEqm, bMap)
            Else
                keyStr = cellStr + "_" + sectorStr + "_" + boardStr + "_" + antnStr + "_" + modelStr
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
    Dim trxListStr As String
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
        trxListStr = Cell_TrxListInfo.GetAt(tVal)
        trxArray = Split(trxListStr, ",")
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
    Dim trxListStr As String
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
        trxListStr = Cell_TrxListInfo.GetAt(tVal)
        trxArray = Split(trxListStr, ",")
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

Private Function checkInputData(sheet As Worksheet, lineNo As Long, Cell As String, _
        sector As String, board As String, antn As String, model As String, baseEqm As String, errRangeCol As Collection) As Boolean
        checkInputData = True
        Dim rangeStr As String
        If checkDataValid(cellValueStr, Cell) = False Then
            rangeStr = sheet.Cells(lineNo, Cell_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
    
        If Trim(sector) = "" Or isAInteger(sector) = False Then
            rangeStr = sheet.Cells(lineNo, Sector_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
        
        If checkDataValid(boardValueStr, board) = False Then
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
        
End Function

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
    '找小区ID所在的列
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Call getCellMocNameAndAttrName(mocName, attrName)
    constCellTempCol = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
        
    '获取天线端口所在列
    Call getAntenneMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antnIndex = findColumnByName(cellsheet, columnName, 2)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    sectorIndex = findColumnByName(cellsheet, columnName, 2)
    
    maxRow = cellsheet.range("a1048576").End(xlUp).row
    
    For rowNum = 3 To maxRow
            baseStationName = cellsheet.Cells(rowNum, 1).value
            cellId = cellsheet.Cells(rowNum, constCellTempCol).value
            sectorStr = ""
            boradStr = ""
            If baseStationName = SITE_NAME Then
                For Each keyVal In cellInfoMap.KeyCollection
                    If cellId = keyVal Then
                        Set tVal = cellInfoMap.GetAt(keyVal)
                        For Each tmpVal In tVal.KeyCollection
                            If sectorStr = "" Then
                                sectorStr = tmpVal
                                boradStr = tVal.GetAt(tmpVal)
                            Else
                                sectorStr = sectorStr & "," & tmpVal
                                boradStr = boradStr & ";" & tVal.GetAt(tmpVal)
                            End If
                        Next
                    End If
               Next
                cellsheet.Cells(rowNum, sectorIndex).value = sectorStr
                cellsheet.Cells(rowNum, antnIndex).value = boradStr
            End If
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

