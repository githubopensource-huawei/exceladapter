Attribute VB_Name = "CapacityGbtsCell"

Option Explicit

Public hyperLintFlag As Boolean
Private Const CELL_BAR_NAME = "SectorEqmBar"
Private Const FINISH_BAR_NAME = "AdjustFinishBar"
Private Const CANCEL_BAR_NAME = "AdjustCancelBar"

Private SITE_NAME As String
Private WRITESUCCESS As Boolean

Private CELL_SHEET_NAME As String
Private CELL_TYPE As Long
Private ROW_COUNT As Long
Private cellValueStr As String
Private boardValueStr As String
Private portNoValueStr As String
Private antnValueStr As String
Private antnGrpIDValueStr As String
Private valueMap As Collection
Private Const CON_SHARP = "#"
Private Const CON_BRACKET_RIGHT = "]"
Private Const CON_BRACKET_LEFT = "["
Private Const CON_COMMA = ","
Private Const NormalPattern = 1
Public Const CUSTOM_SCENARIO_MOC_NAME = "Customization_CME"
Public Const CUSTOM_SCENARIO_ATTR_NAME = "Scenario"
Public Sub AddTrxBinds(siteName As String, CellSheetName As String)
    On Error GoTo ErrorHandler
    Set valueMap = New Collection
    SITE_NAME = siteName
    WRITESUCCESS = False
    CELL_SHEET_NAME = CellSheetName
    CELL_TYPE = cellSheetType(CellSheetName)
    
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
    If Not insertFreqColumn() Then
        Exit Sub
    End If
    'Exit Sub
    If Not insertRxuBoardColumn() Then
        Call delTmpSht
        Exit Sub
    End If
        
    Call insertAntenneColumn
    
    Call insertAntenneGrpColumn
    
    Call writeData
    
    Call AdjustSheetStyle
    
    WRITESUCCESS = True
    Exit Sub
ErrorHandler:
    WRITESUCCESS = False
    
End Sub

Private Sub AdjustSheetStyle()
    Dim tmpSheet As Worksheet
    Dim sheetRange As range
    Set tmpSheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    With tmpSheet.Cells.Font
        .name = "Arial"
        .Size = 10
    End With
    With tmpSheet.range(Cells(2, 1), Cells(1 + ROW_COUNT, 6))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders.LineStyle = xlContinuous
    End With
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


Private Sub writeData()
    Dim Cell As CAntGBts
    Dim row As Variant
    Dim index As Long
    
    Dim boardStr As String
    Dim antnStr As String
    Dim antnGrpIdStr As String
    Dim tempSheet As Worksheet
    Dim antenneCollection As Collection
    Dim rangeStr As String
    Dim portStr As String
    
    
    Set tempSheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    index = 2
    For Each Cell In valueMap
        Set antenneCollection = Cell.getAntenneCollection
        
        For Each row In antenneCollection
            'boardStr = row(2)
            'rangeStr = row(3)
            'rangeStr = row(3)
            'antnStr = row(4)
            'antnGrpIdStr = row(5)
            
            tempSheet.Cells(index, 1).value = row(0)
            tempSheet.Cells(index, 2).value = row(1)
            tempSheet.Cells(index, 3).value = row(2)
            tempSheet.Cells(index, 4).value = row(3)
            tempSheet.Cells(index, 5).value = row(4)
            If row(4) <> "BY_ANTGRP" Then
                tempSheet.Cells(index, 6).Interior.colorIndex = SolidColorIdx
                tempSheet.Cells(index, 6).Interior.Pattern = SolidPattern
                tempSheet.Cells(index, 6).value = ""
                'tempSheet.Cells(index, 6).Validation.ShowInput = False
            Else
                tempSheet.Cells(index, 6).value = row(5)
            End If
            index = index + 1
        Next
    Next
End Sub

Private Function checkDataValid(tValueStr As String, tVal As String, allowEmpty As Boolean) As Boolean
        If allowEmpty And ("" = tVal) Then
            checkDataValid = True
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


Private Function insertCellIdColumn() As Boolean
        Dim myAttrName As String
        Dim myCellMocName As String
        Dim constCellTempCol As Long
        Call getCellMocNameAndAttrName(myCellMocName, myAttrName)
        constCellTempCol = getColNum(CELL_SHEET_NAME, 2, myAttrName, myCellMocName)
        
        Dim cellsheet As Worksheet
        insertCellIdColumn = True
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        cellsheet.Cells(2, constCellTempCol).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, 1)
        ThisWorkbook.ActiveSheet.Cells(1, 1).clearComments

        '仅为了格式
        ThisWorkbook.ActiveSheet.Cells(1, 1).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, 2)
        ThisWorkbook.ActiveSheet.Cells(1, 2).value = ""
        ThisWorkbook.ActiveSheet.Cells(1, 1).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, 3)
        ThisWorkbook.ActiveSheet.Cells(1, 3).value = ""
        ThisWorkbook.ActiveSheet.Cells(1, 1).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, 4)
        ThisWorkbook.ActiveSheet.Cells(1, 4).value = ""
        ThisWorkbook.ActiveSheet.Cells(1, 1).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, 5)
        ThisWorkbook.ActiveSheet.Cells(1, 5).value = ""
        ThisWorkbook.ActiveSheet.Cells(1, 1).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, 6)
        ThisWorkbook.ActiveSheet.Cells(1, 6).value = ""
        Dim cellsStr As String
        cellsStr = ""
        Dim index As Long
        For index = 2 To cellsheet.range("a1048576").End(xlUp).row
            If cellsheet.Cells(index, getGcellBTSNameCol(CELL_SHEET_NAME)).value = SITE_NAME Then
                If cellsStr <> "" Then
                    cellsStr = cellsStr + "," + cellsheet.Cells(index, constCellTempCol).value
                Else
                    cellsStr = cellsheet.Cells(index, constCellTempCol).value
                End If
            End If
        Next
        Dim cellRang As range
        Set cellRang = ThisWorkbook.ActiveSheet.range("A2:A" + CStr(1 + ROW_COUNT))
        If cellsStr <> "" Then
                cellValueStr = cellsStr
                With cellRang.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=cellsStr
                End With
        End If

        insertCellIdColumn = True
End Function

Private Function insertFreqColumn() As Boolean
        
        ThisWorkbook.ActiveSheet.Cells(1, 2).value = getResByKey("GCELLFREQ")
        
        insertFreqColumn = True
End Function
Private Function insertAntenneGrpColumn() As Boolean
    Dim brdStyleSheetName As String
    Dim boardstylesheet As Worksheet
    Dim dspCategory As String
    Dim innerCategory As String
    Dim antGrpIdStr As String
    Dim startRow As Long
    Dim endRow As Long
    Dim index As Long
    
    insertAntenneGrpColumn = True
    
    brdStyleSheetName = findBoardStyleSheet
    dspCategory = findAntGrpName
    innerCategory = "BTSANTENNAGROUP"
    
    ThisWorkbook.ActiveSheet.Cells(1, 6).value = getResByKey("ANTGROUPID")
    
    If Not ("" = brdStyleSheetName) Then
        Set boardstylesheet = ThisWorkbook.Worksheets(brdStyleSheetName)
        If Not ("" = dspCategory) Then
            Call getStartAndEndRowNum(brdStyleSheetName, dspCategory, startRow, endRow)
            Dim antGrpIdColName As String
            Dim mocName As String
            Dim attrName As String
            Dim colNo As Long
                    
            Call getAntGrpIdMocNameAndAttr(mocName, attrName)
            antGrpIdColName = getColNameBaseMapDef(getResByKey("Board Style"), attrName, innerCategory)
            colNo = getColNumByName(brdStyleSheetName, startRow + 1, antGrpIdColName)

            For index = startRow + 2 To endRow
                If antGrpIdStr = "" Then
                    antGrpIdStr = boardstylesheet.Cells(index, colNo).text
                Else
                    antGrpIdStr = antGrpIdStr + "," + boardstylesheet.Cells(index, colNo).text
                End If
            Next
        End If
    Else
        MsgBox getResByKey("brdstylecannotbenull") & SITE_NAME, vbExclamation, getResByKey("Warning")
        insertAntenneGrpColumn = False
        
        Exit Function
    End If
    
    Dim cellRang As range
    Set cellRang = ThisWorkbook.ActiveSheet.range("F2:F" + CStr(1 + ROW_COUNT))
    
    If antGrpIdStr <> "" Then
            antnGrpIDValueStr = antGrpIdStr
            With cellRang.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=antGrpIdStr
            End With
    End If
End Function

Private Function insertRxuBoardColumn() As Boolean
        Dim brdStyleSheetName As String
        Dim grpCollection As Collection
        Dim brdStr As String
        Dim brdGrp
        Dim startRow As Long
        Dim endRow As Long
        Dim index As Long
        Dim colNo As Long
             
        
        insertRxuBoardColumn = True
        
        brdStyleSheetName = findBoardStyleSheet
        brdStr = ""
        Set grpCollection = findBrdGroups
        Dim boardstylesheet As Worksheet
                
        If Not ("" = brdStyleSheetName) Then
            Set boardstylesheet = ThisWorkbook.Worksheets(brdStyleSheetName)
            For Each brdGrp In grpCollection
                Dim strArr() As String
                strArr = Split(CStr(brdGrp), CON_SHARP)
                
                Dim dspCategory As String
                dspCategory = strArr(LBound(strArr))
                Dim innerCategory As String
                innerCategory = strArr(UBound(strArr))
                
                Call getStartAndEndRowNum(brdStyleSheetName, dspCategory, startRow, endRow)
                boardstylesheet.Cells(startRow + 1, 7).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, 3)
                ThisWorkbook.ActiveSheet.Cells(1, 3).clearComments
                
              
                Dim brdNoColName As String
                Dim mocName As String
                Dim attrName As String
                        
                Call getPhybrdMocNameAndAttr(mocName, attrName)
                brdNoColName = getColNameBaseMapDef(getResByKey("Board Style"), attrName, innerCategory)
                colNo = getColNumByName(brdStyleSheetName, startRow + 1, brdNoColName)

                For index = startRow + 2 To endRow
                    If brdStr = "" Then
                        brdStr = boardstylesheet.Cells(index, colNo).value
                    Else
                        brdStr = brdStr + "," + boardstylesheet.Cells(index, colNo).value
                    End If
                Next
            Next
        Else
            MsgBox getResByKey("brdstylecannotbenull") & SITE_NAME, vbExclamation, getResByKey("Warning")
            insertRxuBoardColumn = False
            Exit Function
        End If
        
        
        Dim cellRang As range
        Set cellRang = ThisWorkbook.ActiveSheet.range("C2:C" + CStr(1 + ROW_COUNT))
        If brdStr <> "" Then
                boardValueStr = brdStr
                With cellRang.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=brdStr
                End With
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
        If UCase(relationSheet.Cells(index, 6).value) = "BTSRXUBRD" Then
                If Not Contains(grpCollection, UCase(relationSheet.Cells(index, 6).value)) Then
                    grpCollection.Add Item:=relationSheet.Cells(index, 2).value + CON_SHARP + UCase(relationSheet.Cells(index, 6).value), key:=UCase(relationSheet.Cells(index, 6).value)
                End If
        End If
    Next
    Set findBrdGroups = grpCollection
End Function
Private Function findAntGrpName() As String
    Dim mappingDefSht As Worksheet
    Dim index As Long
    Dim rowCount As Long
    Set mappingDefSht = ThisWorkbook.Worksheets("MAPPING DEF")
    rowCount = mappingDefSht.range("a1048576").End(xlUp).row
    For index = 2 To rowCount
        If UCase(mappingDefSht.Cells(index, 4).value) = "BTSANTENNAGROUP" Then
            findAntGrpName = mappingDefSht.Cells(index, 2).text
            Exit Function
        End If
    Next
End Function


Private Function findBoardStyleSheet() As String
    Dim groupName As String
    Dim columnName As String
    Dim mainSheet As Worksheet
    Dim mainSheetName As String
    mainSheetName = GetMainSheetName
    Set mainSheet = ThisWorkbook.Worksheets(mainSheetName)
    
    Dim siteNameCol As Long
    Dim brdStyleIndex As Long
    'site
    groupName = get_GroupName(mainSheetName, getTransBTSNameCol(mainSheetName))
    columnName = ThisWorkbook.Worksheets(mainSheetName).Cells(2, getTransBTSNameCol(mainSheetName)).value
    siteNameCol = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    'brd
    Call findBrdStyleGrpNameAndColName(mainSheetName, groupName, columnName)
    brdStyleIndex = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    
    Dim row As Long
    For row = 2 To mainSheet.range("b1048576").End(xlUp).row
         If mainSheet.Cells(row, siteNameCol).value = SITE_NAME Then
                findBoardStyleSheet = mainSheet.Cells(row, brdStyleIndex).value
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

Private Sub findBrdStyleGrpNameAndColName(sheetName As String, groupName As String, columnName As String)
    Dim colNum As Long
    
    colNum = getColNum(GetMainSheetName, 2, "BRDSTYLE", "BTS")
    groupName = get_GroupName(GetMainSheetName, colNum)
    columnName = get_ColumnName(GetMainSheetName, colNum)

End Sub

Private Function calculateRow(ByRef chkPassed As Boolean) As Long
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim rowCount As Long
    Dim brdCol As Long
    Dim brdPort As Long
    Dim cellsheet As Worksheet
    Dim index As Long
    Dim antenneIndex As Long
    Dim antenneGrpIdex As Long
    Dim sectorArray As Validation
    Dim trxNumIndex As Long
    Dim bchIndex As Long
    Dim tchIndex As Long
    Dim cellBandIndex As Long
    Dim cellBand As String
    Dim freqs As String
    Dim trxIndex As Long
            
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    '找小区ID所在的列
    Dim constCellTempCol As Long
    Call getCellMocNameAndAttrName(mocName, attrName)
    constCellTempCol = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
        
    Call getMainBcchAttrName(mocName, attrName)
    bchIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    Call getTchAttrName(mocName, attrName)
    tchIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
        
    '获取单板所在列
    Call getPhybrdMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    brdCol = findColumnByName(cellsheet, columnName, 2)
    '获取单板端口所在列
    Call getBrdPortMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    brdPort = findColumnByName(cellsheet, columnName, 2)
    '获取天线端口所在列
    Call getAntPorNoMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antenneIndex = findColumnByName(cellsheet, columnName, 2)
    
    Call getAntGrpIdMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antenneGrpIdex = findColumnByName(cellsheet, columnName, 2)
    '获取载频个数
    Call getTrxNumMocNameAndAttr(mocName, attrName)
    trxNumIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    '获取小区频段
    Call getCellFreqNameAndAttrName(mocName, attrName)
    cellBandIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    rowCount = 0
    If brdCol <= 0 Then
        calculateRow = rowCount
        Exit Function
    End If
    Dim trxNumStr As String
    
    Dim ant As CAntGBts
    For index = 3 To cellsheet.range("a1048576").End(xlUp).row
        If cellsheet.Cells(index, getGcellBTSNameCol(CELL_SHEET_NAME)).value = SITE_NAME Then
            Set ant = New CAntGBts
            trxNumStr = cellsheet.Cells(index, trxNumIndex).value
            cellBand = cellsheet.Cells(index, cellBandIndex).value
                        
            Dim trxNumArr() As String
            Dim elNo As Long
            Dim curCellTrxNum As Long
            curCellTrxNum = 0
            trxNumArr = Split(trxNumStr, ",")
            For elNo = LBound(trxNumArr) To UBound(trxNumArr)
                curCellTrxNum = curCellTrxNum + trxNumArr(elNo)
                trxIndex = elNo
            Next
            ant.trxNum = curCellTrxNum
            
            Dim tchFreqs As String
            tchFreqs = cellsheet.Cells(index, tchIndex).value
            
            If (Trim(tchFreqs) <> "") Then
                If Trim(cellsheet.Cells(index, bchIndex).value) <> "" Then
                    tchFreqs = "," + tchFreqs
                    freqs = cellsheet.Cells(index, bchIndex).value + tchFreqs
                Else
                    freqs = tchFreqs
                End If
            End If
            
            If trxIndex = 1 Then
                Call changeFreqs(freqs, trxNumStr, cellBand)
                ant.freqNos = freqs
            Else
                ant.freqNos = cellsheet.Cells(index, bchIndex).value + tchFreqs
            End If
            
            
            ant.cellId = cellsheet.Cells(index, constCellTempCol).value
            ant.brds = cellsheet.Cells(index, brdCol).value
            
            Dim portNo As String
            portNo = cellsheet.Cells(index, brdPort).value
            ant.portNos = replaceStr(portNo, "NA", "")
            
            Dim antNo As String
            antNo = cellsheet.Cells(index, antenneIndex).value
            ant.antNo = replaceStr(antNo, "NA", "")
            
            Dim antGrpId As String
            antGrpId = cellsheet.Cells(index, antenneGrpIdex).value
            ant.antGrpId = replaceStr(antGrpId, "NA", "")
            
            ant.ranges = cellsheet.Cells(index, brdCol).address(False, False)
            
            chkPassed = ant.commitData()
            If Not chkPassed Then
                Exit Function
            End If
            
            rowCount = rowCount + ant.getAntenneCollection.count
                        
            If Not Contains(valueMap, cellsheet.Cells(index, constCellTempCol).value) Then
                valueMap.Add Item:=ant, key:=cellsheet.Cells(index, constCellTempCol).value
            End If
        End If
    Next
    calculateRow = rowCount
End Function
Private Function replaceStr(str As String, oldStr As String, newStr As String) As String
    Dim strArray() As String
    strArray = Split(str, ",")
    replaceStr = ""
    Dim index As Long
    For index = LBound(strArray) To UBound(strArray)
        Dim number As String
        number = strArray(index)
        If number = oldStr Then
            number = newStr
        End If
        replaceStr = replaceStr + number + ","
    Next
    If Trim(replaceStr) <> "" Then
        replaceStr = Left(replaceStr, Len(replaceStr) - 1)
    End If
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

Private Function findColumnFromSingleRelationDef(sheetName As String, mocName As String, _
    attrName As String, srcSheetName As String) As String
    
    Dim relationSheet As Worksheet
    Dim rowCount As Long
    Dim index As Long
    
    Set relationSheet = ThisWorkbook.Worksheets(srcSheetName)
    rowCount = relationSheet.range("a1048576").End(xlUp).row
    For index = 2 To rowCount
        If relationSheet.Cells(index, 1).value = sheetName And relationSheet.Cells(index, 6).value = mocName And _
            relationSheet.Cells(index, 7).value = attrName Then
            findColumnFromSingleRelationDef = relationSheet.Cells(index, 3).value
            Exit Function
        End If
    Next
    
    findColumnFromSingleRelationDef = ""
End Function


Private Function findColumnFromRelationDef(sheetName As String, mocName As String, attrName As String) As String
    Dim tmpStr As String
    
    tmpStr = findColumnFromSingleRelationDef(sheetName, mocName, attrName, "RELATION DEF")
    If Not "" = tmpStr Then
        findColumnFromRelationDef = tmpStr
    Else
        findColumnFromRelationDef = findColumnFromSingleRelationDef(sheetName, mocName, attrName, "RELATION_EXT")
    End If
End Function

Private Sub getFreqMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        mocName = "TRXINFO"
        attrName = "BCCHFREQ"
    End If
End Sub

Private Sub getPhybrdMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        attrName = "BRDNO"
        mocName = "TRXBIND2PHYBRD"
    End If
End Sub


Private Sub getBrdPortMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        attrName = "TRXPN"
        mocName = "TRXBIND2PHYBRD"
    End If
End Sub

Private Sub getAntPorNoMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        attrName = "ANTPASSNO"
        mocName = "TRXBIND2PHYBRD"
    End If
End Sub
Private Sub getAntGrpIdMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        attrName = "ANTENNAGROUPID"
        mocName = "TRXBIND2PHYBRD"
    End If
End Sub

Private Sub getTrxNumMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        attrName = "TRXNUM"
        mocName = "TRXINFO"
    End If
End Sub

Private Function getMainBcchAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        attrName = "BCCHFREQ"
        mocName = "TRXINFO"
    End If
End Function

Private Function getTchAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        attrName = "NONBCCHFREQLIST"
        mocName = "TRXINFO"
    End If
End Function


Private Function getCellMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        attrName = "CELLNAME"
        mocName = "GCELL"
    End If
End Function

Private Function getCellFreqNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If CELL_TYPE = 21 Then
        attrName = "TYPE"
        mocName = "GCELL"
    End If
End Function
Private Function cellSheetType(sheetName As String) As Long
    If sheetName = "GSM Cell" Or sheetName = getResByKey("GSM Logic Cell") Then
        cellSheetType = 21
    Else
        cellSheetType = -1
    End If
End Function
Private Sub insertAntenneColumn()
    
    Dim tmpSheet As Worksheet
    Dim rowNum As Long
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim portIndex As String
    Dim antnIndex As Long
    Dim cellsheet As Worksheet
    
    portNoValueStr = "0,1,2,3,4,5,6,7"
    antnValueStr = "A,B,C,D,NULL,BY_ANTGRP"
    
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Call getBrdPortMocNameAndAttr(mocName, attrName)
    portIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    
    ThisWorkbook.ActiveSheet.Cells(1, 4).value = getResByKey("TRXBRDPASSNO")
    'cellSheet.Cells(2, portIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, 4)
    Call getAntPorNoMocNameAndAttr(mocName, attrName)
    antnIndex = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
    cellsheet.Cells(2, antnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, 5)
    ThisWorkbook.ActiveSheet.Cells(1, 5).clearComments
   
    
    Dim antnRang As range
    Set antnRang = ThisWorkbook.ActiveSheet.range("E2:E" + CStr(1 + ROW_COUNT))
            
    With antnRang.Validation
       .Delete
       .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=antnValueStr
    End With

    Dim portRang As range
    Set portRang = ThisWorkbook.ActiveSheet.range("D2:D" + CStr(1 + ROW_COUNT))
    With portRang.Validation
       .Delete
       .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=portNoValueStr
    End With

    
End Sub


Public Sub writeGbtsBackData()
    If WRITESUCCESS = True Then
        If chkCellInfoMap() Then
            Exit Sub
        End If
        
        Call writeCellData
    End If
    WRITESUCCESS = False
    Call delTmpSht
End Sub

Public Sub delTmpSht()
    Dim tmpSheet As Worksheet
    Dim cellsheet As Worksheet
    InAdjustAntnPort = False
    If CELL_SHEET_NAME <> "" Then
        Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
        cellsheet.Activate
    End If
    Set tmpSheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    Call changeAlerts(False)
    tmpSheet.Delete
    Call changeAlerts(True)
End Sub


Private Function chkCellInfoMap() As Boolean


    Dim tmpSheet As Worksheet
    Dim maxRow As Long
    Dim rowNum As Long
    Dim cellInfoMap As CMapValueObject
    Dim tmpMap As CMap
    Dim cellId As String
    Dim freqNo As String
    Dim board As String
    Dim portNo As String
    Dim antNo As String
    
    'Dim boardInfo As String
    Dim tVal As CMap
    Dim sVal As String
    Dim isExist As Boolean
    Dim celldes As String
    Dim sectordes As String
    Dim antndes As String
    
    
    
    
    Set tmpSheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))

    
    maxRow = tmpSheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
        cellId = tmpSheet.Cells(rowNum, 1).value
        freqNo = tmpSheet.Cells(rowNum, 2).value
        board = tmpSheet.Cells(rowNum, 3).value
        portNo = tmpSheet.Cells(rowNum, 4).value
        antNo = tmpSheet.Cells(rowNum, 5).value
        'boardInfo = board & "_" & portNo
        
        
        
        If checkInputData(tmpSheet, cellId, freqNo, board, portNo, antNo, rowNum) Then
            chkCellInfoMap = True
            Exit Function
        End If

    Next
    chkCellInfoMap = False
    
End Function

Private Function checkInputData(sheet As Worksheet, ByRef Cell As String, ByRef freqNo As String, _
    ByRef board As String, ByRef portNo As String, ByRef antNo As String, lineNo As Long) As Boolean
        Dim rangeStr As String
        If Trim(Cell) = "" Or Trim(freqNo) = "" Or Trim(board) = "" Then 'Or Trim(portNo) = "" Or Trim(antNo) = "" Then
            sheet.Rows(lineNo).Select
            MsgBox getResByKey("emptyCellExists") & lineNo, vbExclamation, getResByKey("Warning")
            checkInputData = True
            Exit Function
        End If
        
        If checkDataValid(cellValueStr, Cell, False) = False Then
            rangeStr = sheet.Cells(lineNo, 1).address(False, False)
            Call MsgBox(getResByKey("GcellDataWrong") & rangeStr, vbInformation, getResByKey("Warning"))
            sheet.range(rangeStr).Select
            checkInputData = True
            Exit Function
        End If
    
        If isAInteger(freqNo) = False Then
            rangeStr = sheet.Cells(lineNo, 2).address(False, False)
            Call MsgBox(getResByKey("freqDataErrShouldBeInteger") & rangeStr, vbInformation, getResByKey("Warning"))
            sheet.range(rangeStr).Select
            checkInputData = True
            Exit Function
        End If
        
        If checkDataValid(boardValueStr, board, False) = False Then
            rangeStr = sheet.Cells(lineNo, 3).address(False, False)
            Call MsgBox(getResByKey("boardDataWrong") & rangeStr, vbInformation, getResByKey("Warning"))
            sheet.range(rangeStr).Select
            checkInputData = True
            Exit Function
        End If
        
        If checkDataValid(portNoValueStr, portNo, True) = False Then
            rangeStr = sheet.Cells(lineNo, 4).address(False, False)
            Call MsgBox(getResByKey("portNoDataDataWrong") & rangeStr, vbInformation, getResByKey("Warning"))
            sheet.range(rangeStr).Select
            checkInputData = True
            Exit Function
        End If
            
        If checkDataValid(antnValueStr, antNo, True) = False Then
            rangeStr = sheet.Cells(lineNo, 5).address(False, False)
            Call MsgBox(getResByKey("antnPassNoDataDataWrong") & rangeStr, vbInformation, getResByKey("Warning"))
            sheet.range(rangeStr).Select
            checkInputData = True
            Exit Function
        End If
        checkInputData = False
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

Private Function existAntnPort(ByRef antnStr As String, ByRef tVal As String) As Boolean
    Dim antenneArray() As String
    Dim index As Long
    antenneArray = Split(antnStr, ",")
    For index = LBound(antenneArray) To UBound(antenneArray)
        If antenneArray(index) = tVal Then
            existAntnPort = True
            Exit Function
        End If
    Next
    existAntnPort = False
End Function



Private Function shrinkStr(inputStr As String, deliStr As String) As String
        Dim fmtStr As String
        Dim appendStr As String
        
        fmtStr = inputStr
        If CON_BRACKET_RIGHT = deliStr Then
            appendStr = CON_BRACKET_RIGHT
            If Right(inputStr, 1) = CON_BRACKET_RIGHT Then
                fmtStr = Left(inputStr, Len(inputStr) - 1)
            End If
        End If
        
        
        Dim strArray() As String
        strArray = Split(fmtStr, deliStr)

        Dim iNo As Long
        Dim tmpStr As String
        
        For iNo = LBound(strArray) To UBound(strArray)
            If (0 = iNo) Then
                tmpStr = strArray(iNo)
            Else
               If Not (tmpStr = strArray(iNo)) Then
                    shrinkStr = inputStr
                    Exit Function
               End If
            End If
        Next
        
        shrinkStr = tmpStr + appendStr
End Function

Private Function filterEmptyQuoatData(value As String)
    If "[]" = value Then
        filterEmptyQuoatData = ""
    Else
        filterEmptyQuoatData = value
    End If
End Function

Private Sub writeCellData()
    Dim cellsheet As Worksheet
    Dim maxRow As Long
    Dim rowNum As Long
    Dim baseStationName As String
    Dim cellId As String
    Dim keyVal As Variant
    Dim tmpVal As Variant
    
    Dim tVal As CMap
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim brdCol As Long
    Dim brdPort As Long
    Dim constCellTempCol As Long
    Dim antenneIndex As Long
    Dim antenneGrpIdIndex As Long
    
    '找小区ID所在的列
    Set cellsheet = ThisWorkbook.Worksheets(CELL_SHEET_NAME)
    Call getCellMocNameAndAttrName(mocName, attrName)
    constCellTempCol = getColNum(CELL_SHEET_NAME, 2, attrName, mocName)
     
    '获取单板所在列
    Call getPhybrdMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    brdCol = findColumnByName(cellsheet, columnName, 2)
    '获取单板端口所在列
    Call getBrdPortMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    brdPort = findColumnByName(cellsheet, columnName, 2)
    '获取天线端口所在列
    Call getAntPorNoMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antenneIndex = findColumnByName(cellsheet, columnName, 2)
    
    Call getAntGrpIdMocNameAndAttr(mocName, attrName)
    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
    antenneGrpIdIndex = findColumnByName(cellsheet, columnName, 2)
    
    maxRow = cellsheet.range("a1048576").End(xlUp).row
    
    Dim udCollection As Collection
    Call prepareUserDataCollection(udCollection)
    
    For rowNum = 3 To maxRow
            baseStationName = cellsheet.Cells(rowNum, getGcellBTSNameCol(CELL_SHEET_NAME)).value
            cellId = cellsheet.Cells(rowNum, constCellTempCol).value
            If baseStationName = SITE_NAME Then
                Dim brdNoStr As String
                Dim portNoStr As String
                Dim antNoStr As String
                Dim antGrpIdStr As String
                brdNoStr = ""
                portNoStr = ""
                antNoStr = ""
                antGrpIdStr = ""
                Call getAntCfg(udCollection, cellId, brdNoStr, portNoStr, antNoStr, antGrpIdStr)
                cellsheet.Cells(rowNum, brdCol).value = filterEmptyQuoatData(brdNoStr)
                cellsheet.Cells(rowNum, brdPort).value = filterEmptyQuoatData(portNoStr)
                cellsheet.Cells(rowNum, antenneIndex).value = filterEmptyQuoatData(antNoStr)
                cellsheet.Cells(rowNum, antenneGrpIdIndex).value = filterEmptyQuoatData(antGrpIdStr)
            End If
    Next
End Sub

Function conStrValue(ByRef srcStr As String, ByRef curStr As String, conStr As String)
    If ("" = srcStr) Then
        conStrValue = curStr
    Else
        conStrValue = srcStr + conStr + curStr
    End If
End Function

Private Function quoatStr(srcStr As String, rruFlag As Boolean) As String
    If rruFlag Then
        quoatStr = CON_BRACKET_LEFT + srcStr + CON_BRACKET_RIGHT
    Else
        quoatStr = srcStr
    End If
End Function

Private Sub getAntCfg(ByRef udCollection As Collection, ByRef cellId As String, ByRef brdNoStr As String, ByRef portNoStr As String, ByRef antNoStr As String, ByRef antGrpIdStr As String)

    Dim tmpSheet As Worksheet
    Dim maxRow As Long
    Dim rowNum As Long
    Dim tmpCollection As Collection
    Set tmpCollection = New Collection
    
    Dim rruFlag As Boolean
    rruFlag = False
       
    
    Set tmpSheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    If Contains(udCollection, cellId) Then
        '每一个载频
        Dim freq As Collection
        For Each freq In udCollection(cellId)
            '每一个绑定关系，一个载频可以有多个绑定关系
            'RRU共小区场景一个载频的绑定关系使用逗号分隔，中括号括起来
            
            Dim freqBind As CGBtsTrxBind
            If freq.count > 1 Then
                rruFlag = True
                
                Dim newBind As CGBtsTrxBind
                Set newBind = New CGBtsTrxBind
                For Each freqBind In freq
                    newBind.cellId = freqBind.cellId
                    newBind.freqNo = freqBind.freqNo
                    
                    newBind.brdNo = conStrValue(newBind.brdNo, freqBind.brdNo, CON_COMMA)
                    newBind.portNo = conStrValue(newBind.portNo, freqBind.portNo, CON_COMMA)
                    newBind.antNo = conStrValue(newBind.antNo, freqBind.antNo, CON_COMMA)
                    newBind.antGrpId = conStrValue(newBind.antGrpId, freqBind.antGrpId, CON_COMMA)
                    
                Next
                                               
                tmpCollection.Add Item:=newBind
            ElseIf (freq.count = 1) Then
                tmpCollection.Add Item:=freq(1)
            End If
        
        Next
    End If
    
    Dim conStr As String
    Dim deliStr As String
    If rruFlag Then
        conStr = ""
        deliStr = CON_BRACKET_RIGHT
    Else
        conStr = CON_COMMA
        deliStr = CON_COMMA
    End If
    
    Dim bind As CGBtsTrxBind
    For Each bind In tmpCollection
        brdNoStr = conStrValue(brdNoStr, quoatStr(bind.brdNo, rruFlag), conStr)
        portNoStr = conStrValue(portNoStr, quoatStr(bind.portNo, rruFlag), conStr)
        antNoStr = conStrValue(antNoStr, quoatStr(bind.antNo, rruFlag), conStr)
        antGrpIdStr = conStrValue(antGrpIdStr, quoatStr(bind.antGrpId, rruFlag), conStr)
    Next
    
    brdNoStr = shrinkStr(brdNoStr, deliStr)
    portNoStr = shrinkStr(portNoStr, deliStr)
    antNoStr = shrinkStr(antNoStr, deliStr)
    antGrpIdStr = shrinkStr(antGrpIdStr, deliStr)
    
    brdNoStr = replaceStr(brdNoStr, "", "NA")
    portNoStr = replaceStr(portNoStr, "", "NA")
    antNoStr = replaceStr(antNoStr, "", "NA")
    antGrpIdStr = replaceStr(antGrpIdStr, "", "NA")
    
    If rruFlag Then
        Call shrink(brdNoStr, deliStr)
        Call shrink(portNoStr, deliStr)
        Call shrink(antNoStr, deliStr)
        Call shrink(antGrpIdStr, deliStr)
    End If
End Sub
Private Function shrink(ByRef inputStr As String, deliStr As String)
    Dim strArray() As String
    strArray = Split(inputStr, deliStr)
    inputStr = ""
    Dim tmpStr As String
    tmpStr = ""
    
    Dim iNo As Long
    For iNo = LBound(strArray) To UBound(strArray)
        If strArray(iNo) = "" Then
            Exit Function
        End If
        
        tmpStr = Right(strArray(iNo), Len(strArray(iNo)) - 1)
        tmpStr = shrinkStr(tmpStr, CON_COMMA)
        inputStr = inputStr + CON_BRACKET_LEFT + tmpStr + CON_BRACKET_RIGHT
    Next
End Function



'用户填写的数据保存在数据结构中collection(小区名称，collection（小区名称+频点，collection(绑定关系)）)
Private Sub prepareUserDataCollection(ByRef udCollection As Collection)

    Dim tmpSheet As Worksheet
    Dim maxRow As Long
    Dim rowNum As Long
    Set udCollection = New Collection
    
    Set tmpSheet = ThisWorkbook.Worksheets(getResByKey("Temp Sheet"))
    maxRow = tmpSheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
        Dim cellId As String
        Dim freqNo As String
        Dim brdNoStr As String
        Dim portNoStr As String
        Dim antNoStr As String
        Dim antGrpIdStr As String
                
        cellId = Trim(tmpSheet.Cells(rowNum, 1).value)
        freqNo = Trim(tmpSheet.Cells(rowNum, 2).value)
        
        If ("" = Trim(tmpSheet.Cells(rowNum, 3).value)) Then
            brdNoStr = "NA"
        Else
            brdNoStr = Trim(tmpSheet.Cells(rowNum, 3).value)
        End If
        
        If ("" = Trim(tmpSheet.Cells(rowNum, 4).value)) Then
            portNoStr = "NA"
        Else
            portNoStr = Trim(tmpSheet.Cells(rowNum, 4).value)
        End If

        If ("" = Trim(tmpSheet.Cells(rowNum, 5).value)) Then
            antNoStr = "NA"
        Else
            antNoStr = Trim(tmpSheet.Cells(rowNum, 5).value)
        End If

        If ("" = Trim(tmpSheet.Cells(rowNum, 6).value)) Then
            antGrpIdStr = "NA"
        Else
            antGrpIdStr = Trim(tmpSheet.Cells(rowNum, 6).value)
        End If
        
        Dim comKey As String
        comKey = cellId + CON_SHARP + freqNo
        
        Dim bind As CGBtsTrxBind
        Set bind = New CGBtsTrxBind
        bind.cellId = cellId
        bind.freqNo = freqNo
        bind.brdNo = brdNoStr
        bind.portNo = portNoStr
        bind.antNo = antNoStr
        bind.antGrpId = antGrpIdStr
        
        If Contains(udCollection, cellId) Then
            Dim freqCollection As Collection
            Set freqCollection = udCollection(cellId)
            If Contains(freqCollection, comKey) Then
                freqCollection(comKey).Add Item:=bind
            Else
                Dim bdCollection As Collection
                Set bdCollection = New Collection
                bdCollection.Add Item:=bind
                
                freqCollection.Add Item:=bdCollection, key:=comKey
            End If
            
        Else
            Dim bindCollection As Collection
            Set bindCollection = New Collection
            bindCollection.Add Item:=bind
            
            Dim frqCollection As Collection
            Set frqCollection = New Collection
            frqCollection.Add Item:=bindCollection, key:=comKey
            
            udCollection.Add Item:=frqCollection, key:=cellId
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

Private Function getColNameBaseMapDef(sheetName As String, attrName As String, mocName As String) As String
    On Error Resume Next
    Dim m_colNum As Long
    Dim m_rowNum As Long
    Dim mappingDef As Worksheet
    Dim ws As Worksheet
    Dim localCurColName As String
        
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    getColNameBaseMapDef = ""
    For m_rowNum = 2 To mappingDef.range("a1048576").End(xlUp).row
        If UCase(attrName) = UCase(mappingDef.Cells(m_rowNum, 5).value) _
           And UCase(sheetName) = UCase(mappingDef.Cells(m_rowNum, 1).value) _
           And UCase(mocName) = UCase(mappingDef.Cells(m_rowNum, 4).value) Then
            getColNameBaseMapDef = mappingDef.Cells(m_rowNum, 3).value
            
            Exit For
        End If
    Next
End Function

Private Function getColNumByName(sheetName As String, recordRow As Long, ColName As String) As Long
    On Error Resume Next
    Dim m_colNum As Long
    Dim ws As Worksheet
    Dim localCurColName As String
        
    getColNumByName = -1
    
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    m_colNum = 1
    Do
        localCurColName = GetDesStr(ws.Cells(recordRow, m_colNum).value)
        If "" = localCurColName Then
            Exit Function
        End If
        
        If GetDesStr(ColName) = localCurColName Then
            getColNumByName = m_colNum
            Exit Function
        End If
        
        m_colNum = m_colNum + 1
    Loop

End Function

Public Sub BranchControlForTempSht(ByVal sheet As Object, ByVal Target As range)
    Dim cellRange As range, freqRange As range
    Dim antNoStr As String
    
    For Each cellRange In Target
        If cellRange.column = 5 Then
            For Each freqRange In cellRange
                antNoStr = Trim(freqRange.value)
                If cellRange.row > 1 And antNoStr <> "BY_ANTGRP" Then
                    sheet.Cells(cellRange.row, 6).Interior.colorIndex = SolidColorIdx
                    sheet.Cells(cellRange.row, 6).Interior.Pattern = SolidPattern
                    sheet.Cells(cellRange.row, 6).value = ""
                    sheet.Cells(cellRange.row, 6).Validation.ShowInput = False
                
                ElseIf cellRange.row > 1 And antNoStr = "BY_ANTGRP" Then
                    sheet.Cells(cellRange.row, 6).Interior.colorIndex = NormalRangeColorIndex
                    sheet.Cells(cellRange.row, 6).Interior.Pattern = NormalPattern
                    sheet.Cells(cellRange.row, 6).value = ""
                    sheet.Cells(cellRange.row, 6).Validation.ShowInput = True
                End If
            Next
        ElseIf cellRange.column = 6 Then
            For Each freqRange In cellRange
                If freqRange.value <> "" And freqRange.Interior.colorIndex = SolidColorIdx And freqRange.Interior.Pattern = SolidPattern Then
                    MsgBox getResByKey("NoInput"), vbOKOnly + vbExclamation + vbApplicationModal + vbDefaultButton1, getResByKey("Warning")
                    freqRange.value = ""
                    freqRange.Select
                End If
            Next
        End If
    
    Next cellRange
    
End Sub






