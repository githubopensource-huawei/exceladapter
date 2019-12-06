Attribute VB_Name = "AddListRefSub"
Option Explicit
'Private addType as long ' 1-"LIST" 2-"PATTERN" 3-"First NodeB"
Private refreshing As Boolean
Function isRefreshing() As Boolean
        isRefreshing = refreshing
End Function

Function refreshStart()
        refreshing = True
End Function

Function refreshEnd()
         refreshing = False
End Function

Sub popUpIlleagalSelectionErrorMsgbox()
    Dim errorMsg As String
    errorMsg = getResByKey("AddRowSelectionErrorMsg")
    Call MsgBox(errorMsg, vbExclamation, getResByKey("IllegalSelection"))
End Sub

Function checkAddIubRowSelectionRange(ByRef selectionRanges As range, ByRef startRow As Long, ByRef endRow As Long, ByRef mocName As String, ByRef sheetDefIndex As Long) As Boolean
    checkAddIubRowSelectionRange = True
    startRow = 0
    endRow = 0
    
    startRow = selectionRanges(1).row
    
    Dim startRowInSheetDefSheet As Long, endRowInSheetDefSheet As Long
    Dim selectionRange As range
    Dim numberOfCellsNotEmpty As Long
    numberOfCellsNotEmpty = 0
    For Each selectionRange In selectionRanges
        If selectionRange.column <> 1 Or numberOfCellsNotEmpty > 1 Then
            Call popUpIlleagalSelectionErrorMsgbox
            checkAddIubRowSelectionRange = False
            Exit Function
        Else
            endRow = selectionRange.row
            If selectionRange.value <> "" Then
                mocName = selectionRange.value
                numberOfCellsNotEmpty = numberOfCellsNotEmpty + 1
            End If
        End If
    Next selectionRange
    sheetDefIndex = getSheetDefIndex(mocName, startRowInSheetDefSheet, endRowInSheetDefSheet)
    If numberOfCellsNotEmpty = 0 Or startRow <> startRowInSheetDefSheet Or endRow > endRowInSheetDefSheet Then
        Call popUpIlleagalSelectionErrorMsgbox
        checkAddIubRowSelectionRange = False
        Exit Function
    End If
    endRow = endRowInSheetDefSheet
End Function

Sub addIubRowInAllIubSheets(ByVal startRow As Long, ByVal newEndRow As Long)
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        If isIubStyleWorkSheetByParameterWs(sheet) Then
            Call addIubRowInOneSheet(sheet, startRow, newEndRow)
        End If
    Next sheet
End Sub

Function getSheetDefIndex(ByRef sheetName As String, ByRef startRow As Long, ByRef endRow As Long) As Long
    Dim sheetDef As Worksheet
    getSheetDefIndex = -1
    startRow = -1
    endRow = -1
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim cellValue As String
    Dim k As Long
    'For k = 2 To sheetDef.UsedRange.rows.count
    For k = 2 To getUsedRowCount(sheetDef)
        cellValue = sheetDef.Cells(k, innerPositionMgr.sheetDef_shtNameColNo).value
        If cellValue = sheetName Then
            getSheetDefIndex = k
            startRow = CLng(sheetDef.Cells(k, innerPositionMgr.sheetDef_startRowColNo).value)
            endRow = CLng(sheetDef.Cells(k, innerPositionMgr.sheetDef_endRowColNo).value)
            Exit Function
        End If
    Next k
End Function

Sub setBorders(ByRef certainRange As range)
    On Error Resume Next
    certainRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    certainRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    certainRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    certainRange.Borders(xlEdgeRight).LineStyle = xlContinuous
    certainRange.Borders(xlInsideVertical).LineStyle = xlContinuous
    certainRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Sub

Sub setBackGroundColorAndClearContenets(ByRef certainRange As range)
    certainRange.Interior.colorIndex = -4142
    certainRange.ClearContents
End Sub

Sub addIubRowInOneSheet(ByRef sheet As Worksheet, ByVal startRow As Long, ByVal newEndRow As Long)
    sheet.rows(newEndRow - 1).Copy
    sheet.rows(newEndRow).INSERT Shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove
    sheet.range("A" & startRow & ":A" & newEndRow).Merge
    Dim columnNumber As Long
    columnNumber = sheet.range("A" & startRow).End(xlToRight).column
    Dim newRowRange As range
    Set newRowRange = sheet.range(sheet.Cells(newEndRow, 2), sheet.Cells(newEndRow, columnNumber))
    Call setBorders(newRowRange)
    Call setBackGroundColorAndClearContenets(newRowRange)
End Sub

Function calculateColumnName(ByRef columnNumber As Long) As String
    Dim myRange As range
    Set myRange = Cells(1, columnNumber)    '指定该列标号的任意单元格
    calculateColumnName = Left(myRange.range("A1").address(True, False), _
        InStr(1, myRange.range("A1").address(True, False), "$", 1) - 1)
    Set myRange = Nothing
End Function

Function updateNewRefCellArray(ByRef refCellArray() As String, ByVal endRow As Long) As Boolean
    updateNewRefCellArray = False
    If UBound(refCellArray) = -1 Then Exit Function
    Dim refCell As String
    Dim refCellColumnLetter As String
    Dim refCellColumnNumber As Long, refCellRowNumber As Long
    Dim k As Long
    For k = 0 To UBound(refCellArray)
        refCell = refCellArray(k)
        refCellColumnNumber = range(refCell).column
        refCellColumnLetter = calculateColumnName(refCellColumnNumber)
        refCellRowNumber = range(refCell).row
        If refCellRowNumber > endRow Then
            refCellArray(k) = refCellColumnLetter & (refCellRowNumber + 1)
            updateNewRefCellArray = True 'there are ref updated
        End If
    Next k
End Function

Function getNewRefCellValue(ByRef refCellArray() As String) As String
    Dim newRefCellValue As String
    newRefCellValue = ""
    Dim k As Long
    For k = 0 To UBound(refCellArray)
        newRefCellValue = newRefCellValue + refCellArray(k) + ","
    Next k
    getNewRefCellValue = Left(newRefCellValue, Len(newRefCellValue) - 1) 'erase ,
End Function

Sub updateRefStringInBaseStationSheet(ByVal endRow As Long)
    Dim refCellArray() As String
    Dim sheetBaseStation As Worksheet
    Set sheetBaseStation = ThisWorkbook.Worksheets(GetMainSheetName)
    Dim columnNumber As Long
    Dim refCellValue As String, newRefCellValue As String
    'For columnNumber = 1 To sheetBaseStation.UsedRange.columns.count
    For columnNumber = 1 To getUsedColumnCount(sheetBaseStation)
        refCellValue = Trim(sheetBaseStation.Cells(3, columnNumber).value)
        refCellArray() = Split(refCellValue, ",")
        If updateNewRefCellArray(refCellArray(), endRow) Then
            newRefCellValue = getNewRefCellValue(refCellArray())
            sheetBaseStation.Cells(3, columnNumber).value = newRefCellValue
        End If
    Next columnNumber
End Sub

Sub changeAlert(ByRef flag As Boolean)
    Application.DisplayAlerts = flag
    Application.ScreenUpdating = flag
End Sub

Sub addIubRow()
    Call changeAlert(False)
    Dim selectionRanges As range
    Dim startRow As Long
    Dim endRow As Long
    Dim sheetDefIndex As Long
    Dim selectMocName As String
    Dim sheetToAddIubRow As Worksheet
    Set sheetToAddIubRow = ThisWorkbook.ActiveSheet
    Set selectionRanges = Selection
    
    If checkAddIubRowSelectionRange(selectionRanges, startRow, endRow, selectMocName, sheetDefIndex) = False Then Exit Sub
    
    Call addIubRowInAllIubSheets(startRow, endRow + 1)

    Call changeSheetDefRowPos(1, sheetDefIndex)
    
    Call updateRefStringInBaseStationSheet(endRow)
    Call changeAlert(True)
End Sub

Private Sub changeSheetDefRowPos(count As Long, startIndex As Long)
    Dim sheetDef As Worksheet
    Dim iCount As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim index As Long
    Set sheetDef = ThisWorkbook.Sheets("SHEET DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    iCount = getUsedRowCount(sheetDef, 1)
    For index = startIndex To iCount
        If sheetDef.Cells(index, innerPositionMgr.sheetDef_startRowColNo).value <> "" And sheetDef.Cells(index, innerPositionMgr.sheetDef_endRowColNo).value <> "" Then
            startRow = CLng(sheetDef.Cells(index, innerPositionMgr.sheetDef_startRowColNo).value)
            endRow = CLng(sheetDef.Cells(index, innerPositionMgr.sheetDef_endRowColNo).value)
            If index <> startIndex Then
                sheetDef.Cells(index, innerPositionMgr.sheetDef_startRowColNo).value = CStr(startRow + count)
            End If
            sheetDef.Cells(index, innerPositionMgr.sheetDef_endRowColNo).value = CStr(endRow + count)
        End If
    Next
End Sub

Sub iubStyleSheetSelectionChange(sheet As Worksheet, target As range)
    Dim groupName As String
    Dim columnName As String
    Dim sheetName As String
    Dim controlDef As CControlDef
    Dim m_Str As String
    If target.count > 1 Then
        Exit Sub
    End If
    Call getGroupNameShNameAndAttrName(sheet, target, groupName, sheetName, columnName)
    Set controlDef = getControlDefine(sheetName, groupName, columnName)
    If controlDef Is Nothing Then
        Exit Sub
    End If
    Dim mocName As String, attrName As String
    mocName = controlDef.mocName
    attrName = controlDef.attributeName
    
    m_Str = controlDef.lstValue
    If Len(m_Str) > 255 Or isValidDef(mocName + "," + attrName) Then
        Dim valideDef As CValideDef
        Set valideDef = initDefaultDataSub.getInnerValideDef(sheetName + "," + groupName + "," + columnName)
        If valideDef Is Nothing Then
            Set valideDef = addInnerValideDef(sheetName, groupName, columnName, m_Str)
        End If
        m_Str = valideDef.getValidedef
    End If
    
    If Not controlDef Is Nothing Then
        On Error Resume Next
        If UCase(controlDef.dataType) = "ENUM" And controlDef.lstValue <> "" Then
            If target.Validation Is Nothing Then
                With target.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str
                End With
            Else
                If Not targetHasFormula1(target) Then
                    With target.Validation
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str
                        .Modify Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str
                    End With
                End If
            End If
        End If
    End If
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
    If formula1 = "" Then targetHasFormula1 = False
    Exit Function
ErrorHandler:
    targetHasFormula1 = False
End Function

Sub iubStyleBaseStationCheckRang(sheet As Worksheet, target As range)
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
    
    Dim alreadyCheckFlag As Boolean
    alreadyCheckFlag = False
    
    Dim controlDef As CControlDef
    If target.value = "" Then
        Exit Sub
    End If
    If isIubStyleWorkBook() Then
        groupName = get_GroupName(sheet.name, target.column)
        columnName = get_ColumnName(sheet.name, target.column)
        Set mappingDef = getMappingDefine(sheet.name, groupName, columnName)
        If Not mappingDef Is Nothing Then
           If mappingDef.mocName <> "" And mappingDef.attributeName <> "" Then
                Exit Sub
           End If
        End If
        
        If sheet.Cells(3, target.column).value = "" Then
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
        If Not controlDef Is Nothing Then
            Call Check_Value_In_Range(controlDef.dataType, controlDef.bound + controlDef.lstValue, target.value, target, alreadyCheckFlag)
        End If
    End If
End Sub

Public Function getLongValidatastr(ByRef mocSheetName As String, ByRef mocGroupName As String, ByRef mocColumnName As String) As String
    On Error GoTo ErrorHandler
    Dim mappingDef As CMappingDef
     
    Set mappingDef = getMappingDefine(mocSheetName, mocGroupName, mocColumnName)
    Dim currmocName As String, currattrname As String
    currmocName = mappingDef.mocName
    currattrname = mappingDef.attributeName
    If VBA.Trim(currmocName) = "" Or VBA.Trim(currattrname) = "" Then
        getLongValidatastr = ""
        Exit Function
    End If
    Dim laststr As String
    
    Dim validSheet As Worksheet
    Set validSheet = ThisWorkbook.Sheets("VALID DEF")
    Dim index As Long
    For index = 2 To validSheet.range("a65536").End(xlUp).row
        If GetDesStr(validSheet.Cells(index, 1).value) = GetDesStr(currmocName) And GetDesStr(validSheet.Cells(index, 2).value) = GetDesStr(currattrname) Then
            Dim maxColNum As Long
            maxColNum = validSheet.range("IV" + CStr(index)).End(xlToLeft).column
            getLongValidatastr = "=INDIRECT(""'VALID DEF'!C" & CStr(index) & ":" & getColStr(maxColNum) & CStr(index) & """)"
        End If
    Next
    Exit Function
ErrorHandler:
    getLongValidatastr = ""
End Function

Public Sub getGroupNameShNameAndAttrName(sheet As Worksheet, ByRef selectionRange As range, ByRef groupName As String, ByRef sheetName As String, ByRef attributeName As String)
    Dim k As Long
    Dim rangeValue As String
    For k = selectionRange.row To 1 Step -1
        rangeValue = sheet.Cells(k, 1).value
        If rangeValue <> "" Then
            sheetName = rangeValue
            attributeName = sheet.Cells(k, selectionRange.column).value
            Exit For
        End If
    Next k
    If k = selectionRange.row Then
        groupName = ""
        sheetName = ""
        attributeName = ""
        Exit Sub
    End If
    groupName = getGroupNameFromMappingDef(sheetName, attributeName)
End Sub

Private Function getGroupNameFromMappingDef(sheetName As String, attributeName As String) As String
    Dim mappingDef As Worksheet
    Dim index, count As Long
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    'count = mappingDef.UsedRange.rows.count
    count = getUsedRowCount(mappingDef, 1)
    For index = 1 To count
        If mappingDef.Cells(index, 1).value = sheetName And mappingDef.Cells(index, 3).value = attributeName Then
            getGroupNameFromMappingDef = mappingDef.Cells(index, 2)
            Exit Function
        End If
    Next
    getGroupNameFromMappingDef = ""
End Function
Sub destroyMenuStatus()
    With Application
        .CommandBars("Row").reset
        .CommandBars("Column").reset
        .CommandBars("Cell").reset
    End With
End Sub
Sub insertAndDeleteControl(ByRef flag As Boolean)
    On Error Resume Next
    With Application
        With .CommandBars("Row")
            .FindControl(ID:=3183).Enabled = flag '插入
            .FindControl(ID:=296).Enabled = flag '行
            .FindControl(ID:=293).Enabled = flag '删除
        End With
        With .CommandBars("Column")
            .FindControl(ID:=3183).Enabled = flag '插入
            .FindControl(ID:=297).Enabled = flag '行
            .FindControl(ID:=294).Enabled = flag '删除
        End With
        With .CommandBars("Cell")
            .FindControl(ID:=3181).Enabled = flag '插入
            .FindControl(ID:=295).Enabled = flag '行
            .FindControl(ID:=292).Enabled = flag '删除
        End With
    End With
End Sub
Sub initMenuStatus(sh As Worksheet)
    If isIubStyleWorkSheet(sh.name) Then
        Call insertAndDeleteControl(False)
    Else
        Call insertAndDeleteControl(True)
    End If
End Sub

Sub importRef()
    Dim templateFile
    Dim wb As Workbook
    Dim templateSheetDef As Worksheet
    Dim mainSheetName As String
    Dim templateMainSheet As Worksheet
    Dim address As String
    Dim iRange As range
    Dim addressIndex As Long
    Dim groupName As String
    Dim app As Application
    Set app = CreateObject("Excel.Application")
    
    Application.ScreenUpdating = False
    templateFile = Application.GetOpenFilename("Microsoft Excel(*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm")
    mainSheetName = GetMainSheetName()
    If templateFile <> "False" Then
        If templateFile = ThisWorkbook.FullName Then
            Exit Sub
        End If
        'Set wb = Application.Workbooks.Open(templateFile, , True)
        Set wb = app.Workbooks.Open(templateFile, , True)
        ThisWorkbook.Activate
        If checkIubStyle(wb) = True Then
            Dim columns As Long
            Dim index As Long
            Dim arrs() As String
            Set templateMainSheet = wb.Worksheets(mainSheetName)
            Set templateSheetDef = wb.Worksheets("SHEET DEF")
            columns = templateMainSheet.UsedRange.columns.count
            For index = 1 To columns
                If templateMainSheet.Cells(3, index).value <> "" Then
                    arrs = Split(templateMainSheet.Cells(3, index).value, ",")
                    For addressIndex = 0 To UBound(arrs)
                        address = arrs(addressIndex)
                        Set iRange = findMappingRow(wb, templateSheetDef, address)
                        groupName = getGroupNameFromTemplateFile(wb, mainSheetName, index)
                        If Not iRange Is Nothing Then
                            Call addIubStyleCellRef(mainSheetName, groupName, templateMainSheet.Cells(2, index).value, iRange)
                        End If
                    Next
                End If
            Next
        End If
        Application.DisplayAlerts = False
        app.DisplayAlerts = False
        wb.Close (False)
        Application.DisplayAlerts = True
        app.DisplayAlerts = True
        app.Quit
        Set app = Nothing
    End If
    Application.ScreenUpdating = True
End Sub

Private Function getGroupNameFromTemplateFile(wb As Workbook, sheetName, colnum) As String
        Dim index As Long
        Dim ws As Worksheet
        Set ws = wb.Worksheets(sheetName)
        For index = colnum To 1 Step -1
            If Not isEmpty(ws.Cells(1, index).value) And Len(Trim(ws.Cells(1, index).value)) > 0 Then
                getGroupNameFromTemplateFile = ws.Cells(1, index).value
                Exit Function
            End If
        Next
        getGroupNameFromTemplateFile = ""
End Function

Function whetherWbHasCertainSheet(wb As Workbook, sheetName As String) As Boolean
    Dim sheet As Worksheet
    whetherWbHasCertainSheet = False
    For Each sheet In wb.Worksheets
        If sheet.name = sheetName Then
            whetherWbHasCertainSheet = True
            Exit Function
        End If
    Next sheet
End Function

Sub getCoverStyleMap(ws As Worksheet, styleMap As CMap, issueDate As String)
    Dim k As Long
    Dim columnBVal As String, columnDVal As String
    Call styleMap.SetAt("wbName", ws.Cells(2, 2).value)
    For k = 3 To 10 'find Issue Date
        columnBVal = ws.Cells(k, 2).value
        columnDVal = ws.Cells(k, 4).value
        If columnBVal <> issueDate Then
            Call styleMap.SetAt(columnBVal, columnDVal)
        Else
            Exit Sub
        End If
    Next
End Sub

Function checkTwoWbsStyle(dstWbCoverStyleMap As CMap, srcWbCoverStyleMap As CMap) As Boolean
    Dim dstWbKey As Variant, srcWbKey As Variant
    Dim dstVal As Variant, srcVal As Variant
    Dim errorMsg As String
    checkTwoWbsStyle = True
    For Each dstWbKey In dstWbCoverStyleMap.KeyCollection
        For Each srcWbKey In srcWbCoverStyleMap.KeyCollection
            If dstWbKey = srcWbKey Then
                dstVal = dstWbCoverStyleMap.GetAt(dstWbKey)
                srcVal = srcWbCoverStyleMap.GetAt(srcWbKey)
                If dstVal <> srcVal Then
                    If dstWbKey = "wbName" Then
                        errorMsg = getResByKey("WrongWbType") & Chr(10) _
                        & getResByKey("CurrentWb") & dstVal & Chr(10) _
                        & getResByKey("ImportedWb") & srcVal
                        Call MsgBox(errorMsg, vbExclamation, getResByKey("FailToImport"))
                    Else
                        errorMsg = getResByKey("WrongVersion") & Chr(10) _
                        & getResByKey("CurrentWb") & Chr(10) _
                        & dstWbKey & ": " & dstVal & Chr(10) _
                        & getResByKey("ImportedWb") & Chr(10) _
                        & srcWbKey & ": " & srcVal
                        
                        Call MsgBox(errorMsg, vbExclamation, getResByKey("FailToImport"))
                    End If
                    
                    checkTwoWbsStyle = False
                    Exit Function
                End If
            End If
        Next srcWbKey
    Next dstWbKey
End Function

Function isSrcFileIubStyleWorkBook(srcwb As Workbook) As Boolean
    On Error GoTo ErrorHandler:
    If innerPositionMgr Is Nothing Then loadInnerPositions
    Dim ranges As range
    If srcwb.Worksheets("SHEET DEF").Cells(1, innerPositionMgr.sheetDef_startRowColNo).value <> "" Then
        isSrcFileIubStyleWorkBook = True
    Else
        isSrcFileIubStyleWorkBook = False
    End If
    Exit Function
ErrorHandler:
    isSrcFileIubStyleWorkBook = False
End Function

Function checkIubStyle(srcwb As Workbook) As Boolean
    Dim coverName As String, issueDate As String, errorMsg As String
    Dim dstWbCoverStyleMap As New CMap
    Dim srcWbCoverStyleMap As New CMap
    checkIubStyle = True
    coverName = getResByKey("Cover")
    If whetherWbHasCertainSheet(srcwb, coverName) Then
        If Not isSrcFileIubStyleWorkBook(srcwb) Then
            Call MsgBox(getResByKey("FileStyleNotMatch"), vbExclamation, getResByKey("FailToImport"))
            checkIubStyle = False
            Exit Function
        End If
        issueDate = getResByKey("ISSUE_DATE")
        Call getCoverStyleMap(ThisWorkbook.Worksheets(coverName), dstWbCoverStyleMap, issueDate)
        Call getCoverStyleMap(srcwb.Worksheets(coverName), srcWbCoverStyleMap, issueDate)
        checkIubStyle = checkTwoWbsStyle(dstWbCoverStyleMap, srcWbCoverStyleMap)
    Else
        errorMsg = getResByKey("CoverNotFound")
        Call MsgBox(errorMsg, vbExclamation, getResByKey("FailToImport"))
        checkIubStyle = False
    End If
End Function

Function findMappingRow(templateWorkBook As Workbook, templateSheetDef As Worksheet, address As String) As range
    Dim iRow As Long
    Dim column As Long
    Dim recordRow As Long
    iRow = range(address).row
    column = range(address).column
    Dim index As Long
    Dim count As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim sheetName As String
    Dim columnName As String
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    count = templateSheetDef.UsedRange.rows.count
    For index = 2 To count
        If templateSheetDef.Cells(index, innerPositionMgr.sheetDef_startRowColNo).value <> "" And templateSheetDef.Cells(index, innerPositionMgr.sheetDef_endRowColNo).value <> "" Then
            startPos = CLng(templateSheetDef.Cells(index, innerPositionMgr.sheetDef_startRowColNo).value)
            endPos = CLng(templateSheetDef.Cells(index, innerPositionMgr.sheetDef_endRowColNo).value)
            If startPos < iRow And iRow <= endPos Then
                sheetName = templateSheetDef.Cells(index, innerPositionMgr.sheetDef_shtNameColNo).value
                recordRow = iRow - startPos
                Exit For
            End If
        End If
    Next
    Dim mappingDef As CMappingDef
    Set mappingDef = getMappingDefByAddress(templateWorkBook, startPos, column, sheetName)
    Call getSheetNameByMappingDef(sheetName, columnName, mappingDef)
    If sheetName <> "" And columnName <> "" Then
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        count = sheetDef.UsedRange.rows.count
        iRow = -1
        For index = 2 To count
            If sheetDef.Cells(index, innerPositionMgr.sheetDef_shtNameColNo).value = sheetName Then
                If sheetDef.Cells(index, innerPositionMgr.sheetDef_startRowColNo).value <> "" And sheetDef.Cells(index, innerPositionMgr.sheetDef_endRowColNo).value <> "" Then
                     startPos = CLng(sheetDef.Cells(index, innerPositionMgr.sheetDef_startRowColNo).value)
                     endPos = CLng(sheetDef.Cells(index, innerPositionMgr.sheetDef_endRowColNo).value)
                     If startPos + recordRow <= endPos Then
                         iRow = startPos + recordRow
                     Else
                         Set findMappingRow = Nothing
                         Exit Function
                     End If
                End If
                Exit For
            End If
        Next
        If iRow > 0 Then
            Dim iRange As range
            Dim sh As Worksheet
            For Each sh In ThisWorkbook.Worksheets
                If ThisWorkbook.Worksheets(sh.name).Tab.colorIndex = BluePrintSheetColor Then
                    For index = 2 To sh.range("IV" + CStr(startPos)).End(xlToLeft).column
                        If sh.Cells(startPos, index).value = columnName Then
                            Set findMappingRow = sh.Cells(iRow, index)
                            Exit Function
                        End If
                    Next
                End If
            Next
        End If
    End If
End Function

'iub Style专用根据MappingDef获取Sheet
Private Sub getSheetNameByMappingDef(ByRef sheetName As String, ByRef columnName As String, mappingDef As CMappingDef)
    Dim sh As Worksheet
    Dim row As range
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Set sh = ThisWorkbook.Worksheets("MAPPING DEF")
    For Each row In sh.rows
        If mappingDef.mocName = row.Cells(1, innerPositionMgr.mappingDef_mocNameColNo).value And _
            mappingDef.attributeName = row.Cells(1, innerPositionMgr.mappingDef_attrNameColNo).value And mappingDef.neType = row.Cells(1, innerPositionMgr.mappingDef_neTypeColNo).value Then
            sheetName = row.Cells(1, innerPositionMgr.mappingDef_shtNameColNo).value
            columnName = row.Cells(1, innerPositionMgr.mappingDef_colNameColNo).value
            Exit Sub
        End If
    Next
    sheetName = ""
    columnName = ""
End Sub

'iub Style专用根据表格起始行，找到MappingDef
Private Function getMappingDefByAddress(wb As Workbook, startPos As Long, column As Long, sheetName As String) As CMappingDef
    Dim sh As Worksheet
    Dim mappingDef As CMappingDef
    Dim mappingDefSheet As Worksheet
    Dim row As range
    If innerPositionMgr Is Nothing Then loadInnerPositions
    For Each sh In wb.Worksheets
        If wb.Worksheets(sh.name).Tab.colorIndex = BluePrintSheetColor Then
            Set mappingDefSheet = wb.Worksheets("MAPPING DEF")
            For Each row In mappingDefSheet.rows
                If row.Cells(1, innerPositionMgr.mappingDef_shtNameColNo).value = sheetName And row.Cells(1, innerPositionMgr.mappingDef_colNameColNo).value = sh.Cells(startPos, column).value Then
                     Set mappingDef = New CMappingDef
                     mappingDef.sheetName = row.Cells(1, innerPositionMgr.mappingDef_shtNameColNo).value
                     mappingDef.groupName = row.Cells(1, innerPositionMgr.mappingDef_grpNameColNo).value
                     mappingDef.columnName = row.Cells(1, innerPositionMgr.mappingDef_colNameColNo).value
                     mappingDef.mocName = row.Cells(1, innerPositionMgr.mappingDef_mocNameColNo).value
                     mappingDef.attributeName = row.Cells(1, innerPositionMgr.mappingDef_attrNameColNo).value
                     mappingDef.neType = row.Cells(1, innerPositionMgr.mappingDef_neTypeColNo).value
                     Set getMappingDefByAddress = mappingDef
                     Exit Function
                End If
            Next
        End If
    Next
End Function

Sub deleteRef()
    If Not checkDeleteRange() Then
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(GetMainSheetName())
    Dim iRange As range
    For Each iRange In Selection
        Call deleteRefAddress(Replace(iRange.address, "$", ""), ws)
        Call setAllIubSheetCellStyle(Replace(iRange.address, "$", ""), "", "", -4142)
        Call referenceMessages.remmoveReferenceMess(Replace(iRange.address, "$", ""))
    Next
End Sub

Sub deleteRefAddress(address As String, sh As Worksheet)
    Dim cloumNum As Long
    Dim index As Long
    Dim addressIndex As Long
    Dim addressArray As Variant
    Dim newAddress As String
    cloumNum = sh.range("IV3").End(xlToLeft).column
    
    For index = 1 To cloumNum
        If sh.Cells(3, index).value <> "" Then
            addressArray = Split(sh.Cells(3, index).value, ",")
            newAddress = ""
            For addressIndex = LBound(addressArray) To UBound(addressArray)
                If addressArray(addressIndex) <> address Then
                    If newAddress <> "" Then
                       newAddress = newAddress + "," + addressArray(addressIndex)
                    Else
                       newAddress = addressArray(addressIndex)
                    End If
                End If
            Next
            sh.Cells(3, index).value = newAddress
        End If
    Next
End Sub

Private Sub setAllIubSheetCellStyle(address As String, text As String, title As String, colorIndex As Long)
    Dim ws As Worksheet
    Dim iRange As range

    For Each ws In ThisWorkbook.Worksheets
        If isIubStyleWorkSheetByParameterWs(ws) Then
            Set iRange = ws.range(address)
            Call setCellStyle(iRange, text, title, colorIndex)
        End If
    Next
    
End Sub

Private Sub setCellStyle(iRange As range, text As String, title As String, colorIndex As Long)
    Call addValidation(iRange)
    With iRange.Validation
        .inputTitle = title
        .inputmessage = text
        .ShowInput = True
        .ShowError = False
    End With
    With iRange.Interior
        .colorIndex = colorIndex
    End With
End Sub

Private Sub addValidation(iRange As range)
On Error Resume Next
    With iRange.Validation
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
    End With
End Sub


Private Function checkDeleteRange() As Boolean
    Dim iRange As range
    For Each iRange In Selection
        If iRange.Interior.colorIndex <> HyperLinkColorIndex Then
            MsgBox getResByKey(Replace(iRange.address, "$", "") + getResByKey("deleteRefWarning")), vbExclamation + vbOKCancel, getResByKey("Warning")
            checkDeleteRange = False
            Exit Function
        End If
    Next
    checkDeleteRange = True
End Function

Sub addListHyperlinks()
    If Not checkRefRange(Selection) Then
        Exit Sub
    End If
    Load HyperlinksForm
    ListHyperlinksForm.Show
End Sub

Sub addListRef(sheetName As String, groupName As String, coloumName As String)
    Dim selectRange As range
    Set selectRange = Selection
    Dim cell As range
    For Each cell In selectRange
        '在传输页加字段
        Call addIubStyleCellRef(sheetName, groupName, coloumName, cell)
    Next
End Sub

Private Sub addIubStyleCellRef(sheetName As String, groupName As String, columnName As String, cell As range)
    Dim rowNum As Long
    Dim columnNum As Long
    Call addGroupAndColoum(sheetName, groupName, columnName)
    Call getRowNumAndColumnNum(sheetName, groupName, columnName, rowNum, columnNum)
    Call deleteRefAddress(Replace(cell.address, "$", ""), ThisWorkbook.Worksheets(GetMainSheetName()))
    Call setRefAddressValue(cell, sheetName, rowNum + 1, columnNum)
    Dim text As String
    text = sheetName + "\" + groupName + "\" + columnName
    Call setAllIubSheetCellStyle(Replace(cell.address, "$", ""), text, getResByKey("Reference Address"), HyperLinkColorIndex)
    Call referenceMessages.addReferenceMess(Replace(cell.address, "$", ""), text)
End Sub


Sub setRefAddressValue(cell As range, sheetName As String, rowNum As Long, columnNum As Long)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(sheetName)
    If sh.Cells(rowNum, columnNum).value = "" Then
        sh.Cells(rowNum, columnNum).value = Replace(cell.address, "$", "")
    Else
        sh.Cells(rowNum, columnNum).value = sh.Cells(rowNum, columnNum).value + "," + Replace(cell.address, "$", "")
    End If
End Sub

Private Sub getRowNumAndColumnNum(sheetName As String, groupName As String, columnName As String, rowNum As Long, columnNum As Long)
    Dim ws As Worksheet
    Dim m_rowNum As Long
    Dim m_colNum As Long
    Dim m_colNum1 As Long
    Dim columnsNum As Long
    
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If sheetName = getResByKey("Comm Data") Then
        For m_rowNum = 1 To ws.range("a65536").End(xlUp).row
            If groupName = ws.Cells(m_rowNum, 1).value Then
                For m_colNum = 1 To ws.range("IV" + CStr(m_rowNum + 1)).End(xlToLeft).column
                    If columnName = ws.Cells(m_rowNum + 1, m_colNum).value Then
                        rowNum = m_rowNum + 1
                        columnNum = m_colNum
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    Else
        For m_colNum = 1 To ws.range("IV1").End(xlToLeft).column
            If groupName = ws.Cells(1, m_colNum).value Then
                columnsNum = ws.Cells(1, m_colNum).MergeArea.columns.count
                For m_colNum1 = m_colNum To m_colNum + columnsNum - 1
                    If columnName = ws.Cells(2, m_colNum1).value Then
                        rowNum = 2
                        columnNum = m_colNum1
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    End If
End Sub

Sub addGroupAndColoum(sheetName As String, groupName As String, coloumName As String)
    Dim m_colNum As Long
    Dim groupIndex As Long
    Dim coloumStart As Long
    Dim coloumEnd As Long
    Dim columStartChar, columEndChar As String
    Dim coloumIndex As Long
    
    Dim isFound As Boolean
    isFound = False
    groupIndex = -1
    coloumIndex = -1
    Dim index As Long
    Call changeAlert(False)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    For m_colNum = 1 To ws.range("IV" + CStr(2)).End(xlToLeft).column
        If groupName = ws.Cells(1, m_colNum).value Then
            isFound = True
            coloumStart = m_colNum
            groupIndex = m_colNum
        ElseIf ws.Cells(1, m_colNum).value <> "" And isFound = True Then
            Exit For
        End If
        coloumEnd = m_colNum
    Next
    
    Application.CutCopyMode = False
    If groupIndex > 0 Then
        For m_colNum = coloumStart To coloumEnd
            If coloumName = ws.Cells(2, m_colNum).value Then
                coloumIndex = m_colNum
            End If
        Next
    End If
    If coloumIndex > 0 Then
        Exit Sub
    ElseIf coloumIndex <= 0 And groupIndex > 0 Then
        columEndChar = getColStr(coloumEnd + 1)
        columStartChar = getColStr(coloumStart)
        ws.columns(columEndChar + ":" + columEndChar).INSERT Shift:=xlToLeft
        
        Call clearValidationAndResetBackgroundStyle(ws, columEndChar)
        
        ws.Cells(2, coloumEnd + 1).value = coloumName
        ws.range(columStartChar + "1:" + columEndChar + "1").Merge
        Call addGroupNameAndColoumName(sheetName, groupName, coloumName)
        setBoard (sheetName)
    ElseIf coloumIndex <= 0 And groupIndex <= 0 Then
        columEndChar = getColStr(coloumEnd + 1)
        ws.columns(columEndChar + ":" + columEndChar).INSERT Shift:=xlToLeft
        
        Call clearValidationAndResetBackgroundStyle(ws, columEndChar)

        ws.Cells(2, coloumEnd + 1).value = coloumName
        ws.Cells(1, coloumEnd + 1).value = groupName
        Call addGroupNameAndColoumName(sheetName, groupName, coloumName)
        setBoard (sheetName)
    End If
     Call changeAlert(True)
End Sub

Sub clearValidationAndResetBackgroundStyle(ByRef ws As Worksheet, ByRef columEndChar As String)
    Dim newColumnRange As range
    
    Set newColumnRange = ws.columns(columEndChar + ":" + columEndChar)
    Call clearValidation(newColumnRange)
    
    Dim maxRow As Integer
    maxRow = ws.range("a65536").End(xlUp).row
    Set newColumnRange = ws.range(ws.range(columEndChar & "4"), ws.range(columEndChar & maxRow))
    Call resetBackgroundStyle(newColumnRange)
End Sub

Sub clearValidation(ByRef certainRange As range)
    With certainRange.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Sub resetBackgroundStyle(ByRef certainRange As range)
    certainRange.Interior.colorIndex = xlNone
    certainRange.Interior.Pattern = xlNone
End Sub


Sub addGroupNameAndColoumName(sheetName As String, groupName As String, coloumName As String)
    Dim index As Long
    Dim row As Long
    row = -1
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To mappingDef.range("a65536").End(xlUp).row
        row = index
        If mappingDef.Cells(index, 1).value = sheetName _
        And mappingDef.Cells(index, 2).value = groupName _
        And mappingDef.Cells(index, 3).value = coloumName Then
            Exit For
        End If
    Next
    mappingDef.rows(row + 1).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    mappingDef.Cells(row + 1, 1).value = sheetName
    mappingDef.Cells(row + 1, 2).value = groupName
    mappingDef.Cells(row + 1, 3).value = coloumName
    
    Dim mpdef As CMappingDef
    Set mpdef = New CMappingDef
    mpdef.sheetName = sheetName
    mpdef.groupName = groupName
    mpdef.columnName = coloumName
    If Not Contains(mappingDefineMap, mpdef.getKey) And Not mappingDefineMap Is Nothing Then
        mappingDefineMap.Add Item:=mpdef, key:=mpdef.getKey
    End If
End Sub

Sub setRangeBoard(myRange As range)
    With myRange
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .colorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .colorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .colorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .colorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .Weight = xlThin
        End With
        .NumberFormatLocal = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
End Sub

Sub setBoard(sheetName As String)
    Dim maxRow As Long
    Dim maxColomn As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    maxColomn = ws.range("IV2").End(xlToLeft).column
    maxRow = ws.range("a65536").End(xlUp).row
    If maxRow = 0 Then Exit Sub
    Dim myRange As range
    Set myRange = ws.range("A1:" + getColStr(maxColomn) + CStr(maxRow))
    Call setRangeBoard(myRange)
End Sub

Function isListSheet(sheetName As String) As Boolean
    isListSheet = False
    If getSheetType(sheetName) = "LIST" Then isListSheet = True
End Function

Function getSheetType(sheetName As String) As String
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    getSheetType = ""
    
    Dim targetRange As range
    Set targetRange = sheetDef.columns(innerPositionMgr.sheetDef_shtNameColNo).Find(sheetName, lookat:=xlWhole, LookIn:=xlValues)
    If Not targetRange Is Nothing Then getSheetType = UCase(sheetDef.Cells(targetRange.row, innerPositionMgr.sheetDef_shtTypeColNo).value)
End Function

Private Function judgeTwoValues(ByRef value1 As Variant, ByRef value2 As Variant) As Boolean
    Dim response As Variant
    judgeTwoValues = True
    If value1 <> value2 Then
        response = MsgBox(getResByKey("ValueNotEqual") & Chr(13) _
        & Chr(13) & getResByKey("SureToContinue"), vbYesNo + vbQuestion, getResByKey("Comm Data"))
        If response = vbYes Then
            judgeTwoValues = True
        Else
            judgeTwoValues = False
        End If
    End If
End Function
 
 Function checkRefRange(ByRef selectionRanges As range)
    checkRefRange = True
    Dim selectionRange As range
    Dim selectionColor As Variant, selectionColorIndex As Variant
    Dim ValueCollection As New Collection
    For Each selectionRange In selectionRanges
        selectionColorIndex = selectionRange.Interior.colorIndex
        selectionColor = selectionRange.Interior.Color
        ValueCollection.Add Item:=selectionRange.value
        'ColorIndex是一个大范围颜色索引，会有多个颜色的ColorIndex一样，所以这里不能用ColorIndex来判断，需要用Color来判断
        If selectionColor = 12632256 _
            Or selectionColor = 128 _
            Or selectionColor = 10079487 Then
            Call MsgBox(getResByKey("selectedCellsIllegal"), vbExclamation, getResByKey("IllegalSelection"))
            checkRefRange = False
            Exit Function
        End If
    Next selectionRange
    
    Dim selectionValue1 As Variant, selectionValue2 As Variant
    Dim k As Long
    k = 1
    For k = 1 To ValueCollection.count
        selectionValue1 = ValueCollection.Item(k)
        If k <> 1 Then
            checkRefRange = judgeTwoValues(selectionValue1, selectionValue2)
            If selectionValue1 <> selectionValue2 Then Exit Function
        End If
        selectionValue2 = selectionValue1
    Next k
End Function

Function findCertainValColumnNumber(ByRef ws As Worksheet, ByVal rowNumber As Long, ByRef cellVal As Variant) As Long
    Dim k As Long
    Dim currentCellVal As Variant
    Dim maxColumnNumber As Long
    maxColumnNumber = ws.UsedRange.columns.count
    findCertainValColumnNumber = -1
    For k = 1 To maxColumnNumber
        currentCellVal = ws.Cells(rowNumber, k).value
        If currentCellVal = cellVal Then
            findCertainValColumnNumber = k
            Exit For
        End If
    Next
End Function













