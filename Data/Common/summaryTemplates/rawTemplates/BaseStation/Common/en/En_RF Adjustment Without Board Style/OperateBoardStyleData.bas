Attribute VB_Name = "OperateBoardStyleData"
Option Explicit

Public baseStationData As CBaseStationData
Public boardStyleData As CBoardStyleData
Public inAddProcessFlag As Boolean
Public currentSheet As Worksheet
Public currentCellValue As String
Public addBoardStyleMoiInProcess As CAddingBoardStyleMoc
Public moiRowsManager As CMoiRowsManager
Public boardNoManager As CBoardNoManager
Public boardStyleMappingDefMap As CMapValueObject
Public selectedGroupMappingDefData As CBoardStyleMappingDefData
Public cGpsboardNoMap As CMap

Public addBoardStyleButtons As CAddBoardStyleButtons
Public deleteBoardStyleButtons As CDeleteBoardStyleButtons
Public Const NewMoiRangeColorIndex As Long = 43 '浅绿，新增Moi行底色
Public Const NeedFillInRangeColorIndex As Long = 33 '浅蓝，必填单元格底色
Public Const NormalRangeColorIndex As Long = -4142 '白色，正常单元格底色
Public Const BoardNoDelimiter As String = "_"

Public Const PublicMaxRowNumber As Long = 2000

Public Sub addBoardStyleMoiFinishButton()
    On Error GoTo ErrorHandler
    If checkNeedFillInCellsFilled = False Then Exit Sub

    Call makeAutoFillInColumnNameValue
    Call selectCertainCell(currentSheet, "A" & moiRowsManager.groupNameRowNumber)
    Call setNewRangesBackgroundColour(NormalRangeColorIndex)
    Unload BoardStyleForm
    Exit Sub
ErrorHandler:
End Sub

Public Sub setNewRangesBackgroundColour(ByRef colorIndex As Long)
    Dim newMoiRange As range, eachCell As range
    Set newMoiRange = moiRowsManager.getMoiRange
    For Each eachCell In newMoiRange
        '如果单元格不是灰化的，则置为正常底色，是分支控制的灰化，则不变
        If eachCell.Interior.colorIndex <> SolidColorIdx And eachCell.Interior.Pattern <> SolidPattern Then eachCell.Interior.colorIndex = colorIndex
    Next eachCell
End Sub

Public Sub addBoardStyleMoiCancelButton()
    On Error GoTo ErrorHandler
    Dim deletedRowsRange As range
    Set deletedRowsRange = moiRowsManager.getMoiRowsRange
    deletedRowsRange.Delete
    Call selectCertainCell(currentSheet, "A" & moiRowsManager.groupNameRowNumber)
    Unload BoardStyleForm
    Exit Sub
ErrorHandler:
End Sub

Public Function isBoardStyleSheet(ByRef ws As Worksheet) As Boolean
    Dim sheetName As String, boardStyleSheetName As String
    isBoardStyleSheet = False
    boardStyleSheetName = getResByKey("Board Style")
    If InStr(ws.name, boardStyleSheetName) <> 0 Then
        isBoardStyleSheet = True
    End If
End Function

Public Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Dim sheet As Worksheet
    Set sheet = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsASheet = False
End Function

Public Function findCertainValColumnNumber(ByRef ws As Worksheet, ByVal rowNumber As Long, ByRef cellVal As Variant, Optional ByVal startColumn As Long = 1)
    Dim currentCellVal As Variant
    Dim maxColumnNumber As Long, k As Long
    maxColumnNumber = ws.UsedRange.Columns.count
    findCertainValColumnNumber = -1
    For k = startColumn To maxColumnNumber
        currentCellVal = ws.Cells(rowNumber, k).value
        If currentCellVal = cellVal Then
            findCertainValColumnNumber = k
            Exit For
        End If
    Next
End Function

Public Function findCertainValRowNumber(ByRef ws As Worksheet, ByRef cellVal As Variant, Optional ByVal startRow As Long = 1)
    
    Dim maxRowNumber As Long, row As Long, colMax As Long
    Dim curRowRange As range
    
    findCertainValRowNumber = -1
    
    maxRowNumber = ws.UsedRange.rows.count
    For row = startRow To maxRowNumber
        If Not isGroupRow(ws, row) Then GoTo NextLoop
        
        If findColNumByRowAndValue(ws, row, cellVal) <> -1 Then
            findCertainValRowNumber = row
            Exit For
        End If
NextLoop:
    Next
End Function

Public Function containsAToolBar(ByRef barName As String) As Boolean
    On Error GoTo ErrorHandler
    containsAToolBar = True
    Dim bar As CommandBar
    Set bar = CommandBars(barName)
    Exit Function
ErrorHandler:
    containsAToolBar = False
End Function

Public Sub eraseLastChar(ByRef str As String)
    If str <> "" Then str = Left(str, Len(str) - 1)
End Sub

Public Sub selectCertainCell(ByVal ws As Worksheet, ByVal rangeName As String, Optional ByVal scrollFlag As Boolean = True)
    Application.GoTo Reference:=ws.range(rangeName), Scroll:=scrollFlag
End Sub

Public Sub resetAddBoardStyleMoiInfo(ByRef ws As Worksheet)
    'inAddProcessFlag = False
    Set currentSheet = ws
End Sub

Private Function checkNeedFillInCellsFilled() As Boolean
    checkNeedFillInCellsFilled = True
        
    Dim emptyCell As range
    Dim emptyCellAddress As String
    Dim emptyCellAddressString As String
    If moiRowsManager.checkNeedFillInCells(emptyCell, emptyCellAddressString) = False Then
        emptyCellAddress = emptyCell.address(False, False)
        Call MsgBox(getResByKey("EmptyCellFound") & vbCrLf & emptyCellAddressString, vbExclamation)
        Call selectCertainCell(currentSheet, emptyCellAddress, False)
        checkNeedFillInCellsFilled = False
    End If
End Function

Private Sub makeAutoFillInColumnNameValue()
    Dim boardNoName As String
    boardNoName = selectedGroupMappingDefData.autoFillInColumnName
    If boardNoName = "" Then Exit Sub
    
    Dim startRowNumber As Long, endRowNumber As Long
    startRowNumber = moiRowsManager.startRowNumber
    endRowNumber = moiRowsManager.endRowNumber
    
    Call clearBoardNoRanges(startRowNumber, endRowNumber)
    Call initBoardNoManagerPublic
        
    Dim boardNoSourceAttributes As String
    Dim boardNoColumnLetter As String
    boardNoSourceAttributes = selectedGroupMappingDefData.autoFillInColumnNameSourceAttributes
    Dim boardNoSourceAttributeColumnLetterArr As Variant
    boardNoSourceAttributeColumnLetterArr = getBoardNoSourceColumnLetterAttr(boardNoSourceAttributes)
    boardNoColumnLetter = selectedGroupMappingDefData.getColumnNamePositionLetter(boardNoName)

    Call fillInBoardNoRanges(boardNoColumnLetter, startRowNumber, endRowNumber, boardNoSourceAttributeColumnLetterArr)
End Sub

Public Sub fillInBoardNoRanges(ByRef boardNoColumnLetter As String, ByRef startRowNumber As Long, ByRef endRowNumber As Long, ByRef sourceAttributeColumnLetterArr As Variant)
    Dim tempBoardNo As String
    Dim rowNumber As Long
    For rowNumber = startRowNumber To endRowNumber
        tempBoardNo = getTempBoardNo(rowNumber, sourceAttributeColumnLetterArr) '得到0_0_1_这样的编号前缀
        Call boardNoManager.getNewBoardNo(tempBoardNo) '得到最终要生成的单板编号tempBoardNo
        tempBoardNo = tempBoardNo & "(n)"
        Call boardNoManager.addNewBoardNo(selectedGroupMappingDefData.groupName, tempBoardNo)
        currentSheet.range(boardNoColumnLetter & rowNumber).Interior.colorIndex = NullPattern
        currentSheet.range(boardNoColumnLetter & rowNumber).Interior.Pattern = NullPattern
        currentSheet.range(boardNoColumnLetter & rowNumber).value = tempBoardNo
    Next rowNumber

End Sub

Public Function getTempBoardNo(ByRef rowNumber As Long, ByRef sourceAttributeColumnLetterArr As Variant)
    Dim tempBoardNo As String
    tempBoardNo = ""
    Dim index As Long
    Dim attributeValue As String
    For index = LBound(sourceAttributeColumnLetterArr) To UBound(sourceAttributeColumnLetterArr)
        attributeValue = currentSheet.range(sourceAttributeColumnLetterArr(index) & rowNumber).value
        tempBoardNo = tempBoardNo & attributeValue & BoardNoDelimiter
    Next index
    getTempBoardNo = tempBoardNo
End Function

Public Function getBoardNoSourceColumnLetterAttr(ByRef sourceAttr As String) As Variant
    Dim columnLetterArr As Variant
    Dim colunmLetter As String, columnName As String
    columnLetterArr = Split(sourceAttr, ",")
    
    Dim index As Long
    For index = LBound(columnLetterArr) To UBound(columnLetterArr)
        columnName = columnLetterArr(index)
        colunmLetter = selectedGroupMappingDefData.getColumnNamePositionLetter(columnName)
        columnLetterArr(index) = colunmLetter
    Next index
    getBoardNoSourceColumnLetterAttr = columnLetterArr
End Function

Public Sub popUpSheetCannotChangeMsgbox()
    Call MsgBox(getResByKey("CannotChangeSheet"), vbExclamation)
    '"正在进行增加BoardStyle MOC的操作，不能切换页签，请先完成增加操作或退出操作。"
    currentSheet.Select
End Sub

Public Sub popUpWbCannotSaveMsgbox()
    Call MsgBox(getResByKey("CannotSaveWb"), vbExclamation)
End Sub

Public Sub boardStyleSelectionChange(ByRef ws As Worksheet, ByRef target As range)
    On Error GoTo ErrorHandler
    Dim groupName As String, columnName As String
    Dim rowNumber As Long, columnNumber As Long
    If target.rows.count <> 1 Or target.Columns.count <> 1 Then Exit Sub
    
    rowNumber = target.row
    columnNumber = target.column
    
    Call getCgpsBoraNoMap(ws)
    
    Dim targetInRecordsRangeFlag As Boolean, targetIsInListBoxFlag As Boolean, targetInBoardNoFlag As Boolean
    targetInRecordsRangeFlag = getRangeGroupAndColumnName(ws, rowNumber, columnNumber, groupName, columnName)
        
    Dim referencedString As String
    Dim currentBoardStyleMappingDefData As CBoardStyleMappingDefData
    targetIsInListBoxFlag = getReferecedString(groupName, columnName, referencedString, currentBoardStyleMappingDefData) '判断选定的列是否需要增加自动下拉框
    
    targetInBoardNoFlag = judgeWhetherInBoardNoColumn(columnName, currentBoardStyleMappingDefData) '判断选定的列是否是BoardNo
    
    If targetIsInListBoxFlag = False And targetInBoardNoFlag = False Then
        'Target.Validation.Delete
        Exit Sub '如果不是需要添加下拉列表的参数，则直接退出
    End If
    If targetInRecordsRangeFlag = False Then '如果不在数据范围内，则先将有效性清空，再退出
        'Target.Validation.Delete
        Exit Sub
    End If
    
    If targetIsInListBoxFlag = True Then
        Call setBoardStyleListBoxRangeValidation(ws.name, groupName, columnName, referencedString, ws, target)
    ElseIf targetInBoardNoFlag = True Then
        Call setBoardNoRangeValidation(target)
    End If
    Exit Sub
ErrorHandler:
End Sub

Private Function judgeWhetherInBoardNoColumn(ByRef columnName As String, ByRef currentBoardStyleMappingDefData As CBoardStyleMappingDefData)
    If Not currentBoardStyleMappingDefData Is Nothing Then
        If columnName = currentBoardStyleMappingDefData.autoFillInColumnName And columnName <> "" Then
            judgeWhetherInBoardNoColumn = True
        Else
            judgeWhetherInBoardNoColumn = False
        End If
    End If
End Function

Private Sub setBoardNoRangeValidation(ByRef target As range)
    'target.Offset(0, 1).Select
    With target.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation '只给提示信息有效性设置
        .inputTitle = getResByKey("ForbiddenEditTitle") '输入提示Title
        .inputMessage = getResByKey("ForbiddenEditContent") '输入提示内容
        .ShowInput = True 'True显示输入提示信息，False不显示输入提示信息
        .ShowError = False 'True不允许输入非有效性值，False允许输入
    End With
End Sub

Public Sub setBoardStyleListBoxRangeValidation(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, _
    ByRef referencedString As String, ByRef sheet As Worksheet, ByRef target As range)
    If Len(referencedString) > 255 Then
        Dim valideDef As CValideDef
        'Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        Set valideDef = initDefaultDataSub.getInnerValideDef(sheetName + "," + groupName + "," + columnName)
        If valideDef Is Nothing Then
            Set valideDef = addInnerValideDef(sheetName, groupName, columnName, referencedString)
        Else
            Call modiflyInnerValideDef(sheet.name, groupName, columnName, referencedString, valideDef)
        End If
        referencedString = valideDef.getValidedef
    End If
    With target.Validation
        .Delete
        If referencedString <> "" Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
            .ShowError = False
        Else
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=" "
            .ShowError = False
        End If
    End With
End Sub

Public Function getReferecedString(ByRef groupName As String, ByRef columnName As String, ByRef referencedString As String, _
    ByRef currentBoardStyleMappingDefData As CBoardStyleMappingDefData) As Boolean
    getReferecedString = True
    
    If groupName = "" Or columnName = "" Then
        getReferecedString = False
        Exit Function
    End If
    
    If boardStyleData Is Nothing Then Call initBoardStyleMappingDataPublic
    Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    Set currentBoardStyleMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
    
    referencedString = currentBoardStyleMappingDefData.getListBoxReferenceValue(columnName)
    If referencedString = "" Then
        getReferecedString = False
        Exit Function
    End If
    
    referencedString = getReferencedStringByInputString(referencedString)
End Function

Public Function getReferencedStringByInputString(ByRef referencedString As String)
    Call initBoardNoManagerPublic
    Dim mocNameBoardNoArr As Variant
    mocNameBoardNoArr = Split(referencedString, "-")
    getReferencedStringByInputString = getBoardNoString(CStr(mocNameBoardNoArr(0)))
End Function

Private Function getBoardNoString(ByRef mocNameString As String)
    Dim mocNameArr As Variant
    mocNameArr = Split(mocNameString, ",")
    Dim index As Long
    Dim boardNoString As String, mocboardNoString As String, groupName As String
    boardNoString = ""
    For index = LBound(mocNameArr) To UBound(mocNameArr)
        groupName = boardStyleData.getGroupNameByMocName(CStr(mocNameArr(index)))
        mocboardNoString = boardNoManager.getBoardNoStringByGroupName(groupName)
        Call fillterBoradNo(mocboardNoString, groupName)
        If mocboardNoString <> "" Then
            If boardNoString = "" Then
                boardNoString = boardNoString & mocboardNoString
            Else
                boardNoString = boardNoString & "," & mocboardNoString
            End If
        End If
    Next index
    getBoardNoString = boardNoString
End Function

Public Function getGroupRowNumber(ByRef ws As Worksheet, ByVal rowNumber) As Long
    Dim maxRowNumber As Long, k As Long
    
    maxRowNumber = ws.UsedRange.rows.count
    If rowNumber > maxRowNumber Then rowNumber = maxRowNumber
    
    For k = rowNumber To 1 Step -1
        If k = 1 Then Exit For
        If (rowIsBlank(ws, k - 1) = True And rowIsBlank(ws, k) = False) Or (rangeHasBorder(ws.rows(k - 1)) = False And rowIsBlank(ws, k) = False) Then Exit For
    Next k
    getGroupRowNumber = k
End Function

Private Function getNextGroupRowNumber(ByRef ws As Worksheet, ByVal rowNumber) As Long
    Dim nextGroupRowNumber As Long
    nextGroupRowNumber = -1
    Dim maxRowNumber As Long, k As Long
    
    maxRowNumber = ws.UsedRange.rows.count
    
    For k = rowNumber To maxRowNumber
        If (rowIsBlank(ws, k) = True Or rangeHasBorder(ws.rows(k)) = False) And rowIsBlank(ws, k + 1) = False Then
            nextGroupRowNumber = k + 1
            Exit For
        End If
    Next k
    
    '如果是-1，说明是最后一个分组，只能用单元格是否有边框来判断最大行了
    If nextGroupRowNumber = -1 Then
        Dim predefinedMaxRowNumber As Long
        predefinedMaxRowNumber = Application.WorksheetFunction.min(rowNumber + 2000, maxRowNumber) '防止最后一个对象的边框一直设置到1048576，设置一个2000的行数最大限制
    
        For k = rowNumber To predefinedMaxRowNumber
            If rangeHasBorder(ws.rows(k)) Then
                nextGroupRowNumber = k
            Else
                Exit For
            End If
        Next k
        
        nextGroupRowNumber = nextGroupRowNumber + 2 '为了与正常分组的最大值一致，这里加2
    End If
    
    getNextGroupRowNumber = nextGroupRowNumber
End Function

Public Function isOperationExcel() As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.name, getResByKey("Board Style")) <> 0 And InStr(ws.Cells(1, 1), getResByKey("Operation")) <> 0 Then
            isOperationExcel = True
            Exit Function
        End If
    Next ws
    isOperationExcel = False
End Function

Public Function getRangeGroupAndColumnName(ByRef ws As Worksheet, ByVal rowNumber As Long, ByVal columnNumber As Long, _
    ByRef groupName As String, ByRef columnName As String) As Boolean
    
    '这里要校验：选择范围不能跨GroupName或ColumnName
    
    Dim cellRange As range
    Set cellRange = ws.range(getColStr(columnNumber) & rowNumber)
     
    Dim groupStartRow As Long, groupEndRow As Long
    Call getGroupStartAndEndRowByRowNum(ws, rowNumber, groupStartRow, groupEndRow)
       
    If (rowNumber > groupStartRow + 1 And rowNumber <= groupEndRow) Then
        
        If rowNumber = groupEndRow And rowIsBlank(ws, rowNumber + 1) = False Then
            getRangeGroupAndColumnName = False
            Exit Function
        End If
        
        getRangeGroupAndColumnName = True
        Call getGroupAndColumnName(ws, cellRange, groupName, columnName, groupStartRow)
    Else
        getRangeGroupAndColumnName = False
    End If
End Function

Public Function rowIsBlank(ByRef ws As Worksheet, ByRef rowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.range("A" & rowNumber & ":IV" & rowNumber)) = 256 Then
        rowIsBlank = True
    Else
        rowIsBlank = False
    End If
End Function

Public Function rangeHasBorder(ByRef certainRange As range) As Boolean
    If certainRange.Borders.LineStyle = xlLineStyleNone Then '没有边框
        rangeHasBorder = False
    Else '有边框
        rangeHasBorder = True
    End If
End Function

Public Function getGroupNameStartRowNumber(ByRef ws As Worksheet, ByRef groupName As String) As Long
    Dim rowNumber As Long, initialRowNumber As Long
    Dim lastRowEmptyFlag As Boolean

    lastRowEmptyFlag = False
    initialRowNumber = 1
    Do While Not lastRowEmptyFlag
        rowNumber = findCertainValRowNumber(ws, groupName, initialRowNumber)
        If rowNumber = -1 Then Exit Do
        If rowNumber = 1 Then
            lastRowEmptyFlag = True
        ElseIf rowIsBlank(ws, rowNumber - 1) = True Or rangeHasBorder(ws.rows(rowNumber - 1)) = False Then
            lastRowEmptyFlag = True
        End If
        initialRowNumber = rowNumber + 1
    Loop
    getGroupNameStartRowNumber = rowNumber
End Function

Public Sub getGroupNameStartAndEndRowNumber(ByRef groupName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
    
    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    Call getValidGroupRangeRows(currentSheet, groupName, groupNameStartRowNumber, groupNameEndRowNumber)
End Sub

Private Function getCurrentRegionRowsCount(ByRef ws As Worksheet, ByRef startRowNumber As Long) As Long
    Dim rowNumber As Long
    Dim rowscount As Long
    rowscount = 1
    For rowNumber = startRowNumber + 1 To PublicMaxRowNumber
        If rowIsBlank(ws, rowNumber) = True Then
            Exit For
        Else
            rowscount = rowscount + 1
        End If
    Next rowNumber
    getCurrentRegionRowsCount = rowscount
End Function

Public Sub clearBoardNoRanges(ByVal startRowNumber As Long, ByVal endRowNumber As Long)
    Dim autoFillInColumnName As String, columnLetter As String
    Dim boardNoRanges As range
    autoFillInColumnName = selectedGroupMappingDefData.autoFillInColumnName
    If autoFillInColumnName <> "" Then
        columnLetter = selectedGroupMappingDefData.getColumnNamePositionLetter(autoFillInColumnName)
        Set boardNoRanges = currentSheet.range(currentSheet.range(columnLetter & startRowNumber), currentSheet.range(columnLetter & endRowNumber))
        boardNoRanges.ClearContents
        boardNoRanges.Interior.colorIndex = SolidColorIdx
        boardNoRanges.Interior.Pattern = SolidPattern
    End If
End Sub

Public Sub initBoardStyleMappingDataPublic()
    If boardStyleData Is Nothing Then Set boardStyleData = New CBoardStyleData
    Call boardStyleData.init
End Sub

Public Sub initAddBoardStyleButtonsPublic()
    If addBoardStyleButtons Is Nothing Then Set addBoardStyleButtons = New CAddBoardStyleButtons
    Call addBoardStyleButtons.init
End Sub

Public Sub initDeleteBoardStyleButtonsPublic()
    If deleteBoardStyleButtons Is Nothing Then Set deleteBoardStyleButtons = New CDeleteBoardStyleButtons
    Call deleteBoardStyleButtons.init
End Sub

Public Sub initBoardNoManagerPublic()
    If boardNoManager Is Nothing Then Set boardNoManager = New CBoardNoManager
    Call boardNoManager.generateCurrentGroupNameBoardNoMap
End Sub

Public Sub addBoardStyleMoi()
    BoardStyleForm.Show
End Sub

'boardstyle页签删除行按钮
Public Sub deleteBoardStyleMoi()
    Dim rowCollection As New Collection
    Dim groupName As String
    
    If checkSelectRanges(rowCollection, groupName) = False Then
        Call MsgBox(getResByKey("ChooseOneMoc"), vbExclamation)
        Exit Sub
    End If
    
    Call deleteRows(rowCollection, groupName)
    
    Call addBoardStyleHyperlinks_SheetActive(currentSheet) '删除行之后引用单元格会改变，因为需要重新调用一次，以刷新当前页的引用
    'Call MsgBox("selected right-" & groupName)
End Sub

Private Sub keepAtLeastOneRow(ByRef rowCollection As Collection, ByRef groupName As String, ByRef ws As Worksheet)
    Dim groupRowNumber As Long, groupEndRowNumber As Long, dataAreaRowsCount As Long
    Dim groupMaxRow As Long
    
    Call getGroupStartAndEndRowByGroupName(ws, groupName, groupRowNumber, groupEndRowNumber)
    dataAreaRowsCount = groupEndRowNumber - (groupRowNumber + 2) + 1
        
    If rowCollection.count < dataAreaRowsCount Then Exit Sub '如果删除的行数小于数据区域行数，则无需操作
    
    Dim newRowNumber As Long
    newRowNumber = rowCollection(rowCollection.count) + 1
    ws.rows(groupRowNumber + 2).Copy
    ws.rows(newRowNumber).Insert Shift:=xlDown
    ws.range(ws.Cells(newRowNumber, 1), ws.Cells(newRowNumber, ws.UsedRange.Columns.count)).ClearContents
End Sub

Private Sub deleteRows(ByRef rowCollection As Collection, ByRef groupName As String)
    If MsgBox(getResByKey("ConfirmMoiDeletion"), vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Call keepAtLeastOneRow(rowCollection, groupName, currentSheet)
    
    Call initBoardNoManagerPublic
    Dim boardNoColumnLetter As String
    Dim outputString As String
    outputString = ""
    
    boardNoColumnLetter = boardNoManager.getBoardNoColumnLetterByGroupName(groupName)
    If boardNoColumnLetter <> "" Then '如果有关联引用的*单板编号列
        Call deleteRowsWithBoardNoColumn(rowCollection, boardNoColumnLetter, outputString)
        Call MsgBox(outputString, vbInformation)
    Else
        Call deleteRowsWithoutBoardNoColumn(rowCollection)
        Call MsgBox(getResByKey("FinishRowsDeletion"), vbInformation)
    End If
    
    Dim groupNameStartRowNumber As Long
    groupNameStartRowNumber = getGroupNameStartRowNumber(currentSheet, groupName)
    currentSheet.range("A" & groupNameStartRowNumber).Select
End Sub

Private Sub deleteRowsWithBoardNoColumn(ByRef rowCollection As Collection, ByRef boardNoColumnLetter As String, ByRef outputString As String)
    Dim referenceRangeManager As New CReferenceRangeManager
    Call referenceRangeManager.generateBoardNoReferenceAddressMap

    Dim boardNo As String, boardNoReferenceAddressString As String
    Dim rowNumber As Variant
    For Each rowNumber In rowCollection
        boardNo = currentSheet.range(boardNoColumnLetter & rowNumber).value
        boardNoReferenceAddressString = referenceRangeManager.getReferenceAddressString(boardNo)
        If boardNoReferenceAddressString <> "" Then
            Call clearReferenceAddress(boardNoReferenceAddressString, rowCollection, referenceRangeManager, boardNo)
            Call makeBoardNoReferenceAddressOutputString(outputString, boardNo, boardNoReferenceAddressString)
        End If
    Next rowNumber
    
    Call deleteRowsWithoutBoardNoColumn(rowCollection)
    If outputString <> "" Then
        outputString = getResByKey("FinishRowsDeletion") & vbCrLf & vbCrLf & getResByKey("ReferenceAddressCleared") & outputString
    Else
        outputString = getResByKey("FinishRowsDeletion")
    End If
End Sub

Private Sub makeBoardNoReferenceAddressOutputString(ByRef outputString As String, ByRef boardNo As String, ByRef boardNoReferenceAddressString As String)
     outputString = outputString & vbCrLf & boardNo & ": " & boardNoReferenceAddressString
End Sub

Private Sub clearReferenceAddress(ByRef referenceAddressString As String, ByRef rowCollection As Collection, _
    ByRef referenceRangeManager As CReferenceRangeManager, ByRef boardNo As String)
    
    Dim referenceAddressArr As Variant, referenceAddress As Variant
    Dim newReferenceAddressString As String, newReferenceAddress As String
    Dim boardNoStringValue As String, newBoardNoStringValue As String
    
    newReferenceAddressString = referenceAddressString
    referenceAddressArr = Split(referenceAddressString, ",")
    
    For Each referenceAddress In referenceAddressArr
        If getNewAddressAfterDeletion(CStr(referenceAddress), newReferenceAddress, rowCollection) = True Then
            newReferenceAddressString = Replace(newReferenceAddressString, CStr(referenceAddress), newReferenceAddress)
        End If
        If Not referenceRangeManager.isMultiListReferenceAddress(CStr(referenceAddress)) Then
            '如果不是多List的单元格，那就是下拉框引用，就直接清空
            currentSheet.range(referenceAddress).ClearContents '清空引用BoardNo的单元格
        Else
            '如果是多List引用的单元格，就把这个单板编号值从这个单元格的值中删除
            boardNoStringValue = currentSheet.range(referenceAddress)
            newBoardNoStringValue = removeBoardNoFromString(boardNoStringValue, boardNo)
            currentSheet.range(referenceAddress).value = newBoardNoStringValue
        End If
    Next referenceAddress
    referenceAddressString = newReferenceAddressString
End Sub

Private Function removeBoardNoFromString(ByRef boardNoStringValue As String, ByRef boardNo As String) As String
    'On Error Resume Next
    Dim boardNoArray As Variant, eachBoardNo As Variant
    Dim boardNoCol As New Collection
    boardNoArray = Split(boardNoStringValue, BasebandReferenceBoardNoDelimeter)
    For Each eachBoardNo In boardNoArray
        '如果不是要清空的单板编号，则把它加入到容器中
        If eachBoardNo <> boardNo Then
            boardNoCol.Add eachBoardNo
        End If
    Next eachBoardNo
    removeBoardNoFromString = getConnectedStringFromCol(boardNoCol, BasebandReferenceBoardNoDelimeter)
End Function

Private Function getNewAddressAfterDeletion(ByRef oldAddress As String, ByRef newAddress As String, ByRef rowCollection As Collection) As Boolean
    getNewAddressAfterDeletion = False
    Dim rowIndex As Variant
    Dim numberOfRowsToShiftUp As Long
    numberOfRowsToShiftUp = 0
    For Each rowIndex In rowCollection
        If range(oldAddress).row > rowIndex Then
            numberOfRowsToShiftUp = numberOfRowsToShiftUp + 1
            getNewAddressAfterDeletion = True
        End If
    Next rowIndex
    newAddress = range(oldAddress).Offset(-numberOfRowsToShiftUp, 0).address(False, False)
End Function

Private Sub deleteRowsWithoutBoardNoColumn(ByRef rowCollection As Collection)
    Dim rowNumber As Variant
    Dim multiRowsDeletionString As String
    Dim lastMatchRowNumber As Long
    lastMatchRowNumber = -1
    For Each rowNumber In rowCollection
        Call makeRowsString(multiRowsDeletionString, lastMatchRowNumber, CLng(rowNumber))
    Next rowNumber
    currentSheet.range(multiRowsDeletionString).Delete
End Sub

'制作拼接待拷贝行字符串代码
Private Sub makeRowsString(ByRef rowsString As String, ByRef lastMatchRowNumber As Long, ByRef rowIndex As Long)
    If rowsString <> "" Then
        If lastMatchRowNumber = rowIndex - 1 Then
            '如果两个行是相连的，则直接将rowsString的最后一个行号替换成新行，提高效率
            Dim lastMatchRowNumberString As String, prefixRowsString As String
            lastMatchRowNumberString = CStr(lastMatchRowNumber)
            prefixRowsString = Mid(rowsString, 1, Len(rowsString) - Len(lastMatchRowNumberString))
            rowsString = prefixRowsString & rowIndex
        Else
            rowsString = rowsString & "," & rowIndex & ":" & rowIndex
        End If
    Else
        '如果是空，则直接拼接第一个行字符串，如1:1
        rowsString = rowIndex & ":" & rowIndex
    End If
    lastMatchRowNumber = rowIndex
End Sub

Private Function checkSelectRanges(ByRef rowCollection As Collection, ByRef groupName As String) As Boolean
    checkSelectRanges = True
    Dim rowRange As range
    Dim selectionRange As range
    Set selectionRange = Selection
    
    Dim rowNumber As Long, columnNumber As Long
    Dim lastGroupName As String, columnName As String

    columnNumber = 1
    For Each rowRange In selectionRange.rows
        rowNumber = rowRange.row
        rowCollection.Add Item:=rowNumber, key:=CStr(rowNumber)
        If checkLastTwoRow(rowNumber, getModelGroupStarRowByRow(ThisWorkbook.ActiveSheet, rowNumber), groupName, columnName, lastGroupName) = False Then
            checkSelectRanges = False
            Exit Function
        End If
    Next rowRange
End Function

Private Function checkLastTwoRow(ByRef rowNumber As Long, ByRef columnNumber As Long, ByRef groupName As String, _
    ByRef columnName As String, ByRef lastGroupName As String) As Boolean
    checkLastTwoRow = True
    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    If getRangeGroupAndColumnName(currentSheet, rowNumber, columnNumber, groupName, columnName) = True Then
        If groupName <> lastGroupName And lastGroupName <> "" Then
            checkLastTwoRow = False
            Exit Function
        Else
            lastGroupName = groupName
        End If
    Else
        checkLastTwoRow = False
        Exit Function
    End If
End Function

Public Function isLteSectoreqmGrpId() As Boolean
    Dim actSheetName As String
    Dim cellsheet As Worksheet
    Dim mocName As String
    Dim attrName As String
    Dim sectorColumnName As String
    Dim columnIndex As Long
    isLteSectoreqmGrpId = False
    actSheetName = getResByKey("LTECellSheetName")
    For Each cellsheet In ThisWorkbook.Worksheets
        Call getSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
        sectorColumnName = findColumnFromRelationDef(actSheetName, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex > 0 And sectorColumnName <> "" Then
            isLteSectoreqmGrpId = True
            Exit Function
        End If
    Next
End Function

Public Function isLteCellBeamMode() As Boolean
    Dim actSheetName As String
    Dim cellsheet As Worksheet
    Dim mocName As String
    Dim attrName As String
    Dim sectorColumnName As String
    Dim columnIndex As Long
    isLteCellBeamMode = False
    actSheetName = getResByKey("LTECellSheetName")
    For Each cellsheet In ThisWorkbook.Worksheets
        Call getCellbeamModeMocNameAndAttrName(mocName, attrName)
        sectorColumnName = findColumnFromRelationDef(actSheetName, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex > 0 And sectorColumnName <> "" Then
            isLteCellBeamMode = True
            Exit Function
        End If
    Next
End Function


Public Function isUmtsSectoreqmPo() As Boolean
    Dim actSheetName As String
    Dim cellsheet As Worksheet
    Dim mocName As String
    Dim attrName As String
    Dim sectorColumnName As String
    Dim columnIndex As Long
    isUmtsSectoreqmPo = False
    actSheetName = getResByKey("UMTSCellSheetName")
    For Each cellsheet In ThisWorkbook.Worksheets
        Call getSectoreqmMocNameAndAttrName(mocName, attrName)
        sectorColumnName = findColumnFromRelationDef(actSheetName, mocName, attrName)
        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
        If columnIndex > 0 And sectorColumnName <> "" Then
            isUmtsSectoreqmPo = True
            Exit Function
        End If
    Next
End Function

Private Sub getCgpsBoraNoMap(ByRef ws As Worksheet)
    Dim groupName As String
    groupName = getCxugroupName
    If groupName = "" Then
        Exit Sub
    End If
    
    Set cGpsboardNoMap = New CMap
    
    Dim currentBoardStyleMappingDefData As CBoardStyleMappingDefData
    If boardStyleData Is Nothing Then Call initBoardStyleMappingDataPublic
    Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    Set currentBoardStyleMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
    
    Dim autoFillInColumnName As String
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    Dim boardNoColumLetter As String
    Dim boardNoString, eachBoardNo As String
    Dim boardNoTypeColumLetter As String
    Dim cxuTypeColumnName As String
    
    Dim boardNoRange As range
    Dim boardNoTypeRange As range
    
    Dim rowNumber As Long
    
    cxuTypeColumnName = getCxuTypeColum()
    
    
    autoFillInColumnName = currentBoardStyleMappingDefData.autoFillInColumnName
    If autoFillInColumnName <> "" Then
        Call getGroupNameStartAndEndRowNumber(groupName, groupNameStartRowNumber, groupNameEndRowNumber)
        boardNoColumLetter = currentBoardStyleMappingDefData.getColumnNamePositionLetter(autoFillInColumnName)
        boardNoTypeColumLetter = currentBoardStyleMappingDefData.getColumnNamePositionLetter(cxuTypeColumnName)
        
        For rowNumber = groupNameStartRowNumber + 2 To groupNameEndRowNumber
            If ws.Cells(rowNumber, 1) <> "RMV" Then
                Set boardNoRange = ws.range(boardNoColumLetter & rowNumber)
                Set boardNoTypeRange = ws.range(boardNoTypeColumLetter & rowNumber)
                
                eachBoardNo = Trim(boardNoRange.value)
                If eachBoardNo <> "" And (Trim(boardNoTypeRange.value) = "CGPS" Or Trim(boardNoTypeRange.value) = "RPCU") Then
                    boardNoString = boardNoString & eachBoardNo & ","
                End If
            End If
        Next rowNumber
        If boardNoString <> "" Then boardNoString = Left(boardNoString, Len(boardNoString) - 1)

        Call cGpsboardNoMap.SetAt(groupName, boardNoString)
    End If
    
End Sub

Private Function getCxuTypeColum() As String
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    getCxuTypeColum = ""
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mocName = "CXU" And attributeName = "CXUTYPE" Then
            getCxuTypeColum = mappingdefcolumnName
            Exit For
        End If
    Next
End Function

Private Sub fillterBoradNo(ByRef boardNoString As String, ByRef groupName As String)
    Dim cgpsbordNostring As String
    Dim boardNoStringArr() As String
    Dim cgpsboardNoStringArr() As String
    Dim tempboardNoStr As String
    Dim index As Long
    Dim boardNo As String
    
    cgpsbordNostring = ""

    If cGpsboardNoMap Is Nothing Then
        Exit Sub
    End If

    If cGpsboardNoMap.hasKey(groupName) = False Then
        Exit Sub
    Else
        boardNoString = boardNoString + ","
        boardNoStringArr = Split(boardNoString, ",")
        cgpsbordNostring = cGpsboardNoMap.GetAt(groupName)
        cgpsbordNostring = cgpsbordNostring + ","
        cgpsboardNoStringArr = Split(cgpsbordNostring, ",")

        For index = LBound(boardNoStringArr) To UBound(boardNoStringArr)
            boardNo = boardNoStringArr(index)
            If iscgpsBoardNo(boardNo, cgpsboardNoStringArr) = False Then
                tempboardNoStr = tempboardNoStr & boardNo & ","
            End If
        Next index
        If tempboardNoStr <> "" Then tempboardNoStr = Left(tempboardNoStr, Len(tempboardNoStr) - 1)
        boardNoString = tempboardNoStr
    End If

End Sub


Private Function iscgpsBoardNo(ByRef boradNo As String, ByRef cgpsboardNoStringArr() As String) As Boolean
   Dim cgpsbordNo As String
   Dim index As Long
   iscgpsBoardNo = False
   
   For index = LBound(cgpsboardNoStringArr) To UBound(cgpsboardNoStringArr)
        cgpsbordNo = cgpsboardNoStringArr(index)
        If boradNo = cgpsbordNo Then
           iscgpsBoardNo = True
           Exit Function
        End If
    Next index
End Function

Private Function getCxugroupName() As String
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    getCxugroupName = ""
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mocName = sheetDef.Cells(index, 4)
        If mocName = "CXU" Then
            getCxugroupName = mappingdefgroupName
            Exit For
        End If
    Next
End Function


