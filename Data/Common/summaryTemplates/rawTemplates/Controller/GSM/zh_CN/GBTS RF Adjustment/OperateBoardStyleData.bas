Attribute VB_Name = "OperateBoardStyleData"
Option Explicit

Public boardStyleData As CBoardStyleData
Public inAddProcessFlag As Boolean
Public currentSheet As Worksheet
Public currentCellValue As String
Public addBoardStyleMoiInProcess As CAddingBoardStyleMoc
Public moiRowsManager As CMoiRowsManager
Public boardNoManager As CBoardNoManager
Public boardStyleMappingDefMap As CMapValueObject
Public selectedGroupMappingDefData As CBoardStyleMappingDefData

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
    Dim newMoiRange As Range, eachCell As Range
    Set newMoiRange = moiRowsManager.getMoiRange
    For Each eachCell In newMoiRange
        '如果单元格不是灰化的，则置为正常底色，是分支控制的灰化，则不变
        If eachCell.Interior.colorIndex <> SolidColorIdx And eachCell.Interior.Pattern <> SolidPattern Then eachCell.Interior.colorIndex = colorIndex
    Next eachCell
End Sub

Public Sub addBoardStyleMoiCancelButton()
    On Error GoTo ErrorHandler
    Dim deletedRowsRange As Range
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

Public Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String, Optional ByRef ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Set ws = wb.Worksheets(sheetName)
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
Public Function findCertainValRowNumber(ByRef ws As Worksheet, ByVal columnletter As String, ByRef cellVal As Variant, Optional ByVal startRow As Long = 1)
    Dim currentCellVal As Variant
    Dim maxRowNumber As Long, k As Long
    maxRowNumber = ws.UsedRange.Rows.count
    findCertainValRowNumber = -1
    For k = startRow To maxRowNumber
        Dim maxColnum As Long
        maxColnum = ws.Range("XFD" & k).End(xlToLeft).column
        Dim index As Long
        For index = 1 To maxColnum
            currentCellVal = ws.Cells(k, index).value
            If currentCellVal = cellVal Then
                findCertainValRowNumber = k
                Exit For
            End If
        Next
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
    Application.GoTo Reference:=ws.Range(rangeName), Scroll:=scrollFlag
End Sub

Public Sub resetAddBoardStyleMoiInfo(ByRef ws As Worksheet)
    'inAddProcessFlag = False
    Set currentSheet = ws
End Sub

Private Function checkNeedFillInCellsFilled() As Boolean
    checkNeedFillInCellsFilled = True
        
    Dim emptyCell As Range
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
        currentSheet.Range(boardNoColumnLetter & rowNumber).value = tempBoardNo
    Next rowNumber

End Sub

Public Function getTempBoardNo(ByRef rowNumber As Long, ByRef sourceAttributeColumnLetterArr As Variant)
    Dim tempBoardNo As String
    tempBoardNo = ""
    Dim index As Long
    Dim attributeValue As String
    For index = LBound(sourceAttributeColumnLetterArr) To UBound(sourceAttributeColumnLetterArr)
        attributeValue = currentSheet.Range(sourceAttributeColumnLetterArr(index) & rowNumber).value
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

Public Sub boardStyleSelectionChange(ByRef ws As Worksheet, ByRef Target As Range)
    On Error GoTo ErrorHandler
    Dim groupName As String, columnName As String
    Dim rowNumber As Long, columnNumber As Long
    If Target.Rows.count <> 1 Or Target.Columns.count <> 1 Then Exit Sub
    
    rowNumber = Target.row
    columnNumber = Target.column
    
    Dim targetInRecordsRangeFlag As Boolean, targetIsInListBoxFlag As Boolean, targetInBoardNoFlag As Boolean
    targetInRecordsRangeFlag = getRangeGroupAndColumnName(ws, rowNumber, columnNumber, groupName, columnName)
        
    Dim referencedString As String
    Dim currentBoardStyleMappingDefData As CBoardStyleMappingDefData
    targetIsInListBoxFlag = getReferecedString(groupName, columnName, referencedString, currentBoardStyleMappingDefData) '判断选定的列是否需要增加自动下拉框
    
    targetInBoardNoFlag = judgeWhetherInBoardNoColumn(columnName, currentBoardStyleMappingDefData) '判断选定的列是否是BoardNo
    
    If targetIsInListBoxFlag = False And targetInBoardNoFlag = False Then
        Exit Sub '如果不是需要添加下拉列表的参数，则直接退出
    End If
    If targetInRecordsRangeFlag = False Then '如果不在数据范围内，则先将有效性清空，再退出
        'target.Validation.Delete
        Exit Sub
    End If
    
    If targetIsInListBoxFlag = True Then
        Call setBoardStyleListBoxRangeValidation(ws.name, groupName, columnName, referencedString, ws, Target)
    ElseIf targetInBoardNoFlag = True Then
        Call setBoardNoRangeValidation(Target)
    End If
    Exit Sub
ErrorHandler:
End Sub

Private Function judgeWhetherInBoardNoColumn(ByRef columnName As String, ByRef currentBoardStyleMappingDefData As CBoardStyleMappingDefData)
    If columnName = currentBoardStyleMappingDefData.autoFillInColumnName And columnName <> "" Then
        judgeWhetherInBoardNoColumn = True
    Else
        judgeWhetherInBoardNoColumn = False
    End If
End Function

Private Sub setBoardNoRangeValidation(ByRef Target As Range)
    'target.Offset(0, 1).Select
    With Target.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation '只给提示信息有效性设置
        .inputTitle = getResByKey("ForbiddenEditTitle") '输入提示Title
        .inputMessage = getResByKey("ForbiddenEditContent") '输入提示内容
        .ShowInput = True 'True显示输入提示信息，False不显示输入提示信息
        .ShowError = False 'True不允许输入非有效性值，False允许输入
    End With
End Sub

Public Sub setBoardStyleListBoxRangeValidation(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, _
    ByRef referencedString As String, ByRef sheet As Worksheet, ByRef Target As Range)
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
    With Target.Validation
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
    If boardStyleData Is Nothing Then Call initBoardStyleMappingDataPublic
    Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    Set currentBoardStyleMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
    
    referencedString = currentBoardStyleMappingDefData.getListBoxReferenceValue(columnName)
    If referencedString = "" Then
        getReferecedString = False
        Exit Function
    End If
    
    Call initBoardNoManagerPublic
    Dim mocNameBoardNoArr As Variant
    mocNameBoardNoArr = Split(referencedString, "-")
    referencedString = getBoardNoString(CStr(mocNameBoardNoArr(0)))
    'Debug.Print referencedString
    
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
    
    maxRowNumber = ws.UsedRange.Rows.count
    If rowNumber > maxRowNumber Then rowNumber = maxRowNumber
    
    For k = rowNumber To 1 Step -1
        If k = 1 Then Exit For
        If rowIsBlank(ws, k - 1) = True And rowIsBlank(ws, k) = False Then Exit For
    Next k
    getGroupRowNumber = k
End Function

Private Function getNextGroupRowNumber(ByRef ws As Worksheet, ByVal rowNumber) As Long
    Dim nextGroupRowNumber As Long
    nextGroupRowNumber = -1
    Dim maxRowNumber As Long, k As Long
    
    maxRowNumber = ws.UsedRange.Rows.count
    
    For k = rowNumber To maxRowNumber
        If rowIsBlank(ws, k) = True And rowIsBlank(ws, k + 1) = False Then
            nextGroupRowNumber = k + 1
            Exit For
        End If
    Next k
    
    '如果是-1，说明是最后一个分组，只能用单元格是否有边框来判断最大行了
    If nextGroupRowNumber = -1 Then
        Dim predefinedMaxRowNumber As Long
        predefinedMaxRowNumber = Application.WorksheetFunction.min(rowNumber + 2000, maxRowNumber) '防止最后一个对象的边框一直设置到1048576，设置一个2000的行数最大限制
    
        For k = rowNumber To predefinedMaxRowNumber
            If rangeHasBorder(ws.Rows(k)) Then
                nextGroupRowNumber = k
            Else
                Exit For
            End If
        Next k
        
        nextGroupRowNumber = nextGroupRowNumber + 2 '为了与正常分组的最大值一致，这里加2
    End If
    
    getNextGroupRowNumber = nextGroupRowNumber
End Function

Public Function getRangeGroupAndColumnName(ByRef ws As Worksheet, ByVal rowNumber As Long, ByVal columnNumber As Long, _
    ByRef groupName As String, ByRef columnName As String, Optional ByVal deleteBoardStyleFlag As Boolean = False) As Boolean
    
    Dim groupRowNumber As Long
    groupRowNumber = getGroupRowNumber(ws, rowNumber)
    groupName = ws.Range(getColumnNameFromColumnNum(columnNumber) & groupRowNumber).value
    columnName = ws.Cells(groupRowNumber + 1, columnNumber).value
    
    Dim groupMaxRow As Long
    groupMaxRow = getNextGroupRowNumber(ws, rowNumber) - 2
'    If deleteBoardStyleFlag = False Then
'        groupMaxRow = groupRowNumber + getCurrentRegionRowsCount(ws, groupRowNumber)
'    Else
'        groupMaxRow = getNextGroupRowNumber(ws, rowNumber) - 2
'    End If

    'groupMaxRow<0说明选择了最后一个分组的超出边框的行
    If (rowNumber > groupRowNumber + 1 And rowNumber <= groupMaxRow) Then
        getRangeGroupAndColumnName = True
        If rowNumber = groupMaxRow And rowIsBlank(ws, rowNumber + 1) = False Then
            getRangeGroupAndColumnName = False
        End If
    Else
        getRangeGroupAndColumnName = False
    End If
End Function

Public Function getRangeGroupAndColumnName1(ByRef ws As Worksheet, ByVal rowNumber As Long, ByVal columnNumber As Long, _
    ByRef groupName As String, ByRef columnName As String, Optional ByVal deleteBoardStyleFlag As Boolean = False) As Boolean
    
    '这里要校验：选择范围不能跨GroupName或ColumnName
    
    Dim cellRange As Range
    Set cellRange = ws.Range(getColStr(columnNumber) & rowNumber)
    
    Dim groupRowNumber As Long
    groupRowNumber = getGroupRowNumber(ws, rowNumber)
    
    Dim groupMaxRow As Long
    groupMaxRow = getNextGroupRowNumber(ws, rowNumber) - 2
       
    If (rowNumber > groupRowNumber + 1 And rowNumber <= groupMaxRow) Then
        getRangeGroupAndColumnName1 = True
        If rowNumber = groupMaxRow And rowIsBlank(ws, rowNumber + 1) = False Then
            getRangeGroupAndColumnName1 = False
            Exit Function
        End If
        
        Call getGroupAndColumnName(ws, cellRange, groupName, columnName)
    Else
        getRangeGroupAndColumnName1 = False
    End If
End Function

Public Function rowIsBlank(ByRef ws As Worksheet, ByRef rowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.Range("A" & rowNumber & ":IV" & rowNumber)) = 256 Then
        rowIsBlank = True
    Else
        rowIsBlank = False
    End If
End Function

Public Function rangeHasBorder(ByRef certainRange As Range) As Boolean
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
        rowNumber = findCertainValRowNumber(ws, "B", groupName, initialRowNumber)
        If rowNumber = -1 Then Exit Do
        If rowNumber = 1 Then
            lastRowEmptyFlag = True
        ElseIf rowIsBlank(ws, rowNumber - 1) = True Then
            lastRowEmptyFlag = True
        End If
        initialRowNumber = rowNumber + 1
    Loop
    getGroupNameStartRowNumber = rowNumber
End Function

Public Sub getGroupNameStartAndEndRowNumber(ByRef groupName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    groupNameStartRowNumber = getGroupNameStartRowNumber(currentSheet, groupName)
    'groupNameEndRowNumber = findCertainValRowNumber(currentSheet, "A", "", groupNameStartRowNumber)
    'groupNameEndRowNumber = groupNameStartRowNumber + currentSheet.Range("A" & groupNameStartRowNumber).CurrentRegion.rows.count - 1
    groupNameEndRowNumber = groupNameStartRowNumber + getCurrentRegionRowsCount(currentSheet, groupNameStartRowNumber) - 1
    
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
    Dim autoFillInColumnName As String, columnletter As String
    Dim boardNoRanges As Range
    autoFillInColumnName = selectedGroupMappingDefData.autoFillInColumnName
    If autoFillInColumnName <> "" Then
        columnletter = selectedGroupMappingDefData.getColumnNamePositionLetter(autoFillInColumnName)
        Set boardNoRanges = currentSheet.Range(currentSheet.Range(columnletter & startRowNumber), currentSheet.Range(columnletter & endRowNumber))
        boardNoRanges.ClearContents
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
    Dim groupRowNumber As Long, nextGroupRowNumber As Long, dataAreaRowsCount As Long
    Dim groupMaxRow As Long
    groupRowNumber = getGroupNameStartRowNumber(ws, groupName)
    nextGroupRowNumber = getNextGroupRowNumber(ws, groupRowNumber)
    If nextGroupRowNumber = -1 Then '当前的group是最后一个分组
        groupMaxRow = groupRowNumber + getCurrentRegionRowsCount(ws, groupRowNumber) - 1
        dataAreaRowsCount = groupMaxRow - (groupRowNumber + 2) + 1
    Else
        dataAreaRowsCount = (nextGroupRowNumber - 2) - (groupRowNumber + 2) + 1
    End If
    
    If rowCollection.count < dataAreaRowsCount Then Exit Sub '如果删除的行数小于数据区域行数，则无需操作
    
    Dim newRowNumber As Long
    newRowNumber = rowCollection(rowCollection.count) + 1
    ws.Rows(groupRowNumber + 2).Copy
    ws.Rows(newRowNumber).Insert Shift:=xlDown
    ws.Range(ws.Cells(newRowNumber, 1), ws.Cells(newRowNumber, ws.UsedRange.Columns.count)).ClearContents
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
    currentSheet.Range("A" & groupNameStartRowNumber).Select
End Sub

Private Sub deleteRowsWithBoardNoColumn(ByRef rowCollection As Collection, ByRef boardNoColumnLetter As String, ByRef outputString As String)
    Dim referenceRangeManager As New CReferenceRangeManager
    Call referenceRangeManager.generateBoardNoReferenceAddressMap

    Dim boardNo As String, boardNoReferenceAddressString As String
    Dim rowNumber As Variant
    For Each rowNumber In rowCollection
        boardNo = currentSheet.Range(boardNoColumnLetter & rowNumber).value
        boardNoReferenceAddressString = referenceRangeManager.getReferenceAddressString(boardNo)
        If boardNoReferenceAddressString <> "" Then
            Call clearReferenceAddress(boardNoReferenceAddressString, rowCollection)
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

Private Sub clearReferenceAddress(ByRef referenceAddressString As String, ByRef rowCollection As Collection)
    Dim referenceAddressArr As Variant, referenceAddress As Variant
    Dim newReferenceAddressString As String, newReferenceAddress As String
    
    newReferenceAddressString = referenceAddressString
    referenceAddressArr = Split(referenceAddressString, ",")
    
    For Each referenceAddress In referenceAddressArr
        If getNewAddressAfterDeletion(CStr(referenceAddress), newReferenceAddress, rowCollection) = True Then
            newReferenceAddressString = Replace(newReferenceAddressString, CStr(referenceAddress), newReferenceAddress)
        End If
        currentSheet.Range(referenceAddress).ClearContents '清空引用BoardNo的单元格
    Next referenceAddress
    referenceAddressString = newReferenceAddressString
End Sub

Private Function getNewAddressAfterDeletion(ByRef oldAddress As String, ByRef newAddress As String, ByRef rowCollection As Collection) As Boolean
    getNewAddressAfterDeletion = False
    Dim rowIndex As Variant
    Dim numberOfRowsToShiftUp As Long
    numberOfRowsToShiftUp = 0
    For Each rowIndex In rowCollection
        If Range(oldAddress).row > rowIndex Then
            numberOfRowsToShiftUp = numberOfRowsToShiftUp + 1
            getNewAddressAfterDeletion = True
        End If
    Next rowIndex
    newAddress = Range(oldAddress).Offset(-numberOfRowsToShiftUp, 0).address(False, False)
End Function

Private Sub deleteRowsWithoutBoardNoColumn(ByRef rowCollection As Collection)
    Dim rowNumber As Variant
    Dim multiRowsDeletionString As String
    For Each rowNumber In rowCollection
        If multiRowsDeletionString = "" Then
            multiRowsDeletionString = rowNumber & ":" & rowNumber
        Else
            multiRowsDeletionString = multiRowsDeletionString & "," & rowNumber & ":" & rowNumber
        End If
    Next rowNumber
    currentSheet.Range(multiRowsDeletionString).Delete
End Sub

Private Function checkSelectRanges(ByRef rowCollection As Collection, ByRef groupName As String) As Boolean
    checkSelectRanges = True
    Dim rowRange As Range
    Dim selectionRange As Range
    Set selectionRange = Selection
    
    Dim rowNumber As Long, columnNumber As Long
    Dim lastGroupName As String, columnName As String
        
    Dim startColnumLetter As Long
    startColnumLetter = 1
    If isOperationExcel Then
        startColnumLetter = 2
    End If
    
    columnNumber = 1
    For Each rowRange In selectionRange.Rows
        rowNumber = rowRange.row
        rowCollection.Add Item:=rowNumber, key:=CStr(rowNumber)
        If checkLastTwoRow(rowNumber, startColnumLetter, groupName, columnName, lastGroupName) = False Then
            checkSelectRanges = False
            Exit Function
        End If
    Next rowRange
End Function

Private Function checkLastTwoRow(ByRef rowNumber As Long, ByRef columnNumber As Long, ByRef groupName As String, _
    ByRef columnName As String, ByRef lastGroupName As String) As Boolean
    checkLastTwoRow = True
    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    
    '获取该行最后一个组名
    lastGroupName = getLastGroupName(currentSheet, rowNumber)
    
    If lastGroupName = "" Then
        checkLastTwoRow = False
        Exit Function
    End If
    
    If getRangeGroupAndColumnName(currentSheet, rowNumber, columnNumber, groupName, columnName, True) = True Then
        If groupName <> lastGroupName Then
            groupName = lastGroupName
            Exit Function
        End If
    Else
        checkLastTwoRow = False
        Exit Function
    End If
End Function

Private Function getLastGroupName(ByRef currentSheet As Worksheet, ByVal rowNumber As Long) As String
    Dim maxColumn As Long
    Dim index As Long
    
    Dim groupRowNumber As Long
    groupRowNumber = getGroupRowNumber(currentSheet, rowNumber)
    
    getLastGroupName = ""
    maxColumn = currentSheet.Range("XFD" + CStr(groupRowNumber)).End(xlToLeft).column
    
    If maxColumn < 0 Then
        Exit Function
    End If
    
    For index = maxColumn To 1 Step -1
        If currentSheet.Cells(groupRowNumber, index).value <> "" Then
            getLastGroupName = currentSheet.Cells(groupRowNumber, index).value
            Exit For
        End If
    Next
End Function

