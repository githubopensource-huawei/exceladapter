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
Public baseStationData As CBaseStationData
Public allBoardStyleData As CAllBoardStyleData
Public boardStyleNeMap As CMap
Public neBoardStyleMap As CMap
Public cGpsboardNoMap As CMap
Public sourceBoardStyleDataMap As New CMap
Public delSourceBoardStyleDataMap As New CMapValueObject
Public migrationCount As Long

Public addBoardStyleButtons As CAddBoardStyleButtons
Public deleteBoardStyleButtons As CDeleteBoardStyleButtons
Public Const NewMoiRangeColorIndex As Long = 43 '浅绿，新增Moi行底色
Public Const NeedFillInRangeColorIndex As Long = 33 '浅蓝，必填单元格底色
Public Const NormalRangeColorIndex As Long = -4142 '白色，正常单元格底色
Public Const BoardNoDelimiter As String = "_"

Public Const PublicMaxRowNumber As Long = 2000
'Finish按钮点击后
Public Sub addBoardStyleMoiFinishButton()
    On Error GoTo ErrorHandler
    Dim maxColumnNumber As Long
    Dim groupName As String
    migrationCount = migrationCount + 1
    groupName = selectedGroupMappingDefData.groupName
    Dim needFillinColunmEmptyFlag As Boolean
    needFillinColunmEmptyFlag = False
    'maxColumnNumber = selectedGroupMappingDefData.totalColumnNumber
    
    Dim columnNamePositionLetterMap As CMap
    Dim sourceNeNameCol As String
    Dim sourceNeNamePositionLetter As String
    sourceNeNameCol = getResByKey("SOURCEBTSNAME")
    Set columnNamePositionLetterMap = selectedGroupMappingDefData.columnNamePositionLetterMap
    sourceNeNamePositionLetter = columnNamePositionLetterMap.GetAt(sourceNeNameCol)
    
    Call finishCopyMigrationRec(needFillinColunmEmptyFlag, sourceNeNamePositionLetter)
    
    If needFillinColunmEmptyFlag = True Then Exit Sub
    
    If checkNeedFillInCellsFilled = False Then Exit Sub

    
    Call makeAutoFillInColumnNameValue(groupName)
    Call selectCertainCell(currentSheet, sourceNeNamePositionLetter & moiRowsManager.groupNameRowNumber)
    Call setNewRangesBackgroundColour(NormalRangeColorIndex)
    Call copySourceFillInColumnNameValue
    

    Dim columnName As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    groupNameStartRowNumber = moiRowsManager.startRowNumber
    groupNameEndRowNumber = moiRowsManager.endRowNumber
        
    Call getGroupAndColumnName(ws, ws.Range(sourceNeNamePositionLetter & groupNameStartRowNumber), groupName, columnName)
    Call clearSourceBoardNoRangesBoxList(groupName, groupNameStartRowNumber, groupNameEndRowNumber)
        '加底色
        'Call setMigrationRecbackColor(ws, groupNameStartRowNumber, groupNameEndRowNumber, maxColumnNumber)
    Call refreshMigrationSourceData(ws)
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
Public Function findCertainValRowNumber(ByRef ws As Worksheet, ByVal columnLetter As String, ByRef cellVal As Variant, Optional ByVal startRow As Long = 1)
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

Private Sub makeAutoFillInColumnNameValue(ByVal groupName As String)
    Dim boardNoName As String
    boardNoName = selectedGroupMappingDefData.autoFillInColumnName
    If boardNoName = "" Then Exit Sub
    
    Dim startRowNumber As Long, endRowNumber As Long
    startRowNumber = moiRowsManager.startRowNumber
    endRowNumber = moiRowsManager.endRowNumber
    
    Call clearBoardNoRanges(startRowNumber, endRowNumber, groupName)
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
        currentSheet.Range(boardNoColumnLetter & rowNumber).Interior.colorIndex = NullPattern
        currentSheet.Range(boardNoColumnLetter & rowNumber).Interior.Pattern = NullPattern
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
    
    Call getCgpsBoraNoMap(ws)
    
    Dim targetInRecordsRangeFlag As Boolean, targetIsInListBoxFlag As Boolean, targetInBoardNoFlag As Boolean
    targetInRecordsRangeFlag = getRangeGroupAndColumnName(ws, rowNumber, columnNumber, groupName, columnName)
        
    Dim referencedString As String
    Dim currentBoardStyleMappingDefData As CBoardStyleMappingDefData
    targetIsInListBoxFlag = getReferecedString(groupName, columnName, referencedString, currentBoardStyleMappingDefData) '判断选定的列是否需要增加自动下拉框
    
    targetInBoardNoFlag = judgeWhetherInBoardNoColumn(columnName, currentBoardStyleMappingDefData) '判断选定的列是否是BoardNo
    
    If targetIsInListBoxFlag = True Then
        Call setBoardStyleListBoxRangeValidation(ws.name, groupName, columnName, referencedString, ws, Target)
    ElseIf targetInBoardNoFlag = True Then
        Call setBoardNoRangeValidation(Target)
    End If
    '++++++++++++++++++++++
    If Target.Validation.InCellDropdown Then
        Call updateSourceBoradNobyNeName(ws, Target)
    End If
    
    If targetIsInListBoxFlag = False And targetInBoardNoFlag = False Then
        Exit Sub '如果不是需要添加下拉列表的参数，则直接退出
    End If
    If targetInRecordsRangeFlag = False Then '如果不在数据范围内，则先将有效性清空，再退出
        'target.Validation.Delete
        Exit Sub
    End If
    
'    If targetIsInListBoxFlag = True Then
'        Call setBoardStyleListBoxRangeValidation(ws.name, groupName, columnName, referencedString, ws, Target)
'    ElseIf targetInBoardNoFlag = True Then
'        Call setBoardNoRangeValidation(Target)
'    End If
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
    
    '去除Source字段下拉列表字段值重复的情况DTS2016031703941
    Dim valueArr
    Dim valueColl As New Collection
    Dim value As Variant
    Dim i As Integer
    Dim j As Integer
    
    valueArr = Split(referencedString, ",")
        
    For Each value In valueArr
        valueColl.Add value
    Next
        
    referencedString = ""
        
    If valueColl.count >= 2 Then
        For i = 1 To valueColl.count
            For j = i + 1 To valueColl.count
                If valueColl.Item(j) = valueColl.Item(i) Then
                    valueColl.Remove (j)
                    j = j - 1
                End If
                If j = valueColl.count Then
                    Exit For
                End If
            Next
        Next
    End If
    For i = 1 To valueColl.count
        If i = 1 Then
            referencedString = valueColl.Item(i)
        Else
            referencedString = referencedString + "," + valueColl.Item(i)
        End If
    Next
    
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
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=referencedString
            .ShowError = False
        Else
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=" "
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
End Function

Public Function getReferencedStringByInputString(ByRef referencedString As String)
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
    
    '这里要校验：选择范围不能跨GroupName或ColumnName
    
    Dim cellRange As Range
    Set cellRange = ws.Range(getColStr(columnNumber) & rowNumber)
    
    Dim groupRowNumber As Long
    groupRowNumber = getGroupRowNumber(ws, rowNumber)
    
    Dim groupMaxRow As Long
    groupMaxRow = getNextGroupRowNumber(ws, rowNumber) - 2
       
    If (rowNumber > groupRowNumber + 1 And rowNumber <= groupMaxRow) Then
        getRangeGroupAndColumnName = True
        If rowNumber = groupMaxRow And rowIsBlank(ws, rowNumber + 1) = False Then
            getRangeGroupAndColumnName = False
            Exit Function
        End If
        
        Call getGroupAndColumnName(ws, cellRange, groupName, columnName)
    Else
        getRangeGroupAndColumnName = False
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

Public Sub getBoardStyleGroupNameStartAndEndRowNumber(ByRef ws As Worksheet, ByRef groupName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
    groupNameStartRowNumber = getGroupNameStartRowNumber(ws, groupName)
    'groupNameEndRowNumber = findCertainValRowNumber(currentSheet, "A", "", groupNameStartRowNumber)
    'groupNameEndRowNumber = groupNameStartRowNumber + currentSheet.Range("A" & groupNameStartRowNumber).CurrentRegion.rows.count - 1
    groupNameEndRowNumber = groupNameStartRowNumber + getCurrentRegionRowsCount(ws, groupNameStartRowNumber) - 1
    
End Sub

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

Public Sub clearBoardNoRanges(ByVal startRowNumber As Long, ByVal endRowNumber As Long, ByVal groupName As String)
    Dim autoFillInColumnName As String, columnLetter As String
    Dim boardNoRanges As Range
    Set selectedGroupMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
    autoFillInColumnName = selectedGroupMappingDefData.autoFillInColumnName
    If autoFillInColumnName <> "" Then
        columnLetter = selectedGroupMappingDefData.getColumnNamePositionLetter(autoFillInColumnName)
        Set boardNoRanges = currentSheet.Range(currentSheet.Range(columnLetter & startRowNumber), currentSheet.Range(columnLetter & endRowNumber))
        boardNoRanges.ClearContents
    End If
End Sub

Public Sub clearAllBoardNoRanges(ByVal startRowNumber As Long, ByVal endRowNumber As Long, ByVal maxColumnNumber As Long)
    Dim autoFillInColumnName As String, columnLetter As String
    Dim boardNoRanges As Range
    Dim autoCopyFillInColumnName As String, sourceColumnLetter As String
    Dim sourceBoardNoRanges As Range
    Dim autoCopyFillInColumnNameCol As Collection
    Dim currentNeName As String
    Dim groupName As String
    Dim columnName As String
    Dim IndexCol As Integer
    
    For IndexCol = 1 To maxColumnNumber
        columnLetter = getColStr(IndexCol)
        Set boardNoRanges = currentSheet.Range(currentSheet.Range(columnLetter & startRowNumber), currentSheet.Range(columnLetter & endRowNumber))
        boardNoRanges.ClearContents
    Next
    
    autoFillInColumnName = selectedGroupMappingDefData.autoFillInColumnName
    If autoFillInColumnName <> "" Then
        columnLetter = selectedGroupMappingDefData.getColumnNamePositionLetter(autoFillInColumnName)
        Set boardNoRanges = currentSheet.Range(currentSheet.Range(columnLetter & startRowNumber), currentSheet.Range(columnLetter & endRowNumber))
        boardNoRanges.ClearContents
        boardNoRanges.Interior.colorIndex = SolidColorIdx
        boardNoRanges.Interior.Pattern = SolidPattern
    End If
    
    Dim autoCopyFillInColumnNameVar As Variant
    Set autoCopyFillInColumnNameCol = selectedGroupMappingDefData.autoFillInSourceColumnName
    
    For Each autoCopyFillInColumnNameVar In autoCopyFillInColumnNameCol
        autoCopyFillInColumnName = autoCopyFillInColumnNameVar
        If autoCopyFillInColumnName <> "" Then
            sourceColumnLetter = selectedGroupMappingDefData.getColumnNamePositionLetter(autoCopyFillInColumnName)
            Set sourceBoardNoRanges = currentSheet.Range(currentSheet.Range(sourceColumnLetter & startRowNumber), currentSheet.Range(sourceColumnLetter & endRowNumber))
            sourceBoardNoRanges.ClearContents
        End If
    Next autoCopyFillInColumnNameVar
    
    currentNeName = ""
    If boardStyleNeMap.hasKey(currentSheet.name) Then
        currentNeName = boardStyleNeMap.GetAt(currentSheet.name)
    End If
    
    Dim columnNamePositionLetterMap As CMap
    Dim sourceNeNameCol As String
    Dim sourceNeNamePositionLetter As String
    sourceNeNameCol = getResByKey("SOURCEBTSNAME")
    Set columnNamePositionLetterMap = selectedGroupMappingDefData.columnNamePositionLetterMap
    sourceNeNamePositionLetter = columnNamePositionLetterMap.GetAt(sourceNeNameCol)
    
    Call getGroupAndColumnName(currentSheet, currentSheet.Range(sourceNeNamePositionLetter & startRowNumber), groupName, columnName)
    
    If isSourceNeNameColum(groupName, columnName) Then
        Set sourceBoardNoRanges = currentSheet.Range(currentSheet.Range(sourceNeNamePositionLetter & startRowNumber), currentSheet.Range(sourceNeNamePositionLetter & endRowNumber))
        sourceBoardNoRanges.ClearContents
        sourceBoardNoRanges.value = currentNeName
    End If
End Sub
Public Sub addSourceBoardNoRangesBoxList(ByRef groupName As String, ByVal startRowNumber As Long, ByVal endRowNumber As Long)

    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    If allBoardStyleData Is Nothing Then Call initAllBoardStyleDataPublic
    
    Call allBoardStyleData.addSourceColumBoxList(currentSheet, groupName, startRowNumber, endRowNumber)

End Sub
Public Sub clearSourceBoardNoRangesBoxList(ByRef groupName As String, ByVal startRowNumber As Long, ByVal endRowNumber As Long)

    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    If allBoardStyleData Is Nothing Then Call initAllBoardStyleDataPublic
    
    Call allBoardStyleData.clearSourceColumBoxList(currentSheet, groupName, startRowNumber, endRowNumber)

End Sub
Public Sub initBaseStationDataPublic()
    If baseStationData Is Nothing Then Set baseStationData = New CBaseStationData
    Call baseStationData.init
    Set boardStyleNeMap = baseStationData.boardStyleNeMap
    Set neBoardStyleMap = baseStationData.neBoardStyleMap
End Sub
Public Sub initAllBoardStyleDataPublic()
    If allBoardStyleData Is Nothing Then Set allBoardStyleData = New CAllBoardStyleData
    Call allBoardStyleData.initBoardStyleDataMap
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
            'Call clearReferenceAddress(boardNoReferenceAddressString, rowCollection)
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
            '如果不是多list的单元格，那就是下拉框引用，直接清空
            currentSheet.Range(referenceAddress).ClearContents '清空引用Board No.的单元格
        Else
            '如果是多list引用的单元格，就把这个单板编号值从这个单元格的值中删除
'            boardNoStringValue = currentSheet.range(referenceAddress)
'            newBoardNoStringValue = removeBoardNoFromString(boardNoStringValue, boardNo)
'            currentSheet.range(referenceAddress).value = newBoardNoStringValue
        End If
    Next referenceAddress
    referenceAddressString = newReferenceAddressString
End Sub

'Private Function removeBoardNoFromString(ByRef boardNoStringValue As String, ByRef boardNo As String) As String
'    'On Error Resume Next
'    Dim boardNoArray As Variant, eachBoardNo As Variant
'    Dim boardNoCol As New Collection
'    boardNoArray = Split(boardNoStringValue, BasebandReferenceBoardNoDelimeter)
'    For Each eachBoardNo In boardNoArray
'        '如果不是要清空的单板编号，则把它放到容器中
'        If eachBoardNo <> boardNo Then
'            boardNoCol.Add eachBoardNo
'        End If
'    Next eachBoardNo
'    removeBoardNoFromString = getConnectedStringFromCol(boardNoCol, BasebandReferenceBoardNoDelimeter)
'End Function

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

'Private Sub deleteRowsWithoutBoardNoColumn(ByRef rowCollection As Collection)
'    Dim rowNumber As Variant
'    Dim multiRowsDeletionString As String
'    For Each rowNumber In rowCollection
'        If multiRowsDeletionString = "" Then
'            multiRowsDeletionString = rowNumber & ":" & rowNumber
'        Else
'            multiRowsDeletionString = multiRowsDeletionString & "," & rowNumber & ":" & rowNumber
'        End If
'    Next rowNumber
'    currentSheet.Range(multiRowsDeletionString).Delete
'End Sub

Private Sub deleteRowsWithoutBoardNoColumn(ByRef rowCollection As Collection)
    Dim rowNumber As Variant
    Dim multiRowsDeletionString As String
    Dim lastMatchRowNumber As Long
    lastMatchRowNumber = -1
    For Each rowNumber In rowCollection
        Call makeRowsString(multiRowsDeletionString, lastMatchRowNumber, CLng(rowNumber))
    Next rowNumber
    currentSheet.Range(multiRowsDeletionString).Delete
End Sub

'制作拼接待拷贝行字符串代码
Private Sub makeRowsString(ByRef rowsString As String, ByRef lastMatchRowNumber As Long, ByRef rowIndex As Long)
    If rowsString <> "" Then
        If lastMatchRowNumber = rowIndex - 1 Then
            '如果两个行时相连的，则直接将rowsString的最后一个行号替换成新行，提高效率
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
    Dim rowRange As Range
    Dim selectionRange As Range
    Set selectionRange = Selection
    
    Dim rowNumber As Long, columnNumber As Long
    Dim lastGroupName As String, columnName As String

    columnNumber = 1
    For Each rowRange In selectionRange.Rows
        rowNumber = rowRange.row
        rowCollection.Add Item:=rowNumber, key:=CStr(rowNumber)
        If checkLastTwoRow(rowNumber, 2, groupName, columnName, lastGroupName) = False Then
            checkSelectRanges = False
            Exit Function
        End If
    Next rowRange
End Function

Private Function checkLastTwoRow(ByRef rowNumber As Long, ByRef columnNumber As Long, ByRef groupName As String, _
    ByRef columnName As String, ByRef lastGroupName As String) As Boolean
    checkLastTwoRow = True
    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    If getRangeGroupAndColumnName(currentSheet, rowNumber, columnNumber, groupName, columnName, True) = True Then
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

Private Sub copySourceFillInColumnNameValue()
    Dim boardNoNameCol As Collection
    Dim copyboardNoNameCol As Collection
    Set boardNoNameCol = selectedGroupMappingDefData.autoFillInSourceColumnName
    Set copyboardNoNameCol = selectedGroupMappingDefData.copyFillInSourceColumnName
    
    If boardNoNameCol.count = 0 Then Exit Sub
    
    Dim sourceBoardNoName As Variant
    Dim sourceBoardNoNameStr As String
    Dim sourceboardNoColumnLetter As String
    Dim countnum As Long
    Dim copyName As String
    Dim copyboardNoColumnLetter As String
    countnum = 0
    
    Dim startRowNumber As Long, endRowNumber As Long
    startRowNumber = moiRowsManager.startRowNumber
    endRowNumber = moiRowsManager.endRowNumber
    
    For Each sourceBoardNoName In boardNoNameCol
        sourceBoardNoNameStr = sourceBoardNoName
        copyName = getcopyboardNoName(countnum, copyboardNoNameCol)
        sourceboardNoColumnLetter = selectedGroupMappingDefData.getColumnNamePositionLetter(sourceBoardNoNameStr)
        copyboardNoColumnLetter = selectedGroupMappingDefData.getColumnNamePositionLetter(copyName)
        Call copyBoardNoRanges(sourceboardNoColumnLetter, startRowNumber, endRowNumber, copyboardNoColumnLetter)
        countnum = countnum + 1
    Next sourceBoardNoName
    
End Sub

Private Function getcopyboardNoName(ByRef countnum As Long, ByRef copyboardNoNameCol As Collection) As String
    Dim copyboardNoNameVar As Variant
    Dim copyboardNoName As String
    Dim num  As Long
    num = 0
    
    For Each copyboardNoNameVar In copyboardNoNameCol
        If num = countnum Then
            copyboardNoName = copyboardNoNameVar
            getcopyboardNoName = copyboardNoName
            Exit Function
        End If
        num = num + 1
    Next copyboardNoNameVar
    getcopyboardNoName = ""
End Function

Public Sub copyBoardNoRanges(ByRef boardNoColumnLetter As String, ByRef startRowNumber As Long, ByRef endRowNumber As Long, ByRef copyboardNoColumnLetter As String)
    Dim tempBoardNo As String
    Dim rowNumber As Long
    For rowNumber = startRowNumber To endRowNumber
        If currentSheet.Range(boardNoColumnLetter & rowNumber).value = "" Then
            currentSheet.Range(boardNoColumnLetter & rowNumber).value = currentSheet.Range(copyboardNoColumnLetter & rowNumber).value
        End If
    Next rowNumber

End Sub

Public Sub finishCopyMigrationRec(ByRef needFillinColunmEmptyFlag As Boolean, ByRef sourceNeNamePositionLetter As String)
    Dim autoCopySourceNameCol As Collection
    Dim columnNamePositionLetterMap As CMap

    Dim ws As Worksheet
    Dim sheetName As String
    Dim groupName As String
    Dim columnName As String

    Set ws = ThisWorkbook.ActiveSheet
    sheetName = ws.name
    
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    groupNameStartRowNumber = moiRowsManager.startRowNumber
    groupNameEndRowNumber = moiRowsManager.endRowNumber
    
    Call getGroupAndColumnName(ws, ws.Range(sourceNeNamePositionLetter & groupNameStartRowNumber), groupName, columnName)
    
    If isSourceNeNameColum(groupName, columnName) = False Then
        Exit Sub
    End If
    
    
    Dim currentNeName As String
    If baseStationData Is Nothing Then Call initBaseStationDataPublic
    Set boardStyleNeMap = baseStationData.boardStyleNeMap
    Set neBoardStyleMap = baseStationData.neBoardStyleMap
    currentNeName = boardStyleNeMap.GetAt(sheetName)
    

    Dim startColnumNumber As Long
    Dim endColnumNumber As Long
    Dim startcolumnLetter As String
    Dim endColcolumnLetter As String
    
    
    Set columnNamePositionLetterMap = selectedGroupMappingDefData.columnNamePositionLetterMap
    Set autoCopySourceNameCol = selectedGroupMappingDefData.autoFillInSourceColumnName
    
    groupName = selectedGroupMappingDefData.groupName
    
    'startColnumNumber = autoCopySourceNameCol.count + 2
    startColnumNumber = 1
    endColnumNumber = selectedGroupMappingDefData.totalColumnNumber
    startcolumnLetter = getColStr(startColnumNumber)
    endColcolumnLetter = getColStr(endColnumNumber)
    

    Dim sourceNeName As String
    Dim sourceAttrBoardNo As New CMap
    Dim rowNumber As Long
    Dim i As Long
    i = 0
    Dim needCopyReccord As Boolean
    
    For rowNumber = groupNameStartRowNumber To groupNameEndRowNumber
        sourceNeName = ws.Range(sourceNeNamePositionLetter & rowNumber).value
        If sourceNeName <> currentNeName And sourceNeName <> "" And neBoardStyleMap.hasKey(sourceNeName) Then
        
            Call getsourceAttrBoardNo(ws, rowNumber, autoCopySourceNameCol, columnNamePositionLetterMap, sourceAttrBoardNo)
            needCopyReccord = isNeedCopyRecond(ws, rowNumber, columnNamePositionLetterMap)
            If needCopyReccord = False Then
                If checkNeedFillInCellsFilled = False Then
                    needFillinColunmEmptyFlag = True
                    Exit Sub
                End If
            End If
            
            Call copyMigrationBoardStyleLinewtihoutKey(ws, groupName, rowNumber, sourceNeName, sourceAttrBoardNo, startColnumNumber, endColnumNumber)
            
        End If
    Next rowNumber
    
    'Call clearSourceBoardNoRangesBoxList(groupName, groupNameStartRowNumber, groupNameEndRowNumber)
End Sub

Private Sub getsourceAttrBoardNo(ByRef ws As Worksheet, ByRef groupRowNumber As Long, ByRef autoCopySourceNameCol As Collection, ByRef columnNamePositionLetterMap As CMap, ByRef sourceAttrBoardNo As CMap)
    Dim columnNamePositionLetter As String
    Dim sourcecopyBoardNoNameCol As Collection
    Dim valueStr As String
    Dim keyStr As String
    Dim sourcecopyBoardNoName As Variant
    Dim autoCopyName As String
    Dim i As Long
    i = 0
    
    Set sourcecopyBoardNoNameCol = selectedGroupMappingDefData.copyFillInSourceColumnName
    
    For Each sourcecopyBoardNoName In sourcecopyBoardNoNameCol
        keyStr = columnNamePositionLetterMap.GetAt(CStr(sourcecopyBoardNoName))
        autoCopyName = autoFillinCopySourceColumnName(i, autoCopySourceNameCol)
        columnNamePositionLetter = columnNamePositionLetterMap.GetAt(autoCopyName)
        valueStr = ws.Range(columnNamePositionLetter & groupRowNumber).value
        
        Call sourceAttrBoardNo.SetAt(keyStr, valueStr)
        i = i + 1
    Next sourcecopyBoardNoName
    
End Sub

Private Function autoFillinCopySourceColumnName(ByRef copyNamePos As Long, ByRef autoCopySourceNameCol As Collection) As String
    Dim i As Long
    Dim autoCopySourceName As Variant
    
    i = 0
    
    autoFillinCopySourceColumnName = ""
    For Each autoCopySourceName In autoCopySourceNameCol
        If copyNamePos = i Then
            autoFillinCopySourceColumnName = CStr(autoCopySourceName)
            Exit Function
        End If
        i = i + 1
    Next autoCopySourceName
    
End Function


Private Function isNeedCopyRecond(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef columnNamePositionLetterMap As CMap) As Boolean
    Dim mocNeedFillColumnNameCol As Collection
    Dim mocNeedFillColumnNameVar As Variant
    Dim mocColumnLetter As String
    
    isNeedCopyRecond = True
    
    Set mocNeedFillColumnNameCol = selectedGroupMappingDefData.needFillColumnNameCol
    
    For Each mocNeedFillColumnNameVar In mocNeedFillColumnNameCol
        mocColumnLetter = columnNamePositionLetterMap.GetAt(CStr(mocNeedFillColumnNameVar))
        If ws.Range(mocColumnLetter & rowNumber).value <> "" Then
            isNeedCopyRecond = False
            Exit Function
        End If
    Next mocNeedFillColumnNameVar
    
End Function

Private Function isFillinColName(ByRef colnumName As String) As Boolean
    Dim mocNeedFillColumnNameCol As Collection
    Dim mocNeedFillColumnNameVar As Variant
    Dim mocColumnLetter As String
    
    isFillinColName = False
    
    Set mocNeedFillColumnNameCol = selectedGroupMappingDefData.needFillColumnNameCol
    
    For Each mocNeedFillColumnNameVar In mocNeedFillColumnNameCol
        If colnumName = CStr(mocNeedFillColumnNameVar) Then
            isFillinColName = True
            Exit Function
        End If
    Next mocNeedFillColumnNameVar
    
End Function

Private Sub copyMigrationBoardStyleLine(ByRef ws As Worksheet, ByRef groupName As String, ByRef destGroupRowNumber As Long, ByRef sourceNeName As String, ByRef sourceAttrBoardNo As CMap, _
ByRef startcolumnLetter As String, ByRef endColcolumnLetter As String)
    Dim sourceBoardstyleSheetName As String
    Dim sourceBoardstylesheet As Worksheet
    Dim sourceRowNumber As Long
    Dim delKeyValue As String
    
    Set neBoardStyleMap = baseStationData.neBoardStyleMap
    
    sourceBoardstyleSheetName = neBoardStyleMap.GetAt(sourceNeName)
    Set sourceBoardstylesheet = ThisWorkbook.Worksheets(sourceBoardstyleSheetName)
    
    sourceRowNumber = getSourceGroupRowNumber(sourceBoardstylesheet, groupName, sourceAttrBoardNo)
    If sourceRowNumber = -1 Then Exit Sub
    sourceBoardstylesheet.Range(startcolumnLetter & sourceRowNumber & ":" & endColcolumnLetter & sourceRowNumber).Copy
    ws.Range(startcolumnLetter & destGroupRowNumber & ":" & endColcolumnLetter & destGroupRowNumber).PasteSpecial
    delKeyValue = sourceBoardstyleSheetName + "," + groupName + "," + CStr(migrationCount)
    Call delSourceBoardStyleDataMap.SetAt(delKeyValue, sourceAttrBoardNo)
    'sourceBoardstylesheet.Rows(sourceRowNumber).Delete

End Sub
Private Sub copyMigrationBoardStyleLinewtihoutKey(ByRef ws As Worksheet, ByRef groupName As String, ByRef destGroupRowNumber As Long, ByRef sourceNeName As String, ByRef sourceAttrBoardNo As CMap, _
ByRef startcolumnNum As Long, ByRef endColcolumnNum As Long)
    Dim sourceBoardstyleSheetName As String
    Dim sourceBoardstylesheet As Worksheet
    Dim sourceRowNumber As Long
    Dim neBoardStyleMap As CMap
    Dim colunmNum As Long
    Dim groupNameRowNumber As Long
    Dim colnumName As String
    Dim colnumNameLetter As String
    Dim dstRange As Range
    Dim delKeyValue As String
    
    
    groupNameRowNumber = getGroupNameStartRowNumber(ws, groupName)
    
    Set neBoardStyleMap = baseStationData.neBoardStyleMap
    
    sourceBoardstyleSheetName = neBoardStyleMap.GetAt(sourceNeName)
    Set sourceBoardstylesheet = ThisWorkbook.Worksheets(sourceBoardstyleSheetName)
    
    sourceRowNumber = getSourceGroupRowNumber(sourceBoardstylesheet, groupName, sourceAttrBoardNo)
    If sourceRowNumber = -1 Then Exit Sub
    For colunmNum = startcolumnNum To endColcolumnNum
       colnumNameLetter = getColStr(colunmNum)
       colnumName = ws.Range(colnumNameLetter & (groupNameRowNumber + 1)).value
       Set dstRange = ws.Range(colnumNameLetter & destGroupRowNumber & ":" & colnumNameLetter & destGroupRowNumber)
       If dstRange.value = "" And dstRange.Interior.colorIndex <> SolidColorIdx Then
           sourceBoardstylesheet.Range(colnumNameLetter & sourceRowNumber & ":" & colnumNameLetter & sourceRowNumber).Copy
           ws.Range(colnumNameLetter & destGroupRowNumber & ":" & colnumNameLetter & destGroupRowNumber).PasteSpecial
       End If
       
    Next
    
    delKeyValue = sourceBoardstyleSheetName + "," + groupName + "," + CStr(migrationCount)
    Call delSourceBoardStyleDataMap.SetAt(delKeyValue, sourceAttrBoardNo)
    'sourceBoardstylesheet.rows(sourceRowNumber).Delete
    Dim srcKey As String
    srcKey = ws.name + "," + groupName + "," + CStr(destGroupRowNumber)
    Call sourceBoardStyleDataMap.SetAt(srcKey, CStr(endColcolumnNum))
End Sub

Private Function getSourceGroupRowNumber(ByRef ws As Worksheet, ByRef groupName As String, ByRef sourceAttrBoardNo As CMap) As Long

    Dim rowNumber As Long
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    Call getBoardStyleGroupNameStartAndEndRowNumber(ws, groupName, groupNameStartRowNumber, groupNameEndRowNumber)
    
    For rowNumber = groupNameStartRowNumber + 2 To groupNameEndRowNumber
        If isExitGroupRecord(ws, sourceAttrBoardNo, rowNumber) Then
            getSourceGroupRowNumber = rowNumber
            Exit Function
        End If
        
    Next rowNumber
    
    getSourceGroupRowNumber = -1
    
End Function


Private Function isExitGroupRecord(ByRef ws As Worksheet, ByRef sourceAttrBoardNo As CMap, ByRef groupNumber As Long) As Boolean
    Dim sourceColunmNameLeCol As Collection
    Set sourceColunmNameLeCol = sourceAttrBoardNo.KeyCollection
    
    Dim colunmLeterVar As Variant
    Dim colunmLeter As String
    Dim columValue As String
    
    isExitGroupRecord = False
    
    For Each colunmLeterVar In sourceColunmNameLeCol
        colunmLeter = colunmLeterVar
        columValue = sourceAttrBoardNo.GetAt(colunmLeter)
        If ws.Range(colunmLeter & groupNumber).value = columValue Then
            isExitGroupRecord = True
        Else
            isExitGroupRecord = False
            Exit Function
        End If
    Next colunmLeterVar
    
End Function

'??SourceBoxList

Public Sub updateSourceBoradNobyNeName(ByRef ws As Worksheet, ByRef Target As Range)
    Dim cellRange As Range
    Dim groupName As String
    Dim columnName As String
    Dim rowNumber As Long
    Dim neName As String
    
    If Target.count Mod 256 = 0 Then Exit Sub
    If isBoardStyleSheet(ws) = False Then Exit Sub
    If allBoardStyleData Is Nothing Then Call initAllBoardStyleDataPublic
    Call allBoardStyleData.initBoardStyleDataMap
    
    If boardStyleData Is Nothing Then Call initBoardStyleMappingDataPublic
    If boardStyleMappingDefMap Is Nothing Then
        Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    End If
    
    For Each cellRange In Target
        If findAttrName(Trim(cellRange.value)) = True Or findGroupName(Trim(cellRange.value)) = True Or cellRange.Borders.LineStyle = xlLineStyleNone Then
            Exit Sub
        End If
        
        Call getGroupAndColumnName(ws, cellRange, groupName, columnName)
        
        Set selectedGroupMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
        
        If isSourceNeNameColum(groupName, columnName) And boardStyleNeMap.hasKey(ws.name) Then
            neName = boardStyleNeMap.GetAt(ws.name)
            Call updateSourceNeNameBoxList(ws, groupName, neName, columnName, cellRange)
        End If
        
        If isSourceCopyColumName(columnName) Then
            rowNumber = cellRange.row
            neName = ws.Range(getStartLetter(ws, rowNumber) & rowNumber).value
            Call updateSourceGroupBoxList(ws, rowNumber, neName, groupName)
            'Call updateSourceGroupBoxList(ws, cellRange.row, ws.Range("A" & rowNumber), groupName)
        End If
        
        'checkinput
        
    Next cellRange
    
End Sub

Private Sub updateSourceNeNameBoxList(ByRef ws As Worksheet, ByRef groupName As String, ByRef crruentNeName As String, ByRef columnName As String, ByRef cellRange As Range)
    Dim neName As Variant
    Dim neNameStr As String
    Dim srcTarNeMap As CMapValueObject
    Dim onesrcNeMap As CMap
    
    Dim migrationData As CMigrationDataManager
    Set migrationData = New CMigrationDataManager
    Call migrationData.init
    Set srcTarNeMap = migrationData.targetSourceNeMap
    If srcTarNeMap.hasKey(crruentNeName) Then Set onesrcNeMap = srcTarNeMap.GetAt(crruentNeName)
    For Each neName In neBoardStyleMap.KeyCollection
        If CStr(neName) = crruentNeName Or (isgroupDataExit(groupName, CStr(neName)) And isSrcNe(onesrcNeMap, CStr(neName))) Then
            If neNameStr = "" Then
                neNameStr = CStr(neName)
            Else
                neNameStr = neNameStr + "," + CStr(neName)
            End If
        End If
    Next
    
    If neNameStr <> "" Then
        Call setBoardStyleListBoxRangeValidation(ws.name, groupName, columnName, neNameStr, ws, cellRange)
    End If
    
End Sub


Private Sub updateSourceGroupBoxList(ByRef ws As Worksheet, ByRef selectRowNumber As Long, ByRef neName As String, ByRef groupName As String)
    Dim autocopyBoardNoNameCol As Collection
    Set autocopyBoardNoNameCol = selectedGroupMappingDefData.autoFillInSourceColumnName
    
    Dim autocopyBoardNoName As Variant
    
    For Each autocopyBoardNoName In autocopyBoardNoNameCol
        Call updateNeSourceGroupColBoxList(ws, neName, selectRowNumber, groupName, CStr(autocopyBoardNoName))
    Next autocopyBoardNoName
End Sub

Private Sub updateNeSourceGroupColBoxList(ByRef ws As Worksheet, ByRef neName As String, ByRef selectRowNumber As Long, ByRef groupName As String, ByRef copyColName As String)
    Dim columnNamePositionLetterMap As CMap
    Dim columnNamePositionLetter As String
    Dim keyStr As String
    Dim valueStr As String
    Dim boardStyleDataMap As CMap
    Dim allBoardStyleDataMap As CMapValueObject
    
    Set allBoardStyleDataMap = allBoardStyleData.allBoardStyleDataMap
    Set boardStyleDataMap = allBoardStyleDataMap.GetAt(neName)
    
    Set columnNamePositionLetterMap = selectedGroupMappingDefData.columnNamePositionLetterMap
    columnNamePositionLetter = columnNamePositionLetterMap.GetAt(copyColName)
    

    keyStr = neName + "_" + groupName + "_" + copyColName
    valueStr = boardStyleDataMap.GetAt(keyStr)
    
    Call setBoardStyleListBoxRangeValidation(ws.name, groupName, copyColName, valueStr, ws, ws.Range(columnNamePositionLetter & selectRowNumber))
    
End Sub

Private Function isSourceNeNameColum(ByRef groupName As String, ByRef columnName As String) As Boolean
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    isSourceNeNameColum = False
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        attributeName = sheetDef.Cells(index, 5)
        If attributeName = "SOURCEBTSNAME" And groupName = mappingdefgroupName Then
            Exit For
        End If
    Next
    If groupName = mappingdefgroupName And columnName = mappingdefcolumnName Then
        isSourceNeNameColum = True
    End If
End Function

Private Function isSourceCopyColumName(ByRef columnName As String) As Boolean
    Dim autocopyBoardNoNameCol As Collection
    Set autocopyBoardNoNameCol = selectedGroupMappingDefData.autoFillInSourceColumnName
    isSourceCopyColumName = False
    
    Dim autocopyBoardNoName As Variant
    Dim copyName As String
    For Each autocopyBoardNoName In autocopyBoardNoNameCol
        copyName = autocopyBoardNoName
        If columnName = copyName Then
            isSourceCopyColumName = True
            Exit Function
        End If

    Next autocopyBoardNoName
    
End Function

'Public Function isUmtsSectoreqmPo() As Boolean
'    Dim actSheetName As String
'    Dim cellsheet As Worksheet
'    Dim mocName As String
'    Dim attrName As String
'    Dim sectorColumnName As String
'    Dim columnIndex As Long
'    isUmtsSectoreqmPo = False
'    actSheetName = getResByKey("A185")
'    For Each cellsheet In ThisWorkbook.Worksheets
'        Call getSectoreqmMocNameAndAttrName(mocName, attrName)
'        sectorColumnName = findColumnFromRelationDef(actSheetName, mocName, attrName)
'        columnIndex = findColumnByName(cellsheet, sectorColumnName, 2)
'        If columnIndex > 0 And sectorColumnName <> "" Then
'            isUmtsSectoreqmPo = True
'            Exit Function
'        End If
'    Next
'End Function

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
    
    Dim boardNoRange As Range
    Dim boardNoTypeRange As Range
    
    Dim rowNumber As Long
    
    cxuTypeColumnName = getCxuTypeColum()
    
    
    autoFillInColumnName = currentBoardStyleMappingDefData.autoFillInColumnName
    If autoFillInColumnName <> "" Then
        Call getGroupNameStartAndEndRowNumber(groupName, groupNameStartRowNumber, groupNameEndRowNumber)
        boardNoColumLetter = currentBoardStyleMappingDefData.getColumnNamePositionLetter(autoFillInColumnName)
        boardNoTypeColumLetter = currentBoardStyleMappingDefData.getColumnNamePositionLetter(cxuTypeColumnName)
        
        For rowNumber = groupNameStartRowNumber + 2 To groupNameEndRowNumber
            If ws.Cells(rowNumber, 1) <> "RMV" Then
                Set boardNoRange = ws.Range(boardNoColumLetter & rowNumber)
                Set boardNoTypeRange = ws.Range(boardNoTypeColumLetter & rowNumber)
                
                eachBoardNo = Trim(boardNoRange.value)
                If eachBoardNo <> "" And Trim(boardNoTypeRange.value) = "CGPS" Then
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
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
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
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mocName = sheetDef.Cells(index, 4)
        If mocName = "CXU" Then
            getCxugroupName = mappingdefgroupName
            Exit For
        End If
    Next
End Function

Private Function isgroupDataExit(ByRef groupName As String, ByRef neName As String) As Boolean
    Dim keyStr As String
    Dim valueStr As String
    Dim boardStyleDataMap As CMap
    Dim allBoardStyleDataMap As CMapValueObject
    Set allBoardStyleDataMap = allBoardStyleData.allBoardStyleDataMap
    
    Dim autocopyBoardNoNameCol As Collection
    Set autocopyBoardNoNameCol = selectedGroupMappingDefData.autoFillInSourceColumnName
    Dim autocopyBoardNoName As Variant
    
    isgroupDataExit = True
    For Each autocopyBoardNoName In autocopyBoardNoNameCol
        If allBoardStyleDataMap.hasKey(neName) Then
            Set boardStyleDataMap = allBoardStyleDataMap.GetAt(neName)
            keyStr = neName + "_" + groupName + "_" + autocopyBoardNoName
            If boardStyleDataMap.hasKey(keyStr) Then
                            valueStr = boardStyleDataMap.GetAt(keyStr)
                If valueStr = "" Then
                    isgroupDataExit = False
                    Exit Function
                End If
            End If
        End If
    Next autocopyBoardNoName
End Function

Private Function isSrcNe(ByRef srcNeMap As CMap, ByRef srcNe As String) As Boolean
    Dim keyValue As Variant
    Dim valueStr As String
    Dim valueStrArry() As String
    isSrcNe = False
    If srcNeMap Is Nothing Then Exit Function
    For Each keyValue In srcNeMap.KeyCollection
        valueStr = srcNeMap.GetAt(keyValue)
        If valueStr <> "" Then valueStrArry = Split(valueStr, ",")
        If isSrcNeInArry(valueStrArry, srcNe) Then
            isSrcNe = True
            Exit Function
        End If
    Next
End Function

Private Function isSrcNeInArry(ByRef valueStrArry() As String, ByRef srcNe As String) As Boolean
    Dim i As Long
    isSrcNeInArry = False
    
    For i = 0 To UBound(valueStrArry)
        If valueStrArry(i) = srcNe Then
            isSrcNeInArry = True
            Exit Function
        End If
    Next

End Function

Public Sub refreshTargetNeCol(sheetName As String)
    Dim ws As Worksheet
    If sheetName <> getResByKey("BaseTransPort") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    Dim migrationData As CMigrationDataManager
    Dim srcTarNeMap As CMapValueObject
    Set migrationData = New CMigrationDataManager
    Call migrationData.init
    Set srcTarNeMap = migrationData.targetSourceNeMap

    
    Dim migrationColName As String
    Dim migrationColLeter As String
    Dim neNameColName As String
    Dim neNameColLeter As String
    Dim rowNumber As Long
    Dim neName As String
    
    
    migrationColName = getResByKey("Is Migration Target NE")
    migrationColLeter = getColnumNameLeter(ws, migrationColName)
    neNameColName = getNeNameColum
    neNameColLeter = getColnumNameLeter(ws, neNameColName)
    
    If migrationColLeter = "" Or neNameColLeter = "" Then Exit Sub

    'DTS2017011105086
    Dim clipBoardData As DataObject
    Set clipBoardData = New DataObject
    
    clipBoardData.Clear
    clipBoardData.GetFromClipboard
    
    For rowNumber = 3 To ws.Range("B1048576").End(xlUp).row
        neName = ws.Range(neNameColLeter & rowNumber).value
        If srcTarNeMap.hasKey(neName) Then
            ws.Range(migrationColLeter & rowNumber).value = "YES"
        Else
            ws.Range(migrationColLeter & rowNumber).value = "NO"
        End If
        With ws.Range(migrationColLeter & rowNumber)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .NumberFormatLocal = "@" '设置单元格格式为文本
        End With
        
        Call setHyperlinkRangeFont(ws.Range(migrationColLeter & rowNumber))
    Next
    
    'DTS2017011105086
    With clipBoardData
        .SetText ""
        .PutInClipboard
    End With
End Sub
Private Function getColnumNameLeter(ByRef ws As Worksheet, ByRef colnumName As String) As String
    Dim m_colNum As Long
    For m_colNum = 1 To ws.Range("XFD2").End(xlToLeft).column
        If ws.Cells(2, m_colNum).value = colnumName Then
            getColnumNameLeter = getColStr(m_colNum)
            Exit Function
        End If
    Next
    getColnumNameLeter = ""
End Function

Private Function getNeNameColum() As String
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    getNeNameColum = ""
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.Range("B1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mocName = "BTS" And attributeName = "BTSNAME" Then
            getNeNameColum = mappingdefcolumnName
            Exit For
        End If
    Next
End Function
Public Sub setMigrationRecbackColor(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef maxColumnNumber As Long)

    Dim neName As String
    Dim sourceNeName As String
    Dim colNum As Long
    Dim cellRange As Range
    Dim cloltter As String

    If allBoardStyleData Is Nothing Then Call initAllBoardStyleDataPublic
    Call allBoardStyleData.initBoardStyleDataMap

    If boardStyleNeMap.hasKey(ws.name) Then neName = boardStyleNeMap.GetAt(ws.name)
    

        sourceNeName = ws.Range("B" & rowNumber).value
        If neName <> sourceNeName Then
            For colNum = 1 To maxColumnNumber
                cloltter = getColStr(colNum)
                Set cellRange = ws.Range(cloltter & rowNumber)
                If cellRange.Interior.colorIndex <> SolidColorIdx And cellRange.Interior.Pattern <> SolidPattern Then
                    With cellRange.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent4
                        .TintAndShade = 0.599993896298105
                        .PatternTintAndShade = 0
                    End With
                End If
            Next
        End If

End Sub


Public Function getSrcNeNameCol(ByRef ws As Worksheet, ByVal row As Long) As String
    getSrcNeNameCol = ""
    
    Dim sourceNeNameCol As String, sourceNeNameLetter As String
    sourceNeNameCol = getResByKey("SOURCEBTSNAME")
    
    '找到GroupName所在的行
    Dim groupRowNum As Long
    groupRowNum = getGroupRowNum(ws, row)
    If groupRowNum = -1 Then Exit Function
    
    getSrcNeNameCol = findColLetterByRowAndValue(ws, groupRowNum + 1, sourceNeNameCol)

End Function

    
Public Function getStartLetter(ByRef ws As Worksheet, ByVal row As Long) As String
    
    getStartLetter = getSrcNeNameCol(ws, row)
    If getStartLetter = "" Then getStartLetter = "A"
    
End Function
