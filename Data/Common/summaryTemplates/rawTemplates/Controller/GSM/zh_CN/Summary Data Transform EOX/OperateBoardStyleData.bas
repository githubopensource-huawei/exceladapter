Attribute VB_Name = "OperateBoardStyleData"
Option Explicit

Public inAddProcessFlag As Boolean
Public currentSheet As Worksheet
Public currentCellValue As String
Public moiRowsManager As CMoiRowsManager
Public boardStyleMappingDefMap As CMapValueObject

Public Const NewMoiRangeColorIndex As Long = 43 '浅绿，新增Moi行底色
Public Const NeedFillInRangeColorIndex As Long = 33 '浅蓝，必填单元格底色
Public Const NormalRangeColorIndex As Long = -4142 '白色，正常单元格底色
Public Const BoardNoDelimiter As String = "_"

Public Const PublicMaxRowNumber As Long = 2000
Public Sub setNewRangesBackgroundColour(ByRef colorIndex As Long)
    Dim newMoiRange As Range, eachCell As Range
    Set newMoiRange = moiRowsManager.getMoiRange
    For Each eachCell In newMoiRange
        '如果单元格不是灰化的，则置为正常底色，是分支控制的灰化，则不变
        If eachCell.Interior.colorIndex <> SolidColorIdx And eachCell.Interior.Pattern <> SolidPattern Then eachCell.Interior.colorIndex = colorIndex
    Next eachCell
End Sub
Public Function findCertainValColumnNumber(ByRef ws As Worksheet, ByVal RowNumber As Long, ByRef cellVal As Variant, Optional ByVal startColumn As Long = 1)
    Dim currentCellVal As Variant
    Dim maxColumnNumber As Long, k As Long
    maxColumnNumber = ws.UsedRange.Columns.count
    findCertainValColumnNumber = -1
    For k = startColumn To maxColumnNumber
        currentCellVal = ws.Cells(RowNumber, k).value
        If currentCellVal = cellVal Then
            findCertainValColumnNumber = k
            Exit For
        End If
    Next
End Function
Public Function findCertainValRowNumber(ByRef ws As Worksheet, ByVal columnLetter As String, ByRef cellVal As Variant, Optional ByVal startRow As Long = 1)
    Dim currentCellVal As Variant
    Dim maxRowNumber As Long, k As Long
    maxRowNumber = ws.UsedRange.rows.count
    findCertainValRowNumber = -1
    For k = startRow To maxRowNumber
        currentCellVal = ws.Range(columnLetter & k).value
        If currentCellVal = cellVal Then
            findCertainValRowNumber = k
            Exit For
        End If
    Next
End Function



Public Sub eraseLastChar(ByRef str As String)
    If str <> "" Then str = Left(str, Len(str) - 1)
End Sub

Public Sub selectCertainCell(ByVal ws As Worksheet, ByVal rangeName As String, Optional ByVal scrollFlag As Boolean = True)
    Application.GoTo Reference:=ws.Range(rangeName), Scroll:=scrollFlag
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

Public Function getGroupRowNumber(ByRef ws As Worksheet, ByVal RowNumber) As Long
    Dim maxRowNumber As Long, k As Long
    
    maxRowNumber = ws.UsedRange.rows.count
    If RowNumber > maxRowNumber Then RowNumber = maxRowNumber
    
    For k = RowNumber To 1 Step -1
        If k = 1 Then Exit For
        If rowIsBlank(ws, k - 1) = True And rowIsBlank(ws, k) = False Then Exit For
    Next k
    getGroupRowNumber = k
End Function

Private Function getNextGroupRowNumber(ByRef ws As Worksheet, ByVal RowNumber) As Long
    Dim nextGroupRowNumber As Long
    nextGroupRowNumber = -1
    Dim maxRowNumber As Long, k As Long
    
    maxRowNumber = ws.UsedRange.rows.count
    
    For k = RowNumber To maxRowNumber
        If rowIsBlank(ws, k) = True And rowIsBlank(ws, k + 1) = False Then
            nextGroupRowNumber = k + 1
            Exit For
        End If
    Next k
    
    '如果是-1，说明是最后一个分组，只能用单元格是否有边框来判断最大行了
    If nextGroupRowNumber = -1 Then
        Dim predefinedMaxRowNumber As Long
        predefinedMaxRowNumber = Application.WorksheetFunction.min(RowNumber + 2000, maxRowNumber) '防止最后一个对象的边框一直设置到1048576，设置一个2000的行数最大限制
    
        For k = RowNumber To predefinedMaxRowNumber
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

Public Function getRangeGroupAndColumnName(ByRef ws As Worksheet, ByVal RowNumber As Long, ByVal columnNumber As Long, _
    ByRef groupName As String, ByRef columnName As String, Optional ByVal deleteBoardStyleFlag As Boolean = False) As Boolean
    
    Dim groupRowNumber As Long
    groupRowNumber = getGroupRowNumber(ws, RowNumber)
    groupName = ws.Range("A" & groupRowNumber).value
    columnName = ws.Cells(groupRowNumber + 1, columnNumber).value
    
    Dim groupMaxRow As Long
    groupMaxRow = getNextGroupRowNumber(ws, RowNumber) - 2
'    If deleteBoardStyleFlag = False Then
'        groupMaxRow = groupRowNumber + getCurrentRegionRowsCount(ws, groupRowNumber)
'    Else
'        groupMaxRow = getNextGroupRowNumber(ws, rowNumber) - 2
'    End If

    'groupMaxRow<0说明选择了最后一个分组的超出边框的行
    If (RowNumber > groupRowNumber + 1 And RowNumber <= groupMaxRow) Then
        getRangeGroupAndColumnName = True
        If RowNumber = groupMaxRow And rowIsBlank(ws, RowNumber + 1) = False Then
            getRangeGroupAndColumnName = False
        End If
    Else
        getRangeGroupAndColumnName = False
    End If
End Function

Public Function rowIsBlank(ByRef ws As Worksheet, ByRef RowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.Range("A" & RowNumber & ":IV" & RowNumber)) = 256 Then
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
    Dim RowNumber As Long, initialRowNumber As Long
    Dim lastRowEmptyFlag As Boolean
    lastRowEmptyFlag = False
    initialRowNumber = 1
    Do While Not lastRowEmptyFlag
        RowNumber = findCertainValRowNumber(ws, "A", groupName, initialRowNumber)
        If RowNumber = -1 Then Exit Do
        If RowNumber = 1 Then
            lastRowEmptyFlag = True
        ElseIf rowIsBlank(ws, RowNumber - 1) = True Then
            lastRowEmptyFlag = True
        End If
        initialRowNumber = RowNumber + 1
    Loop
    getGroupNameStartRowNumber = RowNumber
End Function

Public Sub getGroupNameStartAndEndRowNumber(ByRef groupName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    groupNameStartRowNumber = getGroupNameStartRowNumber(currentSheet, groupName)
    'groupNameEndRowNumber = findCertainValRowNumber(currentSheet, "A", "", groupNameStartRowNumber)
    'groupNameEndRowNumber = groupNameStartRowNumber + currentSheet.Range("A" & groupNameStartRowNumber).CurrentRegion.rows.count - 1
    groupNameEndRowNumber = groupNameStartRowNumber + getCurrentRegionRowsCount(currentSheet, groupNameStartRowNumber) - 1
    
End Sub

Private Function getCurrentRegionRowsCount(ByRef ws As Worksheet, ByRef startRowNumber As Long) As Long
    Dim RowNumber As Long
    Dim rowscount As Long
    rowscount = 1
    For RowNumber = startRowNumber + 1 To PublicMaxRowNumber
        If rowIsBlank(ws, RowNumber) = True Then
            Exit For
        Else
            rowscount = rowscount + 1
        End If
    Next RowNumber
    getCurrentRegionRowsCount = rowscount
End Function


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
    Dim rowRange As Range
    Dim selectionRange As Range
    Set selectionRange = Selection
    
    Dim RowNumber As Long, columnNumber As Long
    Dim lastGroupName As String, columnName As String

    columnNumber = 1
    For Each rowRange In selectionRange.rows
        RowNumber = rowRange.row
        rowCollection.Add Item:=RowNumber, key:=CStr(RowNumber)
        If checkLastTwoRow(RowNumber, 1, groupName, columnName, lastGroupName) = False Then
            checkSelectRanges = False
            Exit Function
        End If
    Next rowRange
End Function

Private Function checkLastTwoRow(ByRef RowNumber As Long, ByRef columnNumber As Long, ByRef groupName As String, _
    ByRef columnName As String, ByRef lastGroupName As String) As Boolean
    checkLastTwoRow = True
    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    If getRangeGroupAndColumnName(currentSheet, RowNumber, columnNumber, groupName, columnName, True) = True Then
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

