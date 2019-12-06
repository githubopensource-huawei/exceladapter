Attribute VB_Name = "OperateBoardStyleData"
Option Explicit

Public boardStyleData As CBoardStyleData
Public boardNoManager As CBoardNoManager
Public boardStyleMappingDefMap As CMapValueObject
Public Const PublicMaxRowNumber As Long = 2000

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
    maxColumnNumber = ws.UsedRange.columns.count
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
    maxRowNumber = ws.UsedRange.rows.count
    findCertainValRowNumber = -1
    For k = startRow To maxRowNumber
        currentCellVal = ws.range(columnLetter & k).value
        If currentCellVal = cellVal Then
            findCertainValRowNumber = k
            Exit For
        End If
    Next
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

Public Function getGroupNameStartRowNumber(ByRef ws As Worksheet, ByRef groupName As String) As Long
    Dim rowNumber As Long, initialRowNumber As Long
    Dim lastRowEmptyFlag As Boolean
    lastRowEmptyFlag = False
    initialRowNumber = 1
    Do While Not lastRowEmptyFlag
        rowNumber = findCertainValRowNumber(ws, "A", groupName, initialRowNumber)
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
    If g_CurrentSheet Is Nothing Then Set g_CurrentSheet = ThisWorkbook.ActiveSheet
    groupNameStartRowNumber = getGroupNameStartRowNumber(g_CurrentSheet, groupName)
    'groupNameEndRowNumber = findCertainValRowNumber(g_CurrentSheet, "A", "", groupNameStartRowNumber)
    'groupNameEndRowNumber = groupNameStartRowNumber + g_CurrentSheet.Range("A" & groupNameStartRowNumber).CurrentRegion.rows.count - 1
    groupNameEndRowNumber = groupNameStartRowNumber + getCurrentRegionRowsCount(g_CurrentSheet, groupNameStartRowNumber) - 1
    
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

Public Sub initBoardStyleMappingDataPublic()
    If boardStyleData Is Nothing Then Set boardStyleData = New CBoardStyleData
    Call boardStyleData.init
End Sub

Public Sub initBoardNoManagerPublic()
    If boardNoManager Is Nothing Then Set boardNoManager = New CBoardNoManager
    Call boardNoManager.generateCurrentGroupNameBoardNoMap
End Sub

Public Function getBoardStyleInfo(ByRef sheetName As String) As Boolean
    Dim relationDef As Worksheet
    Dim maxRow As Long
    Dim rowNum As Long
    Dim tmpName As String
    Set relationDef = ThisWorkbook.Worksheets("RELATION DEF")
    
    
    maxRow = relationDef.range("a65536").End(xlUp).row
    
    For rowNum = 2 To maxRow
        tmpName = relationDef.Cells(rowNum, 1).value
        If tmpName = sheetName And relationDef.Cells(rowNum, 4).value = "True" And relationDef.Cells(rowNum, 5).value = "False" Then
            board_pattern = relationDef.Cells(rowNum, 2).value
            board_style = relationDef.Cells(rowNum, 3).value
            getBoardStyleInfo = True
            Exit Function
        End If
    Next
    getBoardStyleInfo = False
End Function
