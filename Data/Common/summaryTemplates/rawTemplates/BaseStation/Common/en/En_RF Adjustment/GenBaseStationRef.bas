Attribute VB_Name = "GenBaseStationRef"
Option Explicit
Dim board_pattern As String
Public board_style As String


Public Sub genBaseStationBoardStyleRef(ByRef ws As Worksheet, ByRef target As range)
    Dim groupName As String, columnName As String
    Dim rowNumber As Long, columnNumber As Long
    Dim refString As String
    Dim flag As Boolean
    If target.rows.count <> 1 Or target.Columns.count <> 1 Then Exit Sub
    flag = getBoardStyleInfo(ws.name)
    If flag = False Then
        Exit Sub
    End If
    
    rowNumber = target.row
    columnNumber = target.column
    flag = isBoardStyleCol(ws, rowNumber, columnNumber)
    If flag = False Then Exit Sub
    
    refString = genBoardStyleRefString()
    If refString = "" Then Exit Sub
    
    Call setBoardStyleListBoxRangeValidation(ws.name, board_pattern, board_style, refString, ws, target)
End Sub


Private Function genBoardStyleRefString() As String
    Dim refString As String
    Dim ws As Worksheet
    Dim boardStyleSheetName As String
    refString = ""
    boardStyleSheetName = getResByKey("Board Style")
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.name, boardStyleSheetName) <> 0 Then
            If refString = "" Then
                refString = ws.name
            Else
                refString = refString + "," + ws.name
            End If
        End If
    Next
    genBoardStyleRefString = refString
End Function


Public Function isBoardStyleCol(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef columnNumber As Long) As Boolean
    Dim attrValue As String
    Dim mocValue As String
    Dim maxRow As Long
'    maxRow = ws.Range("a1048576").End(xlUp).row
'    If rowNumber > maxRow Or rowNumber < 3 Then
'        isBoardStyleCol = False
'        Exit Function
'    End If
    
    If rangeHasBorder(ws.rows(rowNumber)) = False Or rowNumber < 3 Then
        isBoardStyleCol = False
        Exit Function
    End If

    attrValue = ws.Cells(2, columnNumber).value
    If attrValue = board_style Then
        mocValue = getMocGroupName(ws, 1, columnNumber)
        If mocValue = board_pattern Then
            isBoardStyleCol = True
            Exit Function
        End If
    End If
    isBoardStyleCol = False
End Function


Public Function getBoardStyleInfo(ByRef sheetName As String) As Boolean
    Dim relationDef As Worksheet
    Dim maxRow As Long
    Dim rowNum As Long
    Dim tmpName As String
    Set relationDef = ThisWorkbook.Worksheets("RELATION DEF")
    
    
    maxRow = relationDef.range("a1048576").End(xlUp).row
    
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


Private Function getMocGroupName(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef columnNumber As Long) As String
    Dim cellValue As String
    cellValue = ws.Cells(rowNumber, columnNumber).value
    Dim k As Long
    If cellValue = "" Then
        For k = columnNumber To 1 Step -1
            cellValue = ws.Cells(rowNumber, k).value
            If cellValue <> "" Then
                getMocGroupName = cellValue
                Exit Function
            End If
        Next k
    Else
        getMocGroupName = cellValue
    End If
End Function


