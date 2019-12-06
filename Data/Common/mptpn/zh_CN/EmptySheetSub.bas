Attribute VB_Name = "EmptySheetSub"
Sub hiddenEmptySheet()
        Dim index As Long
        Dim column As Long
        Dim isEmpty As Boolean
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        For index = 2 To sheetDef.Range("a65536").End(xlUp).row
                isEmpty = True
                If UCase(sheetDef.Cells(index, 2)) <> "MAIN" And UCase(sheetDef.Cells(index, 2)) <> "COMMON" Then
                    For column = 1 To ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value).Range("IV2").End(xlToLeft).column
                             If ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value).Cells(3, column).value <> "" Then
                                    isEmpty = False
                                    Exit For
                             End If
                    Next
                    If isEmpty Then
                        ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value).Visible = 0
                    End If
                End If
        Next
End Sub

Sub showEmptySheet()
        Dim index As Long
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        For index = 2 To sheetDef.Range("a65536").End(xlUp).row
                 If UCase(sheetDef.Cells(index, 2)) <> "MAIN" And UCase(sheetDef.Cells(index, 2)) <> "COMMON" Then
                    ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value).Visible = -1
                End If
        Next
End Sub

Public Sub copySplitSectorAfterLogic()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Call processOneCellSheet(ws)
    Next ws
End Sub

Private Sub processOneCellSheet(ByRef ws As Worksheet)
    Dim maxColumnNumber As Integer
    Dim cellValue As String
    Dim startsplitSectorColumn As Integer
    Dim endsplitSectorColumn As Integer
    Dim startsplitCellColumn As Integer
    Dim endsplitCellColumn As Integer
    Dim wsName As String
    
    wsName = ws.name

    Dim columnNumber As Integer
    Dim emptyColumn As Integer
    
    
    
    startsplitSectorColumn = -1
    startsplitCellColumn = -1
    endsplitSectorColumn = -1
    endsplitCellColumn = -1
    'emptyColumn = findemptyColumn(ws)
    
    maxColumnNumber = ws.Range("A2").End(xlToRight).column
    emptyColumn = maxColumnNumber + 1
    
    If isUmtsCellSheet(ws.name) = True Then
        For columnNumber = 1 To maxColumnNumber
            cellValue = ws.Cells(1, columnNumber)
            If isSplitCellColumn(cellValue) Then
                startsplitCellColumn = columnNumber
            End If
        Next
        endsplitCellColumn = findendColumn(ws, startsplitCellColumn)
        cellValue = ws.Cells(1, endsplitCellColumn + 1)
        
        If cellValue <> "" And startsplitCellColumn <> -1 And endsplitCellColumn <> -1 Then
            ws.Columns(getColStr(startsplitCellColumn) & ":" & getColStr(endsplitCellColumn)).Copy
            ws.Columns(getColStr(emptyColumn)).Insert Shift:=xlToRight
            ws.Columns(getColStr(startsplitCellColumn) & ":" & getColStr(endsplitCellColumn)).Delete Shift:=xlToLeft
        End If
        
    End If
    
    If isBaseTrsportSheet(ws.name) = True Then
        
        For columnNumber = 1 To maxColumnNumber
            cellValue = ws.Cells(1, columnNumber)
            If isSplitSectorColumn(cellValue) Then
                startsplitSectorColumn = columnNumber
            End If
        Next
        endsplitSectorColumn = findendColumn(ws, startsplitSectorColumn)
        cellValue = ws.Cells(1, endsplitSectorColumn + 1)
        
        If cellValue <> "" And startsplitSectorColumn <> -1 And endsplitSectorColumn <> -1 Then
            ws.Columns(getColStr(startsplitSectorColumn) & ":" & getColStr(endsplitSectorColumn)).Copy
            ws.Columns(getColStr(emptyColumn)).Insert Shift:=xlToRight
            ws.Columns(getColStr(startsplitSectorColumn) & ":" & getColStr(endsplitSectorColumn)).Delete Shift:=xlToLeft
        End If
        
    End If
 
End Sub

Private Function isSplitSectorColumn(cellValue As String) As Boolean
    If cellValue = getResByKey("SplitInfo") Or cellValue = "Sector Split Information" Then
        isSplitSectorColumn = True
    Else
        isSplitSectorColumn = False
    End If
    
End Function

Private Function isSplitCellColumn(cellValue As String) As Boolean
    If cellValue = getResByKey("CellSplitInfo") Or cellValue = "Cell Split Information" Then
        isSplitCellColumn = True
    Else
        isSplitCellColumn = False
    End If
    
End Function

Private Function findemptyColumn(ByRef ws As Worksheet) As Integer
    Dim maxColumnNumber As Integer
    Dim cellValue As String
    Dim columnNumber As Integer
    
    findemptyColumn = -1
    
    maxColumnNumber = ws.Range("A2").End(xlToRight).column
    For columnNumber = 1 To maxColumnNumber
        cellValue = ws.Cells(2, columnNumber)
        If cellValue = "" Then
            findemptyColumn = columnNumber
        End If
    Next
    
End Function

Function isUmtsCellSheet(sheetName As String) As Boolean
    If sheetName = "UMTS Cell" Or sheetName = getResByKey("A175") Then
        isUmtsCellSheet = True
        Exit Function
    End If
    isUmtsCellSheet = False
End Function

Function isBaseTrsportSheet(sheetName As String) As Boolean
    If sheetName = "Base Station Transport Data" Or sheetName = getResByKey("BaseTransPort") Then
        isBaseTrsportSheet = True
        Exit Function
    End If
    isBaseTrsportSheet = False
End Function


Private Function findendColumn(ByRef ws As Worksheet, findfirstColumn As Integer) As Integer
    Dim maxColumnNumber As Integer
    Dim cellValue As String
    Dim columnNumber As Integer
    
    findendColumn = findfirstColumn
    
    maxColumnNumber = ws.Range("A2").End(xlToRight).column
    For columnNumber = findfirstColumn + 1 To maxColumnNumber
        cellValue = ws.Cells(1, columnNumber)
        If cellValue <> "" Then
            findendColumn = columnNumber - 1
            Exit Function
        End If
    Next
    findendColumn = maxColumnNumber

End Function
