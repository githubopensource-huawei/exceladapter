Attribute VB_Name = "EmptySheetSub"
Sub hiddenEmptySheet()
        Dim index As Long
        Dim column As Long
        Dim isEmpty As Boolean
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        For index = 2 To sheetDef.range("a1048576").End(xlUp).row
                isEmpty = True
                If UCase(sheetDef.Cells(index, 2)) <> "MAIN" And UCase(sheetDef.Cells(index, 2)) <> "COMMON" Then
                    For column = 1 To ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value).range("XFD2").End(xlToLeft).column
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
        For index = 2 To sheetDef.range("a1048576").End(xlUp).row
                 If UCase(sheetDef.Cells(index, 2)) <> "MAIN" And UCase(sheetDef.Cells(index, 2)) <> "COMMON" Then
                    ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value).Visible = -1
                End If
        Next
End Sub


Public Sub copyDesignCellIdAfterLocellId()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If isCellSheet(ws.name) Then
           Call processOneCellSheet(ws)
        End If
    Next ws
End Sub

Private Sub processOneCellSheet(ByRef ws As Worksheet)
    Dim maxColumnNumber As Integer
    Dim cellValue As String
    Dim lcoellIdColumn As Integer
    Dim designCellIdColumn As Integer
    Dim columnNumber As Integer
    
    lcoellIdColumn = -1
    designCellIdColumn = -1
    
    maxColumnNumber = ws.range("A2").End(xlToRight).column
    For columnNumber = 1 To maxColumnNumber
        cellValue = ws.Cells(2, columnNumber)
        If isLocellIdColumn(cellValue) Then
            lcoellIdColumn = columnNumber
        End If
        If isDesignLocellIdColumn(cellValue) Then
            designCellIdColumn = columnNumber
        End If
    Next
    
    If designCellIdColumn <> -1 And lcoellIdColumn <> -1 And designCellIdColumn <> lcoellIdColumn + 1 Then
        ws.Columns(getColStr(designCellIdColumn)).Copy
        ws.Columns(getColStr(lcoellIdColumn + 1)).Insert Shift:=xlToRight
        ws.Columns(calculateColumnName(designCellIdColumn + 1) & ":" & calculateColumnName(designCellIdColumn + 1)).delete Shift:=xlToLeft
        
    End If
    
End Sub

Private Function isLocellIdColumn(cellValue As String) As Boolean
    If cellValue = getResByKey("A188") Or cellValue = "*Local Cell ID" Then
        isLocellIdColumn = True
    Else
        isLocellIdColumn = False
    End If
    
End Function

Private Function isDesignLocellIdColumn(cellValue As String) As Boolean
    If cellValue = getResByKey("A187") Or cellValue = "BTS Design Cell ID" Then
        isDesignLocellIdColumn = True
    Else
        isDesignLocellIdColumn = False
    End If
    
End Function
