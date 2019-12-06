Attribute VB_Name = "EmptySheetSub"
Sub hiddenEmptySheet()
        Dim index As Long
        Dim column As Long
        Dim isEmpty As Boolean
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
                isEmpty = True
                If UCase(sheetDef.Cells(index, 2)) <> "MAIN" And UCase(sheetDef.Cells(index, 2)) <> "COMMON" Then
                    For column = 1 To ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value).Range("XFD2").End(xlToLeft).column
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
        For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
                 If UCase(sheetDef.Cells(index, 2)) <> "MAIN" And UCase(sheetDef.Cells(index, 2)) <> "COMMON" Then
                    ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value).Visible = -1
                End If
        Next
End Sub


Public Sub copySourceCellIdAfterLocellId()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If isCellSheet(ws.name) Or isCellOpSheet(ws.name) Then
            Call processOneCellSheet(ws)
        End If
    Next ws
End Sub

Private Sub processOneCellSheet(ByRef ws As Worksheet)
    Dim maxColumnNumber As Integer
    Dim cellValue As String
    Dim lcoellIdColumn1 As Integer
    Dim sourceCellIdColumn1 As Integer
    Dim columnNumber As Integer
    Dim sheet As Worksheet
    
    If containsASheet(ThisWorkbook, "LTE Cell") Then
        Set sheet = ThisWorkbook.Worksheets("LTE Cell")
    Else
        Set sheet = ThisWorkbook.Worksheets(getResByKey("A172"))
    End If
    
    lcoellIdColumn1 = lcoellIdColumn(ws)
    sourceCellIdColumn1 = sourceCellIdColumn()
    
    
    If sourceCellIdColumn1 <> -1 And lcoellIdColumn1 <> -1 Then
        If isCellSheet(ws.name) And sourceCellIdColumn1 <> lcoellIdColumn1 + 1 Then
            ws.Columns(getColStr(sourceCellIdColumn1)).Copy
            ws.Columns(getColStr(lcoellIdColumn1 + 1)).Insert Shift:=xlToRight
            ws.Columns(calculateColumnName(sourceCellIdColumn1 + 1) & ":" & calculateColumnName(sourceCellIdColumn1 + 1)).Delete Shift:=xlToLeft
            Range("A1").Select
        End If
        
        If isCellOpSheet(ws.name) And cellopSourceCellIdColumn(ws) = -1 Then
            sheet.Columns(getColStr(sourceCellIdColumn1)).Copy
            ws.Columns(getColStr(lcoellIdColumn1 + 1)).Insert Shift:=xlToRight
            Application.CutCopyMode = False
            Range("A1").Select
        End If
    End If
    
End Sub

Private Function isLocellIdColumn(cellValue As String) As Boolean
    If cellValue = getResByKey("LOCAL_CELL_ID") Then
        isLocellIdColumn = True
    Else
        isLocellIdColumn = False
    End If
    
End Function

Private Function isSourceLocellIdColumn(cellValue As String) As Boolean
    If cellValue = getResByKey("SOURCE_LOCAL_CELL_ID") Then
        isSourceLocellIdColumn = True
    Else
        isSourceLocellIdColumn = False
    End If
    
End Function

Function isCellOpSheet(sheetName As String) As Boolean
    If cellValue = getResByKey("CELL_OP") Then
        isCellOpSheet = True
        Exit Function
    End If
    isCellOpSheet = False
End Function

Private Function sourceCellIdColumn() As Integer
    Dim maxColumnNumber As Integer
    Dim cellValue As String
    Dim columnNumber As Integer
    Dim ws As Worksheet
    
    If containsASheet(ThisWorkbook, "LTE Cell") Then
        Set ws = ThisWorkbook.Worksheets("LTE Cell")
    Else
        Set ws = ThisWorkbook.Worksheets(getResByKey("A172"))
    End If
    
    sourceCellIdColumn = -1
    
    maxColumnNumber = ws.Range("A2").End(xlToRight).column
    For columnNumber = 1 To maxColumnNumber
        cellValue = ws.Cells(2, columnNumber)
        If isSourceLocellIdColumn(cellValue) Then
            sourceCellIdColumn = columnNumber
        End If
    Next
    
End Function

Private Function lcoellIdColumn(ByRef ws As Worksheet) As Integer
    Dim maxColumnNumber As Integer
    Dim cellValue As String
    Dim columnNumber As Integer
    
    lcoellIdColumn = -1
    
    maxColumnNumber = ws.Range("A2").End(xlToRight).column
    For columnNumber = 1 To maxColumnNumber
        cellValue = ws.Cells(2, columnNumber)
        If isLocellIdColumn(cellValue) Then
            lcoellIdColumn = columnNumber
        End If
    Next
    
End Function

Private Function cellopSourceCellIdColumn(ByRef ws As Worksheet) As Integer
    Dim maxColumnNumber As Integer
    Dim cellValue As String
    Dim columnNumber As Integer
    
    cellopSourceCellIdColumn = -1
    
    maxColumnNumber = ws.Range("A2").End(xlToRight).column
    For columnNumber = 1 To maxColumnNumber
        cellValue = ws.Cells(2, columnNumber)
        If isSourceLocellIdColumn(cellValue) Then
            cellopSourceCellIdColumn = columnNumber
        End If
    Next
    
End Function
