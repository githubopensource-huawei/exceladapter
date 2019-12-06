Attribute VB_Name = "BoardStyleSub"
Option Explicit

Public Function findBoardStyleCol(ByRef transportSht As Worksheet) As Integer
    Dim colIdx As Integer
    For colIdx = 1 To getUsedCol(transportSht)
        If transportSht.Cells(2, colIdx) = getResByKey("BoardStyleName") Then
            findBoardStyleCol = colIdx
            Exit Function
        End If
    Next
    findBoardStyleCol = 3
End Function

Public Sub setBgColorWithGray_i(ByRef bStyleSht As Worksheet)
    Dim maxRow As Integer
    Dim rfuGrpRow As Integer
    Dim setBeginRow As Integer
    
    With bStyleSht
        maxRow = .range("a65535").End(xlUp).row
        For rfuGrpRow = 2 To maxRow
            If .Cells(rfuGrpRow, 1) = getResByKey("RFU") Or .Cells(rfuGrpRow, 1) = "RFU" Then Exit For
        Next
        setBeginRow = rfuGrpRow
        If setBeginRow > 1 Then setBeginRow = setBeginRow - 1
        
        .range("A" & setBeginRow & ":A" & (maxRow + 1)).rows.EntireRow.Interior.colorIndex = SolidColorIdx
        
        .Cells(1, 1).Select
    End With
End Sub

Public Sub expandRowInBoardStyleSheet()
    If Selection.EntireRow.Interior.colorIndex = SolidColorIdx Then
        MsgBox (getResByKey("ExpandReadOnlyRange"))
        Exit Sub
    End If
    
    Dim grpBeginRow As Long, grpEndRow As Long, curRow As Long
    grpEndRow = -1
    curRow = Selection.row
    
    Dim curSheet As Worksheet
    Set curSheet = ThisWorkbook.ActiveSheet
    
    grpBeginRow = getGrpRowForPattern(curSheet, curRow)
    
    Dim row As Long
    For row = grpBeginRow To getUsedRow(curSheet, "Pattern")
        If curSheet.Cells(row, 1) <> "" Then
            grpEndRow = row
        Else
            Exit For
        End If
    Next
    
    If grpEndRow - grpBeginRow < 1 Then Exit Sub
    
    Dim newRow As Long, newCol As Long
    newRow = grpEndRow + 1
    newCol = getUsedCol(curSheet, grpBeginRow)
    
    curSheet.rows(CStr(newRow) & ":" & CStr(newRow)).Select
    If grpEndRow - grpBeginRow = 1 Then
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Else
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    Selection.ClearContents
    Selection.Interior.Pattern = xlNone
    
    Dim newRange As range
    Set newRange = range("A" & CStr(newRow), C(newCol) & CStr(newRow))
    Call setBorders(newRange)
    Call setKeyAttrBgColor(curSheet, grpBeginRow, newRow)
End Sub

Public Sub deleteRowInBoardStyleSheet()
    If Selection.EntireRow.Interior.colorIndex = SolidColorIdx Then
        MsgBox (getResByKey("DelReadOnlyRange"))
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    Dim rg As range
    For Each rg In Selection.rows
        If Not isDataRow(ws, rg.row) Then
            MsgBox (getResByKey("DelNotDataRow"))
            Exit Sub
        End If
    Next
    
    Dim beginRow As Long, rowCount As Long
    beginRow = Selection.row
    rowCount = Selection.rows.count
    
    If lastDataRowDeleted(ws) Then
        ws.rows(CStr(beginRow) & ":" & CStr(beginRow)).Select
        Selection.ClearContents
        Selection.Interior.Pattern = xlNone
        beginRow = beginRow + 1
        rowCount = rowCount - 1
    End If
    
     If rowCount < 1 Then Exit Sub

    ws.rows(CStr(beginRow) & ":" & CStr(beginRow + rowCount - 1)).Select
    Selection.EntireRow.Delete
End Sub

Public Sub boardStyleSheetChange(ByRef ws As Worksheet, ByRef target As range)
    On Error Resume Next
    If getGroupName(ws, target.row) <> getResByKey("BBP") Then Exit Sub
    
    Dim colName As String
    colName = getColumnName(ws, target.row, target.column)
    
    Dim boardNoCol As Long, cnCol As Long, srnCol As Long, snCol As Long
    Dim Cn As String, srn As String, sn As String
    
    If colName = getResByKey("CN") Or colName = getResByKey("SRN") Or colName = getResByKey("SN") Then
        boardNoCol = colNumForPatternSht(ws, getResByKey("*BoardNo"), target.row)
        cnCol = colNumForPatternSht(ws, getResByKey("CN"), target.row)
        srnCol = colNumForPatternSht(ws, getResByKey("SRN"), target.row)
        snCol = colNumForPatternSht(ws, getResByKey("SN"), target.row)
        If boardNoCol = -1 Or cnCol = -1 Or srnCol = -1 Or snCol = -1 Then Exit Sub
        
        Cn = ws.Cells(target.row, cnCol)
        srn = ws.Cells(target.row, srnCol)
        sn = ws.Cells(target.row, snCol)
        
        If Cn = "" And srn = "" And sn = "" Then
            ws.Cells(target.row, boardNoCol) = ""
            Exit Sub
        End If
        
        ws.Cells(target.row, boardNoCol) = ws.Cells(target.row, cnCol) & "_" & _
            ws.Cells(target.row, srnCol) & "_" & ws.Cells(target.row, snCol) & "_1"
    End If
End Sub

Public Function isDataRow(ByRef ws As Worksheet, ByVal row As Long) As Boolean
    isDataRow = True
    
    Dim grpNameRow As Long
    grpNameRow = getGrpRowForPattern(ws, row)
    
    Dim colNameRow As Long
    colNameRow = grpNameRow + 1
    
    If row = grpNameRow Or row = colNameRow Then
        isDataRow = False
        Exit Function
    End If
    
    If rowIsBlank(ws, row) And Not rowIsBlank(ws, row + 1) Then
        isDataRow = False
        Exit Function
    End If
End Function

Public Function lastDataRowDeleted(ByRef ws As Worksheet) As Boolean
    lastDataRowDeleted = False
    
    Dim firstRow As Long
    firstRow = Selection.row
    
    Dim grpNameRow As Long
    grpNameRow = getGrpRowForPattern(ws, firstRow)
    
    Dim colNameRow As Long
    colNameRow = grpNameRow + 1
    
    Dim lastDataRow As Long
    For lastDataRow = firstRow To 65535
        If rowIsBlank(ws, lastDataRow + 1) And Not rowIsBlank(ws, lastDataRow + 2) Then 'last data row
            Exit For
        End If
    Next
    
    If (lastDataRow - colNameRow) = Selection.rows.count Then lastDataRowDeleted = True
End Function

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

Private Sub setKeyAttrBgColor(ByRef ws As Worksheet, ByVal grpRow As Integer, ByVal row As Integer)
    Dim grpName As String
    grpName = ws.Cells(grpRow, 1)
    
    If grpName = getResByKey("BBP") Then
        ws.range("B" & CStr(row) & ":G" & CStr(row)).Interior.colorIndex = BLUE_COLOR
    ElseIf grpName = getResByKey("BBEqm") Then
        ws.range("A" & CStr(row) & ":D" & CStr(row)).Interior.colorIndex = BLUE_COLOR
    End If
End Sub






