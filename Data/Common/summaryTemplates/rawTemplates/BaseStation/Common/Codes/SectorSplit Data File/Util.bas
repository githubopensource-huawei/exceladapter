Attribute VB_Name = "Util"
Option Explicit

Public g_CurrentSheet As Worksheet
Public g_CurrentRange As range
Public g_Language As String
Public Const GRAY_COLOR = 16
Public Const g_ListShtColNameRow = 2
Public Const listShtGrpRow = 1
Public Const listShtAttrRow = 2

Function getColStr(ByVal NumVal As Long) As String
    Dim str As String
    Dim strs() As String

    If NumVal > 256 Or NumVal < 1 Then
        getColStr = ""
    Else
        str = Cells(NumVal).address
        strs = Split(str, "$", -1)
        getColStr = strs(1)
    End If
End Function

'Public Function colNumByAttr(ByRef ws As Worksheet, ByRef attrName As String) As Long
'    colNumByAttr = -1
'
'    Dim colIdx As Long
'    For colIdx = 1 To ws.range("IV2").End(xlToLeft).column
'        If ws.Cells(listShtAttrRow, colIdx) = attrName Then
'            colNumByAttr = colIdx
'            Exit Function
'        End If
'    Next
'End Function

Public Function colNumByAttr(sht As Worksheet, attrName As String, Optional row As Integer = 2) As Integer
    Dim target As range
    Set target = sht.range("A" & row & ":" & getColStr(getListShtUsedCol(sht)) & row).Find(what:=attrName, LookIn:=xlValues, lookat:=xlWhole)
    If target Is Nothing Then
        colNumByAttr = -1
    Else
        colNumByAttr = target.column
    End If
End Function

Public Function getListShtUsedCol(listSht As Worksheet) As Integer
    getListShtUsedCol = Application.WorksheetFunction.CountA(listSht.rows(listShtAttrRow))
End Function

Public Function GetMainSheetName() As String
       On Error Resume Next
        Dim name As String
        Dim rowNum As Long
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        For rowNum = 1 To sheetDef.range("a65536").End(xlUp).row
            If sheetDef.Cells(rowNum, 2).value = "MAIN" Then
                name = sheetDef.Cells(rowNum, 1).value
                Exit For
            End If
        Next
        GetMainSheetName = name
End Function

'从普通页取得Group name
Public Function get_GroupName(sheetName As String, colNum As Long) As String
        Dim index As Long
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For index = colNum To 1 Step -1
            If Not isEmpty(ws.Cells(1, index).value) Then
                get_GroupName = ws.Cells(1, index).value
                Exit Function
            End If
        Next
        get_GroupName = ""
End Function

'从普通页取得Colum name
Public Function get_ColumnName(ByVal sheetName As String, colNum As Long) As String
        Dim index As Long
        get_ColumnName = ThisWorkbook.Worksheets(sheetName).Cells(g_ListShtColNameRow, colNum)
End Function

Public Function isNum(contedType As String) As Boolean
    If contedType = "Integer" Or contedType = "UInteger" Then
        isNum = True
    Else
        isNum = False
    End If
End Function

Sub setBorders(ByRef certainRange As range)
    On Error Resume Next
    certainRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    certainRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    certainRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    certainRange.Borders(xlEdgeRight).LineStyle = xlContinuous
    certainRange.Borders(xlInsideVertical).LineStyle = xlContinuous
    certainRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    certainRange.Borders.LineStyle = xlContinuous
End Sub

Function getSheetType(sheetName As String) As String
    Dim m_rowNum As Long
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For m_rowNum = 2 To sheetDef.range("a65536").End(xlUp).row
        If sheetName = sheetDef.Cells(m_rowNum, 1).value Then
            getSheetType = UCase(sheetDef.Cells(m_rowNum, 2).value)
            Exit Function
        End If
    Next
    getSheetType = ""
End Function

Public Sub autoAdjustCommentFrame(ByRef ws As Worksheet)
    On Error Resume Next
    Dim maxCol As Integer
    maxCol = ws.UsedRange.columns.count
    
    Dim idx As Integer, rows As Double, rowHeight As Double
    For idx = 1 To maxCol
        With ws.Cells(2, idx).comment.Shape
            .TextFrame.AutoSize = True
            rows = .Width / 300#
            rowHeight = .Height / 5
            .Width = 300#
            .Height = (rows + 6) * rowHeight
        End With
    Next
End Sub

'将比较字符串整形
Function GetDesStr(srcStr As String) As String
    GetDesStr = UCase(Trim(srcStr))
End Function

Public Function getGroupNameFromMappingDef(ByRef sheetName As String, ByRef attrName As String) As String
    Dim mappingDef As Worksheet
    Dim index, count As Long
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    count = mappingDef.UsedRange.count
    For index = 1 To count
        If mappingDef.Cells(index, 1).value = sheetName And mappingDef.Cells(index, 3).value = attrName Then
            getGroupNameFromMappingDef = mappingDef.Cells(index, 2)
            Exit Function
        End If
    Next
    getGroupNameFromMappingDef = ""
End Function

Function Contains(coll As Collection, key As String) As Boolean
        On Error GoTo NotFound
        Call coll(key)
        Contains = True
        Exit Function
NotFound:
        Contains = False
End Function


'关闭批注的刷新，提高插入删除行的效率
Public Sub refreshComment(ByVal myRange As range)
    On Error Resume Next
    Dim cell As range
    For Each cell In myRange
        cell.comment.Shape.TextFrame.AutoSize = False
    Next
End Sub

Public Function getGroupName(ByRef ws As Worksheet, ByVal row As Integer) As String
    If Not isBoardStyleSheet(ws) Then
        getGroupName = ws.Cells(1, 1)
        Exit Function
    End If
    
    getGroupName = ws.Cells(getGrpRowForPattern(ws, row), 1)
End Function

Public Function getColumnName(ByRef ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long) As String
    If Not isBoardStyleSheet(ws) Then
        getColumnName = ws.Cells(g_ListShtColNameRow, colNum)
        Exit Function
    End If
    
    Dim grpNamePos As Long
    For grpNamePos = rowNum To 1 Step -1
        If grpNamePos = 1 Then Exit For
        If rowIsBlank(ws, grpNamePos - 1) = True And rowIsBlank(ws, grpNamePos) = False Then Exit For
    Next
    
    getColumnName = ws.Cells(grpNamePos + 1, colNum)
End Function

Public Function getGrpRowForPattern(ByRef ws As Worksheet, ByVal row As Long) As Long
    For getGrpRowForPattern = row To 1 Step -1
        If getGrpRowForPattern = 1 Then Exit For
        If rowIsBlank(ws, getGrpRowForPattern - 1) = True And rowIsBlank(ws, getGrpRowForPattern) = False Then Exit For
    Next
End Function

Public Function rowIsBlank(ByRef ws As Worksheet, ByRef rowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.range("A" & rowNumber & ":IV" & rowNumber)) = 256 Then
        rowIsBlank = True
    Else
        rowIsBlank = False
    End If
End Function

Public Function isUMTS() As Boolean
    isUMTS = False
    If existsSheet(getResByKey("UMTSCellSheet")) Then isUMTS = True
End Function

Public Function isLTE() As Boolean
    isLTE = False
    If existsSheet(getResByKey("LTECellSheet")) Then isLTE = True
End Function

Public Function existsSheet(ByRef shtName As String) As Boolean
    existsSheet = False
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = shtName Then
            existsSheet = True
            Exit Function
        End If
    Next
End Function

Public Function siteNameCol(ByRef ws As Worksheet) As Long
    siteNameCol = 1
    
    Dim col As Long
    For col = 1 To getUsedCol(ws)
        If ws.Cells(g_ListShtColNameRow, col) = getResByKey("eNodeBName") Or _
            ws.Cells(g_ListShtColNameRow, col) = getResByKey("NodeBName") Then
            siteNameCol = col
            Exit Function
        End If
    Next
End Function

Public Function colNum(ByRef ws As Worksheet, ByRef colName As String) As Long
    colNum = -1

    Dim col As Long
    For col = 1 To getUsedCol(ws)
        If ws.Cells(g_ListShtColNameRow, col) = colName Then
            colNum = col
            Exit Function
        End If
    Next

End Function

Public Function colNumForPatternSht(ByRef ws As Worksheet, ByRef colName As String, ByVal row As Long) As Long
    colNumForPatternSht = -1
    
    Dim grpNamePos As Long
    For grpNamePos = row To 1 Step -1
        If grpNamePos = 1 Then Exit For
        If rowIsBlank(ws, grpNamePos - 1) = True And rowIsBlank(ws, grpNamePos) = False Then Exit For
    Next
    
    Dim colNameRow As Long
    colNameRow = grpNamePos + 1
    
    Dim col As Long
    For col = 1 To getUsedCol(ws, colNameRow)
        If ws.Cells(colNameRow, col) = colName Then
            colNumForPatternSht = col
            Exit Function
        End If
    Next
End Function

Public Function getUsedRow(ByRef ws As Worksheet, Optional ByRef shtType As String = "LIST") As Long
    Dim row As Long
    Dim col As Long
    
    If UCase(shtType) = "LIST" Then
        col = siteNameCol(ws)
    Else
        col = 1
    End If
    
    getUsedRow = ws.range(getColStr(col) & "65535").End(xlUp).row
End Function

Public Function getUsedCol(ByRef ws As Worksheet, Optional ByVal row As Long = 2) As Long
    'Dim row As Long
    Dim col As Long
    
    'row = 2
    getUsedCol = ws.range("IV" & row).End(xlToLeft).column
End Function

Public Function isCellSheet(ByRef ws As Worksheet) As Boolean
    isCellSheet = False
    If ws.name = getResByKey("UMTSCellSheet") Or ws.name = getResByKey("LTECellSheet") Then isCellSheet = True
End Function

Public Function isBoardStyleSheet(ByRef ws As Worksheet) As Boolean
    Dim sheetName As String, boardStyleSheetName As String
    isBoardStyleSheet = False
    boardStyleSheetName = getResByKey("Board Style")
    If InStr(ws.name, boardStyleSheetName) <> 0 Then
        isBoardStyleSheet = True
    End If
End Function

Public Sub initLanguage()
On Error GoTo ErrorHandler
    Dim coverSht As Worksheet
    Set coverSht = ThisWorkbook.Worksheets("Cover")
    g_Language = "EN"
    Exit Sub
ErrorHandler:
    g_Language = "CN"
End Sub





