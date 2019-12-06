Attribute VB_Name = "Util"
Option Explicit

Public currentSheet As Worksheet
Public currentRange As range
Public freqInfoCol As Integer
Public customFreqCol As Integer
Public bbEqmBoardNoCol As Integer
Public Const boardNoCol = 1

Public Const GRAY_COLOR = 16
Public Const COLUMN_NAME_COLOR = 40
Public Const BLUE_COLOR = 33
Public Const g_ListShtColNameRow = 2

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
        get_ColumnName = ThisWorkbook.Worksheets(sheetName).Cells(2, colNum)
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

Public Function customFreqSelected(ByRef sh As Worksheet, ByRef target As range) As Boolean
    customFreqSelected = False
    Call getCustomFreqCol
    
    If sh.name = getResByKey("USectorSheet") And target.column = customFreqCol Then
        customFreqSelected = True
    End If
End Function

Public Function bbEqmBoardNoSelected(ByRef sh As Worksheet, ByRef target As range) As Boolean
    bbEqmBoardNoSelected = False
    Call getBBEqmBoardNoCol

    Dim bbEqmDataBeginRow As Integer
    Dim bbEqmDataEndRow As Integer
    bbEqmDataBeginRow = getBasebandEqmGrpRow(sh) + 2
    For bbEqmDataEndRow = bbEqmDataBeginRow To sh.UsedRange.rows.count
        If rowIsBlank(sh, bbEqmDataEndRow) Then Exit For
    Next
    bbEqmDataEndRow = bbEqmDataEndRow - 1

    If isBoardStyleSheet(currentSheet) And target.column = bbEqmBoardNoCol And target.row <= bbEqmDataEndRow And target.row >= bbEqmDataBeginRow Then
        bbEqmBoardNoSelected = True
    End If
End Function

Public Sub getBBEqmBoardNoCol()
    Dim bbEqmTitleRow As Integer
    bbEqmTitleRow = getBasebandEqmGrpRow(currentSheet) + 1
    
    Dim col, maxCol As Integer
    maxCol = currentSheet.range("IV" & bbEqmTitleRow).End(xlToLeft).column

    bbEqmBoardNoCol = 4
    For col = 1 To maxCol
        If currentSheet.Cells(bbEqmTitleRow, col) = getResByKey("BBEqmBoardNo") Then
            bbEqmBoardNoCol = col
            Exit Sub
        End If
    Next
End Sub

Public Sub getFreqInfoCol()
    Dim col, maxCol As Integer
    maxCol = currentSheet.range("IV2").End(xlToLeft).column

    freqInfoCol = 3
    For col = 1 To maxCol
        If currentSheet.Cells(2, col) = getResByKey("DLFreq") Then
            freqInfoCol = col
            Exit Sub
        End If
    Next
End Sub

Public Sub getCustomFreqCol()
    Dim col, maxCol As Integer
    maxCol = currentSheet.range("IV2").End(xlToLeft).column
    
    customFreqCol = 5
    For col = maxCol To 1 Step -1
        If currentSheet.Cells(2, col) = getResByKey("DLFreq") Then
            customFreqCol = col
            Exit Sub
        End If
    Next
End Sub

Public Function getBoardNos() As String
    Dim tmpStr As String
    Dim rowIdx As Integer
    For rowIdx = 3 To getBasebandEqmGrpRow(currentSheet) - 1
        Dim BoardNo As String
        BoardNo = currentSheet.Cells(rowIdx, boardNoCol)
        If BoardNo <> "" Then tmpStr = tmpStr & BoardNo & ","
    Next
    
    If tmpStr <> "" Then
        getBoardNos = Left(tmpStr, Len(tmpStr) - 1)
    End If
End Function

Function findBoardStyleCol(ByRef transportSht As Worksheet) As Integer
    Dim colIdx As Integer
    For colIdx = 1 To transportSht.range("IV2").End(xlToLeft).column
        If transportSht.Cells(2, colIdx) = getResByKey("BoardStyleName") Then
            findBoardStyleCol = colIdx
            Exit Function
        End If
    Next
    findBoardStyleCol = 3
End Function

Public Function getBasebandEqmGrpRow(ByRef boardStyleSht As Worksheet) As Integer
    getBasebandEqmGrpRow = 7
    Dim rowIdx As Integer
    
    For rowIdx = 1 To boardStyleSht.UsedRange.rows.count
        If boardStyleSht.Cells(rowIdx, 1) = getResByKey("BBEqm") Then
            getBasebandEqmGrpRow = rowIdx
            Exit Function
        End If
    Next
End Function

Public Function rowIsBlank(ByRef ws As Worksheet, ByVal rowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.range("A" & rowNumber & ":IV" & rowNumber)) = 256 Then
        rowIsBlank = True
    Else
        rowIsBlank = False
    End If
End Function

Sub initMenuStatus(sh As Worksheet)
    If isBoardStyleSheet(sh) Then
        Call insertAndDeleteControl(False)
    ElseIf sh.name = getResByKey("Temp Sheet") Then
        Call initTempSheetControl(False)
    Else
        Call initTempSheetControl(True)
        Call insertAndDeleteControl(True)
    End If
End Sub

Sub destroyMenuStatus()
    With Application
        .CommandBars("Row").Reset
        .CommandBars("Column").Reset
        .CommandBars("Cell").Reset
        .CommandBars("Ply").Reset
    End With
End Sub
 
'从指定sheet页的指定行，查找指定列，返回列号,通过attrName和MocName获取到
Public Function getColNum(sheetName As String, recordRow As Long, attrName As String, mocName As String) As Long
    On Error Resume Next
    Dim m_colNum As Long
    Dim m_rowNum As Long
    Dim colName As String
    Dim colGroupName As String
    
    Dim flag As Boolean
    Dim mappingDef As Worksheet
    Dim ws As Worksheet
    
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    flag = False
    getColNum = -1
    For m_rowNum = 2 To mappingDef.range("a65536").End(xlUp).row
        If UCase(attrName) = UCase(mappingDef.Cells(m_rowNum, 5).value) _
           And UCase(sheetName) = UCase(mappingDef.Cells(m_rowNum, 1).value) _
           And UCase(mocName) = UCase(mappingDef.Cells(m_rowNum, 4).value) Then
            colName = mappingDef.Cells(m_rowNum, 3).value
            colGroupName = mappingDef.Cells(m_rowNum, 2).value
            flag = True
            Exit For
        End If
    Next
    If flag = True Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For m_colNum = 1 To ws.range("IV" + CStr(recordRow)).End(xlToLeft).column
            If get_GroupName(sheetName, m_colNum) = colGroupName Then
                If GetDesStr(colName) = GetDesStr(ws.Cells(recordRow, m_colNum).value) Then
                    getColNum = m_colNum
                    Exit For
                End If
            End If
        Next
    End If
End Function

Public Function max(ByVal a, ByVal b)
    max = IIf(a > b, a, b)
End Function

Function isCellSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("UMTSCellSheet") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function

Public Function getUsedCol(ByRef ws As Worksheet, Optional ByVal row As Long = 2) As Long
    'Dim row As Long
    Dim col As Long
    
    'row = 2
    getUsedCol = ws.range("IV" & row).End(xlToLeft).column
End Function

Public Function getUsedRow(ByRef ws As Worksheet, Optional ByRef shtType As String = "LIST") As Long
    getUsedRow = ws.range("A65535").End(xlUp).row
End Function

Public Function getGrpRowForPattern(ByRef ws As Worksheet, ByVal row As Long) As Long
    For getGrpRowForPattern = row To 1 Step -1
        If getGrpRowForPattern = 1 Then Exit For
        If rowIsBlank(ws, getGrpRowForPattern - 1) = True And rowIsBlank(ws, getGrpRowForPattern) = False Then Exit For
    Next
End Function

'=================================================
'从列数得到列名：1->A，27->AA
'=================================================
Public Function C(ByVal iColumn As Long) As String
  If iColumn >= 257 Or iColumn < 0 Then
    C = ""
    Return
  End If
  
  Dim result As String
  Dim High, Low As Long
  
  High = Int((iColumn - 1) / 26)
  Low = iColumn Mod 26
  
  If High > 0 Then
    result = Chr(High + 64)
  End If

  If Low = 0 Then
    Low = 26
  End If
  
  result = result & Chr(Low + 64)
  C = result
End Function

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

Public Function getNodeName(ratName As String) As String
On Error Resume Next
    getNodeName = ratName
    
    Dim transportSht As Worksheet
    Set transportSht = ThisWorkbook.Worksheets(GetMainSheetName())
    
    Dim nodeNameCol As Integer
    nodeNameCol = getColNumByName(transportSht, getResByKey("*Name"))
    If nodeNameCol = -1 Then Exit Function
    
    Dim ratNameCol As Integer
    ratNameCol = getColNumByName(transportSht, getResByKey("*NodeBName"))
        
    Dim rowIdx As Integer
    For rowIdx = 3 To transportSht.range("a65536").End(xlUp).row
        If transportSht.Cells(rowIdx, ratNameCol) = ratName Then
            getNodeName = transportSht.Cells(rowIdx, nodeNameCol)
            Exit For
        End If
    Next
End Function

Public Function getColNumByName(sht As Worksheet, attrName As String) As Integer
On Error Resume Next
    getColNumByName = -1
    
    Dim col As Integer
    For col = 1 To sht.range("IV" & g_ListShtColNameRow).End(xlToLeft).column
        If sht.Cells(g_ListShtColNameRow, col) = attrName Then
            getColNumByName = col
            Exit Function
        End If
    Next
End Function








