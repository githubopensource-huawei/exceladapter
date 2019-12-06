Attribute VB_Name = "BoardStyleUtil"

'---------------------------------------  判断是否GroupName行  --------------------------------------

Public Function isGroupRow(ByRef ws As Worksheet, ByVal row As Long) As Boolean
    
    isGroupRow = False
    
    If findGroupName(Trim(ws.Cells(row, 1))) = True And Trim(ws.Cells(row, 1)) <> "" Then
        isGroupRow = True
    End If

End Function

Public Function isColumnRow(ByRef ws As Worksheet, ByVal row As Long) As Boolean
    
    isColumnRow = False
    
    If findAttrName(Trim(ws.Cells(row, 1))) = True And Trim(ws.Cells(row, 1)) <> "" Then
        isColumnRow = True
    End If

End Function

Public Function getMaxRow(ByRef ws As Worksheet)
     getMaxRow = ws.UsedRange.Rows.count
End Function

Public Function getMaxCol(ByRef ws As Worksheet, ByVal row As Long)
     getMaxCol = ws.Range("XFD" & row).End(xlToLeft).column
End Function


'--------------  根据行号、单元格内容，找单元格所在列号  ------------------------

Public Function findColNumByRowAndValue(ByRef ws As Worksheet, ByVal row As Long, ByRef cellVal As Variant, Optional ByVal colNumBegin As Long = 1) As Long
    Dim colMax As Long
    Dim curRowRange As Range
    
    findColNumByRowAndValue = -1
    
    colMax = getMaxCol(ws, row)
    If colNumBegin > colMax Then Exit Function
    
    Set curRowRange = ws.Range(ws.Cells(row, colNumBegin), ws.Cells(row, colMax))
        
    Dim colRst
    On Error Resume Next
    colRst = Application.WorksheetFunction.Match(cellVal, curRowRange, 0) '这里找到是设置的查找区域内的相对列数
        
    If Not IsError(colRst) And Not isEmpty(colRst) Then
        findColNumByRowAndValue = colNumBegin + colRst - 1
    End If
    
End Function


Public Function findColLetterByRowAndValue(ByRef ws As Worksheet, ByVal row As Long, ByRef cellVal As Variant, Optional ByVal colNumBegin As Long = 1) As String
    findColLetterByRowAndValue = ""

    Dim colNum As Long
    colNum = findColNumByRowAndValue(ws, row, cellVal, colNumBegin)

    If colNum <> -1 Then
        findColLetterByRowAndValue = getColStr(colNum)
    End If

End Function

'--------------  根据RowNum找到所在的GroupNameRow  ------------------------

Public Function getGroupRowNum(ByVal ws As Worksheet, rowNum As Long) As Long
    getGroupRowNum = -1

    Dim m_rowNum As Long
    If ws.name = getResByKey("Comm Data") Or InStr(ws.name, getResByKey("Board Style")) <> 0 Then
        For m_rowNum = rowNum To 1 Step -1
            If isGroupRow(ws, m_rowNum) Then
                getGroupRowNum = m_rowNum
                Exit Function
            End If
        Next
    Else
        getGroupRowNum = 1
    End If
End Function

'--------------  根据GroupName和ColName，查找ColName属性所在的列  ------------------------

'只在指定的Row所在的Group区域内搜索

Public Function findColNumByGrpAndColName(ByRef ws As Worksheet, ByRef groupName As Variant, ByRef columnName As Variant, ByVal row As Long) As Long
    findColNumByGrpAndColName = -1

    Dim groupRow As Long, columnRow As Long
    groupRow = getGroupRowNum(ws, row)
    If groupRow = -1 Then Exit Function

    columnRow = groupRow + 1

    Dim cellRange As Range
    Dim colMax As Long, index As Long, colNum As Long
    Dim curGroupName As String, curColumnName As String

    index = 1
    colMax = getMaxCol(ws, columnRow)
    Do While index <= colMax

        colNum = findColNumByRowAndValue(ws, columnRow, columnName, index)
        If colNum = -1 Then Exit Function

        curGroupName = ""
        curColumnName = ""
        Call getGroupAndColumnName(ws, ws.Cells(columnRow, colNum), curGroupName, curColumnName)

        If curGroupName <> "" And curGroupName = groupName And ws.Cells(columnRow, colNum) = columnName Then
            findColNumByGrpAndColName = colNum
            Exit Function
        End If

        index = colNum + 1
    Loop

End Function


Public Function findColLetterByGrpAndColName(ByRef ws As Worksheet, ByRef groupName As Variant, ByRef columnName As Variant, ByVal row As Long) As String
    Dim rst As Long
    rst = findColNumByGrpAndColName(ws, groupName, columnName, row)
    
    If rst = -1 Then
        findColLetterByGrpAndColName = ""
    Else
        findColLetterByGrpAndColName = getColStr(rst)
    End If
End Function


'未指定行号时，在整个sheet页内搜索

'TODO:为避免在整个BoardStyle页签内搜索，先尝试在当前单元格所在的Group范围内查找

'Public Function findColNumByGrpAndColNameEx(ByRef ws As Worksheet, ByRef groupName As Variant, ByRef columnName As Variant) As Long
'    findColNumByGrpAndColNameEx = -1
'
'    Dim curRange As Range
'    Dim curGroupNameRowNum As Long
'
''    If ThisWorkbook.ActiveSheet.name = ws.name And Selection.Areas.count = 1 Then
''        Set curRange = Selection.Areas
''
''        curGroupNameRowNum = getGroupRowNum(ws, curRange.row)
''        findColNumByGrpAndColNameEx = findColNumByGrpAndColName(ws, curGroupNameRowNum + 1, groupName, columnName)
''
''        If findColNumByGrpAndColNameEx <> -1 Then Exit Function
''    End If
'
'    Dim m_rowNum As Long
'
'    For m_rowNum = 1 To getMaxRow(ws)
'        If Not isGroupRow(ws, m_rowNum) Then GoTo NextLoop
'
'        findColNumByGrpAndColNameEx = findColNumByGrpAndColName(ws, groupName, columnName, m_rowNum + 1)
'
'        If findColNumByGrpAndColNameEx <> -1 Then Exit Function
'NextLoop:
'    Next
'
'End Function


'-----------------------------   找无线搬迁BoardStyle页签中，根据行号找到SrcName所在列 --------------------------

'指定GroupName时

'Public Function getSrcNeNameColWithGrpName(ByRef ws As Worksheet, ByVal row As Long, ByVal groupName As String) As String
'    getSrcNeNameColWithGrpName = ""
'
'    Dim sourceNeName As String, sourceNeNamePos As String
'    sourceNeName = getResByKey("SOURCENENAME")
'
'    getSrcNeNameColWithGrpName = findColLetterByGrpAndColName(ws, groupName, sourceNeName, row)
'
'End Function
'
''未指定GroupName时
'
'Public Function getSrcNeNameCol(ByRef ws As Worksheet, ByVal row As Long) As String
'    getSrcNeNameCol = ""
'
'    Dim sourceNeNameCol As String, sourceNeNameLetter As String
'    sourceNeNameCol = getResByKey("SOURCENENAME")
'
'    '找到GroupName所在的行
'    Dim groupRowNum As Long
'    groupRowNum = getGroupRowNum(ws, row)
'    If groupRowNum = -1 Then Exit Function
'
'    getSrcNeNameCol = findColLetterByRowAndValue(ws, groupRowNum + 1, sourceNeNameCol)
'
'End Function


'判断BoardStyle页签Group中第rowNum行是否是siteName（FuntionName或NeName）规划的场景

'Public Function isCustomMatchRow(ByVal siteName As String, ByRef boardStyleSheet As Worksheet, ByRef groupName As String, ByRef rowNum As Long) As Boolean
'    isCustomMatchRow = False
'
'    Dim baseStationData As CBaseStationData
'    Set baseStationData = New CBaseStationData
'    Dim boardStyleData As CBoardStyleData
'    Set boardStyleData = New CBoardStyleData
'    Call baseStationData.init
'    Call boardStyleData.init
'
'    '找到归属的neName
'    Dim neName As String
'    neName = baseStationData.getNeNamebyFunctionName(siteName)
'    If neName = "" Then Exit Function
'
'    '找此NeName定制的场景信息, key: CustomizeGroupName\CustomizeColumnName，customColValue
'    Dim neCustomInfoMap As CMap
'    If Not baseStationData.getCustomInfoMapByNeName(neName, neCustomInfoMap) Then
'        isCustomMatchRow = True
'        Exit Function
'    End If
'
'    Dim mocName As String
'    mocName = boardStyleData.getMocNameByGroupName(groupName)
'    If mocName = "" Then Exit Function
'
'    '找BoardStyle页签上定制的场景信息，key: customGroupName+"\"+customColumnName,  columnLetter
'    Dim boardStyleCustomInfoMap As CMap
'    If Not boardStyleData.getBoardStyleCustomInfoMap(mocName, groupName, boardStyleCustomInfoMap) Then
'        isCustomMatchRow = True
'        Exit Function
'    End If
'
'    Dim boardStyleCellVal As String
'    Dim boardStyleCellLetter As String
'    Dim baseStationCellVal As String
'
'    Dim boardStyleCustomInfoVar
'    For Each boardStyleCustomInfoVar In boardStyleCustomInfoMap.KeyCollection
'        '每个循环代表一个场景列判断，必需所有场景列都能与基站传输页的场景列匹配
'
'        '取BoardStyle页签上场景列的值
'        boardStyleCellLetter = boardStyleCustomInfoMap.GetAt(CStr(boardStyleCustomInfoVar))
'        boardStyleCellVal = Trim(boardStyleSheet.Range(boardStyleCellLetter & rowNum).value)
'        If boardStyleCellVal = "" Then GoTo NextLoop
'
'        '取基站传输页上场景列的值
'        If Not neCustomInfoMap.haskey(CStr(boardStyleCustomInfoVar)) Then GoTo NextLoop
'
'        baseStationCellVal = neCustomInfoMap.GetAt(CStr(boardStyleCustomInfoVar))
'        If Not isCustomMatch(boardStyleCellVal, baseStationCellVal) Then Exit Function
'NextLoop:
'    Next
'
'    isCustomMatchRow = True
'End Function


'检查是否场景匹配，原则：boardStyle中任一个值在基站场景信息中

'Public Function isCustomMatch(ByVal boardStyleCustomVal As String, ByVal baseStationCustomVal As String) As Boolean
'    isCustomMatch = False
'
'    Dim baseStationValues As Collection
'    Call SplitWithTrim(baseStationCustomVal, ",", baseStationValues)
'
'    Dim boardStyleValArray() As String
'    boardStyleValArray = Split(boardStyleCustomVal, ",")
'
'    Dim index As Long
'    For index = LBound(boardStyleValArray) To UBound(boardStyleValArray)
'        If isInCollection(baseStationValues, Trim(boardStyleValArray(index))) Then
'            isCustomMatch = True
'            Exit Function
'        End If
'    Next
'
'End Function
    
    
Public Function isReferenceValue(ByRef cellValue As String) As Boolean
    If UBound(Split(cellValue, "\")) = 2 Then
        isReferenceValue = True
    Else
        isReferenceValue = False
    End If
End Function
