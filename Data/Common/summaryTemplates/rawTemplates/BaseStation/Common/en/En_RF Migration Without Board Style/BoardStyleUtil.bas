Attribute VB_Name = "BoardStyleUtil"


Public boardStylePosMgr As CBoardStylePositionMgr
Private boardStyleMappingDefMap_ As CMapValueObject


Public Sub initBoardStylePosMgrPublic()
    If boardStylePosMgr Is Nothing Then
        Set boardStylePosMgr = New CBoardStylePositionMgr
        Call boardStylePosMgr.init
    End If
End Sub


'groupNameEndRowNumber是整个Group最后一行非空数据所在行，用于插入数据时使用
Public Sub getValidGroupRangeRows(ByRef ws As Worksheet, ByRef groupName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
    
    Call getGroupStartAndEndRowByGroupName(ws, groupName, groupNameStartRowNumber, groupNameEndRowNumber)
    
    Dim index As Long
    For index = groupNameEndRowNumber To groupNameStartRowNumber Step -1
        If Not rowIsBlank(ws, index) = True Then
            Exit For
        End If
    Next index

    groupNameEndRowNumber = index
End Sub


'groupNameEndRowNumber是整个Group最后一行位置（与下一个Group行中间间隔一个空行），如果是最后一个Group则是最后一个有边界行，用于检查数据范围使用
Public Sub getGroupStartAndEndRowByGroupName(ByRef ws As Worksheet, ByRef groupName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)

    Call initBoardStylePosMgrPublic
    Call boardStylePosMgr.getGroupStartAndEndRowByGroupName(ws, groupName, groupNameStartRowNumber, groupNameEndRowNumber)

End Sub

'groupNameEndRowNumber是整个Group最后一行位置（与下一个Group行中间间隔一个空行），如果是最后一个Group则是最后一个有边界行，用于检查数据范围使用
Public Sub getGroupStartAndEndRowByRowNum(ByRef ws As Worksheet, ByVal rowNumber As Long, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
    
    Call initBoardStylePosMgrPublic
    Call boardStylePosMgr.getGroupStartAndEndRowByRowNum(ws, rowNumber, groupNameStartRowNumber, groupNameEndRowNumber)

End Sub

'根据RowNumber取模型上的groupName和columnName（排除操作符列或者是定制信息列，返回的是第一个模型属性所在列）
Public Function getModelGroupAndColumnNameByRow(ByRef ws As Worksheet, ByVal rowNumber As Long, ByRef groupName As String, ByRef columnName As String) As Long
    
    Call initBoardStylePosMgrPublic
    getModelGroupAndColumnNameByRow = boardStylePosMgr.getModelGroupAndColumnNameByRow(ws, rowNumber, groupName, columnName)

End Function

'根据RowNumber取模型上的groupName所起始在列
Public Function getModelGroupStarRowByRow(ByRef ws As Worksheet, ByVal rowNumber As Long) As Long
    Dim groupName As String, columnName As String
    
    Call initBoardStylePosMgrPublic
    getModelGroupStarRowByRow = boardStylePosMgr.getModelGroupAndColumnNameByRow(ws, rowNumber, groupName, columnName)

End Function


'---------------------------------------  判断是否GroupName行  --------------------------------------
Public Function isGroupRow(ByRef ws As Worksheet, ByVal row As Long) As Boolean
    Call initBoardStylePosMgrPublic
    isGroupRow = boardStylePosMgr.isGroupRow(ws, row)
End Function

Public Function isColumnRow(ByRef ws As Worksheet, ByVal row As Long) As Boolean
    Call initBoardStylePosMgrPublic
    isColumnRow = boardStylePosMgr.isColumnRow(ws, row)
End Function


Public Function getMaxRow(ByRef ws As Worksheet)
     getMaxRow = ws.UsedRange.Rows.count
End Function

Public Function getMaxCol(ByRef ws As Worksheet, ByVal row As Long)
    getMaxCol = ws.range("XFD" & row).End(xlToLeft).column
End Function


'--------------  根据行号、单元格内容，找单元格所在列号  ------------------------

Public Function findColNumByRowAndValue(ByRef ws As Worksheet, ByVal row As Long, ByRef cellVal As Variant, Optional ByVal colNumBegin As Long = 1) As Long
    Dim colMax As Long
    Dim curRowRange As range
    
    findColNumByRowAndValue = -1
    
    colMax = getMaxCol(ws, row)
    If colNumBegin > colMax Then Exit Function
    
    Set curRowRange = ws.range(ws.Cells(row, colNumBegin), ws.Cells(row, colMax))
        
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
'任何GroupName、AttrName都可以找

Public Function findColNumByGrpAndColName(ByRef ws As Worksheet, ByRef groupName As Variant, ByRef columnName As Variant, ByVal row As Long) As Long
    findColNumByGrpAndColName = -1

    Dim groupRow As Long, columnRow As Long

    '根据row向上找到最近的GroupRowNum
    groupRow = getGroupRowNum(ws, row)
    If groupRow = -1 Then Exit Function

    '最近的columnRowNum=GroupRowNum+1
    columnRow = groupRow + 1

    '找到GroupName所在的ColumnNum
    Dim groupNameCol As Long
    groupNameCol = findColNumByRowAndValue(ws, groupRow, groupName)
    If groupNameCol = -1 Then Exit Function

    '从GroupName所在列开始，找ColumnName所在列
    findColNumByGrpAndColName = findColNumByRowAndValue(ws, columnRow, columnName, groupNameCol)

End Function

'根据GroupName、ColumnName确认ColumnName所在的列号，不需要指定RowNumber
'重要！！！！：只适合查模型MocName,不支持查询定制Moc、属性

Public Function findColNumByGrpAndColNameEx(ByRef ws As Worksheet, ByRef groupName As Variant, ByRef columnName As Variant) As Long
    findColNumByGrpAndColNameEx = -1
    
    If boardStyleData Is Nothing Then Call initBoardStyleMappingDataPublic
    Set boardStyleMappingDefMap_ = boardStyleData.getBoardStyleMappingDefMap
    
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
    
    Dim columnLetter As String
    If boardStyleMappingDefData.columnNamePositionLetterMap.hasKey(CStr(columnName)) Then
    
        columnLetter = boardStyleMappingDefData.getColumnNamePositionLetter(CStr(columnName))
        findColNumByGrpAndColNameEx = CellCol2Int(columnLetter)
    End If
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


'-----------------------------   找无线搬迁BoardStyle页签中，根据行号找到SrcName所在列 --------------------------

'指定GroupName时

Public Function getSrcNeNameColWithGrpName(ByRef ws As Worksheet, ByVal row As Long, ByVal groupName As String) As String
    getSrcNeNameColWithGrpName = ""
    
    Dim sourceNeName As String, sourceNeNamePos As String
    sourceNeName = getResByKey("SOURCENENAME")

    getSrcNeNameColWithGrpName = findColLetterByGrpAndColName(ws, groupName, sourceNeName, row)
       
End Function

'未指定GroupName时

Public Function getSrcNeNameCol(ByRef ws As Worksheet, ByVal row As Long) As String
    getSrcNeNameCol = ""
    
    Dim sourceNeNameCol As String, sourceNeNameLetter As String
    sourceNeNameCol = getResByKey("SOURCENENAME")
    
    '找到GroupName所在的行
    Dim groupRowNum As Long
    groupRowNum = getGroupRowNum(ws, row)
    If groupRowNum = -1 Then Exit Function
    
    getSrcNeNameCol = findColLetterByRowAndValue(ws, groupRowNum + 1, sourceNeNameCol)

End Function

    
Public Function getStartLetter(ByRef ws As Worksheet, ByVal row As Long) As String
    
    getStartLetter = ""
    
    Dim startCol As Long
    startCol = getModelGroupStarRowByRow(ws, row)
    
    getStartLetter = getColStr(startCol)

    
End Function



'判断BoardStyle页签Group中第rowNum行是否是siteName（FuntionName或NeName）规划的场景

Public Function isCustomMatchRow(ByVal siteName As String, ByRef boardStyleSheet As Worksheet, ByRef groupName As String, ByRef rowNum As Long) As Boolean
    isCustomMatchRow = False

    '数据初始化不要放在循环中执行，影响性能
'    Dim baseStationData As CBaseStationData
'    Set baseStationData = New CBaseStationData
'    Dim boardStyleData As CBoardStyleData
'    Set boardStyleData = New CBoardStyleData
'    Call baseStationData.init
'    Call boardStyleData.init
    
    '找到归属的neName
    Dim neName As String
    neName = baseStationData.getNeNamebyFunctionName(siteName)
    If neName = "" Then Exit Function
        
    '找此NeName定制的场景信息, key: CustomizeGroupName\CustomizeColumnName，customColValue
    Dim neCustomInfoMap As CMap
    If Not baseStationData.getCustomInfoMapByNeName(neName, neCustomInfoMap) Then
        isCustomMatchRow = True
        Exit Function
    End If
        
    Dim mocName As String
    mocName = boardStyleData.getMocNameByGroupName(groupName)
    If mocName = "" Then Exit Function
    
    '找BoardStyle页签上定制的场景信息，key: customGroupName+"\"+customColumnName,  columnLetter
    Dim boardStyleCustomInfoMap As CMap
    If Not boardStyleData.getBoardStyleCustomInfoMap(mocName, groupName, boardStyleCustomInfoMap) Then
        isCustomMatchRow = True
        Exit Function
    End If
    
    Dim boardStyleCellVal As String
    Dim boardStyleCellLetter As String
    Dim baseStationCellVal As String
    
    Dim boardStyleCustomInfoVar As Variant
    For Each boardStyleCustomInfoVar In boardStyleCustomInfoMap.KeyCollection
        '每个循环代表一个场景列判断，必需所有场景列都能与基站传输页的场景列匹配
        
        '取BoardStyle页签上场景列的值
        boardStyleCellLetter = boardStyleCustomInfoMap.GetAt(CStr(boardStyleCustomInfoVar))
        boardStyleCellVal = Trim(boardStyleSheet.range(boardStyleCellLetter & rowNum).value)
        If boardStyleCellVal = "" Then GoTo NextLoop
        
        '取基站传输页上场景列的值
        If Not neCustomInfoMap.hasKey(CStr(boardStyleCustomInfoVar)) Then GoTo NextLoop
            
        baseStationCellVal = neCustomInfoMap.GetAt(CStr(boardStyleCustomInfoVar))
        If Not isCustomMatch(boardStyleCellVal, baseStationCellVal) Then Exit Function
NextLoop:
    Next
    
    isCustomMatchRow = True
End Function


'检查是否场景匹配，原则：boardStyle中任一个值在基站场景信息中

Public Function isCustomMatch(ByVal boardStyleCustomVal As String, ByVal baseStationCustomVal As String) As Boolean
    isCustomMatch = False
    
    Dim baseStationValues As Collection
    Call SplitWithTrim(baseStationCustomVal, ",", baseStationValues)

    Dim boardStyleValArray() As String
    boardStyleValArray = Split(boardStyleCustomVal, ",")
    
    Dim index As Long
    For index = LBound(boardStyleValArray) To UBound(boardStyleValArray)
        If isInCollection(baseStationValues, Trim(boardStyleValArray(index))) Then
            isCustomMatch = True
            Exit Function
        End If
    Next
        
End Function
    
    
Public Function isReferenceValue(ByRef cellValue As String) As Boolean
    If UBound(Split(cellValue, "\")) = 2 Then
        isReferenceValue = True
    Else
        isReferenceValue = False
    End If
End Function
