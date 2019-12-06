Attribute VB_Name = "BoardStyleUtil"

'---------------------------------------  �ж��Ƿ�GroupName��  --------------------------------------

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


'--------------  �����кš���Ԫ�����ݣ��ҵ�Ԫ�������к�  ------------------------

Public Function findColNumByRowAndValue(ByRef ws As Worksheet, ByVal row As Long, ByRef cellVal As Variant, Optional ByVal colNumBegin As Long = 1) As Long
    Dim colMax As Long
    Dim curRowRange As Range
    
    findColNumByRowAndValue = -1
    
    colMax = getMaxCol(ws, row)
    If colNumBegin > colMax Then Exit Function
    
    Set curRowRange = ws.Range(ws.Cells(row, colNumBegin), ws.Cells(row, colMax))
        
    Dim colRst
    On Error Resume Next
    colRst = Application.WorksheetFunction.Match(cellVal, curRowRange, 0) '�����ҵ������õĲ��������ڵ��������
        
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

'--------------  ����RowNum�ҵ����ڵ�GroupNameRow  ------------------------

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

'--------------  ����GroupName��ColName������ColName�������ڵ���  ------------------------

'ֻ��ָ����Row���ڵ�Group����������

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


'δָ���к�ʱ��������sheetҳ������

'TODO:Ϊ����������BoardStyleҳǩ���������ȳ����ڵ�ǰ��Ԫ�����ڵ�Group��Χ�ڲ���

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


'-----------------------------   �����߰�ǨBoardStyleҳǩ�У������к��ҵ�SrcName������ --------------------------

'ָ��GroupNameʱ

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
''δָ��GroupNameʱ
'
'Public Function getSrcNeNameCol(ByRef ws As Worksheet, ByVal row As Long) As String
'    getSrcNeNameCol = ""
'
'    Dim sourceNeNameCol As String, sourceNeNameLetter As String
'    sourceNeNameCol = getResByKey("SOURCENENAME")
'
'    '�ҵ�GroupName���ڵ���
'    Dim groupRowNum As Long
'    groupRowNum = getGroupRowNum(ws, row)
'    If groupRowNum = -1 Then Exit Function
'
'    getSrcNeNameCol = findColLetterByRowAndValue(ws, groupRowNum + 1, sourceNeNameCol)
'
'End Function


'�ж�BoardStyleҳǩGroup�е�rowNum���Ƿ���siteName��FuntionName��NeName���滮�ĳ���

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
'    '�ҵ�������neName
'    Dim neName As String
'    neName = baseStationData.getNeNamebyFunctionName(siteName)
'    If neName = "" Then Exit Function
'
'    '�Ҵ�NeName���Ƶĳ�����Ϣ, key: CustomizeGroupName\CustomizeColumnName��customColValue
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
'    '��BoardStyleҳǩ�϶��Ƶĳ�����Ϣ��key: customGroupName+"\"+customColumnName,  columnLetter
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
'        'ÿ��ѭ������һ���������жϣ��������г����ж������վ����ҳ�ĳ�����ƥ��
'
'        'ȡBoardStyleҳǩ�ϳ����е�ֵ
'        boardStyleCellLetter = boardStyleCustomInfoMap.GetAt(CStr(boardStyleCustomInfoVar))
'        boardStyleCellVal = Trim(boardStyleSheet.Range(boardStyleCellLetter & rowNum).value)
'        If boardStyleCellVal = "" Then GoTo NextLoop
'
'        'ȡ��վ����ҳ�ϳ����е�ֵ
'        If Not neCustomInfoMap.haskey(CStr(boardStyleCustomInfoVar)) Then GoTo NextLoop
'
'        baseStationCellVal = neCustomInfoMap.GetAt(CStr(boardStyleCustomInfoVar))
'        If Not isCustomMatch(boardStyleCellVal, baseStationCellVal) Then Exit Function
'NextLoop:
'    Next
'
'    isCustomMatchRow = True
'End Function


'����Ƿ񳡾�ƥ�䣬ԭ��boardStyle����һ��ֵ�ڻ�վ������Ϣ��

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
