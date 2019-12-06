VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBoardStyleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private boardStyleMappingDefMap_ As CMapValueObject
Private mocNameGroupNameMap_ As CMap 'key: mocName-value: groupName
Private groupNameMocNameMap_ As CMap
Private boardStyleCustomMappingDefMap_ As CMapValueObject
Private boardStyleCustomMocInfoMap_ As CMapValueObject

Private Const NeedFillFlag_ As String = "Yes"
Private Const NeedFillTitle_ As String = "ShowInBoardStyle"
Private Const delimiter_ As String = "-"
Private Const customDelimiter_ As String = "\"  '���ڼ��custom�ֶε�GroupName��ColumnName���Զ��幤��ʱ���ǷǷ��ֶ�

'Private needFillColumnLetter_ As String

Public Function getBoardStyleMappingDefMap() As CMapValueObject
    Set getBoardStyleMappingDefMap = boardStyleMappingDefMap_
End Function
Public Function getBoardStyleCustomMappingDefMap() As CMapValueObject
    Set getBoardStyleCustomMappingDefMap = boardStyleCustomMappingDefMap_
End Function
Public Function getBoardStyleCustomMocInfoMap() As CMap
    Set getBoardStyleCustomMocInfoMap = boardStyleCustomMocInfoMap_
End Function

Public Sub init()
    Call initBoardStyleMappingDefMap
    Call initBoardStyleRelationDef
    Call initBoardStyleColumnLetter
    Call initCustomBoardStyleColumnLetter
End Sub

Public Function getGroupNameByMocName(ByRef mocName As String) As String
    getGroupNameByMocName = ""
    If mocNameGroupNameMap_.hasKey(mocName) Then
        getGroupNameByMocName = mocNameGroupNameMap_.GetAt(mocName)
    End If
End Function
Public Function getMocNameByGroupName(ByRef groupName As String) As String
    getMocNameByGroupName = ""
    If groupNameMocNameMap_.hasKey(groupName) Then
        getMocNameByGroupName = groupNameMocNameMap_.GetAt(groupName)
    End If
End Function

Private Sub initCustomBoardStyleColumnLetter()
    Dim boardStyleSheet As Worksheet
    Set boardStyleSheet = getBoardStyleSheet
    
    Dim customattrNameVar As Variant, customGroupColVar As Variant
    Dim customattrNameStr As String
    
    Dim customGroupColName As String
    Dim customGroupColNameArry() As String
    
    Dim mocName As String, groupName As String
    Dim customMappingDefMap As CMap
    Dim customGroupName As String, customColName As String
    Dim groupNameRowNumber As Long
    Dim customColLetter As String

    For Each customattrNameVar In boardStyleCustomMappingDefMap_.KeyCollection
        '??Custom AttrName?MocName
        customattrNameStr = CStr(customattrNameVar)
        
        mocName = SubString(customattrNameStr, InStr(customattrNameStr, ".") + 1, Len(customattrNameStr))
        
        If mocNameGroupNameMap_.hasKey(mocName) Then
            groupName = mocNameGroupNameMap_.GetAt(mocName)
            groupNameRowNumber = getGroupNameStartRowNumber(boardStyleSheet, groupName)
            
            Set customMappingDefMap = boardStyleCustomMappingDefMap_.GetAt(customattrNameVar)
            'ͬһ��MOC�����ж�������ֶΣ�����Ҫѭ��
            For Each customGroupColVar In customMappingDefMap.KeyCollection
                customGroupColNameArry = Split(CStr(customGroupColVar), customDelimiter_)
                customGroupName = customGroupColNameArry(0)
                customColName = customGroupColNameArry(1)
                               
                customColLetter = findColLetterByGrpAndColName(boardStyleSheet, customGroupName, customColName, groupNameRowNumber)
                
                Call addBoardStyleCustomMocInfo(mocName + "," + groupName, CStr(customGroupColVar), customColLetter)
            Next
        End If
    Next
End Sub

'key: mocName+","+mocGroupName, customGroupColStr: customGroupName+"\"+customColumnName
Private Sub addBoardStyleCustomMocInfo(ByRef key As String, ByRef customGroupColStr As String, ByRef customColLetter As String)
    Dim customInfoMap As CMap 'key: GroupName\ColumnName, value: columnLetter
     
    If boardStyleCustomMocInfoMap_.hasKey(key) Then
        Set customInfoMap = boardStyleCustomMocInfoMap_.GetAt(key)
        Call customInfoMap.SetAt(customGroupColStr, customColLetter)
    Else
        Set customInfoMap = New CMap
        Call customInfoMap.SetAt(customGroupColStr, customColLetter)
        Call boardStyleCustomMocInfoMap_.SetAt(key, customInfoMap)
    End If
End Sub

Private Sub initBoardStyleColumnLetter()
    Dim boardStyleSheet As Worksheet
    Set boardStyleSheet = getBoardStyleSheet
    Dim groupName As Variant
    Dim groupNameRowNumber As Long
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    
    For Each groupName In boardStyleMappingDefMap_.KeyCollection
        groupNameRowNumber = getGroupNameStartRowNumber(boardStyleSheet, CStr(groupName))
        Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
        Call boardStyleMappingDefData.updateColumnNamePositionLetterMap(boardStyleSheet, groupNameRowNumber + 1)
    Next groupName
End Sub

Private Sub updateCustomColumnNamePositionLetterMap(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef customInfo As String, ByRef customColName As String)
    Dim columnNumber As Long
    Dim columnName As String, columnLetter As String
    For columnNumber = 1 To ws.range("XFD" & rowNumber).End(xlToLeft).column
        columnName = ws.Cells(rowNumber, columnNumber).value
        If customColName = columnName Then
            columnLetter = getColStr(columnNumber)
            customInfo = customInfo + "," + columnLetter
            Exit Sub
        End If
    Next columnNumber
End Sub

Private Function getBoardStyleSheet() As Worksheet
    Dim boardStyleSheetName As String
    Dim sheet As Worksheet
    Dim sheetName As String
    boardStyleSheetName = getResByKey("Board Style")
    If containsASheet(ThisWorkbook, boardStyleSheetName) Then
        Set getBoardStyleSheet = ThisWorkbook.Worksheets(boardStyleSheetName)
    Else
        For Each sheet In ThisWorkbook.Worksheets
            sheetName = sheet.name
            If InStr(sheetName, boardStyleSheetName) <> 0 Then
                Set getBoardStyleSheet = ThisWorkbook.Worksheets(sheetName)
                Exit Function
            End If
        Next sheet
    End If
End Function

Private Sub initBoardStyleMappingDefMap()
    Dim mappingDefSheet As Worksheet
    Set mappingDefSheet = ThisWorkbook.Worksheets("MAPPING DEF")
    'needFillColumnLetter_ = getColStr(findCertainValColumnNumber(mappingDefSheet, 1, NeedFillTitle_))
    
    Dim sheetName As String, groupName As String, columnName As String, mocName As String, attributeName As String
    Dim neType As String, needFillStr As String
    Dim rowNumber As Long
    
    For rowNumber = 2 To mappingDefSheet.range("a1048576").End(xlUp).row
        Call clearMappingDefEachRowData(sheetName, groupName, columnName, mocName, attributeName)
        Call getMappingDefEachRowData(sheetName, groupName, columnName, mocName, attributeName, mappingDefSheet, rowNumber)
        
        If sheetName = getResByKey("Board Style") And InStr(groupName, getResByKey("Operation")) = 0 And InStr(mocName, "Customization") = 0 Then
            Call addEachRowDataToMappingDefMap(groupName, columnName, mocName, attributeName)
            Call makeMocNameGroupNameMap(groupName, mocName)
        End If
        
        If sheetName = getResByKey("Board Style") And InStr(mocName, "Customization") <> 0 Then
            Call addBoardStyleCustomMapDef(attributeName, groupName, columnName)
        End If
    Next rowNumber
End Sub

Private Sub addBoardStyleCustomMapDef(ByRef attributeName As String, ByRef groupName As String, ByRef columnName As String)
    Dim customInfoMap As CMap 'key: GroupName\ColumnName, value: ""
     
    If boardStyleCustomMappingDefMap_.hasKey(attributeName) Then
        Set customInfoMap = boardStyleCustomMappingDefMap_.GetAt(attributeName)
        Call customInfoMap.SetAt(groupName & customDelimiter_ & columnName, "")
    Else
        Set customInfoMap = New CMap
        Call customInfoMap.SetAt(groupName & customDelimiter_ & columnName, "")
        Call boardStyleCustomMappingDefMap_.SetAt(attributeName, customInfoMap)
    End If
End Sub

Private Sub makeMocNameGroupNameMap(ByRef groupName As String, ByRef mocName As String)
    If mocName <> "" And Not mocNameGroupNameMap_.hasKey(mocName) Then
        Call mocNameGroupNameMap_.SetAt(mocName, groupName)
    End If
    If groupName <> "" And Not groupNameMocNameMap_.hasKey(groupName) Then
        Call groupNameMocNameMap_.SetAt(groupName, mocName)
    End If
End Sub

Private Sub addEachRowDataToMappingDefMap(ByRef groupName As String, ByRef columnName As String, ByRef mocName As String, _
    ByRef attributeName As String)
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    If boardStyleMappingDefMap_.hasKey(groupName) Then
        Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
        Call boardStyleMappingDefData.addData(groupName, columnName, mocName, attributeName)
    Else
        Set boardStyleMappingDefData = New CBoardStyleMappingDefData
        Call boardStyleMappingDefData.addData(groupName, columnName, mocName, attributeName)
        Call boardStyleMappingDefMap_.SetAt(groupName, boardStyleMappingDefData)
    End If
End Sub

Private Sub getMappingDefEachRowData(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, ByRef mocName As String, _
    ByRef attributeName As String, ByRef mappingDefSheet As Worksheet, ByRef rowNumber As Long)
    sheetName = mappingDefSheet.range("A" & rowNumber).value
    groupName = mappingDefSheet.range("B" & rowNumber).value
    columnName = mappingDefSheet.range("C" & rowNumber).value
    mocName = mappingDefSheet.range("D" & rowNumber).value
    attributeName = mappingDefSheet.range("E" & rowNumber).value
    'needFillStr = mappingDefSheet.Range(needFillColumnLetter_ & rowNumber).value
End Sub

Private Sub clearMappingDefEachRowData(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, ByRef mocName As String, _
    ByRef attributeName As String)
    sheetName = ""
    groupName = ""
    columnName = ""
    mocName = ""
    attributeName = ""
End Sub

Private Sub initBoardStyleRelationDef()
    Dim relationDefSheet As Worksheet
    Set relationDefSheet = ThisWorkbook.Worksheets("RELATION DEF")
    Dim sheetName As String, groupName As String, columnName As String, mocName As String, attributeName As String
    Dim isVisible As String, isMocAttribute As String, referenceMoc As String, referenceKey As String, needFillIn As String
    Dim exportMathFunction As String
    Dim rowNumber As Long
    
    For rowNumber = 2 To relationDefSheet.range("a1048576").End(xlUp).row
        Call clearRelationDefEachRowData(sheetName, groupName, columnName, mocName, attributeName, isVisible, isMocAttribute, referenceMoc, referenceKey, needFillIn, exportMathFunction)
        Call getRelationDefEachRowData(sheetName, groupName, columnName, mocName, attributeName, isVisible, isMocAttribute, referenceMoc, referenceKey, needFillIn, relationDefSheet, rowNumber, exportMathFunction)
        
        If sheetName = getResByKey("Board Style") Then
            Call updateRelationDefRowData(groupName, columnName, mocName, attributeName, isVisible, isMocAttribute, referenceMoc, referenceKey, needFillIn, exportMathFunction)
        End If
        
    Next rowNumber
End Sub

Private Sub updateRelationDefRowData(ByRef predefinedGroupName As String, ByRef predefinedColumnName As String, ByRef mocName As String, _
    ByRef attributeName As String, ByRef isVisible As String, ByRef isMocAttribute As String, ByRef referenceMoc As String, _
    ByRef referenceKey As String, ByRef needFillIn As String, ByRef exportMathFunction As String)
    
    'If Not boardStyleMappingDefMap_.haskey(groupName) Then Exit Sub
    If Not mocNameGroupNameMap_.hasKey(mocName) Then Exit Sub
    
    Dim groupName As String, columnName As String
    groupName = mocNameGroupNameMap_.GetAt(mocName)
    
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
    
    columnName = boardStyleMappingDefData.getColumnNameByAttributeName(attributeName)
    
    If columnName = "" Then
        columnName = predefinedColumnName
    End If
    
    If isVisible = "True" And isMocAttribute = "False" And referenceKey = "" Then
        Call boardStyleMappingDefData.updateAutoFillInColumnName(columnName, attributeName)
    ElseIf isVisible = "True" And isMocAttribute = "False" And referenceKey <> "" And UCase(exportMathFunction) = "JOIN" Then
        Call boardStyleMappingDefData.updateListboxReferenceKeyMap(columnName, makeReferenceMocKey(referenceMoc, referenceKey))
    ElseIf isVisible = "True" And isMocAttribute = "False" And referenceKey <> "" And UCase(exportMathFunction) = "LISTJOIN" Then
        '�����LISTJOIN��ʽ������Զ�����ö�����ӵ�һ����Ԫ����
        Call boardStyleMappingDefData.updateMultiListReferenceKeyMap(columnName, makeReferenceMocKey(referenceMoc, referenceKey))
    End If
    
    If needFillIn = "Y" Then
        Call boardStyleMappingDefData.updateNeedFillInColumnNameCol(columnName)
    End If
End Sub

Private Function makeReferenceMocKey(ByRef referenceMoc As String, ByRef referenceKey As String)
    makeReferenceMocKey = referenceMoc & delimiter_ & referenceKey
End Function

Private Sub getRelationDefEachRowData(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, _
    ByRef mocName As String, ByRef attributeName As String, ByRef isVisible As String, ByRef isMocAttribute As String, ByRef referenceMoc As String, _
    ByRef referenceKey As String, ByRef needFillIn As String, ByRef relationDefSheet As Worksheet, ByRef rowNumber As Long, _
    ByRef exportMathFunction As String)
    sheetName = relationDefSheet.range("A" & rowNumber).value
    groupName = relationDefSheet.range("B" & rowNumber).value
    columnName = relationDefSheet.range("C" & rowNumber).value
    mocName = relationDefSheet.range("F" & rowNumber).value
    attributeName = relationDefSheet.range("G" & rowNumber).value
    isVisible = relationDefSheet.range("D" & rowNumber).value
    isMocAttribute = relationDefSheet.range("E" & rowNumber).value
    referenceMoc = relationDefSheet.range("J" & rowNumber).value
    referenceKey = relationDefSheet.range("K" & rowNumber).value
    needFillIn = relationDefSheet.range("M" & rowNumber).value
    exportMathFunction = relationDefSheet.range("I" & rowNumber).value
End Sub

Private Sub clearRelationDefEachRowData(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, ByRef mocName As String, _
    ByRef attributeName As String, ByRef isVisible As String, ByRef isMocAttribute As String, ByRef referenceMoc As String, ByRef referenceKey As String, ByRef needFillIn As String, _
    ByRef exportMathFunction As String)
    sheetName = ""
    groupName = ""
    columnName = ""
    mocName = ""
    attributeName = ""
    isVisible = ""
    isMocAttribute = ""
    referenceMoc = ""
    referenceKey = ""
    needFillIn = ""
    exportMathFunction = ""
End Sub

'����MocName��GroupName�ҵ�������ϢMap
Public Function getBoardStyleCustomInfoMap(ByVal mocName As String, ByVal groupName As String, _
    ByRef customInfoMap As CMap) As Boolean
    
    getBoardStyleCustomInfoMap = False
    
    Dim keyStr As String
    keyStr = mocName + "," + groupName
    
    If boardStyleCustomMocInfoMap_.hasKey(keyStr) Then
        Set customInfoMap = boardStyleCustomMocInfoMap_.GetAt(keyStr)
                If Not customInfoMap.isEmpty Then
                    getBoardStyleCustomInfoMap = True
                End If
    End If

End Function

Private Sub Class_Initialize()
    Set boardStyleMappingDefMap_ = New CMapValueObject
    Set mocNameGroupNameMap_ = New CMap
    Set groupNameMocNameMap_ = New CMap
    Set boardStyleCustomMappingDefMap_ = New CMapValueObject
    Set boardStyleCustomMocInfoMap_ = New CMapValueObject
End Sub