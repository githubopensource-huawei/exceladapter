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
Private Const NeedFillFlag_ As String = "Yes"
Private Const NeedFillTitle_ As String = "ShowInBoardStyle"
Private Const delimiter_ As String = "-"
'Private needFillColumnLetter_ As String

Public Function getBoardStyleMappingDefMap() As CMapValueObject
    Set getBoardStyleMappingDefMap = boardStyleMappingDefMap_
End Function

Public Sub init()
    Call initBoardStyleMappingDefMap
    Call initBoardStyleRelationDef
    Call initBoardStyleColumnLetter
End Sub

Public Function getGroupNameByMocName(ByRef mocName As String) As String
    getGroupNameByMocName = ""
    If mocNameGroupNameMap_.hasKey(mocName) Then
        getGroupNameByMocName = mocNameGroupNameMap_.GetAt(mocName)
    End If
End Function

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

Private Function getBoardStyleSheet() As Worksheet
    Dim boardStyleSheetName As String
    boardStyleSheetName = getResByKey("Board Style")
    If containsASheet(ThisWorkbook, boardStyleSheetName) Then
        Set getBoardStyleSheet = ThisWorkbook.Worksheets(boardStyleSheetName)
    Else
        Set getBoardStyleSheet = ThisWorkbook.Worksheets(boardStyleSheetName & "_1")
    End If
End Function

Private Sub initBoardStyleMappingDefMap()
    Dim mappingDefSheet As Worksheet
    Set mappingDefSheet = ThisWorkbook.Worksheets("MAPPING DEF")
    
    Dim sheetName As String, groupName As String, columnName As String, mocName As String, attributeName As String
    Dim neType As String, needFillStr As String
    Dim rowNumber As Long
    
    For rowNumber = 2 To mappingDefSheet.range("a65536").End(xlUp).row
        Call clearMappingDefEachRowData(sheetName, groupName, columnName, mocName, attributeName)
        Call getMappingDefEachRowData(sheetName, groupName, columnName, mocName, attributeName, mappingDefSheet, rowNumber)
        
        If sheetName = getResByKey("Board Style") Then
            Call addEachRowDataToMappingDefMap(groupName, columnName, mocName, attributeName)
            Call makeMocNameGroupNameMap(groupName, mocName)
        End If
    Next rowNumber
End Sub

Private Sub makeMocNameGroupNameMap(ByRef groupName As String, ByRef mocName As String)
    If mocName <> "" And Not mocNameGroupNameMap_.hasKey(mocName) Then
        Call mocNameGroupNameMap_.SetAt(mocName, groupName)
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
    
    For rowNumber = 2 To relationDefSheet.range("a65536").End(xlUp).row
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
        '如果是LISTJOIN方式，则可以多个引用对象填加到一个单元格中
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

Private Sub Class_Initialize()
    Set boardStyleMappingDefMap_ = New CMapValueObject
    Set mocNameGroupNameMap_ = New CMap
End Sub
