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
Private Const NeedFillFlag_ As String = "Yes"
Private Const NeedFillTitle_ As String = "ShowInBoardStyle"
Private Const delimiter_ As String = "-"
Private boardStyleCustomMappingDefMap_ As CMap
Private boardStyleCustomMocInfoMap_ As CMap
'Private needFillColumnLetter_ As String

Public Function getBoardStyleMappingDefMap() As CMapValueObject
    Set getBoardStyleMappingDefMap = boardStyleMappingDefMap_
End Function

Public Function getBoardStyleCustomMappingDefMap() As CMap
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
    Dim customattrNameVar As Variant
    Dim customattrNameArry() As String
    Dim customgroupColName As String
    Dim customgroupColNameArry() As String
    Dim mocName As String
    Dim groupName As String, customgroupName As String, customColName As String, customInfo As String
    Dim groupNameRowNumber As Long
    
    For Each customattrNameVar In boardStyleCustomMappingDefMap_.KeyCollection
        customattrNameArry = Split(CStr(customattrNameVar), ".")
        customgroupColName = boardStyleCustomMappingDefMap_.GetAt(customattrNameVar)
        customgroupColNameArry = Split(CStr(customgroupColName), ",")
        mocName = customattrNameArry(1)
        customgroupName = customgroupColNameArry(0)
        customColName = customgroupColNameArry(1)
        customInfo = customgroupColName
        If mocNameGroupNameMap_.hasKey(mocName) Then
            groupName = mocNameGroupNameMap_.GetAt(mocName)
            groupNameRowNumber = getGroupNameStartRowNumber(boardStyleSheet, groupName)
            Call updateCustomColumnNamePositionLetterMap(boardStyleSheet, groupNameRowNumber + 1, customInfo, customColName)
            Call boardStyleCustomMocInfoMap_.SetAt(mocName + "," + groupName, customInfo)
        End If
    Next
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
    Dim columnName As String, columnletter As String
    For columnNumber = 1 To ws.Range("XFD" & rowNumber).End(xlToLeft).column
        columnName = ws.Cells(rowNumber, columnNumber).value
        If customColName = columnName Then
            columnletter = getColStr(columnNumber)
            customInfo = customInfo + "," + columnletter
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
    
    For rowNumber = 2 To mappingDefSheet.Range("a1048576").End(xlUp).row
        Call clearMappingDefEachRowData(sheetName, groupName, columnName, mocName, attributeName)
        Call getMappingDefEachRowData(sheetName, groupName, columnName, mocName, attributeName, mappingDefSheet, rowNumber)
        
        If sheetName = getResByKey("Board Style") And InStr(groupName, getResByKey("Operation")) = 0 And InStr(mocName, "Customization") = 0 Then
            Call addEachRowDataToMappingDefMap(groupName, columnName, mocName, attributeName)
            Call makeMocNameGroupNameMap(groupName, mocName)
        End If
        If sheetName = getResByKey("Board Style") And InStr(mocName, "Customization") <> 0 Then
            Call boardStyleCustomMappingDefMap_.SetAt(attributeName, groupName + "," + columnName)
        End If
    Next rowNumber
End Sub

Private Function includeColumnName(ByRef columnName As String) As Boolean
    Dim mappingDefSheet As Worksheet
    Set mappingDefSheet = ThisWorkbook.Worksheets("MAPPING DEF")
    
    Dim columnNameInMappingDef As String
    Dim rowNumber As Long
    includeColumnName = False
    
    For rowNumber = 2 To mappingDefSheet.Range("a1048576").End(xlUp).row
        columnNameInMappingDef = ""
        columnNameInMappingDef = mappingDefSheet.Range("C" & rowNumber).value
        
        If columnNameInMappingDef = columnName Then
            includeColumnName = True
            Exit For
        End If
    Next rowNumber
End Function

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
    If groupName <> "Operation" And groupName <> getResByKey("Operation_Group") Then
        If boardStyleMappingDefMap_.hasKey(groupName) Then
            Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
            Call boardStyleMappingDefData.addData(groupName, columnName, mocName, attributeName)
        Else
            Set boardStyleMappingDefData = New CBoardStyleMappingDefData
            Call boardStyleMappingDefData.addData(groupName, columnName, mocName, attributeName)
            Call boardStyleMappingDefMap_.SetAt(groupName, boardStyleMappingDefData)
        End If
    End If
End Sub

Private Sub getMappingDefEachRowData(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, ByRef mocName As String, _
    ByRef attributeName As String, ByRef mappingDefSheet As Worksheet, ByRef rowNumber As Long)
    sheetName = mappingDefSheet.Range("A" & rowNumber).value
    groupName = mappingDefSheet.Range("B" & rowNumber).value
    columnName = mappingDefSheet.Range("C" & rowNumber).value
    mocName = mappingDefSheet.Range("D" & rowNumber).value
    attributeName = mappingDefSheet.Range("E" & rowNumber).value
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
    Dim isVisible As String, isMocAttribute As String, referenceMoc As String, referenceKey As String, needFillIn As String, needClearIn As String
    
    Dim rowNumber As Long
    
    For rowNumber = 2 To relationDefSheet.Range("a1048576").End(xlUp).row
        Call clearRelationDefEachRowData(sheetName, groupName, columnName, mocName, attributeName, isVisible, isMocAttribute, referenceMoc, referenceKey, needFillIn, needClearIn)
        Call getRelationDefEachRowData(sheetName, groupName, columnName, mocName, attributeName, isVisible, isMocAttribute, referenceMoc, referenceKey, needFillIn, needClearIn, relationDefSheet, rowNumber)
        
        If sheetName = getResByKey("Board Style") Then
            Call updateRelationDefRowData(groupName, columnName, mocName, attributeName, isVisible, isMocAttribute, referenceMoc, referenceKey, needFillIn, needClearIn)
        End If
        
    Next rowNumber
End Sub

Private Sub updateRelationDefRowData(ByRef predefinedGroupName As String, ByRef predefinedColumnName As String, ByRef mocName As String, _
    ByRef attributeName As String, ByRef isVisible As String, ByRef isMocAttribute As String, ByRef referenceMoc As String, _
    ByRef referenceKey As String, ByRef needFillIn As String, ByRef needClearIn As String)
    
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
    ElseIf isVisible = "True" And isMocAttribute = "False" And referenceKey <> "" Then
        Call boardStyleMappingDefData.updateListboxReferenceKeyMap(columnName, makeReferenceMocKey(referenceMoc, referenceKey))
    End If
    
    If needFillIn = "Y" Then
        Call boardStyleMappingDefData.updateNeedFillInColumnNameCol(columnName)
    End If
    
    If needClearIn = "Y" Then
        Call boardStyleMappingDefData.updateNeedClearInColumnNameCol(columnName)
    End If
End Sub

Private Function makeReferenceMocKey(ByRef referenceMoc As String, ByRef referenceKey As String)
    makeReferenceMocKey = referenceMoc & delimiter_ & referenceKey
End Function

Private Sub getRelationDefEachRowData(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, _
    ByRef mocName As String, ByRef attributeName As String, ByRef isVisible As String, ByRef isMocAttribute As String, ByRef referenceMoc As String, _
    ByRef referenceKey As String, ByRef needFillIn As String, ByRef needClearIn As String, ByRef relationDefSheet As Worksheet, ByRef rowNumber As Long)
    sheetName = relationDefSheet.Range("A" & rowNumber).value
    groupName = relationDefSheet.Range("B" & rowNumber).value
    columnName = relationDefSheet.Range("C" & rowNumber).value
    mocName = relationDefSheet.Range("F" & rowNumber).value
    attributeName = relationDefSheet.Range("G" & rowNumber).value
    isVisible = relationDefSheet.Range("D" & rowNumber).value
    isMocAttribute = relationDefSheet.Range("E" & rowNumber).value
    referenceMoc = relationDefSheet.Range("J" & rowNumber).value
    referenceKey = relationDefSheet.Range("K" & rowNumber).value
    needFillIn = relationDefSheet.Range("M" & rowNumber).value
    needClearIn = relationDefSheet.Range("N" & rowNumber).value
End Sub

Private Sub clearRelationDefEachRowData(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, ByRef mocName As String, _
    ByRef attributeName As String, ByRef isVisible As String, ByRef isMocAttribute As String, ByRef referenceMoc As String, ByRef referenceKey As String, _
    ByRef needFillIn As String, ByRef needClearIn As String)
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
    needClearIn = ""
End Sub

Private Sub Class_Initialize()
    Set boardStyleMappingDefMap_ = New CMapValueObject
    Set mocNameGroupNameMap_ = New CMap
    Set groupNameMocNameMap_ = New CMap
    Set boardStyleCustomMappingDefMap_ = New CMap
    Set boardStyleCustomMocInfoMap_ = New CMap
End Sub


