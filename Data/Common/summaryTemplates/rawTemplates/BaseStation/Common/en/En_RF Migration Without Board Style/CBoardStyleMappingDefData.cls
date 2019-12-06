VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBoardStyleMappingDefData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mocName_ As String 'moc短名称
Private groupName_ As String 'groupName
'Private attributeNameCol_ As Collection '属性短名称
'Private displayAttributeNameCol_ As Collection '需要显示的属性短名称
Private columnNameCol_ As Collection '列名全集, 以短名称作为Key
Private columnNamePositionLetterMap_ As CMap '每一个列名所在的ColumnLetter Map
Private needFillColumnNameCol_ As Collection '需要显示的列名全集
'Private hasBoardNoFlag_ As Boolean
Private autoFillInColumnName_ As String
Private autoFillInColumnNameSourceAttributes_ As String
Private listboxReferenceKeyMap_ As CMap '以Reference Moc和Reference Key作为value
Private multiListReferenceKeyMap_ As CMap '如果是多List引用单板编号字段，则以ColumnName为Key，Reference Moc和Reference Key作为value，当前只有"*基带处理板编号"
Private autoFillInSourceColumnName_ As Collection
Private copyFillInSourceColumnName_ As Collection

Public Function getColumnNameByAttributeName(ByRef attributeName As String) As String
    If Contains(columnNameCol_, attributeName) Then
        getColumnNameByAttributeName = columnNameCol_(attributeName)
    Else
        getColumnNameByAttributeName = ""
    End If
End Function

Public Function getColumnNamePositionLetter(ByRef columnName As String) As String
    getColumnNamePositionLetter = columnNamePositionLetterMap_.GetAt(columnName)
End Function

Public Sub updateColumnNamePositionLetterMap(ByRef ws As Worksheet, ByRef rowNumber As Long)
    Dim columnNumber As Long
    Dim columnName As String, columnLetter As String
    For columnNumber = 1 To ws.range("XFD" & rowNumber).End(xlToLeft).column
        columnName = ws.Cells(rowNumber, columnNumber).value
        columnLetter = getColStr(columnNumber)
        Call columnNamePositionLetterMap_.SetAt(columnName, columnLetter)
    Next columnNumber
End Sub

Public Property Get groupName() As String
    groupName = groupName_
End Property

Public Property Get columnNamePositionLetterMap() As CMap
    Set columnNamePositionLetterMap = columnNamePositionLetterMap_
End Property

Public Property Get needFillColumnNameCol() As Collection
    Set needFillColumnNameCol = needFillColumnNameCol_
End Property

Public Property Get totalColumnNumber() As Long
    totalColumnNumber = columnNamePositionLetterMap_.count
End Property

Public Property Get autoFillInColumnName() As String
    autoFillInColumnName = autoFillInColumnName_
End Property

Public Property Get autoFillInColumnNameSourceAttributes() As String
    autoFillInColumnNameSourceAttributes = autoFillInColumnNameSourceAttributes_
End Property
Public Property Get autoFillInSourceColumnName() As Collection
    Set autoFillInSourceColumnName = autoFillInSourceColumnName_
End Property
Public Property Get copyFillInSourceColumnName() As Collection
    Set copyFillInSourceColumnName = copyFillInSourceColumnName_
End Property
Public Sub updateAutoFillInColumnName(ByRef columnName As String, ByRef attributeName As String)
    autoFillInColumnName_ = columnName
    autoFillInColumnNameSourceAttributes_ = getAttributeLongNameString(attributeName)
End Sub
Public Sub updateCopyFillInColumnName(ByRef columnName As String, ByRef copyAttributeName As String)
    autoFillInSourceColumnName_.Add Item:=columnName
    copyFillInSourceColumnName_.Add Item:=copyAttributeName
End Sub

Private Function getAttributeLongNameString(ByRef attruteName As String)
    Dim attributeLongNameString As String
    attributeLongNameString = ""
    Dim delimiter As String
    delimiter = ","
    Dim attributesArray As Variant, eachAttribute As Variant
    attributesArray = Split(attruteName, delimiter)
    For Each eachAttribute In attributesArray
        attributeLongNameString = attributeLongNameString & columnNameCol_(eachAttribute) & delimiter
    Next eachAttribute
    attributeLongNameString = Left(attributeLongNameString, Len(attributeLongNameString) - 1)
    getAttributeLongNameString = attributeLongNameString
End Function

Public Function getListBoxReferenceValue(ByRef columnName) As String
    Dim referenceValue As String
    If listboxReferenceKeyMap_.hasKey(columnName) Then
        referenceValue = listboxReferenceKeyMap_.GetAt(columnName)
    Else
        referenceValue = ""
    End If
    getListBoxReferenceValue = referenceValue
End Function

Public Function getMultiListReferenceValue(ByRef columnName) As String
    Dim referenceValue As String
    If multiListReferenceKeyMap_.hasKey(columnName) Then
        referenceValue = multiListReferenceKeyMap_.GetAt(columnName)
    Else
        referenceValue = ""
    End If
    getMultiListReferenceValue = referenceValue
End Function

Public Sub updateListboxReferenceKeyMap(ByRef columnName As String, ByRef key As String)
    Call listboxReferenceKeyMap_.SetAt(columnName, key)
End Sub

Public Sub updateMultiListReferenceKeyMap(ByRef columnName As String, ByRef key As String)
    Call multiListReferenceKeyMap_.SetAt(columnName, key)
End Sub

Public Function hasListBoxColumns() As Boolean
    hasListBoxColumns = True
    If listboxReferenceKeyMap_.isEmpty Then hasListBoxColumns = False
End Function

Public Function hasMultiListReferenceColumns() As Boolean
    hasMultiListReferenceColumns = True
    If multiListReferenceKeyMap_.isEmpty Then hasMultiListReferenceColumns = False
End Function

Public Function getListBoxReferencKeyMap() As CMap
    Set getListBoxReferencKeyMap = listboxReferenceKeyMap_
End Function

Public Function getMultiListReferenceKeyMap() As CMap
    Set getMultiListReferenceKeyMap = multiListReferenceKeyMap_
End Function

Public Sub updateNeedFillInColumnNameCol(ByRef columnName As String)
    needFillColumnNameCol_.Add Item:=columnName
End Sub

Public Sub addData(ByRef groupName As String, ByRef columnName As String, ByRef mocName As String, _
    ByRef attributeName As String)
    If mocName_ = "" Then mocName_ = mocName
    groupName_ = groupName
    If attributeName = "" Then
        columnNameCol_.Add Item:=columnName, key:=columnName
    Else
        columnNameCol_.Add Item:=columnName, key:=attributeName
    End If
'
'    If UCase(needFillStr) = "YES" Then
'        needFillColumnNameCol_.Add Item:=columnName
'    End If
    
End Sub

Private Sub Class_Initialize()
    Set columnNameCol_ = New Collection
    Set columnNamePositionLetterMap_ = New CMap
    Set needFillColumnNameCol_ = New Collection
    Set listboxReferenceKeyMap_ = New CMap
    Set multiListReferenceKeyMap_ = New CMap
    Set autoFillInSourceColumnName_ = New Collection
    Set copyFillInSourceColumnName_ = New Collection
End Sub
