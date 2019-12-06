VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CControlRelationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const keyDelimeter_ As String = "_"
Private controlledAttributeCollection_ As New Collection '被控参数的容器，简单的字符串key值和value值，当需要时还是要从controlRelationMap_中取
Private controlRelationMap_ As New CMapValueObject

Public Sub addNewAttributeRelation(ByRef mocName As String, ByRef attributeName As String, ByRef controlInfo As String, ByRef neType As String, ByRef sheetName As String, ByRef GroupName As String, ByRef columnName As String, ByRef valueType As String)
    On Error GoTo ErrorHandler
    Dim key As String
    key = getKey(mocName, attributeName, neType, sheetName)
    If Not controlRelationMap_.hasKey(key) Then
        Dim controlRelation As New CControlRelation
        Call controlRelation.init(mocName, attributeName, controlInfo, neType, sheetName, GroupName, columnName, valueType)
        Call controlRelationMap_.SetAt(key, controlRelation)
    End If
    
    If controlInfo <> "" And (Not Contains(controlledAttributeCollection_, key)) Then
        controlledAttributeCollection_.Add Item:=key, key:=key
    End If
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Function getKey(ByRef mocName As String, ByRef attributeName As String, ByRef neType As String, ByRef sheetName As String) As String
    getKey = mocName & keyDelimeter_ & attributeName & keyDelimeter_ & neType & keyDelimeter_ & sheetName
End Function

'Control Def页签中是否有此属性作为被控参数
Public Function containsControlledAttributeRelation(ByRef mocName As String, ByRef attributeName As String, ByRef neType As String, ByRef sheetName As String) As Boolean
    Dim key  As String
    key = getKey(mocName, attributeName, neType, sheetName)
    If Contains(controlledAttributeCollection_, key) Then
        containsControlledAttributeRelation = True
    Else
        containsControlledAttributeRelation = False
    End If
End Function

'Control Def页签中是否有此属性
Public Function containsAttributeRelation(ByRef mocName As String, ByRef attributeName As String, ByRef neType As String, ByRef sheetName As String) As Boolean
    Dim key  As String
    key = getKey(mocName, attributeName, neType, sheetName)
    If controlRelationMap_.hasKey(key) Then
        containsAttributeRelation = True
    Else
        containsAttributeRelation = False
    End If
End Function

Public Function getControlRelation(ByRef mocName As String, ByRef attributeName As String, ByRef neType As String, ByRef sheetName As String) As CControlRelation
    On Error GoTo ErrorHandler
    Dim key As String
    key = getKey(mocName, attributeName, neType, sheetName)
    Set getControlRelation = controlRelationMap_.GetAt(key)
    Exit Function
ErrorHandler:
    Set getControlRelation = Nothing
End Function

