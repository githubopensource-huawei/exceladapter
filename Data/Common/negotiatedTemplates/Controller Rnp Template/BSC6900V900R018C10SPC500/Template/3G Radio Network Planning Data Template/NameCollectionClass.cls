VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NameCollectionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
'类模块名称：NameCollectionClass
'类模块作用：扩展Collection
'***************************************************************
Option Explicit

Private strName As String
Private Values As New Collection

Property Get Name()
    Name = strName
End Property

Property Let Name(Value)
    strName = Value
End Property

Property Get Count()
    Count = Values.Count
End Property

Public Function Item(index As Variant) As Variant
    If IsObject(Values.Item(index)) Then
        Set Item = Values.Item(index)
    Else
        Item = Values.Item(index)
    End If
End Function

Public Sub Add(Item As Variant, Optional Key As Variant)
    If IsMissing(Key) Then
        Values.Add Item:=Item
    Else
        Values.Add Item:=Item, Key:=Key
    End If
End Sub

Public Sub Remove(index As Variant)
    Values.Remove (index)
End Sub
