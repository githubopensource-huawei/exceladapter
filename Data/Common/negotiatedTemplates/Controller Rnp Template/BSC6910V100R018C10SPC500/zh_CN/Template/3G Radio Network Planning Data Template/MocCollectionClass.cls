VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MocCollectionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
'类模块名称：MocCollectionClass
'类模块作用：MocClass对象收集器
'***************************************************************
Option Explicit

Private colMocs As New Collection

Property Get Count()
    Count = colMocs.Count
End Property

Public Sub Add(MocName As String)
    Dim colMoc As New MocClass
    colMoc.Name = MocName
    colMocs.Add Item:=colMoc, Key:=colMoc.Name
End Sub

Public Function Moc(IndexOrName As Variant) As MocClass
    Set Moc = colMocs.Item(IndexOrName)
End Function

Public Sub Clear()
    Dim Moc As MocClass
    Do While colMocs.Count > 0
        Set Moc = colMocs.Item(1)
        Moc.Clear
        colMocs.Remove (1)
    Loop
End Sub

'增加虚拟属性
Public Sub AddVirtualAttr(MocName As String, RefMocName As String, RefAttrName As String)
    Dim Moc As MocClass, RefMoc As MocClass, Attr As AttrClass, RefAttr As AttrClass
    Set Moc = colMocs.Item(MocName)
    Set RefMoc = colMocs.Item(RefMocName)
    Set RefAttr = RefMoc.Attr(RefAttrName)
    Set Attr = RefAttr.Clone
    Attr.RefMocName = RefMocName
    Attr.RefAttrName = RefAttrName
    Moc.Add Attr
End Sub
