VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MocClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
'类模块名称：MocClass
'类模块作用：方便操作RNP模板涉及的MOC及其属性
'***************************************************************
Option Explicit

Private strName As String '存储MOC名称
Private colAttrs As New Collection '存储MOC属性收集器

Private strSheetNameVDF As String '仅为CvtTemplate模块所用

'读取MOC名称
Property Get Name()
    Name = strName
End Property

'设置MOC名称
Property Let Name(Value)
    strName = Value
End Property

Property Get SheetNameVDF()
    SheetNameVDF = strSheetNameVDF
End Property

Property Let SheetNameVDF(Value)
    strSheetNameVDF = Value
End Property

'*******************************************************************************
'读取MOC有多少属性，包含虚拟属性，另外，为方便使用，请保证虚拟属性总是加在最后
'*******************************************************************************
Property Get CountWithVirtualAttr()
    CountWithVirtualAttr = colAttrs.Count
End Property

'*******************************************************************************
'增加一个属性
'*******************************************************************************
Public Sub Add(Attr As AttrClass)
    If Attr.ColIndex <= 0 Then
        Attr.ColIndex = colAttrs.Count + 1 '这个添加顺序产生的列索引作为读取Excel的列索引
    End If
    If Attr.DataType = atUnknown Then
        If Attr.DataTypeName = "INT" Then
            Attr.DataType = atInteger
        ElseIf Attr.DataTypeName = "LIST" Then
            Attr.DataType = atEnum
        ElseIf Attr.DataTypeName = "STRING" Then
            Attr.DataType = atString
        Else
            Attr.DataType = atUnknown
        End If
    End If
    colAttrs.Add Item:=Attr, Key:=Attr.Name
End Sub

'读取一个属性
Public Function Attr(IndexOrName As Variant) As AttrClass
    Set Attr = colAttrs.Item(IndexOrName)
End Function

Public Sub Remove(AttrName As String)
    If Exists(AttrName) Then
        colAttrs.Remove (AttrName)
    End If
End Sub

Public Sub Clear()
    Do While colAttrs.Count > 0
        colAttrs.Remove (1)
    Loop
    Set colAttrs = Nothing
End Sub

'读取MOC有多少属性，不含虚拟属性
Property Get Count()
    Dim iCnt As Integer
    iCnt = 0
    Dim Attr As AttrClass
    Dim i As Integer
    For i = 1 To colAttrs.Count
        Set Attr = colAttrs.Item(i)
        If Not Attr.IsVirtualAttr Then
            iCnt = iCnt + 1
        End If
    Next i
    
    Count = iCnt
End Property

Public Function Exists(AttrName As String) As Boolean
    Exists = IsInCollection(AttrName, colAttrs)
End Function
