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
'��ģ�����ƣ�MocClass
'��ģ�����ã��������RNPģ���漰��MOC��������
'***************************************************************
Option Explicit

Private strName As String '�洢MOC����
Private colAttrs As New Collection '�洢MOC�����ռ���

Private strSheetNameVDF As String '��ΪCvtTemplateģ������

'��ȡMOC����
Property Get Name()
    Name = strName
End Property

'����MOC����
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
'��ȡMOC�ж������ԣ������������ԣ����⣬Ϊ����ʹ�ã��뱣֤�����������Ǽ������
'*******************************************************************************
Property Get CountWithVirtualAttr()
    CountWithVirtualAttr = colAttrs.Count
End Property

'*******************************************************************************
'����һ������
'*******************************************************************************
Public Sub Add(Attr As AttrClass)
    If Attr.ColIndex <= 0 Then
        Attr.ColIndex = colAttrs.Count + 1 '�������˳���������������Ϊ��ȡExcel��������
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

'��ȡһ������
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

'��ȡMOC�ж������ԣ�������������
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