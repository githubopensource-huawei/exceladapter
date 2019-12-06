VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CControlRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mocName_ As String
Private attributeName_ As String '��������Ǳ��ز���
Private controlAttrCol_ As New Collection '�������������ز���
Private neType_ As String
Private sheetName_ As String
Private groupName_ As String
Private columnName_ As String
Private valueType_ As String '���������ͣ�Enum��String��

Public Sub init(ByRef mocName As String, ByRef attributeName As String, ByRef controlInfo As String, ByRef neType As String, ByRef sheetName As String, ByRef GroupName As String, ByRef columnName As String, ByRef valueType As String)
    If controlInfo <> "" Then initControlAttrCol (controlInfo)
    Call initParameters(mocName, attributeName, neType, sheetName, GroupName, columnName, valueType)
End Sub

Property Get GroupName() As String
    GroupName = groupName_
End Property

Property Get columnName() As String
    columnName = columnName_
End Property

Property Get valueType() As String
    valueType = valueType_
End Property

Private Sub initParameters(ByRef mocName As String, ByRef attributeName As String, ByRef neType As String, ByRef sheetName As String, ByRef GroupName As String, ByRef columnName As String, ByRef valueType As String)
    mocName_ = mocName
    attributeName_ = attributeName
    neType_ = neType
    sheetName_ = sheetName
    groupName_ = GroupName
    columnName_ = columnName
    valueType_ = valueType
End Sub

Private Sub initControlAttrCol(ByRef controlInfo As String)
    Dim xmlDoc As Variant
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.LoadXML (controlInfo)
    
    Dim root As Variant '���ڵ�
    Set root = xmlDoc.DocumentElement
    
    Dim controlAttributeNode As Variant
    Dim controlAttributeName As String
    For Each controlAttributeNode In root.GetElementsByTagName("Equals")
        controlAttributeName = controlAttributeNode.getAttribute("attribute")
        If Not IsNull(controlAttributeName) Then
            If Not Contains(controlAttrCol_, controlAttributeName) Then
                controlAttrCol_.Add Item:=controlAttributeName, key:=controlAttributeName '�������ز�������������
            End If
        End If
    Next controlAttributeNode
End Sub

Property Get controlAttrCol() As Collection
    Set controlAttrCol = controlAttrCol_
End Property

