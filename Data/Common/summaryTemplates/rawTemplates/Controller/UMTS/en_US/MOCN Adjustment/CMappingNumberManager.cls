VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMappingNumberManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mocAttrNeTypeNumberMap_ As New CMap '在MappingDef中Moc和Attr个数的Map，如果某个主控Attr个数超过1，则不进行分支控制
Private Const delimeter_ As String = "_"

Public Sub addMocAttrNetype(ByRef sheetName As String, ByRef mocName As String, ByRef attrName As String, ByRef neType As String)
    Dim key As String
    key = getKey(sheetName, mocName, attrName, neType)
    Dim recordNumber As Long
    recordNumber = mocAttrNeTypeNumberMap_.GetAt(key, 0)
    Call mocAttrNeTypeNumberMap_.SetAt(key, recordNumber + 1) '将Moc_Attr_NeType记录个数加1
End Sub

Public Function hasOneMocAttrNetypeRecord(ByRef sheetName As String, ByRef mocName As String, ByRef attrName As String, ByRef neType As String) As Boolean
    If getMocAttrNetypeNumber(sheetName, mocName, attrName, neType) = 1 Then
        '只有在有一条记录的时候才返回True
        hasOneMocAttrNetypeRecord = True
    Else
        hasOneMocAttrNetypeRecord = False
    End If
End Function

Private Function getMocAttrNetypeNumber(ByRef sheetName As String, ByRef mocName As String, ByRef attrName As String, ByRef neType As String) As Long
    Dim key As String
    key = getKey(sheetName, mocName, attrName, neType)
    getMocAttrNetypeNumber = mocAttrNeTypeNumberMap_.GetAt(key, 0)
End Function

Public Function getKey(ByRef sheetName As String, ByRef mocName As String, ByRef attrName As String, ByRef neType As String) As String
    getKey = sheetName & delimeter_ & mocName & delimeter_ & attrName & delimeter_ & neType
End Function

