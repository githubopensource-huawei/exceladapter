VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CControlAttrValueManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private controlAttrValueMap_ As New CMapValueObject '存放所有主控参数值的Map，用主控参数短名称为Key

Public Sub addNewControlAttrValue(ByRef controlAttrValue As CControlAttrValue)
    If Not controlAttrValueMap_.hasKey(controlAttrValue.shortName) Then
        Call controlAttrValueMap_.SetAt(controlAttrValue.shortName, controlAttrValue)
    End If
End Sub

Public Function getControlAttrValue(ByRef shortName As String) As CControlAttrValue
    Set getControlAttrValue = controlAttrValueMap_.GetAt(shortName)
End Function

Public Function hasControlAttr(ByRef shortName As String) As Boolean
    hasControlAttr = controlAttrValueMap_.hasKey(shortName)
End Function

