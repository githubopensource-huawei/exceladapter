VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CControlDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
        Private m_mocName As String
        Private m_attributeName   As String
        Private m_dataType As String
        Private m_bound   As String
        Private m_lstValue   As String
        Private m_sheetName As String
        Private m_groupName As String
        Private m_columnName As String
        Private m_neType As String
        Private m_controlInfo As String
        
Property Get mocName() As String
    mocName = m_mocName
End Property
Property Get attributeName() As String
    attributeName = m_attributeName
End Property
Property Get dataType() As String
    dataType = m_dataType
End Property
Property Get bound() As String
    bound = m_bound
End Property
Property Get lstValue() As String
    lstValue = m_lstValue
End Property

Property Get sheetName() As String
    sheetName = m_sheetName
End Property

Property Get groupName() As String
    groupName = m_groupName
End Property

Property Get columnName() As String
    columnName = m_columnName
End Property

Property Get neType() As String
    neType = m_neType
End Property

Property Get controlInfo() As String
    controlInfo = m_controlInfo
End Property
'let
Property Let mocName(name As String)
    m_mocName = name
End Property
Property Let attributeName(name As String)
    m_attributeName = name
End Property
Property Let dataType(name As String)
    m_dataType = name
End Property
Property Let bound(boundValue As String)
    m_bound = boundValue
End Property
Property Let lstValue(value As String)
    m_lstValue = value
End Property

Property Let sheetName(name As String)
    m_sheetName = name
End Property

Property Let groupName(name As String)
    m_groupName = name
End Property

Property Let columnName(name As String)
    m_columnName = name
End Property

Property Let controlInfo(str As String)
    m_controlInfo = str
End Property

Property Let neType(str As String)
    m_neType = str
End Property

Public Function getKey() As String
    getKey = m_sheetName + "," + m_groupName + "," + m_columnName
End Function



