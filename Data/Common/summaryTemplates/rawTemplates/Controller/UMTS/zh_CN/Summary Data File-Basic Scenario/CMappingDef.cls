VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMappingDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
        Private m_sheetName As String
        Private m_groupName As String
        Private m_columnName As String
        Private m_mocName As String
        Private m_attributeName  As String
        Private m_neType As String
        Private m_neVersion As String
        
' get property
Property Get neType() As String
        neType = m_neType
End Property
Property Get neVersion() As String
        neVersion = m_neVersion
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

Property Get mocName() As String
        mocName = m_mocName
End Property
Property Get attributeName() As String
        attributeName = m_attributeName
End Property


' set property
Property Let sheetName(shName As String)
        m_sheetName = shName
End Property

Property Let groupName(gpName As String)
        m_groupName = gpName
End Property

Property Let columnName(colmName As String)
        m_columnName = colmName
End Property

Property Let mocName(mcName As String)
         m_mocName = mcName
End Property
Property Let attributeName(arrtName As String)
        m_attributeName = arrtName
End Property
Property Let neType(str As String)
        m_neType = str
End Property
Property Let neVersion(str As String)
        m_neVersion = str
End Property

' function

Public Function getKey() As String
    getKey = m_sheetName + "," + m_groupName + "," + m_columnName
End Function


