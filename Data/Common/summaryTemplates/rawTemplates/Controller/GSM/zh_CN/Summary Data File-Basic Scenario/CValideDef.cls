VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CValideDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
        Private m_sheetName As String
        Private m_groupName As String
        Private m_columnName As String
        Private m_row As Long
        Private m_column As Long
        
        Private m_mocName As String
        Private m_attrName As String
        Private m_innerFlag As Boolean
        
    Property Let sheetName(str As String)
        m_innerFlag = True
        m_sheetName = str
    End Property
    
    Property Let groupName(str As String)
        m_groupName = str
    End Property
    
    Property Let columnName(str As String)
        m_columnName = str
    End Property
    
    Property Let mocName(str As String)
        m_mocName = str
        m_innerFlag = False
    End Property
    
    Property Let attrName(str As String)
        m_attrName = str
    End Property
    
    
    Property Let row(str As Long)
        m_row = str
    End Property
    
    Property Let column(str As Long)
        m_column = str
    End Property
    
    Property Get row() As Long
        row = m_row
    End Property
    
    Property Get column() As Long
        column = m_column
    End Property
    
    Function getKey() As String
        getKey = m_sheetName + "," + m_groupName + "," + m_columnName
    End Function
    
    Function getMocKey() As String
        getMocKey = m_mocName + "," + m_attrName
    End Function
    
    Function getValidedef() As String
        If m_innerFlag = True Then
            getValidedef = "=INDIRECT(""'InnerValideDef'!D" & CStr(m_row) & ":" & getColStr(m_column) & CStr(m_row) & """)"
        Else
            getValidedef = "=INDIRECT(""'VALID DEF'!C" & CStr(m_row) & ":" & getColStr(m_column) & CStr(m_row) & """)"
        End If
    End Function
    

