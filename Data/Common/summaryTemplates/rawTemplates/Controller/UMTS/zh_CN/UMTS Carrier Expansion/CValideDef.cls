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
        
    Property Let sheetName(str As String)
        m_sheetName = str
    End Property
    
    Property Let groupName(str As String)
        m_groupName = str
    End Property
    
    Property Let columnName(str As String)
        m_columnName = str
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
        row = m_column
    End Property
    
    Function getKey() As String
        getKey = m_sheetName + "," + m_groupName + "," + m_columnName
    End Function
    
    Function getValidedef() As String
        getValidedef = "=INDIRECT(""'InnerValideDef'!D" & CStr(m_row) & ":" & getColStr(m_column) & CStr(m_row) & """)"
    End Function
    


