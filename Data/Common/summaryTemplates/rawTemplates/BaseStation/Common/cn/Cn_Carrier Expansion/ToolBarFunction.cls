VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolBarFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Private m_siteAddSupport As Boolean
    Private m_cellAddSupport As Boolean
    Private m_radioAddSupport As Boolean
    
Property Get siteAddSupport() As Boolean
    siteAddSupport = m_siteAddSupport
End Property

Property Get cellAddSupport() As Boolean
    cellAddSupport = m_cellAddSupport
End Property

Property Get radioAddSupport() As Boolean
    radioAddSupport = m_radioAddSupport
End Property

Private Sub Class_Initialize()
On Error GoTo ErrExit
            Dim sh As Worksheet
            m_siteAddSupport = True
            m_cellAddSupport = True
            m_radioAddSupport = True
            Set sh = Worksheets("ToolBarConfig")
            
            If sh.Cells(1, 2).value = "N" Then
                  m_siteAddSupport = False
            End If
            If sh.Cells(2, 2).value = "N" Then
                  m_cellAddSupport = False
            End If
            If sh.Cells(3, 2).value = "N" Then
                  m_radioAddSupport = False
            End If
            Exit Sub
ErrExit:
            m_siteAddSupport = True
            m_cellAddSupport = True
            m_radioAddSupport = True
End Sub


