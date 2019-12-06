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
    
    If isMappingDefHasAttrName("SiteTemplateName") = False Then
        m_siteAddSupport = False
        sh.Cells(2, 1) = "N"
    Else
        sh.Cells(2, 1) = "Y"
    End If
    
    If isMappingDefHasAttrName("RadioTemplateName") = False Or isRadioandNodeMerged("SiteTemplateName", "RadioTemplateName") Then
       m_radioAddSupport = False
        sh.Cells(2, 2) = "N"
    Else
        sh.Cells(2, 2) = "Y"
    End If

    
    If isMappingDefHasAttrName("CellTemplateName") = False Then
       m_cellAddSupport = False
        sh.Cells(2, 3) = "N"
    Else
        sh.Cells(2, 3) = "Y"
    End If
    Exit Sub
ErrExit:
    m_siteAddSupport = True
    m_cellAddSupport = True
    m_radioAddSupport = True
End Sub


Private Function isMappingDefHasAttrName(ByRef attrName As String) As Boolean
On Error GoTo ErrExit
     Dim ret As range
     Set ret = Worksheets("MAPPING DEF").range("E:E").Find(what:=attrName, lookat:=xlWhole)
     If Not ret Is Nothing Then
             isMappingDefHasAttrName = True
     Else
             isMappingDefHasAttrName = False
     End If
Exit Function
ErrExit:
     isMappingDefHasAttrName = True
End Function


Private Function isRadioandNodeMerged(ByRef siteTemplateAttrName As String, ByRef radioTemplateAttrName As String) As Boolean
On Error GoTo ErrExit
    Dim nodeRange As range, radioRange As range
    Set nodeRange = Worksheets("MAPPING DEF").range("E:E").Find(what:=siteTemplateAttrName, lookat:=xlWhole)
    Set radioRange = Worksheets("MAPPING DEF").range("E:E").Find(what:=radioTemplateAttrName, lookat:=xlWhole)
    
    If Not nodeRange Is Nothing And Not radioRange Is Nothing Then
        Dim nodeRow As Long, radioRow As Long
        nodeRow = nodeRange.row
        radioRow = radioRange.row
        If Worksheets("MAPPING DEF").Cells(nodeRow, 3).value = Worksheets("MAPPING DEF").Cells(radioRow, 3).value Then
            isRadioandNodeMerged = True
        End If
    Else
            isRadioandNodeMerged = False
    End If
    Exit Function
ErrExit:
     isRadioandNodeMerged = False
End Function
