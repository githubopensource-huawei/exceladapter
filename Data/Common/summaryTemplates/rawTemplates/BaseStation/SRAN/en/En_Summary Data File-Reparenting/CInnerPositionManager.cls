VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInnerPositionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_mappingDef_shtNameColNo As Integer
Private m_mappingDef_grpNameColNo As Integer
Private m_mappingDef_colNameColNo As Integer
Private m_mappingDef_mocNameColNo As Integer
Private m_mappingDef_attrNameColNo As Integer
Private m_mappingDef_isRefColNo As Integer
Private m_mappingDef_readOnlyColNo As Integer
Private m_mappingDef_neTypeColNo As Integer
Private m_mappingDef_neVersionColNo As Integer

Private m_sheetDef_shtNameColNo As Integer
Private m_sheetDef_shtTypeColNo As Integer
Private m_sheetDef_startRowColNo As Integer
Private m_sheetDef_endRowColNo As Integer
Private m_sheetDef_SelectedColNo As Integer

Private m_controlDef_mocNameColNo As Integer
Private m_controlDef_attrNameColNo As Integer
Private m_controlDef_shtNameColNo As Integer
Private m_controlDef_grpNameColNo As Integer
Private m_controlDef_colNameColNo As Integer


Private shtDef As Worksheet
Private mappingDef As Worksheet
Private ctrlDef As Worksheet

Private Const titleRowNo As Integer = 1



Private Sub Class_Initialize()
    Set shtDef = ThisWorkbook.Worksheets("SHEET DEF")
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Set ctrlDef = ThisWorkbook.Worksheets("CONTROL DEF")
    
    Call initMappingDef(mappingDef)
    Call initSheetDef(shtDef)
    Call initControlDef(ctrlDef)
End Sub


Property Get mappingDef_shtNameColNo()
    mappingDef_shtNameColNo = m_mappingDef_shtNameColNo
End Property

Property Get mappingDef_grpNameColNo()
    mappingDef_grpNameColNo = m_mappingDef_grpNameColNo
End Property

Property Get mappingDef_colNameColNo()
    mappingDef_colNameColNo = m_mappingDef_colNameColNo
End Property

Property Get mappingDef_mocNameColNo()
    mappingDef_mocNameColNo = m_mappingDef_mocNameColNo
End Property

Property Get mappingDef_attrNameColNo()
    mappingDef_attrNameColNo = m_mappingDef_attrNameColNo
End Property

Property Get mappingDef_isRefColNo()
    mappingDef_isRefColNo = m_mappingDef_isRefColNo
End Property

Property Get mappingDef_readOnlyColNo()
    mappingDef_readOnlyColNo = m_mappingDef_readOnlyColNo
End Property

Property Get mappingDef_neTypeColNo()
    mappingDef_neTypeColNo = m_mappingDef_neTypeColNo
End Property

Property Get mappingDef_neVersionColNo()
    mappingDef_neVersionColNo = m_mappingDef_neVersionColNo
End Property


Property Get sheetDef_shtNameColNo()
    sheetDef_shtNameColNo = m_sheetDef_shtNameColNo
End Property

Property Get sheetDef_shtTypeColNo()
    sheetDef_shtTypeColNo = m_sheetDef_shtTypeColNo
End Property

Property Get sheetDef_startRowColNo()
    sheetDef_startRowColNo = m_sheetDef_startRowColNo
End Property

Property Get sheetDef_endRowColNo()
    sheetDef_endRowColNo = m_sheetDef_endRowColNo
End Property

Property Get sheetDef_selectedColNo()
    sheetDef_selectedColNo = m_sheetDef_SelectedColNo
End Property


Property Get controlDef_mocNameColNo()
    controlDef_mocNameColNo = m_controlDef_mocNameColNo
End Property


Property Get controlDef_attrNameColNo()
    controlDef_attrNameColNo = m_controlDef_attrNameColNo
End Property

Property Get controlDef_shtNameColNo()
    controlDef_shtNameColNo = m_controlDef_shtNameColNo
End Property

Property Get controlDef_grpNameColNo()
    controlDef_grpNameColNo = m_controlDef_grpNameColNo
End Property

Property Get controlDef_colNameColNo()
    controlDef_colNameColNo = m_controlDef_colNameColNo
End Property


Public Sub updateSheetDefPositions()
    Call initSheetDef(shtDef)
End Sub

Public Sub updateMappingDefPositions()
    Call initMappingDef(mappingDef)
End Sub

Public Sub updateControlDefPositions()
    Call initControlDef(ctrlDef)
End Sub




'=================================================
'SHEET DEF
'=================================================
Private Sub initSheetDef(shtDef As Worksheet)
    m_sheetDef_shtNameColNo = -1
    m_sheetDef_shtTypeColNo = -1
    m_sheetDef_startRowColNo = -1
    m_sheetDef_endRowColNo = -1
    m_sheetDef_SelectedColNo = -1

    With shtDef.rows(titleRowNo)
        Dim shtNameRange As range
        Set shtNameRange = .Find("Sheet Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not shtNameRange Is Nothing Then m_sheetDef_shtNameColNo = shtNameRange.column
        
        Dim shtTypeRange As range
        Set shtTypeRange = .Find("Sheet Type", LookIn:=xlValues, lookat:=xlWhole)
        If Not shtTypeRange Is Nothing Then m_sheetDef_shtTypeColNo = shtTypeRange.column
        
        Dim startRowRange As range
        Set startRowRange = .Find("StartRow", LookIn:=xlValues, lookat:=xlWhole)
        If Not startRowRange Is Nothing Then m_sheetDef_startRowColNo = startRowRange.column
        
        Dim endRowRange As range
        Set endRowRange = .Find("EndRow", LookIn:=xlValues, lookat:=xlWhole)
        If Not endRowRange Is Nothing Then m_sheetDef_endRowColNo = endRowRange.column
        
        Dim selectedRange As range
        Set selectedRange = .Find("Selected", LookIn:=xlValues, lookat:=xlWhole)
        If Not selectedRange Is Nothing Then m_sheetDef_SelectedColNo = selectedRange.column
    End With
End Sub


'=================================================
'MAPPING DEF
'=================================================
Private Sub initMappingDef(mappingDef As Worksheet)
    m_mappingDef_shtNameColNo = -1
    m_mappingDef_grpNameColNo = -1
    m_mappingDef_colNameColNo = -1
    m_mappingDef_mocNameColNo = -1
    m_mappingDef_attrNameColNo = -1
    m_mappingDef_isRefColNo = -1
    m_mappingDef_readOnlyColNo = -1
    m_mappingDef_neTypeColNo = -1
    m_mappingDef_neVersionColNo = -1

    With mappingDef.rows(titleRowNo)
        Dim shtNameRange As range
        Set shtNameRange = .Find("Sheet Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not shtNameRange Is Nothing Then m_mappingDef_shtNameColNo = shtNameRange.column
        
        Dim grpNameRange As range
        Set grpNameRange = .Find("Group Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not grpNameRange Is Nothing Then m_mappingDef_grpNameColNo = grpNameRange.column
        
        Dim colNameRange As range
        Set colNameRange = .Find("Column Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not colNameRange Is Nothing Then m_mappingDef_colNameColNo = colNameRange.column
        
        Dim mocNameRange As range
        Set mocNameRange = .Find("MOC Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not mocNameRange Is Nothing Then m_mappingDef_mocNameColNo = mocNameRange.column
        
        Dim attrNameRange As range
        Set attrNameRange = .Find("Attribute Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not attrNameRange Is Nothing Then m_mappingDef_attrNameColNo = attrNameRange.column
        
        Dim isRefRange As range
        Set isRefRange = .Find("Is Reference", LookIn:=xlValues, lookat:=xlWhole)
        If Not isRefRange Is Nothing Then m_mappingDef_isRefColNo = isRefRange.column
        
        Dim readOnlyRange As range
        Set readOnlyRange = .Find("ReadOnly", LookIn:=xlValues, lookat:=xlWhole)
        If Not readOnlyRange Is Nothing Then m_mappingDef_readOnlyColNo = readOnlyRange.column
        
        Dim neTypeRange As range
        Set neTypeRange = .Find("Ne Type", LookIn:=xlValues, lookat:=xlWhole)
        If Not neTypeRange Is Nothing Then m_mappingDef_neTypeColNo = neTypeRange.column
        
        Dim neVerRange As range
        Set neVerRange = .Find("Ne Version", LookIn:=xlValues, lookat:=xlWhole)
        If Not neVerRange Is Nothing Then m_mappingDef_neVersionColNo = neVerRange.column
    End With
End Sub

'=================================================
'CONTROL DEF
'=================================================
Private Sub initControlDef(ctrlDef As Worksheet)
    m_controlDef_mocNameColNo = -1
    m_controlDef_attrNameColNo = -1
    m_controlDef_shtNameColNo = -1
    m_controlDef_grpNameColNo = -1
    m_controlDef_colNameColNo = -1
    
    With ctrlDef.rows(titleRowNo)
        Dim shtNameRange As range
        Set shtNameRange = .Find("Sheet Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not shtNameRange Is Nothing Then m_controlDef_shtNameColNo = shtNameRange.column
        
        Dim grpNameRange As range
        Set grpNameRange = .Find("Group Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not grpNameRange Is Nothing Then m_controlDef_grpNameColNo = grpNameRange.column
        
        Dim colNameRange As range
        Set colNameRange = .Find("Column Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not colNameRange Is Nothing Then m_controlDef_colNameColNo = colNameRange.column
        
        Dim mocNameRange As range
        Set mocNameRange = .Find("MOC Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not mocNameRange Is Nothing Then m_controlDef_mocNameColNo = mocNameRange.column
        
        Dim attrNameRange As range
        Set attrNameRange = .Find("Attribute Name", LookIn:=xlValues, lookat:=xlWhole)
        If Not attrNameRange Is Nothing Then m_controlDef_attrNameColNo = attrNameRange.column
    End With
End Sub


