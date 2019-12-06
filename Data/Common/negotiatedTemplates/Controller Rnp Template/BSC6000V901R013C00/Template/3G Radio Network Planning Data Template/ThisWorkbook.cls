VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const TABDEF_BAR_NAME = "TabDefBar"
Private Const IMPORT_TMEPLATE = "ImpTempBar"
Private Const IMPORT_RNP_FILE = "ImportRnpFileBar"
Private Const COVERT_DATA = "CovertDataBar"
Private Const CONFIG_INTRANCELL = "ConfigIntraNCellBar"
Private Const CONFIG_GSMCELL = "ConfigGsmCellBar"
Private Const CONFIG_INTRANCELL_SAMESECCTOR = "ConfigInterNCellSameSector"
Private Const CONFIG_INTRANCELL_DIFFSECCTOR = "ConfigInterNCellDiffSector"
Private Const COPYDATA_TO_CELL_LIST = "CopyDataToCELLBar"
Private Const COPYDATA_FROM_CELL_LIST = "CopyDataFromCELL"
Private Const CELL_IDMAP_FROM_ULA = "CellIdMapFromUla"

Private Sub Workbook_Activate()
    'MsgBox "Workbook_Activate"
    On Error Resume Next
    
    InitUIStringResource
        
    Call InsertUserToolBar
    With ActiveWindow
        If .DisplayZeros Then Exit Sub
        .DisplayGridlines = False
        .DisplayZeros = True
    End With
    Call HideExtendFucSheet
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    DeleteUserToolBar
End Sub

Private Sub Workbook_Open()
    'MsgBox "Workbook_Open"
    Call GetSheetDefineData
    'Call GetAllSheetName
    ActiveWorkbook.Sheets("Cover").Select
    iHideSheetFlg = ThisWorkbook.Sheets("TableDef").Range("G11").Text
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    Dim Mocs As Variant
'    Mocs = Array(MOC_NODEB, MOC_CELL, MOC_NRNCCELL, MOC_INTRAFREQNCELL, MOC_INTERFREQNCELL, MOC_GSMCELL, MOC_GSMNCELL, MOC_SMLCCELL, MOC_PHY_NB_RADIO)
'    Dim i As Integer
'    For i = LBound(Mocs) To UBound(Mocs)
'        Call DeleteEmptyRowInSheetEnding(Mocs(i))
'    Next i
End Sub





Private Sub initAujustAntnPortToolBar(ByRef ws As Worksheet)

    Call deleteCellBar
    If ws.Name = "TableDef" Then
        Call createGenRnpTempBar
    End If
    
    If ws.Name = "ConvertTemplate" Then
        Call createConvertTemplateBar
    End If
    
    If ws.Name = "DoubleFrequencyCell" Then
        Call createDoubleFrequencyConfigBar
    End If
    
    If ws.Name = "PhyNBRadio" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "NRNCCELL" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "NODEB" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "LTENCELL" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "LTECELL" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "INTRAFREQNCELL" Then
            Call Do_Worksheet_Activate(ws)
    End If
    If ws.Name = "INTERFREQNCELL" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "GSMCELL" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "GSMNCELL" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "CELL" Then
            Call Do_Worksheet_Activate(ws)
    End If
    If ws.Name = "BSCInfo" Then
            Call Do_Worksheet_Activate(ws)
    End If
    If ws.Name = "WholeNetworkCell" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "DeleteInterNCellRelation" Then
            Call Do_Worksheet_Activate(ws)
    End If
    
    If ws.Name = "DoubleFrequencyCell" Then
            Call DoubleFrequency_Worksheet_Activate(ws)
    End If
    
End Sub

Sub createConvertTemplateBar()
    Dim baseStationChooseBar As CommandBar
    Dim baseStationStyle As CommandBarButton
    
    Call deleteCellBar
    
    Set baseStationChooseBar = Application.CommandBars.Add(IMPORT_TMEPLATE, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Import Template"
            .TooltipText = "ImportTemplate"
            .OnAction = "cmdSelectFileNameForVDF_Click"
            .FaceId = 50
            .Enabled = True
        End With
      End With
    
    Set baseStationChooseBar = Application.CommandBars.Add(IMPORT_RNP_FILE, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Import RNP File"
            .TooltipText = "ImportRnpFile"
            .OnAction = "cmdSelectFileNameForHW_Click"
            .FaceId = 50
            .Enabled = True
        End With
    End With
End Sub


Sub createDoubleFrequencyConfigBar()
    Dim baseStationChooseBar As CommandBar
    Dim baseStationStyle As CommandBarButton
    
    Call deleteCellBar
    Set baseStationChooseBar = Application.CommandBars.Add(CONFIG_INTRANCELL, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Config IntraNCell"
            .TooltipText = "ConfigIntraNCell"
            .OnAction = "cmdConfigIntraNCell_Click"
            .FaceId = 50
            .Enabled = True
        End With
   End With
      
   Set baseStationChooseBar = Application.CommandBars.Add(CONFIG_GSMCELL, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Config GSMCell"
            .TooltipText = "ConfigGSMCell"
            .OnAction = "cmdConfigGSMNCell_Click"
            .FaceId = 50
            .Enabled = True
        End With
   End With
   
   Set baseStationChooseBar = Application.CommandBars.Add(CONFIG_INTRANCELL_SAMESECCTOR, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Config InterNCell SameSector"
            .TooltipText = "ConfigInterNCellSameSector"
            .OnAction = "cmdConfigInterNCellSameSector_Click"
            .FaceId = 50
            .Enabled = True
        End With
   End With
   
   Set baseStationChooseBar = Application.CommandBars.Add(CONFIG_INTRANCELL_DIFFSECCTOR, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Config InterNCell DiffSector"
            .TooltipText = "ConfigInterNCellDiffSector"
            .OnAction = "cmdConfigInterNCellDiffSector_Click"
            .FaceId = 50
            .Enabled = True
        End With
   End With
   
   Set baseStationChooseBar = Application.CommandBars.Add(COPYDATA_TO_CELL_LIST, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Copy Data To Cell List"
            .TooltipText = "CopyDataToCELL"
            .OnAction = "cmdCopyDataToCELL_Click"
            .FaceId = 50
            .Enabled = True
        End With
   End With
   
   Set baseStationChooseBar = Application.CommandBars.Add(COPYDATA_FROM_CELL_LIST, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Copy Data From Cell List"
            .TooltipText = "CopyDataFromCELL"
            .OnAction = "cmdCopyDataFromCELL_Click"
            .FaceId = 50
            .Enabled = True
        End With
   End With
   
   Set baseStationChooseBar = Application.CommandBars.Add(CELL_IDMAP_FROM_ULA, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Cell IdMap From ULA"
            .TooltipText = "CellIdMapFromUla"
            .OnAction = "cmdSetFormula_Click"
            .FaceId = 50
            .Enabled = True
        End With
   End With
End Sub


Sub createGenRnpTempBar()
    Dim baseStationChooseBar As CommandBar
    Dim baseStationStyle As CommandBarButton
    
    Call deleteCellBar
    Set baseStationChooseBar = Application.CommandBars.Add(TABDEF_BAR_NAME, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Gen RNP Template"
            .TooltipText = "GenRnpTemplate"
            .OnAction = "cmdGenNegotiatedFile_Click"
            .FaceId = 50
            .Enabled = True
        End With
      End With
End Sub

Private Function existToolBar(ByRef barName As String) As Boolean
    On Error GoTo ErrorHandler
    existToolBar = True
    Dim bar As CommandBar
    Set bar = Application.CommandBars(barName)
    Exit Function
ErrorHandler:
    existToolBar = False
End Function

Private Sub deleteCellBar()
    If existToolBar(TABDEF_BAR_NAME) Then
        Application.CommandBars(TABDEF_BAR_NAME).Delete
    End If
    If existToolBar(IMPORT_TMEPLATE) Then
        Application.CommandBars(IMPORT_TMEPLATE).Delete
    End If
        
    If existToolBar(IMPORT_RNP_FILE) Then
        Application.CommandBars(IMPORT_RNP_FILE).Delete
    End If
    
    If existToolBar(COVERT_DATA) Then
        Application.CommandBars(COVERT_DATA).Delete
    End If
    
     If existToolBar(CONFIG_INTRANCELL) Then
        Application.CommandBars(CONFIG_INTRANCELL).Delete
    End If
     If existToolBar(CONFIG_GSMCELL) Then
        Application.CommandBars(CONFIG_GSMCELL).Delete
    End If
     If existToolBar(CONFIG_INTRANCELL_SAMESECCTOR) Then
        Application.CommandBars(CONFIG_INTRANCELL_SAMESECCTOR).Delete
    End If
     If existToolBar(CONFIG_INTRANCELL_DIFFSECCTOR) Then
        Application.CommandBars(CONFIG_INTRANCELL_DIFFSECCTOR).Delete
    End If
    
     If existToolBar(COPYDATA_TO_CELL_LIST) Then
        Application.CommandBars(COPYDATA_TO_CELL_LIST).Delete
    End If
     If existToolBar(COPYDATA_FROM_CELL_LIST) Then
        Application.CommandBars(COPYDATA_FROM_CELL_LIST).Delete
    End If
     If existToolBar(CELL_IDMAP_FROM_ULA) Then
        Application.CommandBars(CELL_IDMAP_FROM_ULA).Delete
    End If
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error GoTo ErrorHandler
    If ThisWorkbook.Name <> ActiveWorkbook.Name Or Sh.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    Dim sheet As Worksheet
    Set sheet = Sh
    
    Call initAujustAntnPortToolBar(sheet)
    Exit Sub
ErrorHandler:
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    Dim sheet As Worksheet
    Set sheet = Sh
    
    If sheet.Name = "CELLALGOSWITCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    
    If sheet.Name = "PhyNBRadio" Then
            Call Do_Worksheet_Change(sheet, Target, False, False, True, True, False)
    End If
    
    If sheet.Name = "NRNCCELL" Then
             Call Do_Worksheet_Change(sheet, Target, True, True, True, False, True)
    End If
    
    If sheet.Name = "NODEB" Then
             Call Do_Worksheet_Change(sheet, Target, True, True, True, False, True)
    End If
    
    If sheet.Name = "LTENCELL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    
    If sheet.Name = "LTECELL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    
    If sheet.Name = "INTRAFREQNCELL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "INTERFREQNCELL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    
    If sheet.Name = "GSMCELL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    
    If sheet.Name = "GSMNCELL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    
    If sheet.Name = "CELL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, True, False, True)
    End If
    If sheet.Name = "BSCInfo" Then
            Call Do_Worksheet_Change(sheet, Target, False, False, False, True, False)
    End If
    If sheet.Name = "WholeNetworkCell" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, True, False, True)
    End If
    
    If sheet.Name = "DeleteInterNCellRelation" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    
    If sheet.Name = "DoubleFrequencyCell" Then
            Call DoubleFrequency_Worksheet_Change(sheet, Target)
    End If
       
End Sub
