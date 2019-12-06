VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error GoTo ErrorHandler
    If ThisWorkbook.Name <> ActiveWorkbook.Name Or Sh.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    Dim sheet As Worksheet
    Set sheet = Sh
    
    Call Do_Worksheet_Activate(sheet)
    
    Exit Sub
ErrorHandler:
End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    Dim sheet As Worksheet
    Set sheet = Sh
    
    If sheet.Name = "AICH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "BCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLACCESSSTRICT" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLALGOSWITCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLAMRC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLAMRCWB" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLCAC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLCBS" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLCBSDRX" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLCBSSAC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLCMCF" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLDCCC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLDISTANCEREDIRECTION" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLDRD" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLDRDMIMO" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLDSACMANUALPARA" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLEDRX" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLEFACH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLFRC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLGPSFRMTIMING" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLHCS" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLHCSHO" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLHOCOMM" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLHSDPA" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLINTERFREQHOCOV" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLINTERFREQHONCOV" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLINTERRATHOCOV" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLINTERRATHONCOV" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLINTRAFREQHO" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLIPDL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLLDB" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "CELLLDM" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLLDR" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLLICENSE" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMBDRINTERFREQ" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMBDRINTERRAT" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMBMS" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMBMSFACH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMBMSPARA" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMBMSSA" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMBMSSCCPCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "CELLMBSCCRRM" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMCCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "CELLMCDRD" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMCLDR" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "CELLMEAS" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLMIMO" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLNFREQPRIOINFO" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLOLC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLPUC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLQOSHO" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLQUALITYMEAS" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLREDIRECTION" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLRLACTTIME" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLRLPWR" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "CELLSELRESEL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLSIBSWITCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CELLU2LTEHONCOV" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "CELLULB" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CHPWROFFSET" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "CTCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "ERACHACTOASCMAP" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "ERACHASC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "ERACHBASIC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "FACH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "FACHDYNTFS" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "FACHLOCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "PCCPCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "PCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "PCHDYNTFS" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "PCPICH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "PICH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "PRACHACTOASCMAP" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "PRACHASC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "PRACHBASIC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "PRACHSLOTFORMAT" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "PRACHTFC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "PSCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "RACH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "RACHDYNTFS" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "SCCPCHBASIC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "SCCPCHTFC" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "SCPICH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
        If sheet.Name = "SMLCCELL" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
    If sheet.Name = "SSCH" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
End Sub
Private Sub Workbook_Activate()
  On Error Resume Next
  InsertUserToolBar
 ' ActiveWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
End Sub

Private Sub Workbook_Deactivate()
  On Error Resume Next
  DeleteUserToolBar
 ' ActiveWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
End Sub

