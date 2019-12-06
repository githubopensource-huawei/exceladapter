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



Public Sub cmdGenNegotiatedFile_Click()
    Call T_DefineNegotiatedFile.GenNegotiatedFile
    'added by z00102652 at 2010-04-08, begin
    'Call DoCmdSetCTClick

    'Call DoCmdSetDFClick

    Call DoCmdExportCodeClick
    'added by z00102652 at 2010-04-08, end
End Sub

Private Sub DoCmdExportCodeClick()
    Call ExportCode
    MsgBox "Finished to export code."
End Sub

Private Sub DoCmdSetCTClick()
    Dim sht As Worksheet, sht2 As Worksheet
    Set sht2 = Application.ActiveSheet
    Set sht = Sheets(SHT_CONVERT_TEMPLATE)
    sht.Activate
    sht.Tab.ColorIndex = 6
    Call SetTemplate_CT
    MsgBox "Finished to set sheet '" + SHT_CONVERT_TEMPLATE + "'."

    sht2.Activate
End Sub

Private Sub DoCmdSetDFClick()
    Dim sht As Worksheet, sht2 As Worksheet
    Set sht2 = Application.ActiveSheet
    Set sht = Sheets(SHT_DOUBLE_FREQ_CELL_SETTING)
    sht.Visible = True
    sht.Activate

    Call SetTemplate_DF
    MsgBox "Finished to set sheet '" + SHT_DOUBLE_FREQ_CELL_SETTING + "'."

    sht.Visible = False
    sht2.Activate
End Sub

