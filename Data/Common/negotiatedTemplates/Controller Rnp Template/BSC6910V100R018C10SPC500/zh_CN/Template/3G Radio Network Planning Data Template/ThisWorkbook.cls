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

Public Sub btn_SwitchChs_Click()
   Call SetChineseUI
End Sub

Public Sub btn_SwitchEng_Click()
   Call SetEnglishUI
End Sub


