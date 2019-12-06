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

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error GoTo ErrorHandler
    If ThisWorkbook.Name <> ActiveWorkbook.Name Or Sh.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    Dim sheet As Worksheet
    Set sheet = Sh
    
    If sheet.Name = "DeleteInterNCellRelation" Then
            Call Do_Worksheet_Activate(Me)
    End If
    Exit Sub
ErrorHandler:
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    Dim sheet As Worksheet
    Set sheet = Sh
    
    If sheet.Name = "DeleteInterNCellRelation" Then
            Call Do_Worksheet_Change(sheet, Target, True, True, False, True, False)
    End If
End Sub

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
