VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Activate()
  On Error Resume Next
  InsertUserToolBar
  InsertCostomToolBar
  ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
End Sub

Private Sub Workbook_Deactivate()
  On Error Resume Next
  DeleteCostomToolBar
  DeleteUserToolBar
  ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
  'Call SetSysOption
End Sub

