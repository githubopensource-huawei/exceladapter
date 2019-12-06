VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' 激活时添加自定义工具栏
Private Sub Workbook_Activate()
  On Error Resume Next
  init ThisWorkbook
  BuildToolBar
  'ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
  'init ThisWorkbook
End Sub

' 去激活时删除工具栏
Private Sub Workbook_Deactivate()
  On Error Resume Next
  DelToolBar
  'ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If she.Name = "BTS Transport Layer" Then Call transportLayer_SheetChange(Target)
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If she.Name = "BTS Transport Layer" Then Call transportLayer_SelectionChange(Target)
End Sub
