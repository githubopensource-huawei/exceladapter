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
    ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
  
    InitResource
  
    InsertUserToolBar
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
  
    Call DeleteUserToolBar
    
    DestroyMenuStatus
    Application.StatusBar = False
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
  'Call SetSysOption
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sheet As Object, ByVal Target As Range)
    If sheet.Name <> "HOME" Then
        Exit Sub
    End If
    
    Dim HomeSheet As Worksheet
    Dim TableName As String
    If Target.Row > 100 Or Target.Row < 2 Or Target.Column <> 2 Then
        Exit Sub
    End If
    
    TableName = Trim(Sheets("HOME").Cells(Target.Row, Target.Column))
    If TableName = "" Then
        Exit Sub
    End If
   
    Dim sTmpAddr, sTargetSheetName As String
    sTmpAddr = Sheets("HOME").Range("B" + Trim(str(Target.Row)) + ":" + "B" + Trim(str(Target.Row + 1))).Hyperlinks(1).SubAddress
    
    If "" = sTmpAddr Then
        Exit Sub
    End If
    
    sTargetSheetName = Replace(Left(sTmpAddr, (InStr(sTmpAddr, "!") - 1)), "'", "")
   
    Set HomeSheet = ThisWorkbook.Sheets(sTargetSheetName)
    If HomeSheet.Visible = False Then
       ActiveWorkbook.Unprotect
       HomeSheet.Activate
       HomeSheet.Visible = True
       ActiveWorkbook.Protect
    End If
End Sub

