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
    On Error Resume Next
    InitResource
    
    InsertUserToolBar 
    
End Sub


Private Sub Workbook_Deactivate()
    On Error Resume Next
    
    Call DeleteUserToolBar 
    
    DestroyMenuStatus
    Application.StatusBar = False
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sheet As Object, ByVal Target As range)
    If sheet.Name <> "HOME" Then
        Exit Sub
    End If
    
    Dim HomeSheet As Worksheet
    Dim TableName As String
    
    If Target.row > 100 Or Target.row < 2 Or Target.Column <> 2 Then
        Exit Sub
    End If
    
    TableName = Trim(Sheets("HOME").Cells(Target.row, Target.Column))
    If TableName = "" Then
        Exit Sub
    End If
   
    Dim sTmpAddr, sTargetSheetName As String
    sTmpAddr = Sheets("HOME").range("B" + Trim(Str(Target.row)) + ":" + "B" + Trim(Str(Target.row + 1))).Hyperlinks(1).SubAddress
    
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

