VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal target As range)
    On Error GoTo ErrorHandler
    Dim sht As Worksheet
    Set sht = Sh
    If ThisWorkbook.Name <> ActiveWorkbook.Name Or sht.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    

    If executeDropdown(sht, target) Then Exit Sub
    Exit Sub
ErrorHandler:
End Sub

'分支参数控制响应函数
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As range)
    On Error Resume Next
    If ThisWorkbook.Name <> ActiveWorkbook.Name Or Sh.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    If executeDropdown(Sh, target) Then Exit Sub
End Sub

Public Function executeDropdown(sht As Worksheet, target As range) As Boolean
    On Error GoTo ErrorHandler
    executeDropdown = False
    
    If target.Row < 3 Or target.Column <> 1 Then Exit Function
    
    Dim key As String
    If Trim(target.value) = 3 Or Trim(target.value) = 2 Then
       MsgBox getResByKey("Optiontips"), vbOKOnly + vbExclamation + vbApplicationModal + vbDefaultButton1, getResByKey("Warning")
    End If
    Exit Function
ErrorHandler:
    executeDropdown = False
    Debug.Print "some exception in executeDropdown, " & Err.Description
End Function
