VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public g_SelectionRow As Integer
Public g_SelectionCol As Integer
Const SolidColorIdx = 16
Const SolidPattern = xlGray16
Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal target As Range)
    On Error GoTo ErrorHandler
    Dim sht As Worksheet
    Set sht = sh
    If ThisWorkbook.Name <> ActiveWorkbook.Name Or sht.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    
    If sht.Name <> "Template" Then
        Exit Sub
    End If
    
    If (target.Row < 2) Then
        Exit Sub
    End If
    
    If Ensure_Gray_Cell(target) = False Then
        Exit Sub
    End If
    Exit Sub
ErrorHandler:
End Sub
Public Function Ensure_Gray_Cell(curRange As Range) As Boolean
    If curRange.Value <> "" And curRange.Interior.ColorIndex = SolidColorIdx And curRange.Interior.Pattern = SolidPattern Then
        MsgBox "No input is required!", vbOKOnly + vbExclamation + vbApplicationModal + vbDefaultButton1, "Warning"
        curRange.Select
        Ensure_Gray_Cell = True
    Else
        Ensure_Gray_Cell = False
    End If
End Function
