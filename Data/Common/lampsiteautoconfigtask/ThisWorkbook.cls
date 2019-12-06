VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim sheet As Worksheet

If containsASheet(ThisWorkbook, "System Default Policy") Then
    Set sheet = ThisWorkbook.Sheets("System Default Policy")
    If sheet.Range("B2").Value = "" Or sheet.Range("C2").Value = "" Or sheet.Range("D2").Value = "" Then
        MsgBox "The data is incomplete. Modify the data and save it."
        Cancel = True
    End If
End If

If containsBSheet(ThisWorkbook, "系统默认策略") Then
    Set sheet = ThisWorkbook.Sheets("系统默认策略")
    If sheet.Range("B2").Value = "" Or sheet.Range("C2").Value = "" Or sheet.Range("D2").Value = "" Then
        MsgBox "数据未填写完整，请修改后保存"
        Cancel = True
    End If
End If
End Sub

Public Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String, Optional ByRef ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Set ws = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsASheet = False
End Function

Public Function containsBSheet(ByRef wb As Workbook, ByRef sheetName As String, Optional ByRef ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    containsBSheet = True
    Set ws = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsBSheet = False
End Function



