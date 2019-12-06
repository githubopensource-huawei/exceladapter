VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As Range)
    Dim sht As Worksheet
    Set sht = sh
    
    If sht.Name = "ID Modification of SRAN" Then
        Call sranShtChange(sht, target)
    ElseIf sht.Name = "ID Modification of Cell" Then
        Call cellShtChange(sht, target)
    End If
End Sub


