VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub deleteStyles()
    Dim s As Style
    For Each s In ThisWorkbook.Styles
        If Not s.BuiltIn Then
            Debug.Print s.Name
            s.Delete
        End If
    Next
End Sub

