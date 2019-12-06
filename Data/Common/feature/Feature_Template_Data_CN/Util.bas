Attribute VB_Name = "Util"
Option Explicit

Public Function isNum(contedType As String) As Boolean
    If contedType = "Integer" Or contedType = "UInteger" Then
        isNum = True
    Else
        isNum = False
    End If
End Function

 Function Contains(coll As Collection, key As String) As Boolean
        On Error GoTo NotFound
        Call coll(key)
        Contains = True
        Exit Function
NotFound:
        Contains = False
End Function

Private Sub deleteStyles()
    Dim s As Style
    For Each s In ThisWorkbook.Styles
        If Not s.BuiltIn Then
            Debug.Print s.name
            s.Delete
        End If
    Next
End Sub

