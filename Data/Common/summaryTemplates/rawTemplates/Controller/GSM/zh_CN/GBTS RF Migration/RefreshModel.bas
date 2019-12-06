Attribute VB_Name = "RefreshModel"
Function getColStr(ByVal NumVal As Long) As String
    Dim str As String
    Dim strs() As String
    
    If NumVal > 256 Or NumVal < 1 Then
        getColStr = ""
    Else
        str = Cells(NumVal).address
        strs = Split(str, "$", -1)
        getColStr = strs(1)
    End If
End Function









