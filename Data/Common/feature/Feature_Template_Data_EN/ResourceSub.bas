Attribute VB_Name = "ResourceSub"
Private resource  As CMap

Public Sub initResource()
        Dim index As Integer
        Dim key As String
        Dim value As String
        Set resource = New CMap
        For index = 2 To SummaryRes.Range("a65536").End(xlUp).row
            key = SummaryRes.Cells(index, 1).value
            value = SummaryRes.Cells(index, 2).value
            Call resource.SetAt(key, value)
        Next
        
End Sub

Public Function getResByKey(key As String) As String
   On Error Resume Next
    If resource Is Nothing Then
            Call initResource
    End If
    If (resource.haskey(key)) Then
            getResByKey = resource.GetAt(key)
    Else
            getResByKey = key
    End If
End Function
