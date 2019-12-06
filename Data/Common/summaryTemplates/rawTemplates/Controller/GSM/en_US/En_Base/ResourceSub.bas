Attribute VB_Name = "ResourceSub"
Private resource  As CMap

Public Sub initResource()
        Dim index As Long
        Dim key As String
        Dim value As String
        Set resource = New CMap
        For index = 2 To ThisWorkbook.Worksheets("SummaryRes").Range("a1048576").End(xlUp).row
            key = ThisWorkbook.Worksheets("SummaryRes").Cells(index, 1).value
            value = ThisWorkbook.Worksheets("SummaryRes").Cells(index, 2).value
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