Attribute VB_Name = "ResourceSub"
Private resource  As CMap

Public Sub initResource()
        Dim index As Long
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


Public Function getNeType() As String
    On Error Resume Next
    Dim cover As String
    Dim key As String
    Dim reValue As String
    
    cover = getResByKey("Cover")
    key = ThisWorkbook.Worksheets(cover).Cells(2, 2).value
    reValue = getResByKey(key)
    getNeType = reValue
End Function

Public Function isContainGsmControl() As Boolean
    On Error Resume Next
    Dim cover As String
    Dim key As String
    Dim reValue As String
    Dim index As Long
    
    cover = getResByKey("Cover")
    For index = 3 To 5
        key = ThisWorkbook.Worksheets(cover).Cells(index, 2).value
        reValue = getResByKey(key)
        If reValue = "BSC" Then
            isContainGsmControl = True
            Exit Function
        End If
    Next
    isContainGsmControl = False
End Function

Public Function isContainUmtsControl() As Boolean
    On Error Resume Next
    Dim cover As String
    Dim key As String
    Dim reValue As String
    Dim index As Long
    
    cover = getResByKey("Cover")
    For index = 3 To 5
        key = ThisWorkbook.Worksheets(cover).Cells(index, 2).value
        reValue = getResByKey(key)
        If reValue = "RNC" Then
            isContainUmtsControl = True
            Exit Function
        End If
    Next
    isContainUmtsControl = False
End Function

Public Function isContainBaseStation() As Boolean
    On Error Resume Next
    Dim cover As String
    Dim key As String
    Dim reValue As String
    Dim index As Long
    
    cover = getResByKey("Cover")
    For index = 3 To 5
        key = ThisWorkbook.Worksheets(cover).Cells(index, 2).value
        reValue = getResByKey(key)
        If reValue = "BaseStation" Then
            isContainBaseStation = True
            Exit Function
        End If
    Next
    isContainBaseStation = False
End Function
