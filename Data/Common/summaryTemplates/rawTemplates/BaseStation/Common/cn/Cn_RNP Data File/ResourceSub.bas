Attribute VB_Name = "ResourceSub"
Private resource  As CMap

Private languageType As Integer
Private Const Cn = 0
Private Const En = 1
Private keyCol As Integer
Private valueCol As Integer

Public Sub initResource()
    Dim index As Long
    Dim key As String
    Dim value As String
    
    Call getLanguageType
    
    Set resource = New CMap
    Dim summaryResSheet As Worksheet
    Set summaryResSheet = ThisWorkbook.Worksheets("SummaryRes")
    For index = 2 To summaryResSheet.Range("a65536").End(xlUp).row
        key = summaryResSheet.Cells(index, keyCol).value
        value = summaryResSheet.Cells(index, valueCol).value
        Call resource.SetAt(key, value)
    Next
End Sub

Public Function getResByKey(key As String) As String
   On Error Resume Next
    If resource Is Nothing Then
            Call initResource
    End If
    If (resource.hasKey(key)) Then
            getResByKey = resource.GetAt(key)
    Else
            getResByKey = key
    End If
End Function

Private Sub getLanguageType()
    If containsASheet(ThisWorkbook, "Cover") Then
        languageType = En
        keyCol = 1
        valueCol = 3
    Else
        languageType = Cn
        keyCol = 1
        valueCol = 2
    End If
End Sub

