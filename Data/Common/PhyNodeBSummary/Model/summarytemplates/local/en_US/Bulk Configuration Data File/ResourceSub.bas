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
    For index = 2 To SummaryRes.Range("a65536").End(xlUp).row
        key = SummaryRes.Cells(index, keyCol).value
        value = SummaryRes.Cells(index, valueCol).value
        Call resource.SetAt(key, value)
    Next
End Sub

Public Function getResByKey(key As String) As String
   On Error Resume Next
    If resource Is Nothing Then
        Call initResource
    End If
    If resource.haskey(key) Then
        getResByKey = resource.GetAt(key)
    Else
        getResByKey = key
    End If
End Function

'Public Function isContainGsmControl() As Boolean
'    On Error Resume Next
'    Dim cover As String
'    Dim key As String
'    Dim reValue As String
'    Dim index As Long
'
'    cover = getResByKey("Cover")
'    For index = 3 To 5
'        key = ThisWorkbook.Worksheets(cover).Cells(index, 2).value
'        reValue = getResByKey(key)
'        If reValue = "BSC" Then
'            isContainGsmControl = True
'            Exit Function
'        End If
'    Next
'    isContainGsmControl = False
'End Function

'Public Function isContainUmtsControl() As Boolean
'    On Error Resume Next
'    Dim cover As String
'    Dim key As String
'    Dim reValue As String
'    Dim index As Long
'
'    cover = getResByKey("Cover")
'    For index = 3 To 5
'        key = ThisWorkbook.Worksheets(cover).Cells(index, 2).value
'        reValue = getResByKey(key)
'        If reValue = "RNC" Then
'            isContainUmtsControl = True
'            Exit Function
'        End If
'    Next
'    isContainUmtsControl = False
'End Function
'
'Public Function isContainBaseStation() As Boolean
'    On Error Resume Next
'    Dim cover As String
'    Dim key As String
'    Dim reValue As String
'    Dim index As Long
'
'    cover = getResByKey("Cover")
'    For index = 3 To 5
'        key = ThisWorkbook.Worksheets(cover).Cells(index, 2).value
'        reValue = getResByKey(key)
'        If reValue = "BaseStation" And ThisWorkbook.Worksheets(cover).Cells(index, 4).value <> "NodeBCommon" Then
'            isContainBaseStation = True
'            Exit Function
'        End If
'    Next
'    isContainBaseStation = False
'End Function

Private Sub getLanguageType()
    If containsASheet("Cover") Then
        languageType = En
        keyCol = 1
        valueCol = 3
    Else
        languageType = Cn
        keyCol = 1
        valueCol = 2
    End If
End Sub



