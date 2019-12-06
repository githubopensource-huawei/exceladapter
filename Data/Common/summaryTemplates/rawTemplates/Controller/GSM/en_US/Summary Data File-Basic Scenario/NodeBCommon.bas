Attribute VB_Name = "NodeBCommon"
'NodeB Common的特殊处理

Function isNodeBComm() As Boolean
    Dim coverName As String
    Dim issueDate As String
    
    coverName = getResByKey("Cover")
    issueDate = getResByKey("ISSUE_DATE")
    Dim cover As Worksheet
    Set cover = ThisWorkbook.Worksheets(coverName)
    For k = 3 To 10
        If UCase(cover.Cells(k, 2).value) = UCase(issueDate) Then
            Exit For
        End If
        If UCase(cover.Cells(k, 4).value) = UCase("NodeBCommon") Then
            isNodeBComm = True
            Exit Function
        End If
    Next
    isNodeBComm = False
End Function
