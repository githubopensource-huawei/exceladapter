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
    For index = 2 To Worksheets("SummaryRes").range("a65536").End(xlUp).Row
        key = Worksheets("SummaryRes").Cells(index, keyCol).value
        value = Worksheets("SummaryRes").Cells(index, valueCol).value
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


Public Function getNeType() As String
    On Error Resume Next
    Dim cover As String
    Dim key As String
    Dim reValue As String
    
    cover = getResByKey("Cover")
    key = Worksheets(cover).Cells(2, 2).value
    reValue = getResByKey(key)
    If reValue = key Then
       reValue = "MRAT"
    End If
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
        If reValue = "BaseStation" And ThisWorkbook.Worksheets(cover).Cells(index, 4).value <> "NodeBCommon" Then
            isContainBaseStation = True
            Exit Function
        End If
    Next
    isContainBaseStation = False
End Function

'°üº¬Ä³¸öÒ³Ç©´úÂë
Public Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String) As Boolean
On Error GoTo ErrorHandler
    containsASheet = True
    Dim sheet As Worksheet
    Set sheet = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsASheet = False
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


