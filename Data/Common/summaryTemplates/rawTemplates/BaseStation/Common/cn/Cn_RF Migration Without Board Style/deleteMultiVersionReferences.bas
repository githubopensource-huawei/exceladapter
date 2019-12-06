Attribute VB_Name = "deleteMultiVersionReferences"

Public Const refColInModDiffSht As Integer = 1



Public Function isMultiVersionWb() As Boolean
    isMultiVersionWb = False
    If existsASheet(getResByKey("ModelDiffSht")) Then
        isMultiVersionWb = True
    End If
End Function


Public Sub delModelDiffReferences()
    Dim modelDiffSht As Worksheet
    Set modelDiffSht = ThisWorkbook.Worksheets(getResByKey("ModelDiffSht"))
    
    Dim targetShtName As String
    Dim targetGrpName As String
    Dim targetColName As String
    Dim refArray As Variant
    Dim rowIdx As Integer
    
    With modelDiffSht
        For rowIdx = 2 To .Range("a1048576").End(xlUp).row
            If Not isValidReference(.Cells(rowIdx, refColInModDiffSht), refArray, "\") Then GoTo NextLoop
            
            targetShtName = refArray(0)
            targetGrpName = refArray(1)
            targetColName = refArray(2)
            
            Dim shtType As String
            If getSheetType(targetShtName) = "BOARD" Then
                With .Cells(rowIdx, refColInModDiffSht)
                    .Hyperlinks.Delete
                End With
            End If
NextLoop:
        Next
        Call setBorders(.UsedRange)
    End With
End Sub

'检查是否为合法超链接格式
Public Function isValidReference(refAddr As String, Optional refArray As Variant, Optional delimeter As String) As Boolean
    isValidReference = False
    
    If delimeter <> "" Then
        refArray = Split(refAddr, delimeter)
        If UBound(refArray) <> 2 Then Exit Function
        If refArray(0) = "" Or refArray(1) = "" Or refArray(2) = "" Then Exit Function
        isValidReference = True
        Exit Function
    End If
    
    If isValidReference(refAddr, refArray, "\") Then
        isValidReference = True
        Exit Function
    End If
    
    If isValidReference(refAddr, refArray, ".") Then
        isValidReference = True
        Exit Function
    End If
End Function

Public Function existsASheet(shtName As String) As Boolean
On Error GoTo ErrorHandler:
    existsASheet = True
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(shtName)
    Exit Function
ErrorHandler:
    existsASheet = False
End Function

