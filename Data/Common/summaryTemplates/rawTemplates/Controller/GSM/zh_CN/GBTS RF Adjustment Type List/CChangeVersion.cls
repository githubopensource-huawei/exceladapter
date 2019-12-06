VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CChangeVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function changeVersion(ByRef wb As Workbook, ByRef newVersion As String, Optional ByRef neType As String = "Base Station") As Boolean
    changeVersion = True
    Dim coverName As String
    coverName = getResByKey("Cover")
    With wb.Worksheets(coverName).Cells(3, 4)
        If Trim(.value) = newVersion Then
            changeVersion = False '如果版本号相同，则没有修改版本号，直接退出，标志置为False
            Exit Function
        End If
        .value = newVersion
    End With
    Call changeMappingDefVersion(wb, newVersion, neType)
End Function

'Private Function getCoverName(ByRef wb As Workbook) As String
'    getCoverName = "Cover"
'    If containsASheet(wb, getResByKey("Cover")) Then
'        getCoverName = getResByKey("Cover")
'    End If
'End Function

Private Sub changeMappingDefVersion(ByRef wb As Workbook, ByRef newVersion As String, ByRef neType As String)
    Dim mappingDefSheet As Worksheet
    Set mappingDefSheet = wb.Worksheets("MAPPING DEF")
    
    Dim ucaseNeType As String
    ucaseNeType = UCase(neType)
    
    Dim RowNumber As Long
    For RowNumber = 2 To mappingDefSheet.Range("A1048576").End(xlUp).row
        If UCase(mappingDefSheet.Range("L" & RowNumber).value) = ucaseNeType Then
            mappingDefSheet.Range("M" & RowNumber).value = newVersion
        End If
    Next RowNumber
End Sub
