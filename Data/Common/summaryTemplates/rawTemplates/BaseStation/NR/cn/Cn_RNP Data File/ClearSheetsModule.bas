Attribute VB_Name = "ClearSheetsModule"
Option Explicit

Public Sub clearSheets()
    On Error GoTo ErrorHandler
    
    Dim maxRowNumber As Long
    maxRowNumber = -1
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = getResByKey("Cover") Then
            GoTo NextLoop
        End If
        
        '获取最大行号
        If maxRowNumber = -1 Then
            maxRowNumber = ws.Rows.count
        End If
        
        If ws.Visible = xlSheetVisible Then
            ws.Rows("4:" & maxRowNumber).Clear
        End If
NextLoop:
    Next ws
    
    ThisWorkbook.Save
ErrorHandler:
    Debug.Print "ErrorInfo:" & Err.Description
End Sub
