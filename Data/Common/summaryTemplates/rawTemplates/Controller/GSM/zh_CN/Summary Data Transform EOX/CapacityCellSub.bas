Attribute VB_Name = "CapacityCellSub"
Option Explicit

Private Const DELFREQ_BAR_NAME = "DeleteTrxBar"
Private CELL_SHEET_NAME As String

Sub createCellBar()
    Dim baseStationChooseBar As CommandBar
    Dim baseStationStyle As CommandBarButton
    Dim delChooseBar As CommandBar
    Dim delFreqStyle As CommandBarButton
    
    Call deleteCellBar
    
    Dim actSheetName As String
    actSheetName = ThisWorkbook.ActiveSheet.name
    CELL_SHEET_NAME = actSheetName
    
    Set delChooseBar = Application.CommandBars.Add(DELFREQ_BAR_NAME, msoBarBottom)
    Dim delbarDescLbl As String
    delbarDescLbl = "DeleteFreq"
    With delChooseBar
       .Protection = msoBarNoResize
       .Visible = True
       Set delFreqStyle = .Controls.Add(Type:=msoControlButton)
       With delFreqStyle
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey(delbarDescLbl)
            .TooltipText = getResByKey(delbarDescLbl)
            .OnAction = "deleteFrequency"
            .FaceId = 186
            .Enabled = True
        End With
      End With
End Sub

Public Sub initAujustAntnPortToolBar(ByRef ws As Worksheet)
    Call CapacityCellSub.deleteCellBar
    
    If isCellSheet(ws.name) Then
         Call CapacityCellSub.createCellBar
    End If

End Sub
Private Sub deleteFrequency()
    On Error GoTo ErrorHandler
    
    DeleteFreqForm.Show
    Exit Sub
ErrorHandler:
End Sub

Sub deleteCellBar()
    If existToolBar(DELFREQ_BAR_NAME) Then
        Application.CommandBars(DELFREQ_BAR_NAME).Delete
    End If
End Sub
 
 Sub initTempSheetControl(ByRef flag As Boolean)
    On Error Resume Next
    Dim k As Long
    Dim controlId As Long
    For k = 1 To Application.CommandBars("Ply").Controls.count
        controlId = Application.CommandBars("Ply").Controls(k).ID
        Application.CommandBars("Ply").FindControl(ID:=controlId).Enabled = flag
    Next
    With Application.CommandBars("Column")
        .FindControl(ID:=3183).Enabled = flag
        .FindControl(ID:=297).Enabled = flag
        .FindControl(ID:=294).Enabled = flag
    End With
End Sub
Private Function existToolBar(ByRef barName As String) As Boolean
    On Error GoTo ErrorHandler
    existToolBar = True
    Dim bar As CommandBar
    Set bar = CommandBars(barName)
    Exit Function
ErrorHandler:
    existToolBar = False
End Function

Private Function getCurrentRegionRowsCount(ByRef ws As Worksheet, ByRef startRowNumber As Long) As Long
    Dim RowNumber As Long
    Dim rowscount As Long
    rowscount = 1
    For RowNumber = startRowNumber + 1 To 2000
        If rowIsBlank(ws, RowNumber) = True Then
            Exit For
        Else
            rowscount = rowscount + 1
        End If
    Next RowNumber
    getCurrentRegionRowsCount = rowscount
End Function
