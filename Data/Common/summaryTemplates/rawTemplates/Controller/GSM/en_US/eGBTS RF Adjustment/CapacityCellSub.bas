Attribute VB_Name = "CapacityCellSub"
Option Explicit

Private Const DELFREQ_BAR_NAME = "DeleteTrxBar"
Private CELL_SHEET_NAME As String
Private Const BATCH_DELFREQ_BAR_NAME = "BatchDeleteFreqBar"

Sub createCellBar()
    Dim baseStationChooseBar As CommandBar
    Dim baseStationStyle As CommandBarButton
    Dim delChooseBar As CommandBar
        Dim BatchdelChooseBar As CommandBar
    Dim delFreqStyle As CommandBarButton
        Dim BatchdelFreqStyle As CommandBarButton
    
    Call deleteCellBar
    
    Dim actSheetName As String
    actSheetName = ThisWorkbook.ActiveSheet.name
    CELL_SHEET_NAME = actSheetName
    
    Set delChooseBar = Application.CommandBars.Add(DELFREQ_BAR_NAME, msoBarTop)
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
        Set BatchdelChooseBar = Application.CommandBars.Add(BATCH_DELFREQ_BAR_NAME, msoBarTop)
    Dim BatchdelbarDescLbl As String
    BatchdelbarDescLbl = "BatchDeleteFreq"
    With BatchdelChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set BatchdelFreqStyle = .Controls.Add(Type:=msoControlButton)
        With BatchdelFreqStyle
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey(BatchdelbarDescLbl)
            .TooltipText = getResByKey(BatchdelbarDescLbl)
            .OnAction = "BatchdeleteFrequency"
            .FaceId = 186
            .Enabled = True
        End With
      End With
End Sub

Public Sub initAujustAntnPortToolBar(ByRef ws As Worksheet)
    Call CapacityCellSub.deleteCellBar
    
    If hasFreqColumn And hasNonFreqColumn Then
        If isCellSheet(ws.name) Then
            Call CapacityCellSub.createCellBar
        End If
    End If
End Sub
Private Sub deleteFrequency()
    On Error GoTo ErrorHandler
    
    DeleteFreqForm.Show
    Exit Sub
ErrorHandler:
End Sub

Private Sub BatchdeleteFrequency()
    On Error GoTo ErrorHandler
    
    BatchDeleteFreqForm.Show
    Exit Sub
ErrorHandler:
End Sub

Sub deleteCellBar()
    If existToolBar(DELFREQ_BAR_NAME) Then
        Application.CommandBars(DELFREQ_BAR_NAME).Delete
    End If
        If existToolBar(BATCH_DELFREQ_BAR_NAME) Then
        Application.CommandBars(BATCH_DELFREQ_BAR_NAME).Delete
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


