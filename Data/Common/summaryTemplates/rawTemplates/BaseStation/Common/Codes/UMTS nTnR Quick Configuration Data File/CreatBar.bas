Attribute VB_Name = "CreatBar"
Option Explicit

Private Const UniversalBar = "Universal Bar"
Private Const BoardStyleShtBar = "BoardStyle Sheet Bar"

Public Sub InsertUserToolBar()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    addUniversalBar
    If isCellSheet(ws.name) Then
        createCellBar
    ElseIf isBoardStyleSheet(ws) Then
        insertBoardStyleShtBar
    ElseIf ws.name = getResByKey("Temp Sheet") Then
        createTempBar
    End If
End Sub

Public Sub deleteUserToolBar()
    deleteCellBar
    deleteTempBar
    deleteBoardStyleShtBar
    deleteUniversalBar
End Sub

Private Sub addUniversalBar()
    Dim toolbar2 As CommandBar
    
    For Each toolbar2 In CommandBars
        If toolbar2.name = UniversalBar Then
            Application.CommandBars(UniversalBar).Delete
            Exit For
        End If
    Next
    
    Set toolbar2 = Application.CommandBars.Add(UniversalBar, msoBarTop)
    With toolbar2
        .Protection = msoBarNoResize
        .Visible = True
        With .Controls.Add(Type:=msoControlButton)
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Bar_AddComments")
            .TooltipText = getResByKey("Bar_AddComments")
            .OnAction = "addAllComments"
            .FaceId = 186
        End With
    End With
End Sub

Public Sub insertBoardStyleShtBar()
    Dim toolbar As CommandBar
    
    For Each toolbar In CommandBars
        If toolbar.name = BoardStyleShtBar Then
            Application.CommandBars(BoardStyleShtBar).Delete
            Exit For
        End If
    Next
    
    Set toolbar = Application.CommandBars.Add(BoardStyleShtBar, msoBarTop)
        With toolbar
            .Protection = msoBarNoResize
            .Visible = True
            With .Controls.Add(Type:=msoControlButton)
                .Caption = getResByKey("Bar_ExpandRow")
                .OnAction = "expandRowInBoardStyleSheet"
                .Style = msoButtonIconAndCaption
                .Enabled = True
                .FaceId = 186
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = getResByKey("Bar_DeleteRow")
                .OnAction = "deleteRowInBoardStyleSheet"
                .Style = msoButtonIconAndCaption
                .Enabled = True
                .FaceId = 186
            End With
        End With
End Sub

Private Sub deleteUniversalBar()
    Dim toolbar As CommandBar
    
    For Each toolbar In CommandBars
        If toolbar.name = BoardStyleShtBar Then
            toolbar.Delete
        End If
    Next
End Sub

Public Sub deleteBoardStyleShtBar()
    Dim toolbar As CommandBar
    
    For Each toolbar In CommandBars
        If toolbar.name = UniversalBar Then
            toolbar.Delete
        End If
    Next
End Sub



