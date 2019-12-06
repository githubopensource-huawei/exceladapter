Attribute VB_Name = "CreatBar"
Option Explicit

Private Const UniversalBar = "Universal Bar"
Private Const CellShtBar = "Cell Sheet Bar"
Private Const BoardStyleShtBar = "BoardStyle Sheet Bar"

Public Sub insertUserToolBar()
    addUniversalBar
    If isCellSheet(g_CurrentSheet) Then
        insertCellShtBar
    ElseIf isBoardStyleSheet(g_CurrentSheet) Then
        insertBoardStyleShtBar
    End If
    
End Sub

Public Sub deleteUserToolBar()
    Dim toolbar As CommandBar
    
    For Each toolbar In CommandBars
        If toolbar.name = CellShtBar Or toolbar.name = BoardStyleShtBar Or toolbar.name = UniversalBar Then
            toolbar.Delete
        End If
    Next
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

Public Sub insertCellShtBar()
    Dim toolbar As CommandBar
    
    For Each toolbar In CommandBars
        If toolbar.name = CellShtBar Then
            Application.CommandBars(CellShtBar).Delete
            Exit For
        End If
    Next
    
    Set toolbar = Application.CommandBars.Add(CellShtBar, msoBarTop)
        With toolbar
            .Protection = msoBarNoResize
            .Visible = True
            With .Controls.Add(Type:=msoControlButton)
                .Caption = getResByKey("Bar_HidePara")
                .OnAction = "hideParaInCellSheet"
                .Style = msoButtonIconAndCaption
                .Enabled = True
                .FaceId = 186
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = getResByKey("Bar_ShowPara")
                .OnAction = "showParaInCellSheet"
                .Style = msoButtonIconAndCaption
                .Enabled = True
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



