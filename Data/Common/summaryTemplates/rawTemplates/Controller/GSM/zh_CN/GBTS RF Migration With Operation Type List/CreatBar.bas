Attribute VB_Name = "CreatBar"
Option Explicit

Public Const CapactiyExpansionBarNameAddMoi As String = "CapacityExpansionAddMoi"
Public Const CapactiyExpansionBarNameDeleteMoi As String = "CapacityExpansionDeleteMoi"
Public Const OperationBarName As String = "Operation Bar"

Public Sub initCapacityExpansionToolBar(ByRef ws As Worksheet)
    Call deleteCapacityExpansionToolBar
    
    If isBoardStyleSheet(ws) Then
        Call insertCapacityExpansionToolBar
    End If

End Sub

Public Sub initAujustAntnPortToolBar(ByRef ws As Worksheet)
    Call CapacityCellSub.deleteCellBar
    Call CapacityCellSub.deleteTempBar
    
    If isGTRXSheet(ws.name) Then
         Call CapacityCellSub.createCellBar
    ElseIf ws.name = getResByKey("Temp Sheet") Then
        Call CapacityCellSub.createTempBar
    End If

End Sub

Private Sub insertCapacityExpansionToolBar()
    Call deleteCapacityExpansionToolBar
    
    Call insertAddBoardStyleMoiBar
    
    Call insertDeleteBoardStyleMoiBar
    
    If inAddProcessFlag = True Then
        Call addBoardStyleButtons.setAddBoardStyleButtons
        Call deleteBoardStyleButtons.setDeleteBoardStyleButtons
    End If
End Sub

Private Sub insertAddBoardStyleMoiBar()
    Dim addBoardStyleMoiBar As CommandBar
    Dim addBoardStyleMoiButton As CommandBarButton
    Dim addBoardStyleMoiFinishButton As CommandBarButton
    Dim addBoardStyleMoiCancelButton As CommandBarButton
    
    Set addBoardStyleButtons = New CAddBoardStyleButtons
    'Set addBoardStyleMoiBarCol = New Collection
    Set addBoardStyleMoiBar = Application.CommandBars.Add(CapactiyExpansionBarNameAddMoi, msoBarTop)
    With addBoardStyleMoiBar
        .Protection = msoBarNoResize
        .Visible = True
        Set addBoardStyleMoiButton = .Controls.Add(Type:=msoControlButton)
        'addBoardStyleMoiBarCol.Add Item:=addBoardStyleMoiButton, key:="add"
        With addBoardStyleMoiButton
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("AddBoardStyleMoi")
            .TooltipText = getResByKey("AddBoardStyleMoi")
            .OnAction = "addBoardStyleMoi"
            .FaceId = 3183
            .Enabled = True
        End With
        
        Set addBoardStyleMoiFinishButton = .Controls.Add(Type:=msoControlButton)
        'addBoardStyleMoiBarCol.Add Item:=addBoardStyleMoiFinishButton, key:="finish"
        With addBoardStyleMoiFinishButton
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Finish")
            .TooltipText = getResByKey("Finish")
            .OnAction = "addBoardStyleMoiFinishButton"
            .FaceId = 186
            .Enabled = False
        End With
        
        Set addBoardStyleMoiCancelButton = .Controls.Add(Type:=msoControlButton)
        'addBoardStyleMoiBarCol.Add Item:=addBoardStyleMoiCancelButton, key:="cancel"
        With addBoardStyleMoiCancelButton
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Cancel")
            .TooltipText = getResByKey("Cancel")
            .OnAction = "addBoardStyleMoiCancelButton"
            .FaceId = 186
            .Enabled = False
        End With
    End With
    Call addBoardStyleButtons.initAddBoardStyleButtons(addBoardStyleMoiButton, addBoardStyleMoiFinishButton, addBoardStyleMoiCancelButton)
End Sub

Private Sub insertDeleteBoardStyleMoiBar()
    Dim deleteBoardStyleMoiBar As CommandBar
    Dim deleteBoardStyleMoiButton As CommandBarButton
    Set deleteBoardStyleButtons = New CDeleteBoardStyleButtons
    Set deleteBoardStyleMoiBar = Application.CommandBars.Add(CapactiyExpansionBarNameDeleteMoi, msoBarTop)
    With deleteBoardStyleMoiBar
        .Protection = msoBarNoResize
        .Visible = True
        Set deleteBoardStyleMoiButton = .Controls.Add(Type:=msoControlButton)
        With deleteBoardStyleMoiButton
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("DeleteBoardStyleMoi")
            .TooltipText = getResByKey("DeleteBoardStyleMoi")
            .OnAction = "deleteBoardStyleMoi"
            .FaceId = 293
            .Enabled = True
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Bar_Refrence")
            .TooltipText = getResByKey("Bar_Refrence")
            .OnAction = "addListHyperlinks"
            .FaceId = 186
        End With
    End With
    Call deleteBoardStyleButtons.initDeleteBoardStyleButtons(deleteBoardStyleMoiButton)
End Sub

Public Sub deleteCapacityExpansionToolBar()
    If containsAToolBar(CapactiyExpansionBarNameAddMoi) Then
        Application.CommandBars(CapactiyExpansionBarNameAddMoi).Delete
    End If
    If containsAToolBar(CapactiyExpansionBarNameDeleteMoi) Then
        Application.CommandBars(CapactiyExpansionBarNameDeleteMoi).Delete
    End If
End Sub

Public Sub InsertUserToolBar()
    Dim toolbar As CommandBar
    
    If Not containsAToolBar(OperationBarName) Then
        Set toolbar = Application.CommandBars.Add(OperationBarName, msoBarTop)
        With toolbar
            .Protection = msoBarNoResize
            .Visible = True
'            With .Controls.Add(Type:=msoControlButton)
'                .Caption = getResByKey("Bar_Template")
'                .OnAction = "addTemplate" '增删小区模板
'                .Style = msoButtonIconAndCaption
'                .Enabled = True
'                .FaceId = 186
'            End With
            With .Controls.Add(Type:=msoControlButton)
                .Style = msoButtonIconAndCaption
                .Caption = getResByKey("Bar_AddComments")
                .TooltipText = getResByKey("Bar_AddComments")
                .OnAction = "addAllComments"
                .FaceId = 186
            End With
        End With
    End If
End Sub

Public Sub DeleteUserToolBar()
    If containsAToolBar(OperationBarName) Then
        Application.CommandBars(OperationBarName).Delete
    End If
End Sub







