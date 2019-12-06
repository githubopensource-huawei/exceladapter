Attribute VB_Name = "CreatBar"
Option Explicit

Public Const CapactiyExpansionBarNameAddMoi As String = "CapacityExpansionAddMoi"
Public Const CapactiyExpansionBarNameDeleteMoi As String = "CapacityExpansionDeleteMoi"
Public Const OperationBarName As String = "Operation Bar"
Public Const AddCommentsBarName As String = "AddComments Bar"

Public Sub initCapacityExpansionToolBar(ByRef ws As Worksheet)
    Call deleteCapacityExpansionToolBar
    
    If isBoardStyleSheet(ws) Then
        Call insertCapacityExpansionToolBar
    End If

End Sub

Public Sub initAujustAntnPortToolBar(ByRef ws As Worksheet)
    Call CapacityCellSub.deleteCellBar
    Call CapacityCellSub.deleteTempBar
    
    Call LampsiteAntNoAdjust.deleteAntAdjustBar
    Call LampsiteAntNoAdjust.deleteAntTempBar
    
    If hasVRXUANTNOColum Then
        '�����С��ҳǩ������С��ҳǩ��ť
        If isCellSheet(ws.name) Then
            Call CapacityCellSub.createCellBar
        ElseIf ws.name = getResByKey("Temp Sheet") Then
            Call CapacityCellSub.createTempBar
        End If
        
        '�����С�������豸��ҳǩ������С�������豸��ҳǩ��ť
        If isSectorEqmGroupSht(ws.name) Then
            Call LampsiteAntNoAdjust.createAntAdjustBar
        ElseIf ws.name = getResByKey("Temp_Adjust_Sheet") Then
            Call LampsiteAntNoAdjust.createAntTempBar
        End If
    End If
        '��?1?��D?��Bo����??��B?�̦�?
    If hasFreqColumn And hasNonFreqColumn Then
        'D???��3???��??��?3y???��o��?����?3y???�̡��?��
        If isCellSheet(ws.name) Then
            Call CapacityCellSub.createDelFreqBar
        End If
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
    Dim adjustBasebandEqmButton As CommandBarButton '���������豸�����ŵİ�ť
    Dim addBoardStyleMoiRefBar As CommandBarButton 'BoardStyle���ú�İ�ť
    Dim delBoardStyleMoiRefBar As CommandBarButton 'BoardStyle���ú�İ�ť
    Set deleteBoardStyleButtons = New CDeleteBoardStyleButtons
    Set deleteBoardStyleMoiBar = Application.CommandBars.Add(CapactiyExpansionBarNameDeleteMoi, msoBarTop)
    With deleteBoardStyleMoiBar
        .Protection = msoBarNoResize
        .Visible = True
        'If isOperationExcel = False Then
            Set deleteBoardStyleMoiButton = .Controls.Add(Type:=msoControlButton)
            With deleteBoardStyleMoiButton
                .Style = msoButtonIconAndCaption
                .Caption = getResByKey("DeleteBoardStyleMoi")
                .TooltipText = getResByKey("DeleteBoardStyleMoi")
                .OnAction = "deleteBoardStyleMoi"
                .FaceId = 293
                .Enabled = True
            End With
        'End If
        
        If hasBASEBANDEQMBOARDColum Then
        Set adjustBasebandEqmButton = .Controls.Add(Type:=msoControlButton)
        With adjustBasebandEqmButton
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("AdjustBasebandEqmBoardNo")
            .TooltipText = getResByKey("AdjustBasebandEqmBoardNo")
            .OnAction = "AdjustBasebandEqmBoardNo"
            .FaceId = 855
            .Enabled = True
        End With
        End If
        
        '�������ú�Bar
        Set addBoardStyleMoiRefBar = .Controls.Add(Type:=msoControlButton)
        With addBoardStyleMoiRefBar
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Bar_Refrence")
            .TooltipText = getResByKey("Bar_Refrence")
            .OnAction = "addListHyperlinks"
            .FaceId = 186
        End With
        
        Set delBoardStyleMoiRefBar = .Controls.Add(Type:=msoControlButton)
        With delBoardStyleMoiRefBar
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("deleteRef")
            .TooltipText = getResByKey("deleteRef")
            .OnAction = "deleteRef"
            .FaceId = 186
            .Enabled = True
        End With
    End With
    Call deleteBoardStyleButtons.initDeleteBoardStyleButtons(deleteBoardStyleMoiButton, addBoardStyleMoiRefBar, delBoardStyleMoiRefBar)

End Sub

Public Sub deleteCapacityExpansionToolBar()
    If containsAToolBar(CapactiyExpansionBarNameAddMoi) Then
        Application.CommandBars(CapactiyExpansionBarNameAddMoi).delete
    End If
    If containsAToolBar(CapactiyExpansionBarNameDeleteMoi) Then
        Application.CommandBars(CapactiyExpansionBarNameDeleteMoi).delete
    End If
End Sub

Public Sub InsertUserToolBar()
    Dim addCommentsBar As CommandBar
    
    If Not containsAToolBar(AddCommentsBarName) Then
        Set addCommentsBar = Application.CommandBars.Add(AddCommentsBarName, msoBarTop)
        With addCommentsBar
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
    End If
    
    Dim toolbar As CommandBar
    
    If Not containsAToolBar(OperationBarName) Then
        Set toolbar = Application.CommandBars.Add(OperationBarName, msoBarTop)
        With toolbar
            .Protection = msoBarNoResize
            .Visible = True
            With .Controls.Add(Type:=msoControlButton)
                .Caption = getResByKey("Bar_Template")
                .OnAction = "addTemplate" '��ɾС��ģ��
                .Style = msoButtonIconAndCaption
                .Enabled = True
                .FaceId = 186
            End With
        End With
    End If
End Sub

Public Sub DeleteUserToolBar()
    If containsAToolBar(AddCommentsBarName) Then
        Application.CommandBars(AddCommentsBarName).delete
    End If
    
    If containsAToolBar(OperationBarName) Then
        Application.CommandBars(OperationBarName).delete
    End If
End Sub







