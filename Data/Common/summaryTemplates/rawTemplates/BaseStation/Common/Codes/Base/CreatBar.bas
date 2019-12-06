Attribute VB_Name = "CreatBar"
Public Sub InsertUserToolBar()
    Dim toolbar As CommandBar
    Dim toolBarExist As Boolean
    Dim neType As String
    neType = getNeType()
    toolBarExist = False
    For Each bar In CommandBars
        If bar.name = "Operate Bar" Then
            toolBarExist = True
            Exit For
        End If
    Next

    If Not toolBarExist Then
        Set toolbar = Application.CommandBars.Add("Operate Bar", msoBarTop)
        With toolbar
            .Protection = msoBarNoResize
            .Visible = True
            'Add User Define Template
            Dim toolBarFun As New ToolBarFunction
            If toolBarFun.templateSupport And neType <> "USU" Then
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = getResByKey("Bar_Template")
                    .OnAction = "addTemplate"
                    .Style = msoButtonIconAndCaption
                    .Enabled = True
                    .FaceId = 186
                End With
            End If
            
            If neType = "LTE" Or neType = "USU" Then
                'LLD<-->Summary
                With .Controls.Add(Type:=msoControlButton)
                   .Style = msoButtonIconAndCaption
                    .Caption = getResByKey("Bar_LLD")
                   .TooltipText = getResByKey("Bar_LLD")
                    .OnAction = "Summary2LLD"
                    .FaceId = 186
                End With
            End If
            
            If neType <> "USU" Then
                With .Controls.Add(Type:=msoControlButton)
                    .Style = msoButtonIconAndCaption
                    .Caption = getResByKey("Bar_IPRoute")
                    .TooltipText = getResByKey("Bar_IPRoute")
                    .OnAction = "addIPRoute"
                    .FaceId = 186
                End With
            End If
            'Add Reference
            With .Controls.Add(Type:=msoControlButton)
                .Style = msoButtonIconAndCaption
                .Caption = getResByKey("Bar_Refrence")
                .TooltipText = getResByKey("Bar_Refrence")
                .OnAction = "addHyperlinks"
                .FaceId = 186
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Style = msoButtonIconAndCaption
                .Caption = getResByKey("Bar_Hidden")
                .TooltipText = getResByKey("Bar_Hidden")
                .OnAction = "hiddenEmptySheet"
                .FaceId = 186
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Style = msoButtonIconAndCaption
                .Caption = getResByKey("Bar_Reset")
                .TooltipText = getResByKey("Bar_Reset")
                .OnAction = "showEmptySheet"
                .FaceId = 186
            End With
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
Public Sub HideToolBar()
    Dim toolBarExist As Boolean
    
    toolBarExist = False
    For Each bar In CommandBars
        If bar.name = "Operate Bar" Then
            toolBarExist = True
            Exit For
        End If
    Next

    If toolBarExist Then
        Application.CommandBars("Operate Bar").Protection = msoBarNoResize
        Application.CommandBars("Operate Bar").Visible = False
    End If
End Sub

Public Sub DeleteUserToolBar()
    Dim toolBarExist As Boolean
    
    toolBarExist = False
    For Each bar In CommandBars
        If bar.name = "Operate Bar" Then
            toolBarExist = True
            Exit For
        End If
    Next

    If toolBarExist Then
        Application.CommandBars("Operate Bar").Delete
    End If
End Sub
