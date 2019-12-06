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
            If neType <> "USU" Then
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = getResByKey("Bar_Template")
                    .OnAction = "addTemplate"
                    .Style = msoButtonIconAndCaption
                    .Enabled = True
                    .FaceId = 186
                End With
            End If
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
