Attribute VB_Name = "mod_toolBar"

' 生成自定义工具栏
Public Function BuildToolBar()
    Dim cmbNewBar As CommandBar
    Dim ctlBtn As CommandBarButton
    Dim barName As String
    On Error Resume Next
    
    barName = "Check frequency band"
    Set cmbNewBar = CommandBars.Add(name:="Operate Bar")
    
    'customized template
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = gCaption_BandTitle
            .TooltipText = gCaption_BandTitle
            .OnAction = "frmShow"
            .FaceId = 50
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
    
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .TooltipText = gCaption_TemplateForm
            .FaceId = 28
            .Caption = gCaption_TemplateForm
            .OnAction = "addTemplate"
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
    If isShtExists(gShtNameSpecialFields) Then
        With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .TooltipText = gCaption_CustomizeTemplate
            .FaceId = 28
            .Caption = gCaption_CustomizeTemplate
            .OnAction = "showCustomizeTemplateForm"
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
    End If
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .TooltipText = getResByKey("Bar_AddComments")
            .FaceId = 186
            .Caption = getResByKey("Bar_AddComments")
            .OnAction = "addAllComments"
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
End Function

' 删除自定义工具栏
Public Function DelToolBar()
  On Error Resume Next
  CommandBars("Operate Bar").Delete
End Function





