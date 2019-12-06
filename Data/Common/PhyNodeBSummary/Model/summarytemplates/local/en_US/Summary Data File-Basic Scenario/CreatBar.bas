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
        End If
    Next
    If toolBarExist Then
       Application.CommandBars("Operate Bar").Delete
    End If
    
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
        
        If (Not isIubStyleWorkBook) And neType <> "USU" Then
            With .Controls.Add(Type:=msoControlButton)
                .Style = msoButtonIconAndCaption
                .Caption = getResByKey("Bar_IPRoute")
                .TooltipText = getResByKey("Bar_IPRoute")
                .OnAction = "addIPRoute"
                .FaceId = 186
            End With
        End If
    End With
   initToolBar (ThisWorkbook.ActiveSheet.name)
End Sub

Private Sub addRefreshBar()
    Dim toolbar2 As CommandBar
    Dim refreshBarExist As Boolean
    For Each bar In CommandBars
        If bar.name = "Refresh Bar" Then
            refreshBarExist = True
        End If
    Next
    If refreshBarExist Then
       Application.CommandBars("Refresh Bar").Delete
    End If
    
    Set toolbar2 = Application.CommandBars.Add("Refresh Bar", msoBarTop)
    With toolbar2
        .Protection = msoBarNoResize
        .Visible = True
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
            .Caption = getResByKey("generalMocView")
            .TooltipText = getResByKey("generalMocView")
            .OnAction = "GenIubFormatReport"
            .FaceId = 186
        End With
    End With
End Sub


Public Sub HideToolBar()
    Dim toolBarExist As Boolean
    Dim refreshBarExist As Boolean
    refreshBarExist = False
    toolBarExist = False
    For Each bar In CommandBars
        If bar.name = "Operate Bar" Then
            toolBarExist = True
        End If
        If bar.name = "Refresh Bar" Then
            refreshBarExist = True
        End If
    Next
    If toolBarExist Then
        Application.CommandBars("Operate Bar").Protection = msoBarNoResize
        Application.CommandBars("Operate Bar").Visible = False
    End If
    
    If refreshBarExist Then
        Application.CommandBars("Refresh Bar").Protection = msoBarNoResize
        Application.CommandBars("Refresh Bar").Visible = False
    End If
    
End Sub

Public Sub DeleteUserToolBar()
    Dim toolBarExist As Boolean
    Dim refreshBarExist As Boolean
    
    toolBarExist = False
    refreshBarExist = False
    For Each bar In CommandBars
        If bar.name = "Operate Bar" Then
            toolBarExist = True
        End If
        If bar.name = "Refresh Bar" Then
            refreshBarExist = True
        End If
    Next

    If toolBarExist Then
        Application.CommandBars("Operate Bar").Delete
    End If
    If refreshBarExist Then
        Application.CommandBars("Refresh Bar").Delete
    End If
End Sub

Sub addMainSheetBar()
    Dim refreshBarExist As Boolean
    refreshBarExist = False
    For Each bar In CommandBars
        If bar.name = "Refresh Bar" Then
            refreshBarExist = True
            Exit For
        End If
    Next
    If refreshBarExist Then
         Application.CommandBars("Refresh Bar").Delete
    End If
    Set toolbar2 = Application.CommandBars.Add("Refresh Bar", msoBarTop)
    With toolbar2
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = getResByKey("UpdateSummary")
            .TooltipText = getResByKey("UpdateSummary")
            .OnAction = "UpdateSummaryFromSiteSheet"
            .FaceId = 19
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
      End With
      With toolbar2
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = getResByKey("GenSiteSheet")
            .TooltipText = getResByKey("GenSiteSheet")
            .OnAction = "GenSiteSheetFromSummary"
            .FaceId = 20
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
      End With
      With toolbar2
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = getResByKey("Import Refrence")
            .TooltipText = getResByKey("Import Refrence")
            .OnAction = "importRef"
            .FaceId = 23
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
      End With
End Sub

Sub deleteRefreshBar()
    Dim refreshBarExist As Boolean
    refreshBarExist = False
    For Each bar In CommandBars
        If bar.name = "Refresh Bar" Then
            refreshBarExist = True
            Exit For
        End If
    Next
    If refreshBarExist Then
         Application.CommandBars("Refresh Bar").Delete
    End If
End Sub

Sub initToolBar(shName As String)
    If isIubStyleWorkBook Then
        If shName = GetMainSheetName() Then
            Call addOperateIubBar
            Call addMainSheetBar
        ElseIf isIubStyleWorkSheet(shName) Then
            Call addOperateIubBar
            Call addIubBar
        Else
            Call deleteRefreshBar
        End If
    Else
        Call addRefreshBar
    End If
End Sub

Private Sub addOperateIubBar()
    Dim toolbar2 As CommandBar
    Dim refreshBarExist As Boolean
    For Each bar In CommandBars
        If bar.name = "Operate Bar" Then
            refreshBarExist = True
        End If
    Next
    If refreshBarExist Then
       Application.CommandBars("Operate Bar").Delete
    End If
    Set toolbar2 = Application.CommandBars.Add("Operate Bar", msoBarTop)
    With toolbar2
        .Protection = msoBarNoResize
        .Visible = True
        With .Controls.Add(Type:=msoControlButton)
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Bar_Template")
            .TooltipText = getResByKey("Bar_Template")
            .OnAction = "addTemplate"
            .FaceId = 186
        End With
    End With
End Sub

Function isIubStyleWorkSheet(shName As String) As Boolean
    If ThisWorkbook.Worksheets(shName).Tab.colorIndex = BluePrintSheetColor Then
        isIubStyleWorkSheet = True
    Else
        isIubStyleWorkSheet = False
    End If
End Function

Function isIubStyleWorkSheetByParameterWs(ByRef sheet As Worksheet) As Boolean
    If sheet.Tab.colorIndex = BluePrintSheetColor Then
        isIubStyleWorkSheetByParameterWs = True
    Else
        isIubStyleWorkSheetByParameterWs = False
    End If
End Function

Function isIubStyleWorkBook() As Boolean
    On Error GoTo ErrorHandler:
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets
        If isIubStyleWorkSheet(sht.name) Then
            isIubStyleWorkBook = True
            Exit Function
        End If
    Next
    
    isIubStyleWorkBook = False
    Exit Function
    
ErrorHandler:
    isIubStyleWorkBook = False
End Function

Sub addIubBar()
    Dim refreshBarExist As Boolean
    refreshBarExist = False
    For Each bar In CommandBars
        If bar.name = "Refresh Bar" Then
            refreshBarExist = True
            Exit For
        End If
    Next
    If refreshBarExist Then
        Application.CommandBars("Refresh Bar").Delete
    End If
    Set toolbar2 = Application.CommandBars.Add("Refresh Bar", msoBarTop)
    With toolbar2
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = getResByKey("Bar_AddRow")
            .TooltipText = getResByKey("Bar_AddRow")
            .OnAction = "AddIubRow"  'New Implement
            .FaceId = 3183
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
     End With
    With toolbar2
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = getResByKey("Bar_HideEmptyRow")
            .TooltipText = getResByKey("Bar_HideEmptyRow")
            .OnAction = "HideEmptyRow"  'New Implement
            .FaceId = 54
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
     End With
     With toolbar2
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = getResByKey("Bar_ShowEmptyRow")
            .TooltipText = getResByKey("Bar_ShowEmptyRow")
            .OnAction = "ShowEmptyRow"
            .FaceId = 55
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
        
    With toolbar2
        .Protection = msoBarNoResize
        .Visible = True
        With .Controls.Add(Type:=msoControlButton)
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Bar_Refrence")
            .TooltipText = getResByKey("Bar_Refrence")
            .OnAction = "addListHyperlinks"
            .FaceId = 186
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("deleteRef")
            .TooltipText = getResByKey("deleteRef")
            .OnAction = "deleteRef"
            .FaceId = 186
        End With
    End With
End Sub




