Attribute VB_Name = "CreatBar"
Public Sub InsertUserToolBar()
    Dim toolbar As CommandBar
    Dim toolBarExist As Boolean
    Dim neType As String
    
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
    Dim ranges As Range
    If ThisWorkbook.Worksheets("SHEET DEF").Cells(1, 4).value <> "" Then
        isIubStyleWorkBook = True
    Else
        isIubStyleWorkBook = False
    End If
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






