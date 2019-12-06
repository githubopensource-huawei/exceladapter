Attribute VB_Name = "CreatBar"
Private Const GenMocViewBar = "GenMocView Bar"
Private Const EmptySheetBar = "EmptySheetBar"
Private Const MainSheetBar = "MainSheetBar"
Private Const IubSheetBar = "IubSheetBar"
Private Const PatternSheetBar = "PatternSheetBar"
Private Const ReferenceBar = "Reference Bar"

Private Const SheetType_List = "LIST"
Private Const SheetType_Pattern = "PATTERN"
Private Const SheetType_Main = "MAIN"
Private Const SheetType_Iub = "IUB"
Private g_CurrentSheetType As String
Public Const OperationBarName As String = "Operation Bar"



Public Sub InsertUserToolBar()
    initToolBar (ThisWorkbook.ActiveSheet.name)
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

Public Sub initToolBar(shName As String)
    On Error Resume Next
    Call DeleteUserToolBar
    
    Call initSheetType(shName)

    If isIubStyleWorkBook Then
        If g_CurrentSheetType = SheetType_Pattern Then
            Call addReferenceBar
            Exit Sub
        End If
        
        If g_CurrentSheetType = SheetType_Main Then
            Call addMainSheetBar
            Exit Sub
        End If
        
        If g_CurrentSheetType = SheetType_Iub Then
            Call addIubSheetBar
            Exit Sub
        End If
    Else
        Call addEmptySheetBar
        Call addGenMocViewBar
        If g_CurrentSheetType = SheetType_Pattern Then
            Call addReferenceBar
        End If
    End If
End Sub

Public Sub DeleteUserToolBar()
    On Error Resume Next
    Call deleteGenMocViewBar
    Call deleteEmptySheetBar
    Call deleteMainSheetBar
    Call deleteIubSheetBar
    Call deleteReferenceBar
    Call deleteRefreshBar
    If containsAToolBar(OperationBarName) Then
        Application.CommandBars(OperationBarName).Delete
    End If
End Sub


Private Sub addGenMocViewBar()
    Set toolbar2 = Application.CommandBars.Add(GenMocViewBar, msoBarTop)
    With toolbar2
        .Protection = msoBarNoResize
        .Visible = True
        With .Controls.Add(Type:=msoControlButton)
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("generalMocView")
            .TooltipText = getResByKey("generalMocView")
            .OnAction = "GenIubFormatReport"
            .FaceId = 186
        End With
    End With
End Sub

Private Sub addEmptySheetBar()
    Set toolbar2 = Application.CommandBars.Add(EmptySheetBar, msoBarTop)
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
    End With
End Sub

Private Sub addMainSheetBar()
    Set toolbar2 = Application.CommandBars.Add(MainSheetBar, msoBarTop)
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

Sub addIubSheetBar()
    Set toolbar2 = Application.CommandBars.Add(IubSheetBar, msoBarTop)
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
    
    Call addReferenceBar
End Sub

Private Sub addReferenceBar()
    Set toolbar2 = Application.CommandBars.Add(ReferenceBar, msoBarTop)
    With toolbar2
        .Protection = msoBarNoResize
        .Visible = True
        With .Controls.Add(Type:=msoControlButton)
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Bar_Refrence")
            .TooltipText = getResByKey("Bar_Refrence")
            .OnAction = "addHyperlinks"
            .FaceId = 186
        End With
        If g_CurrentSheetType = SheetType_Iub Then
            With .Controls.Add(Type:=msoControlButton)
                .Style = msoButtonIconAndCaption
                .Caption = getResByKey("deleteRef")
                .TooltipText = getResByKey("deleteRef")
                .OnAction = "deleteRef"
                .FaceId = 186
            End With
        End If
    End With
End Sub

'Private Sub addPatternSheetBar()
'    Set toolbar2 = Application.CommandBars.Add(PatternSheetBar, msoBarTop)
'
'    With toolbar2
'        .Protection = msoBarNoResize
'        .Visible = True
'        With .Controls.Add(Type:=msoControlButton)
'            .Style = msoButtonIconAndCaption
'            .Caption = getResByKey("Bar_Refrence")
'            .TooltipText = getResByKey("Bar_Refrence")
'            .OnAction = "addHyperlinks"
'            .FaceId = 186
'        End With
'    End With
'
'    Call addEmptyShtBar
'End Sub

Private Sub deleteGenMocViewBar()
    Dim genMocViewBarExist As Boolean
    genMocViewBarExist = False
    For Each bar In CommandBars
        If bar.name = GenMocViewBar Then
            genMocViewBarExist = True
            Exit For
        End If
    Next
    If genMocViewBarExist Then
         Application.CommandBars(GenMocViewBar).Delete
    End If
End Sub

Private Sub deleteEmptySheetBar()
    Dim emptySheetBarExist As Boolean
    emptySheetBarExist = False
    For Each bar In CommandBars
        If bar.name = EmptySheetBar Then
            emptySheetBarExist = True
            Exit For
        End If
    Next
    If emptySheetBarExist Then
         Application.CommandBars(EmptySheetBar).Delete
    End If
End Sub

Private Sub deleteMainSheetBar()
    Dim mainSheetBarExist As Boolean
    mainSheetBarExist = False
    For Each bar In CommandBars
        If bar.name = MainSheetBar Then
            mainSheetBarExist = True
            Exit For
        End If
    Next
    If mainSheetBarExist Then
         Application.CommandBars(MainSheetBar).Delete
    End If
End Sub

Private Sub deleteIubSheetBar()
    Dim iubSheetBarExist As Boolean
    iubSheetBarExist = False
    For Each bar In CommandBars
        If bar.name = IubSheetBar Then
            iubSheetBarExist = True
            Exit For
        End If
    Next
    If iubSheetBarExist Then
         Application.CommandBars(IubSheetBar).Delete
    End If
End Sub

Private Sub deleteReferenceBar()
    Dim referenceBarExist As Boolean
    referenceBarExist = False
    For Each bar In CommandBars
        If bar.name = ReferenceBar Then
            referenceBarExist = True
            Exit For
        End If
    Next
    If referenceBarExist Then
         Application.CommandBars(ReferenceBar).Delete
    End If
End Sub

Private Sub deleteRefreshBar()
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

'Private Sub deletePatternShtBar()
'    Dim patternShtBarExist As Boolean
'    patternShtBarExist = False
'    For Each Bar In CommandBars
'        If Bar.name = PatternSheetBar Then
'            patternShtBarExist = True
'            Exit For
'        End If
'    Next
'    If patternShtBarExist Then
'        Application.CommandBars(PatternSheetBar).Delete
'    End If
'End Sub

Private Sub initSheetType(ByRef shtName As String)
    On Error Resume Next
    g_CurrentSheetType = ""
    If UCase(getSheetType(shtName)) = "LIST" Then
        g_CurrentSheetType = SheetType_List
    ElseIf UCase(getSheetType(shtName)) = "PATTERN" Then
        g_CurrentSheetType = SheetType_Pattern
    ElseIf UCase(getSheetType(shtName)) = "MAIN" Then
        g_CurrentSheetType = SheetType_Main
    ElseIf UCase(getSheetType(shtName)) = "" And isIubStyleWorkSheet(shtName) Then
        g_CurrentSheetType = SheetType_Iub
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


Public Function containsAToolBar(ByRef barName As String) As Boolean
    On Error GoTo ErrorHandler
    containsAToolBar = True
    Dim bar As CommandBar
    Set bar = CommandBars(barName)
    Exit Function
ErrorHandler:
    containsAToolBar = False
End Function

