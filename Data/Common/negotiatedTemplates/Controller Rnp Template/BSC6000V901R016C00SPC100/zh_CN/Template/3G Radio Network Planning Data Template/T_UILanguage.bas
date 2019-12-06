Attribute VB_Name = "T_UILanguage"
Option Explicit

Public CustomizeToolbarName As String

Public Sub InitUIStringResource()
    If iLanguageType = 0 Then
        CustomizeToolbarName = "Customize Template"
    Else
        CustomizeToolbarName = "¶¨ÖÆÄ£°å"
    End If
End Sub

Private Sub SwitchUILanguage()
    
    Dim FieldPostil As String, RangeName As String, IsMustGiveFlag As String, MustGivePostil As String
    Dim DisplayName As String, FieldCol As String, FieldRow As String
    Dim SheetID As String, SheetName As String, FoundID As String
    Dim nFieldNameDisplayCol As Integer, nMustGiveCol As Integer
    Dim iSheet As Integer, iDefSheet As Integer
    Dim BField As Boolean
    Dim CurSheet As Worksheet
    Const FieldDefCol = 4
    
    Call GetSheetDefineData
    Call GetAllSheetName
    Call SetCoverLanguage

    If iLanguageType = 0 Then
        nFieldNameDisplayCol = 12
        MustGivePostil = gEngIsMustGive
    Else
        nFieldNameDisplayCol = 13
        MustGivePostil = gChsIsMustGive
    End If
    nMustGiveCol = 16

    For iSheet = 0 To UBound(ArrSheetName) - 1
        BField = False
        SheetID = Trim(ArrSheetName(iSheet, 0))
        SheetName = Trim(ArrSheetName(iSheet, 1))
        Set CurSheet = ThisWorkbook.Sheets(SheetName)
        FieldRow = "2"
        If SheetName = "DoubleFrequencyCell" Then
            FieldRow = "3"
        End If

        For iDefSheet = 0 To UBound(SheetDefine) - 1
            FoundID = Trim(SheetDefine(iDefSheet, 0))
            If SheetID = FoundID Then
                BField = True
                Exit For
            End If
        Next

        If BField Then
            Do
                DisplayName = Trim(SheetDefine(iDefSheet, nFieldNameDisplayCol))
                IsMustGiveFlag = Trim(SheetDefine(iDefSheet, nMustGiveCol))
                FieldCol = Trim(SheetDefine(iDefSheet, FieldDefCol))
                CurSheet.Range(FieldCol + FieldRow) = DisplayName
                
                RangeName = GetRangeInfo(iDefSheet)
                FieldPostil = DisplayName + "(" + RangeName + ")"
                If UCase(IsMustGiveFlag) = "YES" Then
                    FieldPostil = FieldPostil + MustGivePostil
                End If
                CurSheet.Range(FieldCol + FieldRow).ClearComments
                CurSheet.Range(FieldCol + FieldRow).AddComment FieldPostil
                CurSheet.Range(FieldCol + FieldRow).Comment.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
                CurSheet.Range(FieldCol + FieldRow).Comment.Shape.ScaleHeight 1, msoFalse, msoScaleFromTopLeft

                iDefSheet = iDefSheet + 1
                If iDefSheet >= TblRows Then Exit Do
                FoundID = Trim(SheetDefine(iDefSheet, 0))
            Loop While FoundID = ""
        End If

    Next iSheet
End Sub

Private Sub SetCoverLanguage()
    Dim CurSheet As Worksheet
    Set CurSheet = ThisWorkbook.Sheets("Cover")
    
    CurSheet.Unprotect "HWCME"
    
    If iLanguageType = 0 Then
        CurSheet.Range("C4:D4") = gEngTemplateName
        'CurSheet.Range("C9") = gEngCMEVersionTitle
        CurSheet.Range("C10") = gEngNEVersion
        CurSheet.Range("C11") = gEngRNPVersion
        CurSheet.Range("C15:D15") = sEngCoverInfoTitle
        If iHideSheetFlg = 0 Then
            CurSheet.Range("C16:D16") = sEngCoverInfo1
        Else
            CurSheet.Range("C16:D16") = sEngCoverInfo1 + sEngCoverInfo2
        End If
    Else
        CurSheet.Range("C4:D4") = gChsTemplateName
        'CurSheet.Range("C9") = gChsCMEVersionTitle
        CurSheet.Range("C10") = gChsNEVersion
        CurSheet.Range("C11") = gChsRNPVersion
        CurSheet.Range("C15:D15") = sChsCoverInfoTitle
        If iHideSheetFlg = 0 Then
            CurSheet.Range("C16:D16") = sChsCoverInfo1
        Else
            CurSheet.Range("C16:D16") = sChsCoverInfo1 + sChsCoverInfo2
        End If
    End If

    CurSheet.Protect "HWCME", DrawingObjects:=True, Contents:=True, Scenarios:=True
    
End Sub

'-------------------------------------------------
Public Sub SetEnglishUI()
    If iLanguageType = 0 Then
        Exit Sub
    End If
        
    iLanguageType = 0
    
    InitUIStringResource
        
    ThisWorkbook.Sheets("TableDef").Range("H9") = CStr(iLanguageType)
    Call SwitchErrorDefine
    
    ThisWorkbook.Sheets("Cover").Activate
    CommandBars("Operate Bar").Controls.Item(1).caption = "English Version"
    CommandBars("Operate Bar").Controls.Item(1).TooltipText = "English Version"
    CommandBars("Operate Bar").Controls.Item(2).caption = "Chinese Version"
    CommandBars("Operate Bar").Controls.Item(2).TooltipText = "Chinese Version"
    'If iHideSheetFlg = 0 Then
        'CommandBars("Operate Bar").Controls.Item(3).caption = "Show Extension Sheet"
        'CommandBars("Operate Bar").Controls.Item(3).TooltipText = "Show Extension Sheet"
    'Else
        'CommandBars("Operate Bar").Controls.Item(3).caption = "Hide Extension Sheet"
        'CommandBars("Operate Bar").Controls.Item(3).TooltipText = "Hide Extension Sheet"
    'End If
    
    CommandBars("Operate Bar").Controls.Item(3).caption = CustomizeToolbarName
    CommandBars("Operate Bar").Controls.Item(3).TooltipText = CustomizeToolbarName
    
    Call SwitchUILanguage
End Sub

Public Sub SetChineseUI()
    If iLanguageType = 1 Then
        Exit Sub
    End If
        
    iLanguageType = 1
    
    InitUIStringResource
    
    ThisWorkbook.Sheets("TableDef").Range("H9") = CStr(iLanguageType)
    Call SwitchErrorDefine
    
    ThisWorkbook.Sheets("Cover").Activate
    CommandBars("Operate Bar").Controls.Item(1).caption = ThisWorkbook.Sheets("TableDef").Range("H11").Text
    CommandBars("Operate Bar").Controls.Item(1).TooltipText = ThisWorkbook.Sheets("TableDef").Range("H11").Text
    CommandBars("Operate Bar").Controls.Item(2).caption = ThisWorkbook.Sheets("TableDef").Range("H12").Text
    CommandBars("Operate Bar").Controls.Item(2).TooltipText = ThisWorkbook.Sheets("TableDef").Range("H12").Text
    'If iHideSheetFlg = 0 Then
        'CommandBars("Operate Bar").Controls.Item(3).caption = ThisWorkbook.Sheets("TableDef").Range("H13").Text
        'CommandBars("Operate Bar").Controls.Item(3).TooltipText = ThisWorkbook.Sheets("TableDef").Range("H13").Text
    'Else
        'CommandBars("Operate Bar").Controls.Item(3).caption = ThisWorkbook.Sheets("TableDef").Range("H14").Text
        'CommandBars("Operate Bar").Controls.Item(3).TooltipText = ThisWorkbook.Sheets("TableDef").Range("H14").Text
    'End If
    
    CommandBars("Operate Bar").Controls.Item(3).caption = CustomizeToolbarName
    CommandBars("Operate Bar").Controls.Item(3).TooltipText = CustomizeToolbarName
    
    Call SwitchUILanguage
End Sub

Public Sub SwitchErrorDefine()
    If iLanguageType = 1 Then
        ThisWorkbook.Sheets("TableDef").Range("F5") = ThisWorkbook.Sheets("TableDef").Range("D5")
        ThisWorkbook.Sheets("TableDef").Range("F6") = ThisWorkbook.Sheets("TableDef").Range("D6")
        ThisWorkbook.Sheets("TableDef").Range("F7") = ThisWorkbook.Sheets("TableDef").Range("D7")
        
        ThisWorkbook.Sheets("TableDef").Range("E5") = ThisWorkbook.Sheets("TableDef").Range("C5")
        ThisWorkbook.Sheets("TableDef").Range("E6") = ThisWorkbook.Sheets("TableDef").Range("C6")
        ThisWorkbook.Sheets("TableDef").Range("E7") = ThisWorkbook.Sheets("TableDef").Range("C7")
        ThisWorkbook.Sheets("TableDef").Range("E8") = ThisWorkbook.Sheets("TableDef").Range("C8")
    Else
        ThisWorkbook.Sheets("TableDef").Range("F5") = "Range"
        ThisWorkbook.Sheets("TableDef").Range("F6") = "Range"
        ThisWorkbook.Sheets("TableDef").Range("F7") = "Length"
        
        ThisWorkbook.Sheets("TableDef").Range("E5") = "Prompt"
        ThisWorkbook.Sheets("TableDef").Range("E6") = "Prompt"
        ThisWorkbook.Sheets("TableDef").Range("E7") = "Prompt"
        ThisWorkbook.Sheets("TableDef").Range("E8") = "Prompt"
    End If
End Sub

'---------------------------------------------------
Private Sub InsertOneBtnIntoBar(ByRef cmdBar As CommandBar, ByVal caption As String, ByVal action As String)
    Dim btn As CommandBarButton
    
    Set btn = cmdBar.Controls.Add
    
    btn.Style = msoButtonIconAndCaption
    With btn
        .BeginGroup = True
        .caption = caption
        .TooltipText = caption
        .OnAction = action
        .FaceId = 50
    End With
    cmdBar.Protection = msoBarNoCustomize
    cmdBar.Position = msoBarTop
    cmdBar.Visible = True
End Sub

Public Sub InsertUserToolBar()
    Dim cmbNewBar As CommandBar
    Dim ctlBtn As CommandBarButton
    Dim sEngText As String, sChsText As String, sExtFucText As String
    
    If iLanguageType = 1 Then
        sEngText = ThisWorkbook.Sheets("TableDef").Range("H11").Text
        sChsText = ThisWorkbook.Sheets("TableDef").Range("H12").Text
        If iHideSheetFlg = 0 Then
            sExtFucText = ThisWorkbook.Sheets("TableDef").Range("H13").Text
        Else
            sExtFucText = ThisWorkbook.Sheets("TableDef").Range("H14").Text
        End If
    Else
        sEngText = "English Version"
        sChsText = "Chinese Version"
        If iHideSheetFlg = 0 Then
            sExtFucText = "Show Extension Sheet"
        Else
            sExtFucText = "Hide Extension Sheet"
        End If
    End If
    
    On Error Resume Next
    Set cmbNewBar = CommandBars.Add(Name:="Operate Bar")
    'With cmbNewBar
        'Set ctlBtn = .Controls.Add
        'With ctlBtn
          '.Style = msoButtonIconAndCaption
          ''.Style = msoButtonCaption
          '.BeginGroup = True
          '.caption = sEngText
          '.TooltipText = sEngText
          '.OnAction = "SetEnglishUI"
          '.FaceId = 50
        'End With
        '.Protection = msoBarNoCustomize
        '.Position = msoBarTop
        '.Visible = True
    'End With
    'With cmbNewBar
        'Set ctlBtn = .Controls.Add
        'With ctlBtn
          '.Style = msoButtonIconAndCaption
          ''.Style = msoButtonCaption
          '.BeginGroup = True
          '.caption = sChsText
          '.TooltipText = sChsText
          '.OnAction = "SetChineseUI"
          '.FaceId = 50
        'End With
        '.Protection = msoBarNoCustomize
        '.Position = msoBarTop
        '.Visible = True
    'End With
    'With cmbNewBar
        'Set ctlBtn = .Controls.Add
        'With ctlBtn
          '.Style = msoButtonIconAndCaption
          ''.Style = msoButtonCaption
          '.BeginGroup = True
          '.caption = sExtFucText
          '.TooltipText = sExtFucText
          '.OnAction = "HideExtendFucSheet"
          '.FaceId = 50
        'End With
        '.Protection = msoBarNoCustomize
        '.Position = msoBarTop
        '.Visible = True
    'End With
    
    InsertOneBtnIntoBar cmbNewBar, CustomizeToolbarName, "CustomizeTemplateSub"

End Sub
Public Sub CustomizeTemplateSub()
    CustomizeTemplate.Show vbModeless
    CustomizeTemplate.initCheckBox iLanguageType
End Sub

Public Sub HideExtendFucSheet()
    ThisWorkbook.Sheets("Cover").Unprotect "HWCME"
    If iLanguageType = 0 Then
        If iHideSheetFlg = 1 Then
            CommandBars("Operate Bar").Controls.Item(3).caption = "Show Extension Sheet"
            CommandBars("Operate Bar").Controls.Item(3).TooltipText = "Show Extension Sheet"
            ThisWorkbook.Sheets("Cover").Range("C16:D16") = sEngCoverInfo1
            ThisWorkbook.Sheets("Cover").Range("C16:D16").RowHeight = 170
        Else
            CommandBars("Operate Bar").Controls.Item(3).caption = "Hide Extension Sheet"
            CommandBars("Operate Bar").Controls.Item(3).TooltipText = "Hide Extension Sheet"
            ThisWorkbook.Sheets("Cover").Range("C16:D16") = sEngCoverInfo1 + sEngCoverInfo2
            ThisWorkbook.Sheets("Cover").Range("C16:D16").RowHeight = 270
        End If
    Else
        If iHideSheetFlg = 1 Then
            CommandBars("Operate Bar").Controls.Item(3).caption = ThisWorkbook.Sheets("TableDef").Range("H13").Text
            CommandBars("Operate Bar").Controls.Item(3).TooltipText = ThisWorkbook.Sheets("TableDef").Range("H13").Text
            ThisWorkbook.Sheets("Cover").Range("C16:D16") = sChsCoverInfo1
            ThisWorkbook.Sheets("Cover").Range("C16:D16").RowHeight = 170
        Else
            CommandBars("Operate Bar").Controls.Item(3).caption = ThisWorkbook.Sheets("TableDef").Range("H14").Text
            CommandBars("Operate Bar").Controls.Item(3).TooltipText = ThisWorkbook.Sheets("TableDef").Range("H14").Text
            ThisWorkbook.Sheets("Cover").Range("C16:D16") = sChsCoverInfo1 + sChsCoverInfo2
            ThisWorkbook.Sheets("Cover").Range("C16:D16").RowHeight = 270
        End If
    End If
    ThisWorkbook.Sheets("Cover").Protect "HWCME", DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    If iHideSheetFlg = 1 Then
        ThisWorkbook.Sheets("BSCInfo").Visible = False
        ThisWorkbook.Sheets("DoubleFrequencyCell").Visible = False
        ThisWorkbook.Sheets("WholeNetworkCell").Visible = False
        'ThisWorkbook.Sheets("DeleteInterNCellRelation").Visible = False
        ThisWorkbook.Sheets("ConvertTemplate").Visible = False
        iHideSheetFlg = 0
    Else
        ThisWorkbook.Sheets("BSCInfo").Visible = True
        ThisWorkbook.Sheets("DoubleFrequencyCell").Visible = True
        ThisWorkbook.Sheets("WholeNetworkCell").Visible = True
        'ThisWorkbook.Sheets("DeleteInterNCellRelation").Visible = True
        ThisWorkbook.Sheets("ConvertTemplate").Visible = True
        iHideSheetFlg = 1
    End If
    ThisWorkbook.Sheets("TableDef").Range("G11") = CStr(iHideSheetFlg)
End Sub

Public Sub DeleteUserToolBar()
    CommandBars("Operate Bar").Delete
End Sub


