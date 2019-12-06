VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NoneLteTemplateForm 
   Caption         =   "Add/Delete User-defined Template"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   OleObjectBlob   =   "NoneLteTemplateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NoneLteTemplateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'以下用于添加和删除「MappingSiteTemplate」,「MappingCellTemplate」和「MappingRadioTemplate」页的模板
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                         '
'1.添加和删除「MappingSiteTemplate」页的模板。
'「*Site Type」「Cabinet Type」「FDD/TDD Mode」「*Site Patten」列侯选值的窗体    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'激活事件,默认首先显示Site标签页
Private Sub UserForm_Activate()

    Dim func As ToolBarFunction
    Set func = New ToolBarFunction
    If func.cellAddSupport = True Then
        MultiPage.Pages.Item(0).Visible = True
        setCellTypePattern
    End If
End Sub



'取消此次操作
Private Sub CancelSiteButton_Click()
    Unload Me
End Sub

'「DeleteSite」按钮事件,删除模板
Private Sub DeleteSite()
    Dim TemplateName As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim neType As String
    neType = getNeType()
    
    TemplateName = Trim(SitePatternList.text)
    
    '当前数据行数
    rowscount = MappingSiteTemplate.Range("a1048576").End(xlUp).row
    
    '存在flag
    existflg = 0
      
    '「Site Pattern」为空时报错
    If TemplateName = "" Then
        MsgBox TemplateName & getResByKey("sitePatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 2
        Exit Sub
    End If
    
    '遍历已有默认侯选值，存在与输入值重复项时删除。
    For m_rowNum = 2 To rowscount
        If MappingSiteTemplate.Cells(m_rowNum, 4) = TemplateName _
                And MappingSiteTemplate.Cells(m_rowNum, 1) = SiteType.value _
                And MappingSiteTemplate.Cells(m_rowNum, 5) = neType Then
            MappingSiteTemplate.rows(m_rowNum).Delete
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
        MsgBox TemplateName & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    '删除成功时将焦点移动到「Base Station Transport Data」页的第一条记录
    If existflg = 1 Then
        SitePatternList.text = ""
        For iIndex = 0 To SitePatternList.ListCount - 1
                    If (SitePatternList.List(iIndex, 0) = TemplateName) Then
                        SitePatternList.RemoveItem (iIndex)
                        Exit For
                    End If
        Next
        

        'SitePatternList.SetFocus
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    End If
End Sub

'「AddSite」按钮事件,添加模板
Private Sub AddSite()
    Dim TemplateName As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    Dim neType As String
    
    neType = getNeType()

    '用户输入值
    TemplateName = Trim(SitePattern.text)
    
    '当前数据行数
    rowscount = MappingSiteTemplate.Range("a1048576").End(xlUp).row
    
    '重复存在flag
    existflg = 0
      
    '「Site Pattern」为空时报错
    If TemplateName = "" Then
        MsgBox TemplateName & getResByKey("sitePatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 1
        Exit Sub
    End If
    
    '遍历已有默认侯选值，存在与输入值重复项时报错。
    For m_rowNum = 2 To rowscount
        If MappingSiteTemplate.Cells(m_rowNum, 4) = TemplateName _
                And MappingSiteTemplate.Cells(m_rowNum, 1) = SiteType.value _
                And MappingSiteTemplate.Cells(m_rowNum, 5) = neType Then
            MsgBox TemplateName & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            SitePattern.SetFocus
            existflg = 1
            Exit Sub
        End If
    Next
        
    '查找当前组的最后一行
    'lastLineofGroup = rowscount
    'For n_RowNum = 2 To rowscount
    '    If (MappingSiteTemplate.Cells(n_RowNum, 1) = SiteType.value And MappingSiteTemplate.Cells(n_RowNum, 5) = neType) _
    '        Or MappingSiteTemplate.Cells(n_RowNum, 1) = "" Then
    '        lastLineofGroup = n_RowNum
    '     End If
    'Next
    'lastLineofGroup = lastLineofGroup + 1
    lastLineofGroup = rowscount + 1
    Worksheets("MappingSiteTemplate").rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    '不重复时将输入值追加到候选值列表，并将焦点移动到「Base Station Transport Data」页的第一条记录
    If existflg = 0 Then
        MappingSiteTemplate.Cells(lastLineofGroup, 1).value = SiteType.value
        MappingSiteTemplate.Cells(lastLineofGroup, 2).value = ""
        MappingSiteTemplate.Cells(lastLineofGroup, 3).value = ""
        MappingSiteTemplate.Cells(lastLineofGroup, 4).value = TemplateName
        MappingSiteTemplate.Cells(lastLineofGroup, 5).value = neType
        'Worksheets("Base Station Transport Data").Activate
        'Worksheets("Base Station Transport Data").range("A3").Select
        
        SitePattern.value = ""
        SitePattern.SetFocus
        Load Me
        MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

'「Site Type」选择事件
Private Sub SiteType_Change()
    Set_Template_Related
End Sub




'从「MappingSiteTemplate」页获取「*Site Patten」列侯选值
Private Sub Set_Template_Related()
    Dim m_rowNum As Long
    Dim neType As String
    neType = getNeType()
    '清除「Cabinet Type」旧的侯选值
    SitePatternList.Clear
    For m_rowNum = 2 To MappingSiteTemplate.Range("a1048576").End(xlUp).row
        If SiteType.text = MappingSiteTemplate.Cells(m_rowNum, 1).value _
        And neType = MappingSiteTemplate.Cells(m_rowNum, 5).value Then
            SitePatternList.AddItem (MappingSiteTemplate.Cells(m_rowNum, 4).value)
        End If
    Next
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'添加和删除「MappingCellTemplate」页的模板                             '
'「*Band Width」「FddTddIdd」「FDD/TDD Mode」「*Cell Mode」列侯选值的窗体    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub AddCellTemplate_Click()
    Me.CellPatternText.Visible = True
    Me.CellPattern.Visible = False
    Me.AddCellButton.Caption = getResByKey("Add")
End Sub
Private Sub DeleteCellTemplate_Click()
    Me.CellPattern.Visible = True
    Me.CellPatternText.Visible = False
    Me.AddCellButton.Caption = getResByKey("Delete")
    SetCellPattern
End Sub
'提交进行Add/Delete操作
Private Sub AddCellButton_Click()
    If Me.AddCellTemplate.value = True Then
        AddCell
    Else
        DeleteCell
    End If
    Call refreshCell
End Sub
'取消此次操作
Private Sub CancelCellButton_Click()
    Unload Me
End Sub

'「Delete」按钮事件,删除模板
Private Sub DeleteCell()
    Dim TemplateName As String
    Dim CellType, neType As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    
    neType = "GSM"
    '用户输入值
    TemplateName = Trim(CellPattern.text)
    CellType = Me.CellType.value
    '当前数据行数
    rowscount = MappingCellTemplate.Range("a1048576").End(xlUp).row
    
    '存在flag
    existflg = 0
      
    '「Cell Pattern」为空时报错
    If TemplateName = "" Then
        MsgBox TemplateName & getResByKey("cellPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 2
        Exit Sub
    End If
    
    '遍历已有默认侯选值，找到相应项时删除。
    If Not isLteCellType Then
        For m_rowNum = 2 To rowscount
            If MappingCellTemplate.Cells(m_rowNum, 1) = TemplateName _
            And MappingCellTemplate.Cells(m_rowNum, 2) = CellType _
            And MappingCellTemplate.Cells(m_rowNum, 3) = neType Then
                MappingCellTemplate.rows(m_rowNum).Delete
                Call SetCellPattern
                existflg = 1
            End If
        Next
    Else
        For m_rowNum = 2 To rowscount
            If MappingCellTemplate.Cells(m_rowNum, 1) = TemplateName _
                And MappingCellTemplate.Cells(m_rowNum, 2) = CellType _
                And MappingCellTemplate.Cells(m_rowNum, 3) = neType _
                And (MappingCellTemplate.Cells(m_rowNum, 7) = SA.value) _
                And MappingCellTemplate.Cells(m_rowNum, 6) = FddTddIdd.value _
                And MappingCellTemplate.Cells(m_rowNum, 5) = TxRxMode.value _
                And MappingCellTemplate.Cells(m_rowNum, 4) = BandWidth.value Then
                MappingCellTemplate.rows(m_rowNum).Delete
                SetCellPattern
                existflg = 1
            End If
        Next
    End If
    
    If existflg = 0 Then
        MsgBox TemplateName & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    If existflg = 1 Then
        'CellPattern.text = ""
'        For iIndex = 0 To CellPattern.ListCount - 1
'            If (CellPattern.List(iIndex, 0) = templatename) Then
'                CellPattern.RemoveItem (iIndex)
'                Exit For
'            End If
'        Next
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    End If
End Sub

'「Add」按钮事件,添加模板
Private Sub AddCell()
    Dim TemplateName As String
    Dim CellType, neType As String
    
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    Dim flag4, flag3, flag2, flag1 As Boolean
    
    flag4 = False
    flag3 = False
    flag2 = False
    falg1 = False
    
    neType = "GSM"
    
    '用户输入值
    TemplateName = Trim(CellPatternText.text)
    
    CellType = Me.CellType.value
    '当前数据行数
    rowscount = MappingCellTemplate.Range("a1048576").End(xlUp).row
    
    '重复存在flag
    existflg = 0
      
    '「Cell Pattern」为空时报错
    If TemplateName = "" Then
        MsgBox TemplateName & getResByKey("cellPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 1
        Exit Sub
    End If
    
    If Not isLteCellType Then
    '遍历已有默认侯选值，存在与输入值重复项时报错。
        For m_rowNum = rowscount To 2 Step -1
            If MappingCellTemplate.Cells(m_rowNum, 1) = TemplateName _
            And MappingCellTemplate.Cells(m_rowNum, 2) = CellType _
            And MappingCellTemplate.Cells(m_rowNum, 3) = neType Then
                MsgBox TemplateName & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
                CellPatternText.SetFocus
                existflg = 1
                Exit Sub
            End If
        Next
    Else
        For m_rowNum = rowscount To 2 Step -1
            If MappingCellTemplate.Cells(m_rowNum, 1) = TemplateName _
                And MappingCellTemplate.Cells(m_rowNum, 2) = CellType _
                And MappingCellTemplate.Cells(m_rowNum, 3) = neType _
                And (MappingCellTemplate.Cells(m_rowNum, 7) = SA.value) _
                And MappingCellTemplate.Cells(m_rowNum, 6) = FddTddIdd.value _
                And MappingCellTemplate.Cells(m_rowNum, 5) = TxRxMode.value _
                And MappingCellTemplate.Cells(m_rowNum, 4) = BandWidth.value Then
                MsgBox TemplateName & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
                CellPatternText.SetFocus
                existflg = 1
                Exit For
            End If
    Next
    End If
            
    '不重复时将输入值追加到候选值列表，并将焦点移动到「Base Station Transport Data」页的第一条记录
    If existflg = 0 Then
        '查找当前组的最后一行
        lastLineofGroup = rowscount + 1
        Worksheets("MappingCellTemplate").rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        MappingCellTemplate.rows(lastLineofGroup).NumberFormatLocal = "@" '设置单元格格式为文本
        MappingCellTemplate.Cells(lastLineofGroup, 1).value = TemplateName
        MappingCellTemplate.Cells(lastLineofGroup, 2).value = CellType
        MappingCellTemplate.Cells(lastLineofGroup, 3).value = neType
        If isLteCellType Then
            MappingCellTemplate.Cells(lastLineofGroup, 4).value = BandWidth.value
            MappingCellTemplate.Cells(lastLineofGroup, 5).value = TxRxMode.value
            MappingCellTemplate.Cells(lastLineofGroup, 6).value = FddTddIdd.value
            MappingCellTemplate.Cells(lastLineofGroup, 7).value = SA.value
        End If
        Me.CellPatternText.value = ""
        Me.CellPatternText.SetFocus
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
        Load Me
        MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

Private Function containsCellSheet(ByRef neType As String) As Boolean
    Dim cellSheetNameEn As String, cellSheetNameCn As String
    cellSheetNameEn = neType & " Cell"
    cellSheetNameCn = neType & getResByKey("A159")
    If containsASheet(ThisWorkbook, cellSheetNameEn) Or containsASheet(ThisWorkbook, cellSheetNameCn) Then
        containsCellSheet = True
    Else
        containsCellSheet = False
    End If
End Function

Private Sub setCellTypePattern()
        Dim neType As String
        neType = "GSM"
        CellType.Clear
        If neType = "MRAT" Then
            With CellType
                If isContainBaseStation() Then
                    If containsCellSheet("GSM") Then .AddItem getResByKey("GSM Local Cell")
                    If containsCellSheet("UMTS") Then .AddItem getResByKey("UMTS Local Cell")
                    If containsCellSheet("LTE") Then .AddItem getResByKey("LTE Cell")
                End If
                If isContainGsmControl() Then
                    .AddItem getResByKey("GSM Logic Cell")
                End If
                If isContainUmtsControl() Then
                    .AddItem getResByKey("UMTS Logic Cell")
                End If
            End With
        ElseIf neType = "UMTS" Then
            With CellType
                        If isContainBaseStation() Then
                            .AddItem getResByKey("UMTS Local Cell")
                         End If
                         If isContainUmtsControl() Then
                            .AddItem getResByKey("UMTS Logic Cell")
                         End If
            End With
        ElseIf neType = "GSM" Then
            With CellType
                         If isContainBaseStation() Then
                            .AddItem "GSM"
                        End If
                        If isContainGsmControl() Then
                            .AddItem "GSM"
                        End If
            End With
        ElseIf neType = "LTE" Then
            With CellType
                        .AddItem getResByKey("LTE Cell")
            End With
        End If
        If Me.CellType.ListCount Then
            CellType.ListIndex = 0
        End If
End Sub
Private Sub BandWidth_Change()
    SetCellPattern
End Sub

Private Sub TxRxMode_Change()
    SetCellPattern
End Sub

Private Sub SA_Change()
    SetCellPattern
End Sub

Private Function isLteCellType() As Boolean
    isLteCellType = False
End Function

Private Sub initLteCellFilterPatermeters()
    If isLteCellType Then
        Call initLteCellFilterParameters(True)
        Call adjustCellTemplateForL
    Else
        Call initLteCellFilterParameters(False)
        Call adjustCellTemplateForGU
    End If
End Sub

Private Sub adjustCellTemplateForGU()
    Me.CellPatternLabel.Top = 92
    Me.CellPattern.Top = 90
    Me.CellPatternText.Top = 90
    
    Me.AddCellButton.Top = 138
    Me.CancelCellButton.Top = 138
End Sub

Private Sub adjustCellTemplateForL()
    Me.CellPatternLabel.Top = 165
    Me.CellPattern.Top = 162
    Me.CellPatternText.Top = 162
    
    Me.AddCellButton.Top = 186
    Me.CancelCellButton.Top = 186
End Sub

'从「MappingCellTemplate」页获取「*Cell Pattern」列侯选值
Private Sub SetCellPattern()
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    Dim CellType, neType As String
    
    neType = "GSM"
    CellType = "GSM"
    
    flag = True
    
    '清除「Cell Pattern」旧的侯选值
    Me.CellPattern.Clear
    If Not isLteCellType Then
        For m_rowNum = 2 To MappingCellTemplate.Range("a1048576").End(xlUp).row
            If CellType = MappingCellTemplate.Cells(m_rowNum, 2).value _
            And neType = MappingCellTemplate.Cells(m_rowNum, 3).value Then
                Me.CellPattern.AddItem (MappingCellTemplate.Cells(m_rowNum, 1).value)
                 If flag = True Then
                     CellPattern.value = MappingCellTemplate.Cells(m_rowNum, 1).value
                     flag = False
                 End If
            End If
        Next
    Else
       
    End If
End Sub

Private Sub initLteCellFilterParameters(ByRef flag As Boolean)
    'Call grayLteCellFilterParameters(flag)
    Call displayLteCellFilterParameters(flag)
End Sub

Private Sub displayLteCellFilterParameters(ByRef flag As Boolean)
    FddTddIddLabel.Visible = flag
    FddTddIdd.Visible = flag
    
    BandWidthLabel.Visible = flag
    BandWidth.Visible = flag
    
    TxRxModeLabel.Visible = flag
    TxRxMode.Visible = flag
    
    SALabel.Visible = flag
    SA.Visible = flag
End Sub

Private Sub grayLteCellFilterParameters(ByRef flag As Boolean)
    FddTddIddLabel.Enabled = flag
    FddTddIdd.Enabled = flag
    
    BandWidthLabel.Enabled = flag
    BandWidth.Enabled = flag
    
    TxRxModeLabel.Enabled = flag
    TxRxMode.Enabled = flag
    
    SALabel.Enabled = flag
    SA.Enabled = flag
End Sub

Private Sub clearLteCellFilterParameters()

End Sub

Private Sub CancelRadioButton_Click()
    Unload Me
End Sub

'「Delete」按钮事件,删除模板
Private Sub DeleteRadio()
    Dim TemplateName As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim radioTp, neType As String
    
    neType = getNeType()
    
    '用户输入值
    TemplateName = Trim(RadioPattern.text)
    radioTp = RadioType.value
    '当前数据行数
    rowscount = MappingRadioTemplate.Range("a1048576").End(xlUp).row
    
    '存在flag
    existflg = 0
      
    '「Radio Pattern」为空时报错
    If TemplateName = "" Then
        MsgBox TemplateName & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 2
        Exit Sub
    End If
    
    '遍历已有默认侯选值，找到删除项，进行删除。
    For m_rowNum = 2 To rowscount
        If MappingRadioTemplate.Cells(m_rowNum, 1) = TemplateName _
        And MappingRadioTemplate.Cells(m_rowNum, 2) = radioTp _
        And MappingRadioTemplate.Cells(m_rowNum, 3) = neType Then
            MappingRadioTemplate.rows(m_rowNum).Delete
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
        MsgBox TemplateName & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    '删除成功时将焦点移动到「Base Station Transport Data」页的第一条记录
    If existflg = 1 Then
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
        
        RadioPattern.text = ""
        For iIndex = 0 To RadioPattern.ListCount - 1
                    If (RadioPattern.List(iIndex, 0) = TemplateName) Then
                        RadioPattern.RemoveItem (iIndex)
                        Exit For
                    End If
        Next
        RadioPattern.SetFocus
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    End If
End Sub

'「Add」按钮事件,添加模板
Private Sub AddRadio()
    Dim TemplateName As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    Dim flag As Boolean
    Dim radioTp, neType As String
    radioTp = RadioType.value
    flag = False
    
    neType = getNeType()
    '用户输入值
    TemplateName = Trim(RadioPatternText.text)
    
    '当前数据行数
    rowscount = MappingRadioTemplate.Range("a1048576").End(xlUp).row
    
    '重复存在flag
    existflg = 0
      
    '「Radio Pattern」为空时报错
    If TemplateName = "" Then
        MsgBox TemplateName & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 1
        Exit Sub
    End If
    
    '遍历已有默认侯选值，若存在与输入值重复项时报错。
    For m_rowNum = 2 To rowscount
        If MappingRadioTemplate.Cells(m_rowNum, 1) = TemplateName _
           And MappingRadioTemplate.Cells(m_rowNum, 2) = radioTp _
           And MappingRadioTemplate.Cells(m_rowNum, 3) = neType Then
            MsgBox TemplateName & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            RadioPatternText.SetFocus
            existflg = 1
            Exit Sub
        End If
    Next

    'lastLineofGroup = rowscount
    'For n_RowNum = 2 To rowscount
    '    If (MappingRadioTemplate.Cells(n_RowNum, 3) = neType And MappingRadioTemplate.Cells(n_RowNum, 2) = radioTp) _
    '    Or MappingRadioTemplate.Cells(n_RowNum, 1) = "" Then
    '        lastLineofGroup = n_RowNum
    '    End If
    ' Next
    
    '查找当前组的最后一行
   ' lastLineofGroup = lastLineofGroup + 1
   lastLineofGroup = rowscount + 1
    Worksheets("MappingRadioTemplate").rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    '不重复时将输入值追加到候选值列表，并将焦点移动到「Base Station Transport Data」页的第一条记录
    If existflg = 0 Then
        MappingRadioTemplate.Cells(lastLineofGroup, 1).value = TemplateName
        MappingRadioTemplate.Cells(lastLineofGroup, 2).value = radioTp
        MappingRadioTemplate.Cells(lastLineofGroup, 3).value = neType
        RadioPatternText.value = ""
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
        Load Me
        MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub



Private Sub refreshCell()
    Dim rangeHis As Range
    Dim row, columen As Long
    Set rangeHis = Selection
    ActiveSheet.Cells(Selection.row + 1, Selection.column).Select
    rangeHis.Select
End Sub






