VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LteTemplateForm 
   Caption         =   "Add/Delete User-defined Template"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   OleObjectBlob   =   "LteTemplateForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "LteTemplateForm"
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
    MultiPage.Font.Size = 10
    Me.Caption = getResByKey("Bar_Template")
'    If MultiPage.value = 0 Then
'        'Me.SitePattern.SetFocus
'        'SetSiteType
'    Else
    MultiPage.value = 1
    SetFddTddIdd
'    End If
    Dim func As ToolBarFunction
    Set func = New ToolBarFunction
    If func.siteAddSupport = False Then
        MultiPage.Pages.Item(0).Visible = False
    End If
    If func.cellAddSupport = False Then
        MultiPage.Pages.Item(1).Visible = False
    End If
    If func.radioAddSupport = False Then
        MultiPage.Pages.Item(2).Visible = False
    End If
End Sub
'选择事件,选择不同的标签(包括Site,Cell和Radio)
Private Sub MultiPage_Change()
 If MultiPage.value = 0 Then
    'SetSiteType
 ElseIf MultiPage.value = 1 Then
    SetFddTddIdd
 ElseIf MultiPage.value = 2 Then
    'SetFTMode
    'SetRSA
 End If
End Sub
'选择事件,选择Add 选项
Private Sub AddSiteTemplate_Click()
    Me.SitePattern.Visible = True
    Me.SitePatternList.Visible = False
    Me.AddSiteButton.Caption = getResByKey("Add")
End Sub
'选择事件,选择Delete选项
Private Sub DeleteSiteTemplate_Click()
    Me.SitePattern.Visible = False
    Me.SitePatternList.Visible = True
    Me.AddSiteButton.Caption = getResByKey("Delete")
    Set_Template_Related
End Sub
'提交进行Add/Delete操作
Private Sub AddSiteButton_Click()
    If Me.AddSiteTemplate.value = True Then
        AddSite
    Else
        DeleteSite
    End If
    Call refreshCell
End Sub
'取消此次操作
Private Sub CancelSiteButton_Click()
    Unload Me
End Sub

'「DeleteSite」按钮事件,删除模板
Private Sub DeleteSite()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    
    templatename = SitePatternList.text
    
    '当前数据行数
    rowscount = MappingSiteTemplate.Range("a1048576").End(xlUp).row
    
    '存在flag
    existflg = 0
      
    '「Site Pattern」为空时报错
    If templatename = "" Then
       MsgBox templatename & getResByKey("sitePatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 2
        Exit Sub
    End If
    
    '遍历已有默认侯选值，存在与输入值重复项时删除。
    For m_rowNum = 2 To rowscount
        If MappingSiteTemplate.Cells(m_rowNum, 4) = templatename _
                And MappingSiteTemplate.Cells(m_rowNum, 3) = FDDTDDMode.value _
                And MappingSiteTemplate.Cells(m_rowNum, 2) = CabinetType.value _
                And MappingSiteTemplate.Cells(m_rowNum, 1) = SiteType.value Then
            MappingSiteTemplate.Rows(m_rowNum).Delete
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
        MsgBox templatename & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    If existflg = 1 Then
        SitePatternList.text = ""
        For iIndex = 0 To SitePatternList.ListCount - 1
                    If (SitePatternList.List(iIndex, 0) = templatename) Then
                        SitePatternList.RemoveItem (iIndex)
                        Exit For
                    End If
        Next
        Load Me
                MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    End If
    
End Sub

'「AddSite」按钮事件,添加模板
Private Sub AddSite()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    
    '用户输入值
    templatename = SitePattern.text
    
    '当前数据行数
    rowscount = MappingSiteTemplate.Range("a1048576").End(xlUp).row
    
    '重复存在flag
    existflg = 0
      
    '「Site Pattern」为空时报错
    If templatename = "" Then
        MsgBox templatename & getResByKey("sitePatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 1
        Exit Sub
    End If
    
    '遍历已有默认侯选值，存在与输入值重复项时报错。
    For m_rowNum = 2 To rowscount
        If MappingSiteTemplate.Cells(m_rowNum, 4) = templatename _
                And MappingSiteTemplate.Cells(m_rowNum, 3) = FDDTDDMode.value _
                And MappingSiteTemplate.Cells(m_rowNum, 2) = CabinetType.value _
                And MappingSiteTemplate.Cells(m_rowNum, 1) = SiteType.value Then
            MsgBox templatename & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            SitePattern.SetFocus
            existflg = 1
            Exit For
        End If
    Next
        
    '查找当前组的最后一行
    lastLineofGroup = rowscount + 1
    Worksheets("MappingSiteTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    '不重复时将输入值追加到候选值列表，并将焦点移动到「eNodeB Transport Data」页的第一条记录
    If existflg = 0 Then
        MappingSiteTemplate.Cells(lastLineofGroup, 1).value = SiteType.value
        MappingSiteTemplate.Cells(lastLineofGroup, 2).value = CabinetType.value
        MappingSiteTemplate.Cells(lastLineofGroup, 3).value = FDDTDDMode.value
        MappingSiteTemplate.Cells(lastLineofGroup, 4).value = templatename
        Load Me
         MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

'「Site Type」选择事件
Private Sub SiteType_Change()
    SetCabinetType (Me.SiteType.text)
    
    'DBS3900_LTE支持FDD/TDD/FDDTDD三种，其它站型只支持FDD
    Me.FDDTDDMode.Clear
    If Me.SiteType.text = "DBS3900_LTE" Then
        Me.FDDTDDMode.AddItem ("FDD")
        Me.FDDTDDMode.AddItem ("TDD")
        Me.FDDTDDMode.AddItem ("FDDTDD")
    Else
        Me.FDDTDDMode.AddItem ("FDD")
    End If
    
    FDDTDDMode.value = "FDD"
    Set_Template_Related
    
End Sub

'「Cabinet Type」选择事件
Private Sub CabinetType_Change()
    Set_Template_Related
End Sub

'「FDDTDD Mode」选择事件
Private Sub FDDTDDMode_Change()
    Set_Template_Related
End Sub

'从「MappingSiteTypeCabinetType」页获取「*Site Type」列侯选值
Private Sub SetSiteType()
    Dim m_rowNum As Long
    Dim m_RowNum_Inner As Long
    Dim m_Str As String
    Dim flag As Boolean
    
    Me.SiteType.Clear
    Me.SiteType.AddItem (MappingSiteTypeCabinetType.Cells(2, 1).value)
    SiteType.value = MappingSiteTypeCabinetType.Cells(2, 1).value
    For m_rowNum = 3 To MappingSiteTypeCabinetType.Range("a1048576").End(xlUp).row
        For m_RowNum_Inner = 2 To m_rowNum - 1
            If MappingSiteTypeCabinetType.Cells(m_rowNum, 1) <> MappingSiteTypeCabinetType.Cells(m_RowNum_Inner, 1) Then
                flag = False
            Else
                flag = True
                Exit For
            End If
        Next
        If flag = False Then
            Me.SiteType.AddItem (MappingSiteTypeCabinetType.Cells(m_rowNum, 1).value)
        End If
    Next
    
End Sub

'从「MappingSiteTypeCabinetType」页获取「Cabinet Type」列侯选值
Private Sub SetCabinetType(SiteType As String)
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    flag = True
    
    '清除「Cabinet Type」旧的侯选值
    Me.CabinetType.Clear
    
    For m_rowNum = 2 To MappingSiteTypeCabinetType.Range("a1048576").End(xlUp).row
        If SiteType = MappingSiteTypeCabinetType.Cells(m_rowNum, 1).value Then
            Me.CabinetType.AddItem (MappingSiteTypeCabinetType.Cells(m_rowNum, 2).value)
            If flag = True Then
                CabinetType.value = MappingSiteTypeCabinetType.Cells(m_rowNum, 2).value
                flag = False
            End If
        End If
    Next
    
End Sub
'从「MappingSiteTemplate」页获取「*Site Patten」列侯选值
Private Sub Set_Template_Related()
    Dim m_rowNum As Long
    
    '清除「Cabinet Type」旧的侯选值
    SitePatternList.Clear
    For m_rowNum = 2 To MappingSiteTemplate.Range("a1048576").End(xlUp).row
        If SiteType.text = MappingSiteTemplate.Cells(m_rowNum, 1).value _
            And CabinetType.text = MappingSiteTemplate.Cells(m_rowNum, 2).value _
            And FDDTDDMode.text = MappingSiteTemplate.Cells(m_rowNum, 3).value Then
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
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
        
        
    '用户输入值
    templatename = CellPattern.text
    
    '当前数据行数
    rowscount = MappingCellTemplate.Range("a1048576").End(xlUp).row
    
    '存在flag
    existflg = 0
      
    '「Cell Pattern」为空时报错
    If templatename = "" Then
        MsgBox templatename & getResByKey("cellPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 2
        Exit Sub
    End If
    
    '遍历已有默认侯选值，找到相应项时删除。
    For m_rowNum = 2 To rowscount
        If MappingCellTemplate.Cells(m_rowNum, 5) = templatename _
                And (MappingCellTemplate.Cells(m_rowNum, 4) = SA.value Or Me.SA.Enabled = False) _
                And MappingCellTemplate.Cells(m_rowNum, 3) = FddTddIdd.value _
                And MappingCellTemplate.Cells(m_rowNum, 2) = TxRxMode.value _
                And MappingCellTemplate.Cells(m_rowNum, 1) = BandWidth.value Then
            MappingCellTemplate.Rows(m_rowNum).Delete
            SetCellPattern
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
         MsgBox templatename & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    If existflg = 1 Then
        'CellPattern.text = ""
        For iIndex = 0 To CellPattern.ListCount - 1
                    If (CellPattern.List(iIndex, 0) = templatename) Then
                        CellPattern.RemoveItem (iIndex)
                        Exit For
                    End If
        Next
        Load Me
                MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    End If
End Sub

'「Add」按钮事件,添加模板
Private Sub AddCell()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    Dim flag4, flag3, flag2, flag1 As Boolean
    
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    flag4 = False
    flag3 = False
    flag2 = False
    falg1 = False
    
    '用户输入值
    templatename = CellPatternText.text
    
    '当前数据行数
    rowscount = MappingCellTemplate.Range("E1048576").End(xlUp).row
    
    '重复存在flag
    existflg = 0
      
    '「Cell Pattern」为空时报错
    If templatename = "" Then
        MsgBox templatename & getResByKey("cellPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 1
        Exit Sub
    End If
    
    '遍历已有默认侯选值，存在与输入值重复项时报错。
    For m_rowNum = rowscount To 2 Step -1
        If MappingCellTemplate.Cells(m_rowNum, 5) = templatename _
                And (MappingCellTemplate.Cells(m_rowNum, 4) = SA.value Or Me.SA.Enabled = False) _
                And MappingCellTemplate.Cells(m_rowNum, 3) = FddTddIdd.value _
                And MappingCellTemplate.Cells(m_rowNum, 2) = TxRxMode.value _
                And MappingCellTemplate.Cells(m_rowNum, 1) = BandWidth.value Then
            MsgBox templatename & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            CellPatternText.SetFocus
            existflg = 1
            Exit For
        End If
    Next
        
    '查找当前组的最后一行
    lastLineofGroup = rowscount + 1
    Worksheets("MappingCellTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
    '不重复时将输入值追加到候选值列表，并将焦点移动到「Base Station Transport Data」页的第一条记录
    If existflg = 0 Then
        MappingCellTemplate.Cells(lastLineofGroup, 1).value = BandWidth.value
        MappingCellTemplate.Cells(lastLineofGroup, 2).value = TxRxMode.value
        MappingCellTemplate.Cells(lastLineofGroup, 3).value = FddTddIdd.value
        MappingCellTemplate.Cells(lastLineofGroup, 4).value = SA.value
        MappingCellTemplate.Cells(lastLineofGroup, 5).value = templatename
        Me.CellPatternText.value = ""
        Me.CellPatternText.SetFocus
        Load Me
        MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

Private Sub FddTddIdd_Change()
    SetBandWidth
    SetTxRxMode
    SetSA
    SetCellPattern
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
'设置「*FddTddIdd」列侯选值
Private Sub SetFddTddIdd()

    Me.FddTddIdd.Clear
    Me.FddTddIdd.AddItem ("FDD")
    Me.FddTddIdd.AddItem ("TDD")
    Me.FddTddIdd.text = "TDD"
    
End Sub

'设置「*BandWidth」列侯选值
Private Sub SetBandWidth()
    
    Me.BandWidth.Clear
    
    If Me.FddTddIdd.text = "FDD" Then
        Me.BandWidth.AddItem ("1.4M")
        Me.BandWidth.AddItem ("3M")
        Me.BandWidth.AddItem ("5M")
        Me.BandWidth.AddItem ("10M")
        Me.BandWidth.AddItem ("15M")
        Me.BandWidth.AddItem ("20M")
        Me.BandWidth.text = ("1.4M")
    ElseIf Me.FddTddIdd.text = "TDD" Then
        Me.BandWidth.AddItem ("5M")
        Me.BandWidth.AddItem ("10M")
        Me.BandWidth.AddItem ("15M")
        Me.BandWidth.AddItem ("20M")
        Me.BandWidth.text = "10M"
    End If

End Sub
'设置「*TxRxMode」列侯选值
Private Sub SetTxRxMode()
    
    Me.TxRxMode.Clear
    
    If Me.FddTddIdd.text = "FDD" Then
        Me.TxRxMode.AddItem ("1T1R")
        Me.TxRxMode.AddItem ("1T2R")
        Me.TxRxMode.AddItem ("2T2R")
        Me.TxRxMode.AddItem ("2T4R")
        Me.TxRxMode.AddItem ("4T4R")
        Me.TxRxMode.text = "1T1R"
    ElseIf Me.FddTddIdd.text = "TDD" Then
        Me.TxRxMode.AddItem ("1T1R")
        Me.TxRxMode.AddItem ("2T2R")
        Me.TxRxMode.AddItem ("4T4R")
        Me.TxRxMode.AddItem ("8T8R")
        Me.TxRxMode.text = "1T1R"
    End If
    
End Sub

'设置「*SA」列侯选值
Private Sub SetSA()
    
    Me.SA.Clear
    If Me.FddTddIdd.text = "TDD" Then
        Me.SA.Enabled = True
        Me.SA.AddItem ("SA0")
        Me.SA.AddItem ("SA1")
        Me.SA.AddItem ("SA2")
        Me.SA.AddItem ("SA3")
        Me.SA.AddItem ("SA4")
        Me.SA.AddItem ("SA5")
        Me.SA.AddItem ("SA6")
        Me.SA.text = "SA0"
    Else
        'Me.SA.text = " "
        Me.SA.Enabled = False
    End If
    
End Sub
'从「MappingCellTemplate」页获取「*Cell Pattern」列侯选值
Private Sub SetCellPattern()
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    flag = True
    
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    FddTddIdd = Me.FddTddIdd.text
    BandWidth = Me.BandWidth.text
    TxRxMode = Me.TxRxMode.text
    SA = Me.SA.text
    
    '清除「Cell Pattern」旧的侯选值
    Me.CellPattern.Clear
    
    For m_rowNum = 2 To MappingCellTemplate.Range("a1048576").End(xlUp).row
         If BandWidth = MappingCellTemplate.Cells(m_rowNum, 1).value And FddTddIdd = MappingCellTemplate.Cells(m_rowNum, 3).value And TxRxMode = MappingCellTemplate.Cells(m_rowNum, 2).value Then
            If (Me.SA.Enabled = False) Or (SA = MappingCellTemplate.Cells(m_rowNum, 4).value) Then
                Me.CellPattern.AddItem (MappingCellTemplate.Cells(m_rowNum, 5).value)
                If flag = True Then
                    CellPattern.value = MappingCellTemplate.Cells(m_rowNum, 5).value
                    flag = False
                End If
             End If
        End If
    Next
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'添加和删除「MappingRadioTemplate」页的模板                             '
'「FTMode」「RSA」「Radio Pattern」列侯选值的窗体    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRadioTemplate_Click()
    Me.RadioPatternText.Visible = True
    Me.RadioPattern.Visible = False
    Me.AddRadioButton.Caption = "Add"
End Sub
Private Sub DeleteRadioTemplate_Click()
    Me.RadioPattern.Visible = True
    Me.RadioPatternText.Visible = False
    Me.AddRadioButton.Caption = "Delete"
    SetRadioPattern
End Sub
'提交进行Add/Delete操作
Private Sub AddRadioButton_Click()
    If Me.AddRadioTemplate.value = True Then
        AddRadio
    Else
        DeleteRadio
    End If
End Sub
Private Sub CancelRadioButton_Click()
    Unload Me
End Sub
Private Sub SetFTMode()
    
    Me.FTMode.Clear
    Me.FTMode.AddItem ("FDD")
    Me.FTMode.AddItem ("TDD")
    Me.FTMode.AddItem ("FDDTDD")
    Me.FTMode.text = "TDD"

End Sub
Private Sub SetRSA()
    
    Me.RSA.Clear
    If Me.FTMode.text = "TDD" Then
        Me.RSA.Enabled = True
        Me.RSA.AddItem ("SA0")
        Me.RSA.AddItem ("SA1")
        Me.RSA.AddItem ("SA2")
        Me.RSA.AddItem ("SA3")
        Me.RSA.AddItem ("SA4")
        Me.RSA.AddItem ("SA5")
        Me.RSA.AddItem ("SA6")
        Me.RSA.text = "SA0"
    Else
        Me.RSA.text = ""
        Me.RSA.Enabled = False
    End If

End Sub
Private Sub FTMode_Change()
    SetRSA
    SetRadioPattern
End Sub

Private Sub RSA_Change()
    SetRadioPattern
End Sub
Private Sub SetRadioPattern()
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    flag = True
    
    FTMode = Me.FTMode.text
    RSA = Me.RSA.text
    
    '清除「Radio Pattern」旧的侯选值
    Me.RadioPattern.Clear
    
    For m_rowNum = 2 To MappingRadioTemplate.Range("a1048576").End(xlUp).row
         If FTMode = MappingRadioTemplate.Cells(m_rowNum, 1).value Then
            If ((Me.RSA.Enabled = False) Or (RSA = "") Or (RSA = MappingRadioTemplate.Cells(m_rowNum, 2).value)) Then
                Me.RadioPattern.AddItem (MappingRadioTemplate.Cells(m_rowNum, 3).value)
                If flag = True Then
                    RadioPattern.value = MappingRadioTemplate.Cells(m_rowNum, 3).value
                    flag = False
                End If
             End If
        End If
    Next

End Sub
'「Delete」按钮事件,删除模板
Private Sub DeleteRadio()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    
    '用户输入值
    templatename = RadioPattern.text
    
    '当前数据行数
    rowscount = MappingRadioTemplate.Range("a1048576").End(xlUp).row
    
    '存在flag
    existflg = 0
      
    '「Radio Pattern」为空时报错
    If templatename = "" Then
        MsgBox templatename & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")

        existflg = 2
        Exit Sub
    End If
    
    '遍历已有默认侯选值，找到删除项，进行删除。
    For m_rowNum = 2 To rowscount
        If MappingRadioTemplate.Cells(m_rowNum, 3) = templatename _
                And MappingRadioTemplate.Cells(m_rowNum, 1) = FTMode.value _
                And (Me.RSA.Enabled = False Or MappingRadioTemplate.Cells(m_rowNum, 2) = RSA.value) Then
            MappingRadioTemplate.Rows(m_rowNum).Delete
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
         MsgBox templatename & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    If existflg = 1 Then
        RadioPattern.text = ""
        For iIndex = 0 To RadioPattern.ListCount - 1
                    If (RadioPattern.List(iIndex, 0) = templatename) Then
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
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    Dim flag As Boolean
    flag = False
    
    '用户输入值
    templatename = RadioPatternText.text
    
    '当前数据行数
    rowscount = MappingRadioTemplate.Range("a1048576").End(xlUp).row
    
    '重复存在flag
    existflg = 0
      
    '「Radio Pattern」为空时报错
    If templatename = "" Then
        MsgBox templatename & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 1
        Exit Sub
    End If
    
    '遍历已有默认侯选值，若存在与输入值重复项时报错。
    For m_rowNum = 2 To rowscount
        If MappingRadioTemplate.Cells(m_rowNum, 3) = templatename _
                And MappingRadioTemplate.Cells(m_rowNum, 2) = RSA.value _
                And MappingRadioTemplate.Cells(m_rowNum, 1) = FTMode.value Then
            MsgBox templatename & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            RadioPatternText.SetFocus
            existflg = 1
            Exit For
        End If
    Next
        
   lastLineofGroup = rowscount + 1
    Worksheets("MappingRadioTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
 
    '不重复时将输入值追加到候选值列表，并将焦点移动到「eNodeB Transport Data」页的第一条记录
    If existflg = 0 Then
        MappingRadioTemplate.Cells(lastLineofGroup, 1).value = FTMode.value
        MappingRadioTemplate.Cells(lastLineofGroup, 2).value = RSA.value
        MappingRadioTemplate.Cells(lastLineofGroup, 3).value = templatename
        RadioPatternText.value = ""
        Load Me
         MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

Private Sub refreshCell()
    Dim rangeHis As Range
    Dim row, columen As Long
    Dim selectionSheet As Worksheet
    Dim cellsheet As Worksheet
    Dim sheetName As String
    
    Set selectionSheet = ThisWorkbook.ActiveSheet
    
    For Each sheet In ThisWorkbook.Worksheets
        sheetName = sheet.name
        If isCellSheet(sheetName) Then
            Set cellsheet = ThisWorkbook.Worksheets(sheetName)
            cellsheet.Select
            Set rangeHis = Selection
            ActiveSheet.Cells(Selection.row + 1, Selection.column).Select
            rangeHis.Select
        End If
    Next sheet
    
    selectionSheet.Select
    
End Sub

