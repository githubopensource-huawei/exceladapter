VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateForm 
   Caption         =   "增加/删除用户自定义模板"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   OleObjectBlob   =   "TemplateForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "TemplateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CabinetAttr As String = "TYPE"
Private Const CtrlDef_MocCol As Integer = 1
Private Const CtrlDef_AttrCol As Integer = 2
Private Const CtrlDef_ListValueCol As Integer = 5

Private Const MapSite_SiteTypeCol As Integer = 1
Private Const MapSite_CabinetTypeCol As Integer = 2
Private Const MapSite_FDDTDDModeCol As Integer = 3
Private Const MapSite_TempNameCol As Integer = 4

Private Const MapSiteCabinet_SiteTypeCol As Integer = 1
Private Const MapSiteCabinet_CabinetTypeCol As Integer = 2

Private Const MapCell_BandwidthCol As Integer = 1
Private Const MapCell_TxRxModeCol As Integer = 2
Private Const MapCell_FDDTDDCol As Integer = 3
Private Const MapCell_SACol As Integer = 4
Private Const MapCell_TempNameCol As Integer = 5
Private Const MapCell_CellTypeCol As Integer = 6

Private Const MapRadio_FDDTDDCol As Integer = 1
Private Const MapRadio_SACol As Integer = 2
Private Const MapRadio_TempNameCol As Integer = 3
Private Const MapRadio_RadioTypeCol As Integer = 4

Dim extCabinetTypes As Collection


Private Sub UserForm_Activate()
    MultiPage.Font.Size = 10
    Call setFrameCaptions
    
    If MultiPage.value = 0 Then
        Me.SiteTemplateArea.SetFocus
        SetSiteType
    Else
        MultiPage.value = 0
    End If
    
    If getNeType = "USU" Then
        Call grayCabinetFDDTDD
    End If
    
    Dim func As ToolBarFunction
    Set func = New ToolBarFunction
    If func.siteAddSupport = False Then
        MultiPage.Pages.item(0).Visible = False
    End If
    If func.cellAddSupport = False Then
        MultiPage.Pages.item(1).Visible = False
    End If
    If func.radioAddSupport = False Then
        MultiPage.Pages.item(2).Visible = False
    End If
End Sub

Private Sub setFrameCaptions()
    Me.Caption = getResByKey("Bar_Template")
    Me.MultiPage.Pages.item(0).Caption = getResByKey("SiteCaption")
    Me.MultiPage.Pages.item(1).Caption = getResByKey("CellCaption")
    Me.MultiPage.Pages.item(2).Caption = getResByKey("RadioCaption")
    
    Me.AddSiteButton.Caption = getResByKey("AddButtonCaption")
    Me.AddCellButton.Caption = getResByKey("AddButtonCaption")
    Me.AddRadioButton.Caption = getResByKey("AddButtonCaption")
    Me.CancelSiteButton.Caption = getResByKey("CancelButtonCaption")
    Me.CancelCellButton.Caption = getResByKey("CancelButtonCaption")
    Me.CancelRadioButton.Caption = getResByKey("CancelButtonCaption")
    
    
    Me.AddSiteTemplate.Caption = getResByKey("Add")
    Me.AddCellTemplate.Caption = getResByKey("Add")
    Me.AddRadioTemplate.Caption = getResByKey("Add")
    Me.DeleteSiteTemplate.Caption = getResByKey("Delete")
    Me.DeleteCellTemplate.Caption = getResByKey("Delete")
    Me.DeleteRadioTemplate.Caption = getResByKey("Delete")
    
    Me.SiteTypeLabel.Caption = getResByKey("SiteTypeCaption")
    Me.CabinetTypeLabel.Caption = getResByKey("CabinetTypeCaption")
    Me.SiteFDDTDDModeLabel.Caption = getResByKey("FDD/TDD Mode")
    Me.SiteTemplateLabel.Caption = getResByKey("SiteTemplateCaption")
    Me.CellFDDTDDModeLabel.Caption = getResByKey("FDD/TDD Mode")
    Me.BandWidthLabel.Caption = getResByKey("BandWidthCaption")
    Me.TxRxModeLabel.Caption = getResByKey("TxRxModeCaption")
    Me.CellSALabel.Caption = getResByKey("SACaption")
    Me.CellTemplateLabel.Caption = getResByKey("CellTemplateCaption")
    Me.RadioFDDTDDModeLabel.Caption = getResByKey("FDD/TDD Mode")
    Me.RadioSALabel.Caption = getResByKey("SACaption")
    Me.RadioTemplateLabel.Caption = getResByKey("RadioTemplateCaption")
    
    If getNBIOTFlag = True Then
        Me.SiteFDDTDDModeLabel.Caption = getResByKey("Mode")
        Me.CellFDDTDDModeLabel.Caption = getResByKey("Mode")
        Me.RadioFDDTDDModeLabel.Caption = getResByKey("Mode")
    End If
End Sub

Private Sub grayCabinetFDDTDD()
    With Me.CabinetType
        .Clear
        .Enabled = False
    End With
    
    With Me.FDDTDDMode
        .Clear
        .Enabled = False
    End With
End Sub

'选择事件,选择不同的标签(包括Site,Cell和Radio)
Private Sub MultiPage_Change()
 If MultiPage.value = 0 Then
    SetSiteType
 ElseIf MultiPage.value = 1 Then
    setCellFddTddMode
 ElseIf MultiPage.value = 2 Then
    SetFTMode
    SetRSA
 End If
End Sub


'=============================Site========================
'选择事件,选择Add 选项
Private Sub AddSiteTemplate_Click()
    Me.SiteTemplateArea.Visible = True
    Me.SiteTemplateList.Visible = False
    Me.AddSiteButton.Caption = getResByKey("Add")
End Sub

'选择事件,选择Delete选项
Private Sub DeleteSiteTemplate_Click()
    Me.SiteTemplateArea.Visible = False
    Me.SiteTemplateList.Visible = True
    Me.AddSiteButton.Caption = getResByKey("Delete")
    setSiteTemplateList
End Sub

'提交进行Add/Delete操作
Private Sub AddSiteButton_Click()
    If Me.AddSiteTemplate.value = True Then
        addSite
    Else
        deleteSite
    End If
    Call refreshCell
End Sub

'取消此次操作
Private Sub CancelSiteButton_Click()
    Unload Me
End Sub

'「DeleteSite」按钮事件,删除模板
Private Sub deleteSite()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = SiteTemplateList.text
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("sitePatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim mapSiteTemplate As Worksheet
    Set mapSiteTemplate = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    Dim cabinetTypeValue As String
    If getNeType <> "USU" Then
        cabinetTypeValue = Me.CabinetType.value
    Else
        cabinetTypeValue = getUsuCabinetType() 'USU当前只有这一种柜型，由于在表格中不体现这一列，所以给它默认值
    End If
    
    Dim targetRows As New Collection
    Dim firstAddr As String
    Dim targetRange As range
    With mapSiteTemplate.columns(MapSite_TempNameCol)
        Set targetRange = .Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If targetRange.Offset(0, MapSite_FDDTDDModeCol - MapSite_TempNameCol) = FDDTDDMode.value _
                    And targetRange.Offset(0, MapSite_CabinetTypeCol - MapSite_TempNameCol) = cabinetTypeValue _
                    And targetRange.Offset(0, MapSite_SiteTypeCol - MapSite_TempNameCol) = SiteType.value Then
                        targetRows.Add item:=CStr(targetRange.row), key:=CStr(targetRange.row)
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    If targetRows.count > 0 Then
        Dim rowIdx As Integer
        With mapSiteTemplate
            For rowIdx = targetRows.count To 1 Step -1
                .Cells(CInt(targetRows.item(rowIdx)), MapSite_TempNameCol).value = ""
            Next
        End With
        
        Call setSiteTemplateList
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    Else
        MsgBox templateName & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in deleteSite," & Err.Description
End Sub

'「AddSite」按钮事件,添加模板
Private Sub addSite()
    Dim templateName As String
    templateName = Trim(SiteTemplateArea.text)
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("sitePatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    If InStr(templateName, ",") <> 0 Then
        MsgBox templateName & getResByKey("invalidTemplateName_Comma"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim mapSiteTemplate As Worksheet
    Set mapSiteTemplate = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    Dim cabinetTypeValue As String
    If getNeType <> "USU" Then
        cabinetTypeValue = Me.CabinetType.value
    Else
        cabinetTypeValue = getUsuCabinetType() 'USU当前只有这一种柜型，由于在表格中不体现这一列，所以给它默认值
    End If
    
    Dim firstAddr As String
    Dim targetRange As range
    With mapSiteTemplate.columns(MapSite_TempNameCol)
        Set targetRange = .Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If targetRange.Offset(0, MapSite_FDDTDDModeCol - MapSite_TempNameCol) = FDDTDDMode.value _
                    And targetRange.Offset(0, MapSite_CabinetTypeCol - MapSite_TempNameCol) = cabinetTypeValue _
                    And targetRange.Offset(0, MapSite_SiteTypeCol - MapSite_TempNameCol) = SiteType.value Then
                        MsgBox templateName & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
                        SiteTemplateArea.SetFocus
                        Exit Sub
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    Dim insertPos As Long
    With mapSiteTemplate
        insertPos = .range("a65536").End(xlUp).row + 1
        .rows(insertPos).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        .rows(insertPos).NumberFormatLocal = "@" '设置单元格格式为文本
        .Cells(insertPos, MapSite_SiteTypeCol).value = SiteType.value
        .Cells(insertPos, MapSite_CabinetTypeCol).value = cabinetTypeValue
        .Cells(insertPos, MapSite_FDDTDDModeCol).value = FDDTDDMode.value
        .Cells(insertPos, MapSite_TempNameCol).value = templateName
    End With
    
    Me.SiteTemplateArea.value = ""
    Me.SiteTemplateArea.SetFocus
    Load Me
    MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    
    '将不重复时的site Type 与 CabinetType 插入SiteTypeCabinetType
    Dim mapSiteTypeCabinetType As Worksheet
    Set mapSiteTypeCabinetType = ThisWorkbook.Worksheets("Mapping SiteType_CabinetType")
    
    Set targetRange = Nothing
    firstAddr = ""
    With mapSiteTypeCabinetType.columns(MapSiteCabinet_CabinetTypeCol)
        Set targetRange = .Find(cabinetTypeValue, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If targetRange.Offset(0, MapSiteCabinet_SiteTypeCol - MapSiteCabinet_CabinetTypeCol) = SiteType.value Then Exit Sub
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With

    With mapSiteTypeCabinetType
        insertPos = .range("a65536").End(xlUp).row + 1
        .rows(insertPos).NumberFormatLocal = "@"
        .Cells(insertPos, MapSiteCabinet_SiteTypeCol).value = SiteType.value
        .Cells(insertPos, MapSiteCabinet_CabinetTypeCol).value = cabinetTypeValue
    End With
End Sub

'「Site Type」选择事件
Private Sub SiteType_Change()
On Error GoTo ErrorHandler
    If getNeType = "USU" Then
        Call grayCabinetFDDTDD
        Call setSiteTemplateList(False)
        Exit Sub
    End If
    
    Me.CabinetType.Clear
    
    If Me.AddSiteTemplate.value = True Then
        Call setCabinetType("")
    Else
        Call setCabinetType(Me.SiteType.text)
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in SiteType_Change, " & Err.Description
End Sub

Private Sub setCabinetType(siteTypeTxt As String)
On Error GoTo ErrorHandler
    Dim cabinetTypes As New Collection
    
    Dim mapSiteTemplate As Worksheet
    Set mapSiteTemplate = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    Dim cabinetTypeTxt As String
    
    If siteTypeTxt = "" Then
        Dim rowIdx As Integer
        With mapSiteTemplate
            For rowIdx = 2 To .range(getColStr(MapSite_CabinetTypeCol) & "65535").End(xlUp).row
                cabinetTypeTxt = .Cells(rowIdx, MapSite_CabinetTypeCol)
                If cabinetTypeTxt <> "" And Not Contains(cabinetTypes, cabinetTypeTxt) Then
                    cabinetTypes.Add item:=cabinetTypeTxt, key:=cabinetTypeTxt
                    Me.CabinetType.AddItem (cabinetTypeTxt)
                End If
            Next
        End With
    Else
        Dim targetRange As range
        Dim firstAddr As String
        
        With mapSiteTemplate.columns(MapSite_SiteTypeCol)
            Set targetRange = .Find(siteTypeTxt, lookat:=xlWhole, LookIn:=xlValues)
            If Not targetRange Is Nothing Then
                firstAddr = targetRange.address
                Do
                    cabinetTypeTxt = targetRange.Offset(0, MapSite_CabinetTypeCol - MapSite_SiteTypeCol).value
                    
                    If cabinetTypeTxt <> "" And Not Contains(cabinetTypes, cabinetTypeTxt) Then
                        cabinetTypes.Add item:=cabinetTypeTxt, key:=cabinetTypeTxt
                        Me.CabinetType.AddItem (cabinetTypeTxt)
                    End If
                    Set targetRange = .FindNext(targetRange)
                Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
            End If
        End With
    End If
    
    Set extCabinetTypes = getExtCabinetTypes()
    If Not extCabinetTypes Is Nothing Then
        Dim cabinetName As Variant
        For Each cabinetName In extCabinetTypes
            If Not Contains(cabinetTypes, CStr(cabinetName)) Then
                cabinetTypes.Add item:=CStr(cabinetName), key:=CStr(cabinetName)
                Me.CabinetType.AddItem (CStr(cabinetName))
            End If
        Next
    End If
    
    If Me.CabinetType.ListCount <> 0 Then Me.CabinetType.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setCabinetType, " & Err.Description
End Sub

'「Cabinet Type」选择事件
Private Sub CabinetType_Change()
On Error GoTo ErrorHandler
    If Me.FDDTDDMode.Enabled = False Then Exit Sub
    
    Me.FDDTDDMode.Clear

    If Me.AddSiteTemplate.value = True Then
        Call setFDDTDDMode("", "")
    Else
        Call setFDDTDDMode(Me.SiteType.text, Me.CabinetType.text)
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in CabinetType_Change, " & Err.Description
End Sub

Private Sub setFDDTDDMode(siteTypeTxt As String, cabinetTypeTxt As String)
On Error GoTo ErrorHandler
    Dim modes As New Collection
    
    Dim mapSiteTemplate As Worksheet
    Set mapSiteTemplate = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    Dim mode As String
    
    If siteTypeTxt = "" Or cabinetTypeTxt = "" Then
        Dim rowIdx As Integer
        With mapSiteTemplate
            For rowIdx = 2 To .range(getColStr(MapSite_FDDTDDModeCol) & "65535").End(xlUp).row
                mode = .Cells(rowIdx, MapSite_FDDTDDModeCol)
                If mode <> "" And Not Contains(modes, mode) Then
                    modes.Add item:=mode, key:=mode
                    Me.FDDTDDMode.AddItem (mode)
                End If
            Next
        End With
    Else
        Dim targetRange As range
        Dim firstAddr As String
        
        With mapSiteTemplate.columns(MapSite_CabinetTypeCol)
            Set targetRange = .Find(cabinetTypeTxt, lookat:=xlWhole, LookIn:=xlValues)
            If Not targetRange Is Nothing Then
                firstAddr = targetRange.address
                Do
                    If targetRange.Offset(0, MapSite_SiteTypeCol - MapSite_CabinetTypeCol).value = siteTypeTxt Then
                        mode = targetRange.Offset(0, MapSite_FDDTDDModeCol - MapSite_CabinetTypeCol).value
                        If mode <> "" And Not Contains(modes, mode) Then
                            modes.Add item:=mode, key:=mode
                            Me.FDDTDDMode.AddItem (mode)
                        End If
                    End If
                    Set targetRange = .FindNext(targetRange)
                Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
            End If
        End With
    End If
    
    If Me.FDDTDDMode.ListCount <> 0 Then Me.FDDTDDMode.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setFDDTDDMode, " & Err.Description
End Sub

'「FDDTDD Mode」选择事件
Private Sub FDDTDDMode_Change()
On Error GoTo ErrorHandler
    If Me.CabinetType.Enabled = False Then Exit Sub
    
    Call setSiteTemplateList
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in SiteType_Change, " & Err.Description
End Sub

'从「MappingSiteTypeCabinetType」页获取「*Site Type」列侯选值
Private Sub SetSiteType()
    Dim m_rowNum As Long
    Dim m_RowNum_Inner As Long
    Dim m_Str As String
    Dim flag As Boolean
    
    Me.SiteType.Clear

    Dim mapSiteTemplate As Worksheet
    Set mapSiteTemplate = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    Dim siteTypes As New Collection
    Dim siteTypeValue As String
    Dim rowIndex As Long
    With mapSiteTemplate
        For rowIndex = 2 To .range("a65536").End(xlUp).row
            siteTypeValue = .Cells(rowIndex, MapSite_SiteTypeCol).value
            If siteTypeValue <> "" And Not Contains(siteTypes, siteTypeValue) Then
                siteTypes.Add item:=siteTypeValue, key:=siteTypeValue
                Me.SiteType.AddItem siteTypeValue
            End If
        Next
    End With
    
    If Me.SiteType.ListCount <> 0 Then Me.SiteType.ListIndex = 0 '显示第一个对象
End Sub

Private Sub setSiteTemplateList(Optional withCbnt As Boolean = True)
On Error GoTo ErrorHandler
    Me.SiteTemplateList.Clear
    
    If getNeType = "USU" Then withCbnt = False
    
    Dim siteTypeTxt As String
    Dim cabinetTypeTxt As String
    Dim modeTxt As String
    
    siteTypeTxt = Me.SiteType.text
    If withCbnt = True Then
        cabinetTypeTxt = Me.CabinetType.text
        modeTxt = Me.FDDTDDMode.text
    Else
        cabinetTypeTxt = getUsuCabinetType
    End If
    
    Dim mapSiteTemplate As Worksheet
    Set mapSiteTemplate = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    Dim targetRange As range
    Dim firstAddr As String
    Dim siteTemplateName As String
    
    With mapSiteTemplate.columns(MapSite_SiteTypeCol)
        Set targetRange = .Find(siteTypeTxt, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If withCbnt = True Then
                    If cabinetTypeTxt = targetRange.Offset(0, MapSite_CabinetTypeCol - MapSite_SiteTypeCol).value _
                        And modeTxt = targetRange.Offset(0, MapSite_FDDTDDModeCol - MapSite_SiteTypeCol).value Then
                            siteTemplateName = targetRange.Offset(0, MapSite_TempNameCol - MapSite_SiteTypeCol).value
                            If siteTemplateName <> "" Then SiteTemplateList.AddItem (siteTemplateName)
                    End If
                Else
                    siteTemplateName = targetRange.Offset(0, MapSite_TempNameCol - MapSite_SiteTypeCol).value
                    If siteTemplateName <> "" Then SiteTemplateList.AddItem (siteTemplateName)
                End If
                
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    If Me.SiteTemplateList.ListCount <> 0 Then Me.SiteTemplateList.ListIndex = 0
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setSiteTemplateList, " & Err.Description
End Sub

Private Function getUsuCabinetType() As String
    Dim usuCabinetType As String, eachUsuCabinetType As String
    usuCabinetType = ""
    Dim rowIndex As Long
    For rowIndex = 2 To Worksheets("MappingSiteTemplate").range("B65536").End(xlUp).row
        eachUsuCabinetType = Worksheets("MappingSiteTemplate").range("B" & rowIndex).value
        If eachUsuCabinetType <> "" Then
            usuCabinetType = eachUsuCabinetType
            Exit For '如果不为空，则找到一个机柜类型即可
        End If
    Next rowIndex
    
    If usuCabinetType = "" Then usuCabinetType = "VIRTUAL"
    getUsuCabinetType = usuCabinetType
End Function

Public Function getExtCabinetTypes() As Collection
On Error GoTo ErrorHandler
    Dim cabinetTypes As String
    Dim v() As String

    Dim targetRange As range
    With ThisWorkbook.Worksheets("CONTROL DEF").columns(CtrlDef_AttrCol)
        Set targetRange = .Find(CabinetAttr, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            If targetRange.Offset(0, CtrlDef_MocCol - CtrlDef_AttrCol).value = "NODE" Then
                cabinetTypes = targetRange.Offset(0, CtrlDef_ListValueCol - CtrlDef_AttrCol).value
                If cabinetTypes = "" Then
                    Exit Function
                End If
                v = Split(cabinetTypes, ",")
            End If
        Else
            Exit Function
        End If
    End With

    Set getExtCabinetTypes = New Collection
    Dim i As Integer
    For i = 1 To UBound(v)
       getExtCabinetTypes.Add item:=v(i), key:=v(i)
    Next
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getExtCabinetTypes, " & Err.Description
End Function



'==============================Cell=====================================
'添加和删除「MappingCellTemplate」页的模板
'「*Band Width」「CellFddTddMode」「FDD/TDD Mode」「*Cell Mode」列侯选值的窗体
Private Sub AddCellTemplate_Click()
    Me.CellTemplateArea.Visible = True
    Me.CellTemplateList.Visible = False
    Me.AddCellButton.Caption = getResByKey("Add")
End Sub

Private Sub DeleteCellTemplate_Click()
    Me.CellTemplateList.Visible = True
    Me.CellTemplateArea.Visible = False
    Me.AddCellButton.Caption = getResByKey("Delete")
    setCellTemplateList
End Sub

'提交进行Add/Delete操作
Private Sub AddCellButton_Click()
    If Me.AddCellTemplate.value = True Then
        addCell
    Else
        deleteCell
    End If
    Call refreshCell
End Sub

Private Sub CancelCellButton_Click()
    Unload Me
End Sub

'「Delete」按钮事件,删除模板
Private Sub deleteCell()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = CellTemplateList.text
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("cellPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Dim targetRows As New Collection
    
    Dim firstAddr As String
    Dim targetRange As range
    With mapCellTemplate.columns(MapCell_TempNameCol)
        Set targetRange = .Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If (targetRange.Offset(0, MapCell_SACol - MapCell_TempNameCol) = SA.value Or Me.SA.Enabled = False) _
                    And targetRange.Offset(0, MapCell_FDDTDDCol - MapCell_TempNameCol) = CellFddTddMode.value _
                    And (targetRange.Offset(0, MapCell_TxRxModeCol - MapCell_TempNameCol) = TxRxMode.value Or Me.TxRxMode.Enabled = False) _
                    And (targetRange.Offset(0, MapCell_BandwidthCol - MapCell_TempNameCol) = BandWidth.value Or Me.BandWidth.Enabled = False) Then
                        targetRows.Add item:=CStr(targetRange.row), key:=CStr(targetRange.row)
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    If targetRows.count > 0 Then
        Dim rowIdx As Integer
        With mapCellTemplate
            For rowIdx = targetRows.count To 1 Step -1
                .Cells(CInt(targetRows.item(rowIdx)), MapCell_TempNameCol).value = ""
            Next
        End With
        
        setCellTemplateList
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    Else
        MsgBox templateName & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in deleteCell, " & Err.Description
End Sub

'「Add」按钮事件,添加模板
Private Sub addCell()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = Trim(CellTemplateArea.text)
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("cellPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    If InStr(templateName, ",") <> 0 Then
        MsgBox templateName & getResByKey("invalidTemplateName_Comma"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Dim firstAddr As String
    Dim targetRange As range
    With mapCellTemplate.columns(MapCell_TempNameCol)
        Set targetRange = .Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If (targetRange.Offset(0, MapCell_SACol - MapCell_TempNameCol) = SA.value Or Me.SA.Enabled = False) _
                    And targetRange.Offset(0, MapCell_FDDTDDCol - MapCell_TempNameCol) = CellFddTddMode.value _
                    And (targetRange.Offset(0, MapCell_TxRxModeCol - MapCell_TempNameCol) = TxRxMode.value Or Me.TxRxMode.Enabled = False) _
                    And (targetRange.Offset(0, MapCell_BandwidthCol - MapCell_TempNameCol) = BandWidth.value Or Me.BandWidth.Enabled = False) Then
                        MsgBox templateName & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
                        CellTemplateArea.SetFocus
                        Exit Sub
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    Dim insertPos As Long
    With mapCellTemplate
        insertPos = .range("E65536").End(xlUp).row + 1
        .rows(insertPos).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        .rows(insertPos).NumberFormatLocal = "@" '设置单元格格式为文本
        .Cells(insertPos, MapCell_BandwidthCol).value = BandWidth.value
        .Cells(insertPos, MapCell_TxRxModeCol).value = TxRxMode.value
        .Cells(insertPos, MapCell_FDDTDDCol).value = CellFddTddMode.value
        .Cells(insertPos, MapCell_SACol).value = SA.value
        .Cells(insertPos, MapCell_TempNameCol).value = templateName
        .Cells(insertPos, MapCell_CellTypeCol).value = getResByKey("LTE Cell")
    End With

    Me.CellTemplateArea.value = ""
    Me.CellTemplateArea.SetFocus
    Load Me
    MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in addCell, " & Err.Description
End Sub

Private Sub CellFddTddMode_Change()
    SetBandWidth
    SetTxRxMode
    SetSA
    setCellTemplateList
End Sub

Private Sub BandWidth_Change()
    setCellTemplateList
End Sub

Private Sub TxRxMode_Change()
    setCellTemplateList
End Sub

Private Sub SA_Change()
    setCellTemplateList
End Sub

'设置「*CellFddTddMode」列侯选值
Private Sub setCellFddTddMode()
    Dim baseStationVersion As String
    baseStationVersion = UCase(getBaseStationVersion)
    
    Me.CellFddTddMode.Clear
    '根据基站版本增加FDD/TDD Mode，宏站增加FDD/TDD
    If InStr(baseStationVersion, "BTS3900") <> 0 Or InStr(baseStationVersion, "BTS5900") <> 0 Then
        Me.CellFddTddMode.AddItem ("FDD")
        Me.CellFddTddMode.AddItem ("TDD")
        If getNBIOTFlag = True Then
             Me.CellFddTddMode.AddItem ("NB-IoT")
        End If
    ElseIf InStr(baseStationVersion, "BTS3205E") <> 0 Then
        '3205E只支持TDD
        Me.CellFddTddMode.AddItem ("TDD")
    Else
        '其余小站只支持FDD
        Me.CellFddTddMode.AddItem ("FDD")
    End If

    Me.CellFddTddMode.ListIndex = 0
End Sub

'设置「*BandWidth」列侯选值
Private Sub SetBandWidth()
    Me.BandWidth.Clear
    
    If Me.CellFddTddMode.text = "FDD" Then
        Me.BandWidth.Enabled = False
    ElseIf Me.CellFddTddMode.text = "TDD" Then
        Me.BandWidth.Enabled = True
        Me.BandWidth.AddItem ("5M")
        Me.BandWidth.AddItem ("10M")
        Me.BandWidth.AddItem ("15M")
        Me.BandWidth.AddItem ("20M")
        Me.BandWidth.text = "10M"
    Else
        Me.BandWidth.Enabled = False
    End If
End Sub

'设置「*TxRxMode」列侯选值
Private Sub SetTxRxMode()
    
    Me.TxRxMode.Clear
    
    If Me.CellFddTddMode.text = "FDD" Then
        Me.TxRxMode.Enabled = False
    ElseIf Me.CellFddTddMode.text = "TDD" Then
        Me.TxRxMode.Enabled = True
        Me.TxRxMode.AddItem ("1T1R")
        Me.TxRxMode.AddItem ("2T2R")
        Me.TxRxMode.AddItem ("4T4R")
        Me.TxRxMode.AddItem ("8T8R")
        Me.TxRxMode.AddItem ("64T64R")
        Me.TxRxMode.text = "1T1R"
    Else
        Me.TxRxMode.Enabled = False
    End If
    
End Sub

'设置「*SA」列侯选值
Private Sub SetSA()
    Me.SA.Clear
    If Me.CellFddTddMode.text = "TDD" Then
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
        Me.SA.Enabled = False
    End If
End Sub

'从「MappingCellTemplate」页获取「*Cell Pattern」列侯选值
Private Sub setCellTemplateList()
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    flag = True
    
    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    CellFddTddMode = Me.CellFddTddMode.text
    BandWidth = Me.BandWidth.text
    TxRxMode = Me.TxRxMode.text
    SA = Me.SA.text
    
    '清除「Cell Pattern」旧的侯选值
    Me.CellTemplateList.Clear
    
    With mapCellTemplate
        For m_rowNum = 2 To .range("E65536").End(xlUp).row
            If BandWidth = .Cells(m_rowNum, 1).value And CellFddTddMode = .Cells(m_rowNum, 3).value And TxRxMode = .Cells(m_rowNum, 2).value _
                Or (Me.BandWidth.Enabled = False And Me.TxRxMode.Enabled = False And CellFddTddMode = .Cells(m_rowNum, 3).value) Then
                    If (Me.SA.Enabled = False) Or (SA = .Cells(m_rowNum, 4).value) Then
                        If Len(.Cells(m_rowNum, 5).value) > 0 Then
                            Me.CellTemplateList.AddItem (.Cells(m_rowNum, 5).value)
                            If flag = True Then
                                CellTemplateList.value = .Cells(m_rowNum, 5).value
                                flag = False
                             End If
                        End If

                    End If
            End If
        Next
    End With
End Sub



'=============================Radio=====================================
'添加和删除「MappingRadioTemplate」页的模板
'「RadioMode」「RadioSA」「Radio Pattern」列侯选值的窗体
Private Sub AddRadioTemplate_Click()
    Me.RadioTemplateArea.Visible = True
    Me.RadioTemplateList.Visible = False
    Me.AddRadioButton.Caption = getResByKey("Add")
End Sub

Private Sub DeleteRadioTemplate_Click()
    Me.RadioTemplateList.Visible = True
    Me.RadioTemplateArea.Visible = False
    Me.AddRadioButton.Caption = getResByKey("Delete")
    setRadioTemplateList
End Sub

'提交进行Add/Delete操作
Private Sub AddRadioButton_Click()
    If Me.AddRadioTemplate.value = True Then
        addRadio
    Else
        deleteRadio
    End If
    Call refreshCell
End Sub

Private Sub CancelRadioButton_Click()
    Unload Me
End Sub

Private Sub SetFTMode()
    Dim baseStationVersion As String
    baseStationVersion = UCase(getBaseStationVersion)
    
    Me.RadioMode.Clear
    If InStr(baseStationVersion, "BTS3900") <> 0 Or InStr(baseStationVersion, "BTS5900") <> 0 Then
        Me.RadioMode.AddItem ("FDD")
        Me.RadioMode.AddItem ("TDD")
        Me.RadioMode.AddItem ("FDDTDD")
        If getNBIOTFlag = True Then
            Me.RadioMode.AddItem ("NB-IoT")
            Me.RadioMode.AddItem ("FDDNB-IoT")
        End If
    ElseIf InStr(baseStationVersion, "BTS3205E") <> 0 Then
        '3205E只支持TDD
        Me.RadioMode.AddItem ("TDD")
    Else
        '其余小站只支持FDD
        Me.RadioMode.AddItem ("FDD")
    End If
    Me.RadioMode.ListIndex = 0
End Sub

Private Sub SetRSA()
    Me.RadioSA.Clear
    If Me.RadioMode.text = "TDD" Then
        Me.RadioSA.Enabled = True
        Me.RadioSA.AddItem ("SA0")
        Me.RadioSA.AddItem ("SA1")
        Me.RadioSA.AddItem ("SA2")
        Me.RadioSA.AddItem ("SA3")
        Me.RadioSA.AddItem ("SA4")
        Me.RadioSA.AddItem ("SA5")
        Me.RadioSA.AddItem ("SA6")
        Me.RadioSA.text = "SA0"
    Else
        Me.RadioSA.Enabled = False
    End If
End Sub

Private Sub RadioMode_Change()
    SetRSA
    setRadioTemplateList
End Sub

Private Sub RadioSA_Change()
    setRadioTemplateList
End Sub

Private Sub setRadioTemplateList()
On Error GoTo ErrorHandler
    Dim radioModeTxt As String
    Dim radioSaTxt As String
    radioModeTxt = Me.RadioMode.text
    radioSaTxt = Me.RadioSA.text
    
    Dim mapRadioTemplate As Worksheet
    Set mapRadioTemplate = ThisWorkbook.Worksheets("MappingRadioTemplate")

    Me.RadioTemplateList.Clear
    
    Dim firstAddr As String
    Dim targetRange As range
    With mapRadioTemplate.columns(MapRadio_FDDTDDCol)
        Set targetRange = .Find(radioModeTxt, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If (targetRange.Offset(0, MapRadio_SACol - MapRadio_FDDTDDCol).value = radioSaTxt Or Me.RadioSA.Enabled = False Or radioSaTxt = "") Then
                    If Len(targetRange.Offset(0, MapRadio_TempNameCol - MapRadio_FDDTDDCol).value) > 0 Then
                       Me.RadioTemplateList.AddItem (targetRange.Offset(0, MapRadio_TempNameCol - MapRadio_FDDTDDCol).value)
                    End If
                End If
                
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    If Me.RadioTemplateList.ListCount <> 0 Then Me.RadioTemplateList.ListIndex = 0
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setRadioTemplateList, " & Err.Description
End Sub

'「Delete」按钮事件,删除模板
Private Sub deleteRadio()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = RadioTemplateList.text
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim radioModeTxt As String
    Dim radioSaTxt As String
    radioModeTxt = Me.RadioMode.text
    radioSaTxt = Me.RadioSA.text
    
    Dim mapRadioTemplate As Worksheet
    Set mapRadioTemplate = ThisWorkbook.Worksheets("MappingRadioTemplate")
    
    Dim targetRows As New Collection
    
    Dim targetRange As range
    Dim firstAddr As String
    With mapRadioTemplate
        Set targetRange = .columns(MapRadio_TempNameCol).Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If targetRange.Offset(0, MapRadio_FDDTDDCol - MapRadio_TempNameCol).value = radioModeTxt _
                    And (targetRange.Offset(0, MapRadio_SACol - MapRadio_TempNameCol).value = radioSaTxt Or Me.RadioSA.Enabled = False Or radioSaTxt = "") Then
                        targetRows.Add item:=CStr(targetRange.row), key:=CStr(targetRange.row)
                End If
                targetRange = .columns(MapRadio_TempNameCol).FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    If targetRows.count > 0 Then
        Dim rowIdx As Integer
        With mapRadioTemplate
            For rowIdx = targetRows.count To 1 Step -1
                .Cells(CInt(targetRows.item(rowIdx)), MapRadio_TempNameCol).value = ""
            Next
        End With
        
        setRadioTemplateList
        RadioTemplateList.SetFocus
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    Else
        MsgBox templateName & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in deleteRadio, " & Err.Description
End Sub

'「Add」按钮事件,添加模板
Private Sub addRadio()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = Trim(RadioTemplateArea.text)
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    If InStr(templateName, ",") <> 0 Then
        MsgBox templateName & getResByKey("invalidTemplateName_Comma"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim radioModeTxt As String
    Dim radioSaTxt As String
    radioModeTxt = Me.RadioMode.text
    radioSaTxt = Me.RadioSA.text
    
    Dim mapRadioTemplate As Worksheet
    Set mapRadioTemplate = ThisWorkbook.Worksheets("MappingRadioTemplate")
    
    Dim existFlag As Boolean
    existFlag = False
    
    Dim targetRange As range
    Dim firstAddr As String
    With mapRadioTemplate
        Set targetRange = .columns(MapRadio_TempNameCol).Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If targetRange.Offset(0, MapRadio_FDDTDDCol - MapRadio_TempNameCol).value = radioModeTxt _
                    And (targetRange.Offset(0, MapRadio_SACol - MapRadio_TempNameCol).value = radioSaTxt Or Me.RadioSA.Enabled = False) Then
                        MsgBox templateName & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
                        RadioTemplateArea.SetFocus
                        Exit Sub
                End If
                targetRange = .columns(MapRadio_TempNameCol).FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    Dim insertPos As Long
    With mapRadioTemplate
        insertPos = .range("a65536").End(xlUp).row + 1
        .rows(insertPos).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        .rows(insertPos).NumberFormatLocal = "@" '设置单元格格式为文本
        .Cells(insertPos, MapRadio_FDDTDDCol).value = radioModeTxt
        .Cells(insertPos, MapRadio_SACol).value = radioSaTxt
        .Cells(insertPos, MapRadio_TempNameCol).value = templateName
    End With

    RadioTemplateArea.value = ""
    Load Me
    MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in addRadio, " & Err.Description
End Sub

Private Sub refreshCell()
    Dim rangeHis As range
    Dim row, columen As Long
    Set rangeHis = Selection
    ActiveSheet.Cells(Selection.row + 1, Selection.column).Select
    rangeHis.Select
End Sub




