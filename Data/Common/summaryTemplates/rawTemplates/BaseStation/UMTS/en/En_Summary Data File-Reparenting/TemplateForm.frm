VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateForm 
   Caption         =   "Add/Delete User-defined Template"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   OleObjectBlob   =   "TemplateForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "TemplateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Const ProductType_SiteTypeCol As Integer = 1
Private Const ProductType_NeTypeCol As Integer = 2

Private Const MapSite_SiteTypeCol As Integer = 1
Private Const MapSite_CabinetTypeCol As Integer = 2
Private Const MapSite_FDDTDDModeCol As Integer = 3
Private Const MapSite_TempNameCol As Integer = 4
Private Const MapSite_NeTypeCol As Integer = 5

Private Const MapCell_TempNameCol As Integer = 1
Private Const MapCell_CellTypeCol As Integer = 2
Private Const MapCell_NeTypeCol As Integer = 3
Private Const MapCell_BandwidthCol As Integer = 4
Private Const MapCell_TxRxModeCol As Integer = 5
Private Const MapCell_FDDTDDCol As Integer = 6
Private Const MapCell_SACol As Integer = 7

Private Const MapRadio_TempNameCol As Integer = 1
Private Const MapRadio_RadioTypeCol As Integer = 2
Private Const MapRadio_NeTypeCol As Integer = 3

Private Const OpMode_Add As String = "Add"
Private Const OpMode_Del As String = "Delete"
Private OpMode As String

Private neType As String


'激活事件,默认首先显示Site标签页
Private Sub UserForm_Activate()
    MultiPage.Font.Size = 10
    Call setFrameCaptions
    neType = getNeType
    
    Me.Caption = getResByKey("Bar_Template")
    If MultiPage.value = 0 Then
        Me.SiteTemplateArea.SetFocus
        setSiteTypeListValue
    Else
        Me.MultiPage.value = 0
    End If
    Dim func As ToolBarFunction
    Set func = New ToolBarFunction
    If func.siteAddSupport = False Then
        Me.MultiPage.Pages.Item(0).Visible = False
    End If
    If func.cellAddSupport = False Then
        Me.MultiPage.Pages.Item(1).Visible = False
    End If
    If func.radioAddSupport = False Then
        Me.MultiPage.Pages.Item(2).Visible = False
    End If
    
    OpMode = OpMode_Add
End Sub

Private Sub setFrameCaptions()
    Me.Caption = getResByKey("Bar_Template")
    Me.MultiPage.Pages.Item(0).Caption = getResByKey("SiteCaption")
    Me.MultiPage.Pages.Item(1).Caption = getResByKey("CellCaption")
    Me.MultiPage.Pages.Item(2).Caption = getResByKey("RadioCaption")
    Me.AddSiteTemplateRadio.Caption = getResByKey("AddButtonCaption")
    Me.AddCellTemplateRadio.Caption = getResByKey("AddButtonCaption")
    Me.AddRadioTemplateRadio.Caption = getResByKey("AddButtonCaption")
       Me.AddSiteButton.Caption = getResByKey("AddButtonCaption")
    Me.CancelSiteButton.Caption = getResByKey("CancelButtonCaption")
    
    Me.AddCellButton.Caption = getResByKey("AddButtonCaption")
    Me.CancelCellButton.Caption = getResByKey("CancelButtonCaption")
    
    Me.AddRadioButton.Caption = getResByKey("AddButtonCaption")
    Me.CancelRadioButton.Caption = getResByKey("CancelButtonCaption")
    Me.DeleteSiteTemplateRadio.Caption = getResByKey("DeleteButtonCaption")
    Me.DeleteCellTemplateRadio.Caption = getResByKey("DeleteButtonCaption")
    Me.DeleteRadioTemplateRadio.Caption = getResByKey("DeleteButtonCaption")
    
    Me.SiteTypeLabel.Caption = getResByKey("SiteTypeCaption")
    Me.CellTypeLabel.Caption = getResByKey("CellTypeCaption")
    Me.RadioTemplateTypeLabel.Caption = getResByKey("RadioTemplateTypeCaption")
    
    Me.SiteTemplateLabel.Caption = getResByKey("SiteTemplateCaption")
    Me.CellFDDTDDModeLabel.Caption = getResByKey("FDD/TDD Mode")
    Me.BandWidthLabel.Caption = getResByKey("BandWidthCaption")
    Me.TxRxModeLabel.Caption = getResByKey("TxRxModeCaption")
    Me.CellSALabel.Caption = getResByKey("SACaption")
    Me.CellTemplateLabel.Caption = getResByKey("CellTemplateCaption")
    Me.RadioTemplateLabel.Caption = getResByKey("RadioTemplateCaption")
    
    If getNBIOTFlag = True Then
        Me.CellFDDTDDModeLabel.Caption = getResByKey("Mode")
    End If
End Sub

'选择事件,选择不同的标签(包括Site,Cell和Radio)
Private Sub MultiPage_Change()
    OpMode = OpMode_Add
    If Me.MultiPage.value = 0 Then
        If Me.DeleteSiteTemplateRadio = True Then OpMode = OpMode_Del
        setSiteTypeListValue
    ElseIf Me.MultiPage.value = 1 Then
        If Me.DeleteCellTemplateRadio = True Then OpMode = OpMode_Del
        If getNBIOTFlag = True Then
           Me.CellFDDTDDModeLabel.Caption = getResByKey("Mode")
        End If
        setCellTypeListValue
    ElseIf Me.MultiPage.value = 2 Then
        If Me.DeleteRadioTemplateRadio = True Then OpMode = OpMode_Del
        setRadioTypeListValue
    End If
End Sub



'=============================Site========================
'选择事件,选择Add 选项
Private Sub AddSiteTemplateRadio_Click()
    Me.SiteTemplateArea.Visible = True
    Me.SiteTemplateList.Visible = False
    Me.AddSiteButton.Caption = getResByKey("Add")
    OpMode = OpMode_Add
    setSiteTypeListValue
End Sub

'选择事件,选择Delete选项
Private Sub DeleteSiteTemplateRadio_Click()
    Me.SiteTemplateArea.Visible = False
    Me.SiteTemplateList.Visible = True
    Me.AddSiteButton.Caption = getResByKey("Delete")
    OpMode = OpMode_Del
    setSiteTypeListValue
End Sub

'提交进行Add/Delete操作
Private Sub AddSiteButton_Click()
    If OpMode = OpMode_Add Then
        addSiteTemplate
    Else
        deleteSiteTemplate
    End If
    Call refreshCell
End Sub

'取消此次操作
Private Sub CancelSiteButton_Click()
    Unload Me
End Sub

'从「MappingSiteTemplate」页获取「*Site Type」列侯选值
Private Sub setSiteTypeListValue()
On Error GoTo ErrorHandler
    Dim productTypeSht As Worksheet
    Set productTypeSht = ThisWorkbook.Worksheets("ProductType")
    
    Me.SiteTypeList.Clear
    
    Dim productTypes As New Collection
    
    Dim productTypeText As String
    Dim rowIdx As Integer
    With productTypeSht
        For rowIdx = 2 To .range("a65536").End(xlUp).row
            If .Cells(rowIdx, ProductType_NeTypeCol).value = neType Then
                productTypeText = .Cells(rowIdx, ProductType_SiteTypeCol).value
                If productTypeText <> "" And Not Contains(productTypes, productTypeText) Then Me.SiteTypeList.AddItem (productTypeText)
            End If
        Next
    End With
    
    If Me.SiteTypeList.ListCount Then Me.SiteTypeList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setSiteTypeListValue," & Err.Description
End Sub

'「Site Type」选择事件
Private Sub SiteTypeList_Change()
    setSiteTemplateListValue
End Sub

'从「MappingSiteTemplate」页获取「*Site Patten」列侯选值
Private Sub setSiteTemplateListValue()
On Error GoTo ErrorHandler
    Dim siteType As String
    siteType = Me.SiteTypeList.text
    
    Dim mapSiteTemplate As Worksheet
    Set mapSiteTemplate = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    SiteTemplateList.Clear
    Dim templates As New Collection
    
    Dim rowIdx As Integer
    With mapSiteTemplate
        For rowIdx = 2 To .range(getColStr(MapSite_TempNameCol) & "65536").End(xlUp).row
            If siteType = .Cells(rowIdx, MapSite_SiteTypeCol).value And neType = .Cells(rowIdx, MapSite_NeTypeCol).value Then
                Dim siteTemplateName As String
                siteTemplateName = .Cells(rowIdx, MapSite_TempNameCol).value
                If siteTemplateName <> "" And Not Contains(templates, siteTemplateName) Then
                    templates.Add Item:=siteTemplateName, key:=siteTemplateName
                    Me.SiteTemplateList.AddItem (siteTemplateName)
                End If
            End If
        Next
    End With
    
    If Me.SiteTemplateList.ListCount > 0 Then Me.SiteTemplateList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setSiteTemplateListValue," & Err.Description
End Sub

'「AddSite」按钮事件,添加模板
Private Sub addSiteTemplate()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = Trim(Me.SiteTemplateArea.text)
    
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
    
    Dim firstAddr As String
    Dim targetRange As range
    With mapSiteTemplate.columns(MapSite_TempNameCol)
        Set targetRange = .Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If targetRange.Offset(0, MapSite_SiteTypeCol - MapSite_TempNameCol) = SiteTypeList.value _
                    And targetRange.Offset(0, MapSite_NeTypeCol - MapSite_TempNameCol) = neType Then
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
        .Cells(insertPos, MapSite_SiteTypeCol).value = SiteTypeList.value
        .Cells(insertPos, MapSite_CabinetTypeCol).value = ""
        .Cells(insertPos, MapSite_FDDTDDModeCol).value = ""
        .Cells(insertPos, MapSite_TempNameCol).value = templateName
        .Cells(insertPos, MapSite_NeTypeCol).value = neType
    End With
    
    Me.SiteTemplateArea.value = ""
    Me.SiteTemplateArea.SetFocus
    Load Me
    MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in addSiteTemplate," & Err.Description
End Sub

'「DeleteSite」按钮事件,删除模板
Private Sub deleteSiteTemplate()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = Trim(Me.SiteTemplateList.text)
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("sitePatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim mapSiteTemplate As Worksheet
    Set mapSiteTemplate = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    Dim targetRows As New Collection
    Dim firstAddr As String
    Dim targetRange As range
    With mapSiteTemplate.columns(MapSite_TempNameCol)
        Set targetRange = .Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If targetRange.Offset(0, MapSite_SiteTypeCol - MapSite_TempNameCol) = SiteTypeList.value _
                    And targetRange.Offset(0, MapSite_NeTypeCol - MapSite_TempNameCol) = neType Then
                        targetRows.Add Item:=CStr(targetRange.row), key:=CStr(targetRange.row)
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    If targetRows.count > 0 Then
        Dim rowIdx As Integer
        With mapSiteTemplate
            For rowIdx = targetRows.count To 1 Step -1
                .rows(CInt(targetRows.Item(rowIdx))).Delete
            Next
        End With
        
        Call setSiteTemplateListValue
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    Else
        MsgBox templateName & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in deleteSiteTemplate," & Err.Description
End Sub




'=============================Cell========================
Private Sub AddCellTemplateRadio_Click()
    Me.CellTemplateTextArea.Visible = True
    Me.CellTemplateList.Visible = False
    Me.AddCellButton.Caption = getResByKey("Add")
    OpMode = OpMode_Add
    setCellTypeListValue
End Sub

Private Sub DeleteCellTemplateRadio_Click()
    Me.CellTemplateList.Visible = True
    Me.CellTemplateTextArea.Visible = False
    Me.AddCellButton.Caption = getResByKey("Delete")
    OpMode = OpMode_Del
    setCellTypeListValue
End Sub

'提交进行Add/Delete操作
Private Sub AddCellButton_Click()
    If OpMode = OpMode_Add Then
        addCellTemplate
    Else
        deleteCellTemplate
    End If
    Call refreshCell
End Sub

'取消此次操作
Private Sub CancelCellButton_Click()
    Unload Me
End Sub

Private Sub setCellTypeListValue()
On Error GoTo ErrorHandler
    Me.CellTypeList.Clear
    
    With Me.CellTypeList
        If neType = "MRAT" Then
            If isContainBaseStation() Then
                If containsCellSheet("GSM") Then .AddItem getResByKey("GSM Local Cell")
                If containsCellSheet("UMTS") Then .AddItem getResByKey("UMTS Local Cell")
                If containsCellSheet("LTE") Then .AddItem getResByKey("LTE Cell")
                If containsCellSheet("NB-IoT") Then .AddItem getResByKey("NB-IoT Cell")
                If containsCellSheet("RFA") Then .AddItem getResByKey("RFA Cell")
                If containsASheet(ThisWorkbook, getResByKey("NR Cell")) Then .AddItem getResByKey("NR Cell")
                If containsASheet(ThisWorkbook, getResByKey("NR DU Cell")) Then .AddItem getResByKey("NR DU Cell")
                If containsASheet(ThisWorkbook, getResByKey("DCell")) Then .AddItem getResByKey("DCell")
            End If
            If isContainGsmControl() Then
                .AddItem getResByKey("GSM Logic Cell")
            End If
            If isContainUmtsControl() Then
                .AddItem getResByKey("UMTS Logic Cell")
            End If
        ElseIf neType = "UMTS" Then
            If isContainBaseStation() Then
                .AddItem getResByKey("UMTS Local Cell")
            End If
            If isContainUmtsControl() Then
                .AddItem getResByKey("UMTS Logic Cell")
            End If
        ElseIf neType = "GSM" Then
            If isContainBaseStation() Then
                .AddItem getResByKey("GSM Local Cell")
            End If
            If isContainGsmControl() Then
                .AddItem getResByKey("GSM Logic Cell")
            End If
        ElseIf neType = "LTE" Then
            .AddItem getResByKey("LTE Cell")
        End If
    End With
    
    If Me.CellTypeList.ListCount > 0 Then Me.CellTypeList.ListIndex = 0

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setCellTypeListValue," & Err.Description
End Sub

Private Sub CellTypeList_Change()
    Call initCellFilterParameters
    
    Call clearLteCellFilterParameters
    
    If Me.FddTddModeList.Enabled = True Then
        setFddTddModeListValue
    ElseIf Me.BandwidthList.Enabled = True Then
        setBandwidthListValue
    ElseIf Me.TxRxModeList.Enabled = True Then
        setTxRxModeListValue
    ElseIf Me.SAList.Enabled = True Then
        setSAListValue
    Else
        setCellTemplateListValue
    End If
End Sub

Private Sub setFddTddModeListValue()
On Error GoTo ErrorHandler
    Me.FddTddModeList.Clear
    
    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Dim fddTddModes As New Collection
    
    Dim cellType As String
    cellType = Me.CellTypeList.value
    
    Dim cellTypeValue As String
    Dim rowIdx As Integer
    With mapCellTemplate
        For rowIdx = 2 To .range("a65536").End(xlUp).row
            cellTypeValue = .Cells(rowIdx, MapCell_CellTypeCol).value
            If (cellTypeValue = cellType Or cellTypeValue = "") And .Cells(rowIdx, MapCell_NeTypeCol).value = neType Then
                Dim fddTddMode As String
                fddTddMode = .Cells(rowIdx, MapCell_FDDTDDCol).value
                If fddTddMode <> "" And Not Contains(fddTddModes, fddTddMode) Then
                    fddTddModes.Add Item:=fddTddMode, key:=fddTddMode
                    Me.FddTddModeList.AddItem (fddTddMode)
                End If
            End If
        Next
    End With
    
    If Me.FddTddModeList.ListCount > 0 Then Me.FddTddModeList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setFddTddModeListValue," & Err.Description
End Sub

Private Sub FddTddModeList_Change()
    Call enableDisableFilters
    
    If Me.BandwidthList.Enabled = True Then
        setBandwidthListValue
    ElseIf Me.TxRxModeList.Enabled = True Then
        setTxRxModeListValue
    ElseIf Me.SAList.Enabled = True Then
        setSAListValue
    Else
        setCellTemplateListValue
    End If
End Sub

Private Sub enableDisableFilters()
    Dim cellType As String, fddTddMode As String
    cellType = Me.CellTypeList.value
    fddTddMode = Me.FddTddModeList.value
    
    Me.BandwidthList.Enabled = False
    Me.BandwidthList.Clear
    Me.TxRxModeList.Enabled = False
    Me.TxRxModeList.Clear
    Me.SAList.Enabled = False
    Me.SAList.Clear
        
    If fddTddMode = "TDD" And cellType = getResByKey("LTE Cell") Then
        Me.BandwidthList.Enabled = True
        Me.TxRxModeList.Enabled = True
        Me.SAList.Enabled = True
    End If
End Sub

Private Sub setBandwidthListValue()
On Error GoTo ErrorHandler
    Me.BandwidthList.Clear

    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Dim bandwidthes As Collection
    Set bandwidthes = getCellBandwidthes(mapCellTemplate, neType, Me.CellTypeList.value)
    
    Dim bandwidth As Variant
    For Each bandwidth In bandwidthes
        Me.BandwidthList.AddItem (CStr(bandwidth))
    Next
    
    If Me.BandwidthList.ListCount > 0 Then Me.BandwidthList.ListIndex = 0

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setBandwidthListValue," & Err.Description
End Sub

Private Sub BandwidthList_Change()
    If Me.TxRxModeList.Enabled = True Then
        setTxRxModeListValue
    ElseIf Me.SAList.Enabled = True Then
        setSAListValue
    Else
        setCellTemplateListValue
    End If
End Sub

Private Sub setTxRxModeListValue()
On Error GoTo ErrorHandler
    Me.TxRxModeList.Clear

    Dim cellType As String, bandwidth As String
    cellType = Me.CellTypeList.value
    bandwidth = Me.BandwidthList.value
    
    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Dim txrxModes As Collection
    Set txrxModes = getCellTxRxModes(mapCellTemplate, neType, cellType, bandwidth)
    
    Dim txrxMode As Variant
    For Each txrxMode In txrxModes
        Me.TxRxModeList.AddItem (CStr(txrxMode))
    Next
    
    If Me.TxRxModeList.ListCount > 0 Then Me.TxRxModeList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setTxRxModeListValue," & Err.Description
End Sub

Private Sub TxRxModeList_Change()
    If Me.SAList.Enabled = True Then
        setSAListValue
    Else
        setCellTemplateListValue
    End If
End Sub

Private Sub setSAListValue()
On Error GoTo ErrorHandler
    Me.SAList.Clear

    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Dim cellType As String, bandwidth As String, txrxMode As String
    cellType = Me.CellTypeList.value
    bandwidth = Me.BandwidthList.value
    txrxMode = Me.TxRxModeList.value
    
    Dim saListValues As Collection
    Set saListValues = getSAListValues(mapCellTemplate, neType, cellType, bandwidth, txrxMode)
    
    Dim sa As Variant
    For Each sa In saListValues
        Me.SAList.AddItem (CStr(sa))
    Next
    
    If Me.SAList.ListCount > 0 Then Me.SAList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setSAListValue," & Err.Description
End Sub

Private Sub SAList_Change()
    setCellTemplateListValue
End Sub

'从「MappingCellTemplate」页获取「*Cell Pattern」列侯选值
Private Sub setCellTemplateListValue()
On Error GoTo ErrorHandler
    Dim cellType As String, fddTddMode As String, bandwidth As String
    cellType = Me.CellTypeList.value
    fddTddMode = Me.FddTddModeList.value
    bandwidth = Me.BandwidthList.value

    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Me.CellTemplateList.Clear
    Dim cellTemplates As New Collection
    
    Dim cellTemplateName As String, cellTypeValue As String, fddTddModeValue As String, bandwidthValue As String, txrxModeValue As String, saValue As String
    Dim rowIdx As Integer
    With mapCellTemplate
        If isLteCellType Then
            setLTECellTemplateListValue
        ElseIf isNRCellType Then
            setNRCellTemplateListValue
        Else
            setMRATCellTemplateListValue
        End If
    End With
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setCellTemplateListValue," & Err.Description
End Sub

Private Sub setLTECellTemplateListValue()
On Error GoTo ErrorHandler
    Dim cellType As String, fddTddMode As String, bandwidth As String, txrxMode As String, sa As String
    cellType = Me.CellTypeList.value
    fddTddMode = Me.FddTddModeList.value
    bandwidth = Me.BandwidthList.value
    txrxMode = Me.TxRxModeList.value
    sa = Me.SAList.value

    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Me.CellTemplateList.Clear
    
    Dim cellTemplates As New Collection
    
    Dim cellTemplateName As String, cellTypeValue As String, fddTddModeValue As String
    Dim bandwidthValue As String, txrxModeValue As String, saValue As String
    
    Dim matched As Boolean
    Dim rowIdx As Integer
    With mapCellTemplate
        For rowIdx = 2 To .range(getColStr(MapCell_TempNameCol) & "65535").End(xlUp).row
            cellTypeValue = .Cells(rowIdx, MapCell_CellTypeCol).value
            fddTddModeValue = .Cells(rowIdx, MapCell_FDDTDDCol).value

            If .Cells(rowIdx, MapCell_NeTypeCol).value = neType And (cellTypeValue = cellType Or cellTypeValue = "") _
                And (fddTddModeValue = fddTddMode Or fddTddModeValue = "") Then
                    matched = True
                    cellTemplateName = .Cells(rowIdx, MapCell_TempNameCol)
                    
                    If fddTddMode = "TDD" Then
                        matched = False
                        bandwidthValue = .Cells(rowIdx, MapCell_BandwidthCol).value
                        txrxModeValue = .Cells(rowIdx, MapCell_TxRxModeCol).value
                        saValue = .Cells(rowIdx, MapCell_SACol).value
                        If (bandwidthValue = bandwidth Or bandwidthValue = "") And (txrxModeValue = txrxMode Or txrxModeValue = "") _
                            And (saValue = sa Or saValue = "") Then
                                matched = True
                        End If
                    End If
                    
                    If matched And cellTemplateName <> "" And Not Contains(cellTemplates, cellTemplateName) Then
                        Me.CellTemplateList.AddItem (cellTemplateName)
                    End If
            End If
        Next
    End With
    
    If Me.CellTemplateList.ListCount > 0 Then Me.CellTemplateList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setLTECellTemplateListValue," & Err.Description
End Sub

Private Sub setNRCellTemplateListValue()
On Error GoTo ErrorHandler
    Dim cellType As String, fddTddMode As String, bandwidth As String
    cellType = Me.CellTypeList.value
    fddTddMode = Me.FddTddModeList.value
    bandwidth = Me.BandwidthList.value

    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Me.CellTemplateList.Clear
    
    Dim cellTemplates As New Collection
    
    Dim cellTemplateName As String, cellTypeValue As String, fddTddModeValue As String, bandwidthValue As String
    Dim rowIdx As Integer
    With mapCellTemplate
        For rowIdx = 2 To .range(getColStr(MapCell_TempNameCol) & "65535").End(xlUp).row
            cellTypeValue = .Cells(rowIdx, MapCell_CellTypeCol).value
            fddTddModeValue = .Cells(rowIdx, MapCell_FDDTDDCol).value
            bandwidthValue = .Cells(rowIdx, MapCell_BandwidthCol).value
            
            If .Cells(rowIdx, MapCell_NeTypeCol).value = neType And (cellTypeValue = cellType Or cellTypeValue = "") _
                And (fddTddModeValue = fddTddMode Or fddTddModeValue = "") And (bandwidthValue = bandwidth Or bandwidthValue = "" Or bandwidth = "") Then
                    cellTemplateName = .Cells(rowIdx, MapCell_TempNameCol)
                    If cellTemplateName <> "" And Not Contains(cellTemplates, cellTemplateName) Then
                        Me.CellTemplateList.AddItem (cellTemplateName)
                    End If
            End If
        Next
    End With
    
    If Me.CellTemplateList.ListCount > 0 Then Me.CellTemplateList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setNRCellTemplateListValue," & Err.Description
End Sub

Private Sub setMRATCellTemplateListValue()
On Error GoTo ErrorHandler
    Dim cellType As String
    cellType = Me.CellTypeList.value

    Dim mapCellTemplate As Worksheet
    Set mapCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Me.CellTemplateList.Clear
    
    Dim cellTemplates As New Collection
    
    Dim cellTemplateName As String, cellTypeValue As String
    Dim rowIdx As Integer
    With mapCellTemplate
        For rowIdx = 2 To .range(getColStr(MapCell_TempNameCol) & "65535").End(xlUp).row
            cellTypeValue = .Cells(rowIdx, MapCell_CellTypeCol).value
            If .Cells(rowIdx, MapCell_NeTypeCol).value = neType And (cellTypeValue = cellType Or cellTypeValue = "") Then
                cellTemplateName = .Cells(rowIdx, MapCell_TempNameCol)
                If cellTemplateName <> "" And Not Contains(cellTemplates, cellTemplateName) Then
                    Me.CellTemplateList.AddItem (cellTemplateName)
                End If
            End If
        Next
    End With
    
    If Me.CellTemplateList.ListCount > 0 Then Me.CellTemplateList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setMRATCellTemplateListValue," & Err.Description
End Sub

'「Add」按钮事件,添加模板
Private Sub addCellTemplate()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = Trim(Me.CellTemplateTextArea.text)
    
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
    
    Dim cellTypeText As String
    cellTypeText = Me.CellTypeList.value
    
    Dim firstAddr As String
    Dim targetRange As range
    With mapCellTemplate.columns(MapCell_TempNameCol)
        Set targetRange = .Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If cellTemplateConditionMatch(targetRange, neType, cellTypeText) Then
                    MsgBox templateName & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
                    Me.CellTemplateTextArea.SetFocus
                    Exit Sub
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With

    Dim insertPos As Long
    With mapCellTemplate
        insertPos = .range("a65536").End(xlUp).row + 1
        .rows(insertPos).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        .rows(insertPos).NumberFormatLocal = "@" '设置单元格格式为文本
        .Cells(insertPos, MapCell_TempNameCol).value = templateName
        .Cells(insertPos, MapCell_CellTypeCol).value = cellTypeText
        .Cells(insertPos, MapCell_NeTypeCol).value = neType
        If isLteCellType Or isNRCellType Then
            .Cells(insertPos, MapCell_BandwidthCol).value = BandwidthList.value
            .Cells(insertPos, MapCell_TxRxModeCol).value = TxRxModeList.value
            .Cells(insertPos, MapCell_FDDTDDCol).value = FddTddModeList.value
            .Cells(insertPos, MapCell_SACol).value = SAList.value
        End If
    End With
    
    Me.CellTemplateTextArea.value = ""
    Me.CellTemplateTextArea.SetFocus
    Load Me
    MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in addCellTemplate," & Err.Description
End Sub

'「Delete」按钮事件,删除模板
Private Sub deleteCellTemplate()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = Trim(Me.CellTemplateList.text)
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("cellPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim cellTypeText As String
    cellTypeText = Me.CellTypeList.value

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
                If cellTemplateConditionMatch(targetRange, neType, cellTypeText) Then
                    targetRows.Add Item:=CStr(targetRange.row), key:=CStr(targetRange.row)
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    If targetRows.count > 0 Then
        Dim rowIdx As Integer
        With mapCellTemplate
            For rowIdx = targetRows.count To 1 Step -1
                .rows(CInt(targetRows.Item(rowIdx))).Delete
            Next
        End With
        
        setCellTemplateListValue
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    Else
        MsgBox templateName & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in deleteCellTemplate," & Err.Description
End Sub

Private Function cellTemplateConditionMatch(targetRange As range, neType As String, cellTypeText As String) As Boolean
On Error GoTo ErrorHandler
    cellTemplateConditionMatch = False
    
    Dim cellTypeValue As String
    cellTypeValue = Trim(targetRange.Offset(0, MapCell_CellTypeCol - MapCell_TempNameCol).value)
    
    If cellTypeValue <> "" And cellTypeValue <> cellTypeText Then Exit Function
    
    If targetRange.Offset(0, MapCell_NeTypeCol - MapCell_TempNameCol).value <> neType Then Exit Function
    
    If isLteCellType Or isNRCellType Then
        Dim saValue As String, fddtddValue As String, txrxModeValue As String, bandwidthValue As String
        With targetRange
            If Me.SAList.Enabled = True Then
                saValue = Trim(.Offset(0, MapCell_SACol - MapCell_TempNameCol).value)
                If saValue <> "" And saValue <> Me.SAList.value Then Exit Function
            End If
            
            fddtddValue = Trim(.Offset(0, MapCell_FDDTDDCol - MapCell_TempNameCol).value)
            If fddtddValue <> "" And fddtddValue <> Me.FddTddModeList.value Then Exit Function
            
            If Me.TxRxModeList.Enabled = True Then
                txrxModeValue = Trim(.Offset(0, MapCell_TxRxModeCol - MapCell_TempNameCol).value)
                If txrxModeValue <> "" And txrxModeValue <> Me.TxRxModeList.value Then Exit Function
            End If
            
            If Me.BandwidthList.Enabled = True Then
                bandwidthValue = Trim(.Offset(0, MapCell_BandwidthCol - MapCell_TempNameCol).value)
                If bandwidthValue <> "" And bandwidthValue <> Me.BandwidthList.value Then Exit Function
            End If
        End With
    End If
    
    cellTemplateConditionMatch = True
    
    Exit Function
ErrorHandler:
    cellTemplateConditionMatch = True
    Debug.Print "some exception in cellTemplateConditionMatch, " & Err.Description
End Function

Private Function getCellBandwidthes(mapCellTemplate As Worksheet, neType As String, cellType As String) As Collection
On Error GoTo ErrorHandler
    Set getCellBandwidthes = New Collection
    
    If OpMode = OpMode_Add Then
        If cellType = getResByKey("LTE Cell") Then
            getCellBandwidthes.Add Item:="5M", key:="5M"
            getCellBandwidthes.Add Item:="10M", key:="10M"
            getCellBandwidthes.Add Item:="15M", key:="15M"
            getCellBandwidthes.Add Item:="20M", key:="20M"
        ElseIf cellType = getResByKey("NR Cell") Or cellType = getResByKey("NR DU Cell") Then
            getCellBandwidthes.Add Item:="10M", key:="10M"
            getCellBandwidthes.Add Item:="15M", key:="15M"
            getCellBandwidthes.Add Item:="20M", key:="20M"
            getCellBandwidthes.Add Item:="40M", key:="40M"
            getCellBandwidthes.Add Item:="60M", key:="60M"
            getCellBandwidthes.Add Item:="80M", key:="80M"
            getCellBandwidthes.Add Item:="100M", key:="100M"
            getCellBandwidthes.Add Item:="200M", key:="200M"
        End If
    End If
    
    Dim bandwidth As String
    Dim rowIdx As Integer
    With mapCellTemplate
        For rowIdx = 2 To .range(getColStr(MapCell_BandwidthCol) & "65535").End(xlUp).row
            If .Cells(rowIdx, MapCell_NeTypeCol).value = neType And .Cells(rowIdx, MapCell_CellTypeCol).value = cellType Then
                bandwidth = .Cells(rowIdx, MapCell_BandwidthCol).value
                If bandwidth <> "" And Not Contains(getCellBandwidthes, bandwidth) Then getCellBandwidthes.Add Item:=bandwidth, key:=bandwidth
            End If
        Next
    End With
        
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getCellBandwidthes, " & Err.Description
End Function

Private Function getCellTxRxModes(mapCellTemplate As Worksheet, neType As String, cellType As String, bandwidth As String) As Collection
On Error GoTo ErrorHandler
    Set getCellTxRxModes = New Collection
    
    If OpMode = OpMode_Add And cellType = getResByKey("LTE Cell") Then
        getCellTxRxModes.Add Item:="1T1R", key:="1T1R"
        getCellTxRxModes.Add Item:="2T2R", key:="2T2R"
        getCellTxRxModes.Add Item:="4T4R", key:="4T4R"
        getCellTxRxModes.Add Item:="8T8R", key:="8T8R"
        getCellTxRxModes.Add Item:="64T64R", key:="64T64R"
    End If
    
    Dim txrxMode As String
    Dim rowIdx As Integer
    With mapCellTemplate
        For rowIdx = 2 To .range(getColStr(MapCell_TxRxModeCol) & "65535").End(xlUp).row
            If .Cells(rowIdx, MapCell_NeTypeCol).value = neType And .Cells(rowIdx, MapCell_CellTypeCol).value = cellType Then
                If OpMode = OpMode_Add Or (OpMode = OpMode_Del And (.Cells(rowIdx, MapCell_BandwidthCol).value = bandwidth Or bandwidth = "")) Then
                    txrxMode = .Cells(rowIdx, MapCell_TxRxModeCol).value
                    If txrxMode <> "" And Not Contains(getCellTxRxModes, txrxMode) Then getCellTxRxModes.Add Item:=txrxMode, key:=txrxMode
                End If
            End If
        Next
    End With
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getCellTxRxModes, " & Err.Description
End Function

Private Function getSAListValues(mapCellTemplate As Worksheet, neType As String, cellType As String, bandwidth As String, txrxMode As String) As Collection
On Error GoTo ErrorHandler
    Set getSAListValues = New Collection

    If OpMode = OpMode_Add And cellType = getResByKey("LTE Cell") Then
        getSAListValues.Add Item:="SA0", key:="SA0"
        getSAListValues.Add Item:="SA1", key:="SA1"
        getSAListValues.Add Item:="SA2", key:="SA2"
        getSAListValues.Add Item:="SA3", key:="SA3"
        getSAListValues.Add Item:="SA4", key:="SA4"
        getSAListValues.Add Item:="SA5", key:="SA5"
        getSAListValues.Add Item:="SA6", key:="SA6"
    End If
    
    Dim sa As String
    Dim rowIdx As Integer
    With mapCellTemplate
        For rowIdx = 2 To .range(getColStr(MapCell_SACol) & "65535").End(xlUp).row
            If .Cells(rowIdx, MapCell_NeTypeCol).value = neType And .Cells(rowIdx, MapCell_CellTypeCol).value = cellType Then
                If OpMode = OpMode_Add Or _
                    (OpMode = OpMode_Del And _
                        (.Cells(rowIdx, MapCell_BandwidthCol).value = bandwidth Or bandwidth = "") And _
                        (.Cells(rowIdx, MapCell_TxRxModeCol).value = txrxMode Or txrxMode = "")) Then
                            sa = .Cells(rowIdx, MapCell_SACol).value
                            If sa <> "" And Not Contains(getSAListValues, sa) Then getSAListValues.Add Item:=sa, key:=sa
                End If
            End If
        Next
    End With
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getSAListValues, " & Err.Description
End Function

Private Function isLteCellType() As Boolean
    If Me.CellTypeList.value = getResByKey("LTE Cell") Then
        isLteCellType = True
    Else
        isLteCellType = False
    End If
End Function

Private Function isNRCellType() As Boolean
    If Me.CellTypeList.value = getResByKey("NR Cell") Or Me.CellTypeList.value = getResByKey("NR DU Cell") Then
        isNRCellType = True
    Else
        isNRCellType = False
    End If
End Function

Private Sub initCellFilterParameters()
On Error GoTo ErrorHandler
    If isLteCellType Then
        Call displayCellFilterParameters(True, True, True, True)
        Call adjustCellTemplateForL
    ElseIf isNRCellType Then
        Call displayCellFilterParameters(True, False, False, False)
        Call adjustCellTemplateForNR
    Else
        Call displayCellFilterParameters(False, False, False, False)
        Call adjustCellTemplateForMRAT
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in initCellFilterParameters, " & Err.Description
End Sub

Private Sub displayCellFilterParameters(ByVal fddTddFlag As Boolean, ByVal bandwidthFlag As Boolean, ByVal txrxModeFlag As Boolean, ByVal saFlag As Boolean)
    Me.CellFDDTDDModeLabel.Visible = fddTddFlag
    Me.FddTddModeList.Enabled = fddTddFlag
    Me.FddTddModeList.Visible = fddTddFlag
    
    Me.BandWidthLabel.Visible = bandwidthFlag
    Me.BandwidthList.Enabled = bandwidthFlag
    Me.BandwidthList.Visible = bandwidthFlag
    
    Me.TxRxModeLabel.Visible = txrxModeFlag
    Me.TxRxModeList.Enabled = txrxModeFlag
    Me.TxRxModeList.Visible = txrxModeFlag
    
    Me.CellSALabel.Visible = saFlag
    Me.SAList.Enabled = saFlag
    Me.SAList.Visible = saFlag
End Sub

Private Sub adjustCellTemplateForMRAT()
    Me.CellTemplateLabel.Top = 92
    Me.CellTemplateList.Top = 90
    Me.CellTemplateTextArea.Top = 90
    
    Me.AddCellButton.Top = 138
    Me.CancelCellButton.Top = 138
End Sub

Private Sub adjustCellTemplateForL()
    Me.CellFDDTDDModeLabel.Top = 65
    Me.FddTddModeList.Top = 65
    
    Me.CellTemplateLabel.Top = 165
    Me.CellTemplateList.Top = 162
    Me.CellTemplateTextArea.Top = 162
    
    Me.AddCellButton.Top = 186
    Me.CancelCellButton.Top = 186
End Sub

Private Sub adjustCellTemplateForNR()
    Me.CellFDDTDDModeLabel.Top = 80
    Me.FddTddModeList.Top = 80
    
    Me.CellTemplateLabel.Top = 118
    Me.CellTemplateList.Top = 116
    Me.CellTemplateTextArea.Top = 116
    
    Me.AddCellButton.Top = 156
    Me.CancelCellButton.Top = 156
End Sub

Private Sub clearLteCellFilterParameters()
    Me.FddTddModeList.Clear
    Me.BandwidthList.Clear
    Me.TxRxModeList.Clear
    Me.SAList.Clear
End Sub

Private Function containsCellSheet(ByRef neType As String) As Boolean
    Dim cellSheetName As String
    cellSheetName = neType & getResByKey("Cell")
    If containsASheet(ThisWorkbook, cellSheetName) Then
        containsCellSheet = True
    Else
        containsCellSheet = False
    End If
End Function



'=============================Radio========================
'添加和删除「MappingRadioTemplate」页的模板
Private Sub AddRadioTemplateRadio_Click()
    Me.RadioTemplateTextArea.Visible = True
    Me.RadioTemplateList.Visible = False
    Me.AddRadioButton.Caption = getResByKey("Add")
    OpMode = OpMode_Add
    setRadioTemplateListValue
End Sub

Private Sub DeleteRadioTemplateRadio_Click()
    Me.RadioTemplateList.Visible = True
    Me.RadioTemplateTextArea.Visible = False
    Me.AddRadioButton.Caption = getResByKey("Delete")
    OpMode = OpMode_Del
    setRadioTemplateListValue
End Sub

'提交进行Add/Delete操作
Private Sub AddRadioButton_Click()
    If OpMode = OpMode_Add Then
        addRadioTemplate
    Else
        deleteRadioTemplate
    End If
    Call refreshCell
End Sub

Private Sub CancelRadioButton_Click()
    Unload Me
End Sub

Private Function containsRadioTemplate(ByRef neType As String, ByRef mappingDefSheet As Worksheet) As Boolean
    If findCertainValRowNumberByTwoKeys(mappingDefSheet, "D", neType, "E", "RadioTemplateName") <> -1 Then
        containsRadioTemplate = True
    Else
        containsRadioTemplate = False
    End If
End Function

Private Sub setRadioTypeListValue()
On Error GoTo ErrorHandler
    Me.RadioTypeList.Clear
    If neType = "MRAT" Then
        With Me.RadioTypeList
            Dim mappingDefSheet As Worksheet
            Set mappingDefSheet = ThisWorkbook.Worksheets("MAPPING DEF")
            If containsRadioTemplate("GBTSFUNCTION", mappingDefSheet) Then .AddItem getResByKey("GSM Radio Template")
            If containsRadioTemplate("NODEBFUNCTION", mappingDefSheet) Then .AddItem getResByKey("UMTS Radio Template")
            If containsRadioTemplate("eNodeBFunction", mappingDefSheet) Then .AddItem getResByKey("LTE Radio Template")
            If containsRadioTemplate("NBBSFunction", mappingDefSheet) Then .AddItem getResByKey("NB-IoT Radio Template")
            If containsRadioTemplate("gNodeBFunction", mappingDefSheet) Then .AddItem getResByKey("NR Radio Template")
            If containsRadioTemplate("DsaFunction", mappingDefSheet) Then .AddItem getResByKey("DSA Radio Template")
        End With
    ElseIf neType = "UMTS" Then
        Me.RadioTypeList.AddItem getResByKey("UMTS Radio Template")
    ElseIf neType = "GSM" Then
        Me.RadioTypeList.AddItem getResByKey("GSM Radio Template")
    ElseIf neType = "LTE" Then
        Me.RadioTypeList.AddItem getResByKey("LTE Radio Template")
    End If
    
    If Me.RadioTypeList.ListCount > 0 Then Me.RadioTypeList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setRadioTypeListValue, " & Err.Description
End Sub

Private Sub RadioTypeList_Change()
    setRadioTemplateListValue
End Sub

Private Sub setRadioTemplateListValue()
On Error GoTo ErrorHandler
    Dim radioType As String
    radioType = Me.RadioTypeList.value

    Me.RadioTemplateList.Clear
    
    Dim mapRadioTemplate As Worksheet
    Set mapRadioTemplate = ThisWorkbook.Worksheets("MappingRadioTemplate")
    
    Dim radioTemplates As New Collection
    
    Dim radioTemplate As String
    Dim rowIdx As Integer
    With mapRadioTemplate
        For rowIdx = 2 To .range("a65536").End(xlUp).row
            If (.Cells(rowIdx, MapRadio_RadioTypeCol).value = radioType Or Trim(.Cells(rowIdx, MapRadio_RadioTypeCol).value) = "") _
                And .Cells(rowIdx, MapRadio_NeTypeCol).value = neType Then
                    radioTemplate = .Cells(rowIdx, MapRadio_TempNameCol).value
                    If radioTemplate <> "" And Not Contains(radioTemplates, radioTemplate) Then Me.RadioTemplateList.AddItem (radioTemplate)
            End If
        Next
    End With
    
    If Me.RadioTemplateList.ListCount > 0 Then Me.RadioTemplateList.ListIndex = 0
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setRadioTemplateListValue, " & Err.Description
End Sub

'「Add」按钮事件,添加模板
Private Sub addRadioTemplate()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = Trim(Me.RadioTemplateTextArea.text)
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    If InStr(templateName, ",") <> 0 Then
        MsgBox templateName & getResByKey("invalidTemplateName_Comma"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim radioType As String
    radioType = Me.RadioTypeList.value
    
    Dim mapRadioTemplate As Worksheet
    Set mapRadioTemplate = ThisWorkbook.Worksheets("MappingRadioTemplate")
    
    Dim targetRange As range
    Dim firstAddr As String
    With mapRadioTemplate.columns(MapRadio_TempNameCol)
        Set targetRange = .Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If targetRange.Offset(0, MapRadio_RadioTypeCol - MapRadio_TempNameCol).value = radioType _
                    And targetRange.Offset(0, MapRadio_NeTypeCol - MapRadio_TempNameCol).value = neType Then
                        MsgBox templateName & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
                        Me.RadioTemplateTextArea.SetFocus
                        Exit Sub
                End If
                targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    Dim insertPos As Long
    With mapRadioTemplate
        insertPos = .range("a65536").End(xlUp).row + 1
        .rows(insertPos).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        .rows(insertPos).NumberFormatLocal = "@" '设置单元格格式为文本
        .Cells(insertPos, MapRadio_RadioTypeCol).value = radioType
        .Cells(insertPos, MapRadio_NeTypeCol).value = neType
        .Cells(insertPos, MapRadio_TempNameCol).value = templateName
    End With
    
    Me.RadioTemplateTextArea.value = ""
    Load Me
    MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in addRadioTemplate, " & Err.Description
End Sub

'「Delete」按钮事件,删除模板
Private Sub deleteRadioTemplate()
On Error GoTo ErrorHandler
    Dim templateName As String
    templateName = Trim(Me.RadioTemplateList.text)
    
    If templateName = "" Then
        MsgBox templateName & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        Exit Sub
    End If
    
    Dim radioType As String
    radioType = Me.RadioTypeList.value
    
    Dim mapRadioTemplate As Worksheet
    Set mapRadioTemplate = ThisWorkbook.Worksheets("MappingRadioTemplate")
    
    Dim targetRows As New Collection
    
    Dim targetRange As range
    Dim firstAddr As String
    With mapRadioTemplate.columns(MapRadio_TempNameCol)
        Set targetRange = .Find(templateName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If targetRange.Offset(0, MapRadio_RadioTypeCol - MapRadio_TempNameCol).value = radioType _
                    And targetRange.Offset(0, MapRadio_NeTypeCol - MapRadio_TempNameCol).value = neType Then
                        targetRows.Add Item:=CStr(targetRange.row), key:=CStr(targetRange.row)
                End If
                targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    If targetRows.count > 0 Then
        Dim rowIdx As Integer
        With mapRadioTemplate
            For rowIdx = targetRows.count To 1 Step -1
                .rows(CInt(targetRows.Item(rowIdx))).Delete
            Next
        End With
        
        setRadioTemplateListValue
        Me.RadioTemplateList.SetFocus
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    Else
        MsgBox templateName & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in deleteRadioTemplate, " & Err.Description
End Sub


Private Sub refreshCell()
    Dim rangeHis As range
    Dim row, columen As Long
    Set rangeHis = Selection
    ActiveSheet.Cells(Selection.row + 1, Selection.column).Select
    rangeHis.Select
End Sub










