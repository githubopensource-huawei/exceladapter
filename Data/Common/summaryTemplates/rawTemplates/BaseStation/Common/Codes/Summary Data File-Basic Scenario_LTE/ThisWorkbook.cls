VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public g_SelectionRow As Long
Public g_SelectionCol As Long
Public g_NBIOTFlag As Boolean

Private Sub Workbook_Activate()
    On Error Resume Next
    initResource
    InsertUserToolBar
    buildControlRelations
    Call buildMappingNumberRelations
    initAddRef
    initInnerValideDef
    loadInnerPositions
    
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.ActiveSheet
    Call initMenuStatus(sheet)
    Call changeIubSheetDeleteButtonFunction(sheet)
    Call setFileType("0")
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
    Dim sheet As Worksheet
    Set sheet = sh
    Call initToolBar(sheet.name)
    Call initMenuStatus(sheet)
    Call changeIubSheetDeleteButtonFunction(sheet)
End Sub

'分支参数控制响应函数
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Set sheet = sh

    If ThisWorkbook.name <> ActiveWorkbook.name Or sheet.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    If sheet.name = getResByKey("autoDeployment") Then
        Call AutoDeploySheetChange(sheet, target)
        Exit Sub
    End If
    If getLLDModelFlag <> "" Then
        Exit Sub
    End If
    If isControlDefSheetExist = False Then
        Exit Sub
    End If
    
     If referenceMessages Is Nothing Then
      Set referenceMessages = New CReferenceMessage
      Call referenceMessages.generateCurrentReferenceStr
    End If
    
    If isIubStyleWorkSheet(sheet.name) Then
        Call IubStyleControl(sheet, target)
        Call referenceMessages.refreshReferenceMess(sheet, target.row)
        Exit Sub
    End If
    
    Dim sheetType As String
    Dim currentNeType As String
    sheetType = getSheetType(sheet.name)
    If (sheet.name <> "MappingSiteTemplate") And (sheet.name <> "ProductType") _
        And (sheet.name <> "MappingCellTemplate") And (sheet.name <> "MappingRadioTemplate") _
        And (sheet.name <> "MAPPING DEF") And (sheet.name <> "SHEET DEF") And (sheet.name <> "CONTROL DEF") _
        And (sheet.name <> "Help for USB Parameter") And (sheet.name <> getResByKey("help")) _
        And (sheet.name <> "IPRouteMap") And (sheet.name <> getResByKey("Cover")) And (sheet.name <> "COMMON") _
        And (sheet.name <> "Qos") And (sheet.name <> "USB Parameter for Sites") And (sheet.name <> "SummaryRes") _
        And (sheetType <> "") Then
        If sheetType = "LIST" Or sheetType = "PATTERN" Then
            Call setLinkRange(sh, target)
        End If
        If sheet.name = GetMainSheetName() And (Not (isIubStyleWorkBook() And target.row = 3)) Then
            Call transportSheetChange(sheet, target)
        End If
        If target.count Mod 256 = 0 Then
            Exit Sub
        End If
        For Each cellRange In target
            If sheet.name = getResByKey("Comm Data") Then
                If findAttrName(Trim(cellRange.value)) = True Or findGroupName(Trim(cellRange.value)) = True Or cellRange.Borders.LineStyle = xlLineStyleNone Then
                    Exit Sub
                End If
            Else
                If (cellRange.row < 3) Or (cellRange.column > sheet.range("IV2").End(xlToLeft).column) Then
                    Exit Sub
                End If
            End If
            
            If (Not (sheet.name = GetMainSheetName() And isIubStyleWorkBook() And cellRange.row = 3)) And Ensure_Gray_Cell(cellRange) = False Then
                'isConAttr表明是否是主控参数
                If UBound(Split(cellRange.value, "\")) <> 2 And cellRange.Hyperlinks.count = 1 Then
                    cellRange.Hyperlinks.Delete
                End If
                isConAttr = False
                
               If (Check_Value_Validation(sheet, cellRange, isConAttr, contRel, currentNeType) = 1) And (isConAttr = True) Then
                    Call Execute_Branch_Control(sheet, cellRange, contRel, currentNeType)
                End If
                Call iubStyleBaseStationCheckRang(sheet, cellRange)
                '校验是否是引用参数
                Call CheckRefedCellValidation(sheet, cellRange)
            End If
        Next cellRange
    End If
End Sub
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    On Error Resume Next
'    Call HideToolBar
'End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    DeleteUserToolBar
    destroyMenuStatus
    Application.StatusBar = False
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    'Call InsertUserToolBar
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal target As range)
    On Error GoTo ErrorHandler
    Dim sht As New Worksheet
    Set sht = sh
    If isRefreshing() Or ThisWorkbook.name <> ActiveWorkbook.name Or sht.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    
    Dim copyPasteGuard As New CCopyPasteGuard
    copyPasteGuard.init
    
    If sheetSelectionShouldCheck(sht) Then
        Call currentParameterBranchCheck(sht, target) '当前参数的分支校验
    End If
    
    If isIubStyleWorkSheet(sht.name) Then
        Call iubStyleSheetSelectionChange(sht, target)
        Call referenceMessages.refreshReferenceMess(sht, target.row)
        Exit Sub
    End If
    If sht.name = GetMainSheetName() Then
        If target.row < 3 Or (isIubStyleWorkBook And target.row = 3) Then Exit Sub
        Call transportSheetSelectionChange(sht, target)
    ElseIf isCellSheet(sht.name) Then
        Call Cell_Worksheet_SelectionChange(sht, target)
    End If
    If (sht.name = "USB Parameter for Sites") Then
        g_SelectionRow = target.Cells.row
        g_SelectionCol = target.Cells.column
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in Workbook_SheetSelectionChange, " & Err.Description
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not Worksheets("SummaryRes").Visible Then
        addLink4NormalShts
        If Not isIubStyleWorkBook Then deleteStyles
    End If
End Sub

Private Sub IubStyleControl(ByVal sh As Object, ByVal target As range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As range
    
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Set sheet = sh
    If target.count Mod 256 = 0 Then
        Exit Sub
    End If
    
    Dim currentNeType As String
    
    For Each cellRange In target
        If isAttrRow_IUB(sheet, cellRange.row) = True Or cellRange.Borders.LineStyle = xlLineStyleNone Or cellRange.column <= 1 Then
            Exit Sub
        End If
        If Ensure_Gray_Cell(cellRange) = False Then
            'isConAttr表明是否是主控参数
            If UBound(Split(cellRange.value, "\")) <> 2 And cellRange.Hyperlinks.count = 1 Then
                cellRange.Hyperlinks.Delete
            End If
            isConAttr = False
           If (Check_Value_Validation(sheet, cellRange, isConAttr, contRel, currentNeType) = 1) And (isConAttr = True) Then
                Call Execute_Branch_Control(sheet, cellRange, contRel, currentNeType)
            End If
        End If
    Next cellRange
End Sub

Private Sub deleteStyles()
    On Error Resume Next
    Dim s As Style
    For Each s In ThisWorkbook.Styles
        If Not s.BuiltIn Then
            Debug.Print s.name
            s.Delete
        End If
    Next
End Sub


