VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Activate()
    On Error Resume Next
    initResource
    buildControlRelations
    Call buildMappingNumberRelations
    InsertUserToolBar
    initAddRef
    initInnerValideDef
    
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.ActiveSheet
    Call initMenuStatus(sheet)
    Call changeIubSheetDeleteButtonFunction(sheet)
    Call setFileType("1")
End Sub


Private Sub Workbook_Deactivate()
    On Error Resume Next
    DeleteUserToolBar
    destroyMenuStatus
    Application.StatusBar = False
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
        
        Dim copyPasteGuard As New CCopyPasteGuard
    copyPasteGuard.init
        
    If isRefreshing() Or ThisWorkbook.name <> ActiveWorkbook.name Or sheet.Visible <> xlSheetVisible Then
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
        Call IubStyleControl(sh, target)
        Call referenceMessages.refreshReferenceMess(sheet, target.row)
        Exit Sub
    End If
    
    Dim currentNeType As String
    If (sheet.name <> "MappingSiteTemplate") And (sheet.name <> "ProductType") _
        And (sheet.name <> "MappingCellTemplate") And (sheet.name <> "MappingRadioTemplate") _
        And (sheet.name <> "MAPPING DEF") And (sheet.name <> "SHEET DEF") And (sheet.name <> "CONTROL DEF") _
        And (sheet.name <> "Help for USB Parameter") And (sheet.name <> getResByKey("help")) _
        And (sheet.name <> "IPRouteMap") And (sheet.name <> getResByKey("Cover")) And (sheet.name <> "COMMON") _
        And (sheet.name <> "Qos") And (sheet.name <> "USB Parameter for Sites") And (sheet.name <> "SummaryRes") Then
        If target.count Mod 256 = 0 Then
            Exit Sub
        End If
        For Each cellRange In target
            If sheet.name = getResByKey("Comm Data") Then
                If cellRange.Interior.colorIndex = 34 Or cellRange.Interior.colorIndex = 40 Or cellRange.Borders.LineStyle = xlLineStyleNone Then
                    Exit Sub
                End If
            Else
                If (cellRange.row < 3) Or (cellRange.column > sheet.range("XFD2").End(xlToLeft).column) Then
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
            End If
        Next cellRange
    End If
End Sub
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    Dim sheet As Worksheet
    Set sheet = sh
    Call initToolBar(sheet.name)
    Call InsertUserToolBar
    Call initMenuStatus(sheet)
    Call changeIubSheetDeleteButtonFunction(sheet)
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not isIubStyleWorkBook() And Not ThisWorkbook.Worksheets("SummaryRes").Visible Then
        AddLink
    End If
End Sub

'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    On Error Resume Next
'    Call HideToolBar
'End Sub


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
        If cellRange.Interior.colorIndex = 34 Or cellRange.Interior.colorIndex = 40 Or cellRange.Borders.LineStyle = xlLineStyleNone Or cellRange.column <= 1 Then
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

Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal target As range)
    On Error GoTo ErrorHandler
    Dim sht As Worksheet
    Set sht = sh
    
    If isRefreshing() Or ThisWorkbook.name <> ActiveWorkbook.name Or sht.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    
    If sheetSelectionShouldCheck(sht) Then
        Call currentParameterBranchCheck(sht, target) '当前参数的分支校验
    End If

    If isIubStyleWorkSheet(sht.name) Then
        Call iubStyleSheetSelectionChange(sht, target)
        Call referenceMessages.refreshReferenceMess(sht, target.row)
    End If
    
     If sht.name = GetMainSheetName() And getNeType() <> "USU" And (Not (isIubStyleWorkBook() And target.row < 4)) Then
        Call iubStyleWorkBookSelectionChange(sht, target)
    End If
    Exit Sub
ErrorHandler:
End Sub


