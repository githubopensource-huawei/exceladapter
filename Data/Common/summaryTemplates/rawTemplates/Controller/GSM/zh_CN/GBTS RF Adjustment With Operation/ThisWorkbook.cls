VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Activate()
    On Error Resume Next
    initResource
    buildControlRelations
    hyperLintFlag = True
    Set currentSheet = ThisWorkbook.ActiveSheet
    Call initBoardStyleMap
    Call InsertUserToolBar
    Call initCapacityExpansionToolBar(currentSheet)
    Call initAujustAntnPortToolBar(currentSheet)
    Call initMenuStatus(currentSheet)
    
    Call addHyperlinksForSheets(currentSheet) '给小区页，单板样式页，传输页增加超链接
End Sub


Private Sub Workbook_Deactivate()
    On Error Resume Next
    
    Call DeleteUserToolBar '增删小区等按钮
    
    Call deleteCapacityExpansionToolBar
    Call CapacityCellSub.deleteCellBar
    Call CapacityCellSub.deleteTempBar
    destroyMenuStatus
    Application.StatusBar = False
End Sub

'分支参数控制响应函数
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As Range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Dim sheetName As String
    Set sheet = sh
    sheetName = sheet.name
    
        Dim copyPasteGuard As New CCopyPasteGuard
    copyPasteGuard.init
        
    If ThisWorkbook.name <> ActiveWorkbook.name Or sheet.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    If isControlDefSheetExist = False Then
        Exit Sub
    End If
    
    If isBoardStyleSheet(sheet) Then
        Call boardStyleSheetControl(sheet, Target) '单板样式页签分支控制
        Call addReferenceRangeHyperlinks_SheetChange(sheet, Target) '添加引用单元格超链接
        Exit Sub
    End If
    
    If sheetName = GetMainSheetName() Then
        Call addBoardStyleNameHyperlinks_SheetChange(sheet, Target) '添加传输页单板样式名称超链接
    ElseIf isCellSheet(sheetName) Then
        Call addRxuAntNoHyperlinks_SheetChange(sheet, Target) '给小区页签的RXU Ant No.添加超链接
    End If
    
    If sheetName = getResByKey("Temp Sheet") Then
        Call changeRangeColor(sheet, Target)
        Call BranchControlForTempSht(sheet, Target)
    End If
    
    Dim sheetType As String
    Dim currentNeType As String
    sheetType = getSheetType(sheetName)
   If (isGsmCellSheet(sheetName) = False) And (isBaseStationSheet(sheetName) = False) And (isGTRXSheet(sheetName) = False) And (sheetName <> "MappingSiteTemplate") And (sheetName <> "ProductType") _
        And (sheetName <> "MappingCellTemplate") And (sheetName <> "MappingRadioTemplate") _
        And (sheetName <> "MAPPING DEF") And (sheetName <> "SHEET DEF") And (sheetName <> "CONTROL DEF") And (sheetName <> "RELATION DEF") _
        And (sheetName <> "Help for USB Parameter") And (sheetName <> getResByKey("help")) _
        And (sheetName <> "IPRouteMap") And (sheetName <> getResByKey("Cover")) And (sheetName <> "COMMON") _
        And (sheetName <> "Qos") And (sheetName <> "USB Parameter for Sites") And (sheetName <> "SummaryRes") _
        And (sheetName <> getResByKey("Temp Sheet")) And (sheetType <> "") _
        And Not isGsmCellSheet(sheet.name) Then
        If sheetName = GetMainSheetName() And getNeType() <> "USU" And (Not (cellRange.row = 3)) Then
            Call TransPortSheetChange(sheet, Target)
        End If
        If Target.count Mod 256 = 0 Then
            Exit Sub
        End If
        For Each cellRange In Target
            If sheetName = getResByKey("Comm Data") Then
                If cellRange.Interior.colorIndex = 34 Or cellRange.Interior.colorIndex = 40 Or cellRange.Borders.LineStyle = xlLineStyleNone Then
                    Exit Sub
                End If
            Else
                If (cellRange.row < 3) Or (cellRange.column > sheet.Range("XFD2").End(xlToLeft).column) Then
                    Exit Sub
                End If
            End If
            
            If (Not (sheetName = GetMainSheetName() And cellRange.row = 3)) And Ensure_Gray_Cell(cellRange) = False Then
                'isConAttr表明是否是主控参数
'                If UBound(Split(cellRange.value, "\")) <> 2 And cellRange.Hyperlinks.count = 1 Then
'                    cellRange.Hyperlinks.Delete
'                End If
                isConAttr = False
                
               If (Check_Value_Validation(sheet, cellRange, isConAttr, contRel, currentNeType) = 1) And (isConAttr = True) Then
                    Call Execute_Branch_Control(sheet, cellRange, contRel, currentNeType)
                End If
                
                '校验是否是引用参数
                'Call CheckRefedCellValidation(sheet, cellRange)
            End If
        Next cellRange
    End If
End Sub
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    On Error GoTo ErrorHandler
    If ThisWorkbook.name <> ActiveWorkbook.name Or sh.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    Dim sheet As Worksheet
    Set sheet = sh
    'If addBoardStyleMoiInProcess Is Nothing Then
    If inAddProcessFlag = False Then
        Call resetAddBoardStyleMoiInfo(sheet)
    Else
        Exit Sub
    End If
    
    If InAdjustAntnPort = True Then
        Exit Sub
    End If
    
    Call InsertUserToolBar
    Call initAujustAntnPortToolBar(sheet)
    Call initCapacityExpansionToolBar(sheet)
    Call initMenuStatus(sheet)
    
    Call addHyperlinksForSheets(sheet) '给小区页，单板样式页，传输页增加超链接
    Exit Sub
ErrorHandler:
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    If Not isIubStyleWorkBook() And Not ThisWorkbook.Worksheets("SummaryRes").Visible Then
'        AddLink
'    End If
    If inAddProcessFlag = True Then
        Call popUpWbCannotSaveMsgbox
        Cancel = True
        SaveAsUI = True
    End If
    If InAdjustAntnPort = True Then
        Call popUpWbCannotSaveMsgbox
        Cancel = True
        SaveAsUI = True
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If inAddProcessFlag = True Then
        Call popUpWbCannotSaveMsgbox
        Cancel = True
    End If
    If InAdjustAntnPort = True Then
        Call popUpTempSheetCannotSaveMsgbox
        Cancel = True
    End If
End Sub

Private Sub Workbook_SheetDeactivate(ByVal sh As Object)
    On Error GoTo ErrorHandler
    If inAddProcessFlag = True And sh.name = currentSheet.name Then
        Call popUpSheetCannotChangeMsgbox
    End If
    If InAdjustAntnPort = True And sh.name = getResByKey("Temp Sheet") Then
        Call popUpTempSheetCannotChangeMsgbox
    End If
    Exit Sub
ErrorHandler:
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler
    'currentCellValue = ActiveCell.value
    Dim sht As Worksheet
    Set sht = sh
    
    If ThisWorkbook.name <> ActiveWorkbook.name Or sht.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    
    If sheetSelectionShouldCheck(sht) Then
        Call currentParameterBranchCheck(sht, Target) '当前字段分支校验
    End If
    
    If isBoardStyleSheet(sht) Then
        Call boardStyleSelectionChange(sht, Target)
        Call rowControl(sht, Target)
    End If
    
    If sht.name = GetMainSheetName() Then
        Call genBaseStationBoardStyleRef(sht, Target)
    ElseIf isCellSheet(sht.name) And Not IsGBTSTemplate() Then
        Call cellSheetSelectionEvent(sht, Target)
    End If
    
'    If isIubStyleWorkSheet(sht.name) Then
'        Call iubStyleSheetSelectionChange(sht, Target)
'    End If
'
'     If sht.name = GetMainSheetName() And getNeType() <> "USU" And (Not (isIubStyleWorkBook() And Target.row = 3)) Then
'        Call iubStyleWorkBookSelectionChange(sht, Target)
'    End If
    Exit Sub
ErrorHandler:
End Sub




