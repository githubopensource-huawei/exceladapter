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
    initLanguage
    buildControlRelations
    Call buildMappingNumberRelations

    Set g_CurrentSheet = ThisWorkbook.ActiveSheet
    
    Call adjustColumnPosition

    Call addHyperlinksForSheets(g_CurrentSheet) '给单板样式页，传输页增加超链接
    
    If isCellSheet(g_CurrentSheet) Then
        Set g_CellSheet = g_CurrentSheet
        
        If isUMTS Then
            Set g_ExpParaSheet = ThisWorkbook.Worksheets("ExposeUMTSCellParas")
        Else
            Set g_ExpParaSheet = ThisWorkbook.Worksheets("ExposeLTECellParas")
        End If
    End If

    If isBoardStyleSheet(g_CurrentSheet) Then
        Call setBgColorWithGray_i(g_CurrentSheet)
    End If
    insertUserToolBar
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
    On Error Resume Next
    deleteUserToolBar
    
    Set g_CurrentSheet = sh
    
    If isCellSheet(g_CurrentSheet) Then
        Set g_CellSheet = g_CurrentSheet
        
        If isUMTS Then
            Set g_ExpParaSheet = ThisWorkbook.Worksheets("ExposeUMTSCellParas")
        Else
            Set g_ExpParaSheet = ThisWorkbook.Worksheets("ExposeLTECellParas")
        End If
    End If

    insertUserToolBar
    
    Call addHyperlinksForSheets(sh) '给小区页，单板样式页，传输页增加超链接
    
    If isBoardStyleSheet(sh) Then
        Call setBgColorWithGray_i(sh)
        insertAndDeleteControl (False)
    Else
        insertAndDeleteControl (True)
    End If
ErrorHandle:
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    deleteUserToolBar
ErrorHandler:
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal target As range, cancel As Boolean)
    On Error GoTo ErrorHandler
    If rowIsBlank(sh, target.row) Then Exit Sub
    
    If target.Interior.colorIndex <> GRAY_COLOR Then
        Set g_CurrentSheet = sh
        Set g_CurrentRange = target
        If splitSectorSelected(sh, target) Then
            SectorInfo.Show
        ElseIf splitFreqSelected(sh, target) Then
            freqInfo.Show
        ElseIf cellAntSelected(sh, target) Then
            CellAntInfo.Show
        End If
    End If

    Exit Sub
ErrorHandler:
End Sub

'分支参数控制响应函数
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Dim sheetName As String
    Set sheet = sh
    sheetName = sheet.name

    If ThisWorkbook.name <> ActiveWorkbook.name Or sheet.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    If isControlDefSheetExist = False Then
        Exit Sub
    End If
    
    If isBoardStyleSheet(sheet) And Ensure_Gray_Cell(target) = False Then
        Call boardStyleSheetControl(sheet, target) '单板样式页签分支控制
        Application.EnableEvents = False '避免更新boardNo之后，又触发sheetchange而陷入死循环
        Call boardStyleSheetChange(sheet, target) 'update board no.
        Application.EnableEvents = True
        Exit Sub
    End If
    
    Dim sheetType As String
    Dim currentNeType As String
    sheetType = getSheetType(sheetName)
    
    If sheetName = GetMainSheetName() Or sheetName = getResByKey("UMTSCellSheet") Or sheetName = getResByKey("LTECellSheet") Or sheetName = getResByKey("LTECellMMSheet") Then
        For Each cellRange In target
            If (cellRange.row < 3) Or (cellRange.column > sheet.range("IV2").End(xlToLeft).column) Then
                Exit Sub
            End If
            
            If Ensure_Gray_Cell(cellRange) = False Then
                'isConAttr表明是否是主控参数
                isConAttr = False
                
               If (Check_Value_Validation(sheet, cellRange, isConAttr, contRel, currentNeType) = 1) And (isConAttr = True) Then
                    Call Execute_Branch_Control(sheet, cellRange, contRel, currentNeType)
                End If

            End If
        Next cellRange
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, cancel As Boolean)

End Sub

Private Sub Workbook_BeforeClose(cancel As Boolean)
    On Error Resume Next
    deleteUserToolBar
ErrorHandler:
End Sub

Private Sub Workbook_SheetDeactivate(ByVal sh As Object)

End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal target As range)
    On Error GoTo ErrorHandler
    Dim sht As Worksheet
    Set sht = sh
    Set g_CurrentRange = target
    
    If ThisWorkbook.name <> ActiveWorkbook.name Or sht.Visible <> xlSheetVisible Or target.count > 1 Then
        Exit Sub
    End If
    
    Dim copyPasteGuard As New CCopyPasteGuard
    copyPasteGuard.init
    
    If sheetSelectionShouldCheck(sht) Then
        Call currentParameterBranchCheck(sht, target) '当前字段分支校验
    End If
    
    If g_CurrentSheet.name = getResByKey("LTECellSheet") Or g_CurrentSheet.name = getResByKey("UMTSCellSheet") Or g_CurrentSheet.name = getResByKey("LTECellMMSheet") Then 'gen new head board No.
        If get_ColumnName(g_CurrentSheet.name, target.column) = getResByKey("NewHeadBoardNo") Or _
            get_ColumnName(g_CurrentSheet.name, target.column) = getResByKey("NewTailBoardNo") Or _
            get_ColumnName(g_CurrentSheet.name, target.column) = getResByKey("NewHeadBoardNo2") Or _
            get_ColumnName(g_CurrentSheet.name, target.column) = getResByKey("NewTailBoardNo2") Then
            Call genBoardNoList(sht, target)
        End If
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in Workbook_SheetSelectionChange, " & Err.Description
End Sub





