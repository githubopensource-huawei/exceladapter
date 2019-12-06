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
    Call buildMappingNumberRelations

    Set currentSheet = ThisWorkbook.ActiveSheet
    Call addHyperlinksForSheets(currentSheet) '给单板样式页，传输页增加超链接

    InsertUserToolBar
    'insertAndDeleteControl (True)
    Call initMenuStatus(currentSheet)

    If isBoardStyleSheet(currentSheet) Then
        Call setBgColorWithGray_i(currentSheet)
        'insertAndDeleteControl (False)
    End If
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
    On Error Resume Next
    If ThisWorkbook.name <> ActiveWorkbook.name Or sh.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    
    If InAdjustAntnPort = True Then
        Exit Sub
    End If
    
    deleteUserToolBar
    InsertUserToolBar
    
    Dim sheet As Worksheet
    Set sheet = sh
    Call initMenuStatus(sheet)
    
    Call addHyperlinksForSheets(sh) '给单板样式页，传输页增加超链接
    
    If isBoardStyleSheet(sheet) Then
        Call setBgColorWithGray_i(sheet)
        'insertAndDeleteControl (False)
    End If
errorhandle:
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    deleteUserToolBar
    destroyMenuStatus
ErrorHandler:
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal target As range, cancel As Boolean)
    On Error GoTo ErrorHandler
    If target.Interior.colorIndex <> GRAY_COLOR Then
        Set currentSheet = sh
        Set currentRange = target
        If customFreqSelected(sh, target) Then
            freqInfo.Show
        ElseIf bbEqmBoardNoSelected(sh, target) Then
            BoardNo.Show
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
    
    If sheetName = GetMainSheetName() Or sheetName = getResByKey("UMTSCellSheet") Or sheetName = getResByKey("USectorSheet") Then
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
    On Error Resume Next
    If InAdjustAntnPort = True Then
        Call popUpTempSheetCannotSaveMsgbox
        cancel = True
        SaveAsUI = True
    End If
End Sub

Private Sub Workbook_BeforeClose(cancel As Boolean)
    On Error Resume Next
    If InAdjustAntnPort = True Then
        Call popUpTempSheetCannotSaveMsgbox
        cancel = True
    End If
ErrorHandler:
End Sub

Private Sub Workbook_SheetDeactivate(ByVal sh As Object)
    On Error Resume Next
    If InAdjustAntnPort = True And sh.name = getResByKey("Temp Sheet") Then
        Call popUpTempSheetCannotChangeMsgbox
    End If
    'insertAndDeleteControl (True)
ErrorHandler:
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal target As range)
    On Error GoTo ErrorHandler
    'currentCellValue = ActiveCell.value
    Dim sht As Worksheet
    Set sht = sh
    
    Dim copyPasteGuard As New CCopyPasteGuard
    copyPasteGuard.init
    
    If ThisWorkbook.name <> ActiveWorkbook.name Or sht.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    
    If sheetSelectionShouldCheck(sht) Then
        Call currentParameterBranchCheck(sht, target) '当前字段分支校验
    End If

    Exit Sub
ErrorHandler:
    Debug.Print "some exception in Workbook_SheetSelectionChange, " & Err.Description
End Sub







