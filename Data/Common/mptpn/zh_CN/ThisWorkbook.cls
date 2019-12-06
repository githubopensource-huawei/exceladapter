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
End Sub


Private Sub Workbook_Deactivate()
    On Error Resume Next
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
    
    If ThisWorkbook.name <> ActiveWorkbook.name Or sheet.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    If isControlDefSheetExist = False Then
        Exit Sub
    End If

    Dim currentNeType As String

    If (sheetName <> "MAPPING DEF") And (sheetName <> "SHEET DEF") And (sheetName <> "CONTROL DEF") _
        And (sheetName <> getResByKey("help")) And (sheetName <> getResByKey("Cover")) Then
        If sheetName = GetMainSheetName() Then
            Call TransPortSheetChange(sheet, Target)
        End If
        If Target.count Mod 256 = 0 Then
            Exit Sub
        End If
        For Each cellRange In Target
            If (cellRange.row < 3) Or (cellRange.column > sheet.Range("IV2").End(xlToLeft).column) Then
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


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    If Not isIubStyleWorkBook() And Not SummaryRes.Visible Then
'        AddLink
'    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

End Sub

Private Sub Workbook_SheetDeactivate(ByVal sh As Object)

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

    Exit Sub
ErrorHandler:
End Sub



