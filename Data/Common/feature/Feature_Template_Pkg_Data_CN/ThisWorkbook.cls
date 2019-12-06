VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public g_SelectionRow As Integer
Public g_SelectionCol As Integer

Private Sub Workbook_Activate()
    On Error Resume Next
    initResource
    buildControlRelations
End Sub

'分支参数控制响应函数
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As Range)
    On Error Resume Next
    Dim sheet As Worksheet
    Dim cellRange As Range
    Dim isConAttr As Boolean
    Dim rowNum As Integer
    Dim contRel As controlRelation
    Set sheet = sh
    If ThisWorkbook.name <> ActiveWorkbook.name Or sheet.Visible <> xlSheetVisible Then
        Exit Sub
    End If

    If isControlDefSheetExist = False Then
        Exit Sub
    End If
    
    Dim currentNeType As String
    
    If (sheet.name <> MappingSheetName) And (sheetName <> SheetDefName) And (sheet.name <> ControllSheetName) And (sheet.name <> FeatureListSheetName) _
        And (sheet.name <> getResByKey("help")) And (sheet.name <> getResByKey("Cover")) And (sheet.name <> "SummaryRes") And (sheet.name <> getResByKey("PackageCustomTemplate")) Then
        If target.count Mod 256 = 0 Then
            Exit Sub
        End If
        For Each cellRange In target
            If (cellRange.row < DataSheetAttrRow + 2) Or (cellRange.Column > sheet.Range("IV" + CStr(DataSheetAttrRow)).End(xlToLeft).Column) Then
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
    ElseIf (sheet.name = getResByKey("PackageCustomTemplate")) Then
        Call executeTemplateBranchControl(sheet, target)
    End If
    
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal target As Range)
    On Error GoTo ErrorHandler
    Dim sht As Worksheet
    Set sht = sh
    If ThisWorkbook.name <> ActiveWorkbook.name Or sht.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    
    If (target.row < DataSheetAttrRow + 2) Then
        Exit Sub
    End If
    
    If sheetSelectionShouldCheck(sht) Then
        Call currentParameterBranchCheck(sht, target) '当前参数的分支校验
    End If
    If Ensure_Gray_Cell(target) = False Then
        Exit Sub
    End If
    Exit Sub
ErrorHandler:
End Sub

