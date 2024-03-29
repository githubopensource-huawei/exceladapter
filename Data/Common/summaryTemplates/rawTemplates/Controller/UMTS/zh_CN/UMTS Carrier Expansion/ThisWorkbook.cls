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

Private Sub Workbook_Activate()
    On Error Resume Next
    initResource
    InsertUserToolBar
    buildControlRelations
    Call buildMappingNumberRelations
    'initAddRef
    initInnerValideDef
    Call initBandDefData
End Sub
'分支参数控制响应函数
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As range)
    On Error Resume Next
    Dim sheet As Worksheet
    Dim cellRange As range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Dim currentNeType As String
    
        Dim copyPasteGuard As New CCopyPasteGuard
    copyPasteGuard.init
        
    currentNeType = "UMTS"
    Set sheet = sh
    If ThisWorkbook.name <> ActiveWorkbook.name Or sheet.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    If getLLDModelFlag <> "" Then
        Exit Sub
    End If
    If isControlDefSheetExist = False Then
        Exit Sub
    End If
    
     
    If (isGsmCellSheet(sheet.name) <> True) And (sheet.name <> "MappingSiteTemplate") And (sheet.name <> "ProductType") _
        And (sheet.name <> "MappingCellTemplate") And (sheet.name <> "MappingRadioTemplate") _
        And (sheet.name <> "MAPPING DEF") And (sheet.name <> "SHEET DEF") And (sheet.name <> "CONTROL DEF") _
        And (sheet.name <> "Help for USB Parameter") And (sheet.name <> getResByKey("help")) _
        And (sheet.name <> "IPRouteMap") And (sheet.name <> getResByKey("Cover")) And (sheet.name <> "COMMON") _
        And (sheet.name <> "Qos") And (sheet.name <> "USB Parameter for Sites") And (sheet.name <> "SummaryRes") Then
        If Target.count Mod 256 = 0 Then
            Exit Sub
        End If
        For Each cellRange In Target
            If sheet.name = getResByKey("Comm Data") Then
                If cellRange.Interior.colorIndex = 34 Or cellRange.Interior.colorIndex = 40 Or cellRange.Borders.LineStyle = xlLineStyleNone Then
                    Exit Sub
                End If
            Else
                If (cellRange.row < 3) Or (cellRange.column > sheet.range("XFD2").End(xlToLeft).column) Then
                    Exit Sub
                End If
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
    End If
End Sub
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    On Error Resume Next
'    Call HideToolBar
'End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    DeleteUserToolBar
    Application.StatusBar = False
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    'Call InsertUserToolBar
End Sub
Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal Target As range)
    On Error GoTo ErrorHandler
    Dim sht As Worksheet
    Set sht = sh
    If ThisWorkbook.name <> ActiveWorkbook.name Or sht.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    
    If sheetSelectionShouldCheck(sht) Then
        Call currentParameterBranchCheck(sht, Target) '当前参数的分支校验
    End If

     If isCellSheet(sht.name) And getNeType() <> "USU" Then
        Call CellSelectionChange(sht, Target)
    End If
    If isNeighborPolicySht(sht.name) Then
        Call neighborPolicyShtSelectionChange(sht, Target)
    End If
    Exit Sub
ErrorHandler:
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
    On Error Resume Next
    If isBandDefinitionSht(sht.name) Then
        Call initBandDefData
    End If
    
End Sub

Public Function getSelRow() As Long
    If (g_SelectionRow = 0) Then
        getSelRow = 8
    Else
        getSelRow = g_SelectionRow
    End If
End Function
Public Function getSelCol() As Long
    If (g_SelectionCol = 0) Then
        getSelCol = 3
    Else
        getSelCol = g_SelectionCol
    End If
End Function




