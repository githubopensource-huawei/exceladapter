VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private IsCommonSheetHasRefresh As Boolean
Private IsIUCSSheetHasRefresh As Boolean
Private IsIUPSSheetHasRefresh As Boolean
Private IsIURSheetHasRefresh As Boolean

Private Sub Workbook_Activate()
    On Error Resume Next
    bIsEng = CBool(Sheets("TableDef").Range("P10").Text)
    InsertUserToolBar
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    DeleteUserToolBar
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    Call GetValidDefineData
    ThisWorkbook.Worksheets("Cover").Select
End Sub

Private Sub Workbook_SheetActivate(ByVal sheet As Object)
    If sheet.Name = "COMMON" Then
        If Not IsCommonSheetHasRefresh Then
          IsCommonSheetHasRefresh = True
          Dim Common_AllRange As String
          Common_AllRange = GetAllRange("COMMON")
          Call Common_RefreshThisSheet(Range(Common_AllRange), IsCommonSheetHasRefresh)
        End If
    End If
    
    If sheet.Name = "IUCS" Then
        If Not IsIUCSSheetHasRefresh Then
          IsIUCSSheetHasRefresh = True
          Dim IUCS_AllRange As String
          IUCS_AllRange = GetAllRange("IUCS")
          Call IUCS_RefreshThisSheet(Range(IUCS_AllRange), IsIUCSSheetHasRefresh)
        End If
    End If
    
    If sheet.Name = "IUPS" Then
        If Not IsIUPSSheetHasRefresh Then
          IsIUPSSheetHasRefresh = True
          Dim IUPS_AllRange As String
          IUPS_AllRange = GetAllRange("IUPS")
          Call IUPS_RefreshThisSheet(Range(IUPS_AllRange), IsIUPSSheetHasRefresh)
        End If

    End If
    
    If sheet.Name = "IUR" Then
        If Not IsIURSheetHasRefresh Then
          IsIURSheetHasRefresh = True
          Dim IUR_AllRange As String
          IUR_AllRange = GetAllRange("IUR")
          Call IUR_RefreshThisSheet(Range(IUR_AllRange), IsIURSheetHasRefresh)
        End If
    End If
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sheet As Object, ByVal Target As Range)
    If sheet.Name = "COMMON" Then
        Call Common_RefreshThisSheet(Target, False)
    End If
    
    If sheet.Name = "IUCS" Then
        Call IUCS_RefreshThisSheet(Target, False)
    End If
    
    If sheet.Name = "IUPS" Then
        Call IUPS_RefreshThisSheet(Target, False)
    End If
    
    If sheet.Name = "IUR" Then
        Call IUR_RefreshThisSheet(Target, False)
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal sheet As Object, ByVal Target As Range)
    If sheet.Name = "COMMON" Then
        Call Common_RefreshThisSheet(Target, False)
    End If
    
    If sheet.Name = "IUCS" Then
        Call IUCS_RefreshThisSheet(Target, False)
    End If
    
    If sheet.Name = "IUPS" Then
        Call IUPS_RefreshThisSheet(Target, False)
    End If
    
    If sheet.Name = "IUR" Then
        Call IUR_RefreshThisSheet(Target, False)
    End If
End Sub
