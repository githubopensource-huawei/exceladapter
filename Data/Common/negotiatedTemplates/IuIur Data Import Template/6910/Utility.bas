Attribute VB_Name = "Utility"
'用以设置颜色
Const SolidColorIdx = 16
Const SolidPattern = xlGray16
Const NullPattern = xlNone
Const ValidShtSheetNameColumn = 2

Public Sub UnprotectSheet(CurSheet As Worksheet)
    CurSheet.Unprotect (GetSheetsPass)
   ' Application.ScreenUpdating = True
End Sub

Public Sub ProtectSheet(CurSheet As Worksheet)
   ' CurSheet.Protect Password:=GetSheetsPass, AllowFormattingCells:=True, AllowFormattingColumns:=True
   ' Application.ScreenUpdating = True
End Sub
Public Sub InsertUserToolBar()
    Dim cmbNewBar As CommandBar
    Dim ctlBtn As CommandBarButton
    
    On Error Resume Next
    Set cmbNewBar = CommandBars.Add(Name:="Operate Bar")
         'With cmbNewBar
    '    Set ctlBtn = .Controls.Add
    '    With ctlBtn
    '        .Style = msoButtonIconAndCaption
    '        .BeginGroup = True
    '
    '        If bIsEng Then
    '            .Caption = "&Customize Template"
    '            .TooltipText = "Customize Template"
    '        Else
    '            .Caption = Sheets("TableDef").Range("P3").Text
    '            .TooltipText = Sheets("TableDef").Range("P3").Text
    '        End If
    '
    '        .OnAction = "ShowCfgForm"
    '        .FaceId = 50
    '    End With
    '    .Protection = msoBarNoCustomize
    '    .Position = msoBarTop
    '    .Visible = True
    'End With
      With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
          .Style = msoButtonIconAndCaption
          .BeginGroup = True
          
          If bIsEng Then
              .Caption = "&English Version&"
              .TooltipText = "English Version"
          Else
              .Caption = Sheets("TableDef").Range("P4").Text
              .TooltipText = Sheets("TableDef").Range("P4").Text
          End If
             
          .OnAction = "SetEnglishVersion"
          .FaceId = 50
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
      With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
          .Style = msoButtonIconAndCaption
          .BeginGroup = True
          
          If bIsEng Then
              .Caption = "&Chinese Version&"
              .TooltipText = "Chinese Version"
          Else
              .Caption = Sheets("TableDef").Range("P5").Text
              .TooltipText = Sheets("TableDef").Range("P5").Text
          End If
          
          .OnAction = "SetChineseVersion"
          .FaceId = 50
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
    
      With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            
            If bIsEng Then
                .Caption = "&Hide Empty Row"
                .TooltipText = "Hide Empty Row"
            Else
                .Caption = Sheets("TableDef").Range("P6").Text
                .TooltipText = Sheets("TableDef").Range("P6").Text
            End If
            
            .OnAction = "HideTemplateEmptyRow"
            .FaceId = 54
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
     With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            
            If bIsEng Then
                .Caption = "&Reset Row"
                .TooltipText = "Reset Row"
            Else
                .Caption = Sheets("TableDef").Range("P7").Text
                .TooltipText = Sheets("TableDef").Range("P7").Text
            End If
            
            .OnAction = "ToolButtoUnHideTemplate"
            .FaceId = 55
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            
            If bIsEng Then
                .Caption = "&Expand Row"
                .TooltipText = "Expand Row"
            Else
                .Caption = Sheets("TableDef").Range("R8").Text
                .TooltipText = Sheets("TableDef").Range("R8").Text
            End If
            
            .OnAction = "ExpandRow"
            .FaceId = 56
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
End Sub

Public Sub DeleteUserToolBar()
    Application.CommandBars("Operate Bar").Delete
End Sub

Public Sub SetValidation(FieldRange As String, AddType As Long, Formula1 As String, Formula2 As String, RangeStr As String)
    With Range(FieldRange).Validation
        .Delete
        .Add Type:=AddType, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=Formula1, Formula2:=Formula2
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "CME"
        .InputMessage = ""
        .ErrorMessage = RangeStr
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Public Sub ClearValidation(FieldRange As String)
    Call SetValidation(FieldRange, xlValidateTextLength, "0", "0", "No input is required.")
End Sub

Public Function IsSheetNameExists(NewSheetName As String) As Boolean
    Dim ShtIdx As Integer
    Dim IsExist As Boolean
    
    IsSheetNameExists = False
    ShtIdx = 1
    Do While ShtIdx <= ActiveWorkbook.Sheets.Count
        If ActiveWorkbook.Sheets(ShtIdx).Name = NewSheetName Then
           IsSheetNameExists = True
           Exit Function
        End If
        
        ShtIdx = ShtIdx + 1
    Loop
End Function

Public Sub SetInValidationMode(FieldRange As String)
    Range(FieldRange).Interior.ColorIndex = SolidColorIdx
    Range(FieldRange).Interior.Pattern = SolidPattern
    Call ClearCell(Range(FieldRange))
End Sub
Public Sub SetValidationMode(FieldRange As String)
    Range(FieldRange).Interior.ColorIndex = NullPattern
    Range(FieldRange).Interior.Pattern = NullPattern
End Sub

Private Sub RefreshWorkSheets()
    Dim CurSheet As Worksheet
    Dim ShtIdx As Integer
    
    On Error Resume Next
    ShtIdx = 1
    Do While (ShtIdx <= ActiveWorkbook.Sheets.Count)
        If Not IsSystemSheet(ActiveWorkbook.Sheets(ShtIdx)) Then
            ActiveWorkbook.Sheets(ShtIdx).Select
            ActiveWorkbook.Sheets(ShtIdx).RefreshThisSheet
        End If
        ShtIdx = ShtIdx + 1
    Loop
    ThisWorkbook.Worksheets("Cover").Select
End Sub

Public Sub RefreshWorkbook()
    RefreshWorkSheets
    ProtectedAll
End Sub

Public Sub UnProtectedAll()
    Dim ShtIdx As Integer
    Dim OpSht As Worksheet
    
    On Error Resume Next
    Call UnprotectWorkBook
    ShtIdx = 1
    Do While (ShtIdx <= ActiveWorkbook.Sheets.Count)
        Set OpSht = ActiveWorkbook.Sheets(ShtIdx)
        If Not IsSystemSheet(OpSht) Then
            Call UnprotectWorkSheet(OpSht)
        End If
        ShtIdx = ShtIdx + 1
    Loop
End Sub
Private Sub ProtectedAll()
    Dim ShtIdx As Integer
    Dim OpSht As Worksheet
    
    On Error Resume Next
    ShtIdx = 1
    Do While (ShtIdx <= ActiveWorkbook.Sheets.Count)
        Set OpSht = ActiveWorkbook.Sheets(ShtIdx)
        If Not IsSystemSheet(OpSht) Then
            Call ProtectWorkSheet(OpSht)
        End If
        ShtIdx = ShtIdx + 1
    Loop
    Call ProtectWorkBook
End Sub

Public Sub NoValueNeeded(TargetFieldRange As String)
    MsgBox "No input is required.", vbOKOnly, "CME"
    Range(TargetFieldRange).Select
End Sub

Public Function IsCellInRange(SmallRange As Range, BigRange As Range) As Boolean
    IsCellInRange = False
    If (SmallRange.Column >= BigRange.Column) And (SmallRange.Column < BigRange.Column + BigRange.Columns.Count) _
       And (SmallRange.Row >= BigRange.Row) And (SmallRange.Row < BigRange.Row + BigRange.Rows.Count) Then
        IsCellInRange = True
    End If
End Function

Public Function IsSubStr(substr As String, str As String) As Boolean
    Dim ArrData() As String
    Dim i As Integer
    
    IsSubStr = False
    
    ArrData = Split(str, ",")
    For i = 0 To UBound(ArrData)
        If Trim(ArrData(i)) = Trim(substr) Then
            IsSubStr = True
            Exit Function
        End If
    Next
End Function

Public Sub SetInvalidateField(ByVal Target As Range, isActivateFlag As Boolean, CurSheetName As String)
    'validdef
    Const ValidSheetNameIndex = 0
    Const ValidDefBranchFieldIndex = 6
    Const ValidDefBeginRowIndex = 8
    Const ValidDefEndRowIndex = 9
    Const ValidDefValueIndex = 4
    Const ValidDefFieldIndex = 7
     
    'color
    Const SolidColorIdx = 16
    Const SolidPattern = xlSolid
    Const NullPattern = xlNone
    Dim CurRange As Range, DefRow As Integer, BeginRow As String, EndRow As String
    Dim BranchRange As String, BranchValue As String, Valid As String, FieldRange As String, sSheetName As String
    
    For DefRow = 0 To UBound(ValidDefine) - 1
        BeginRow = ValidDefine(DefRow, ValidDefBeginRowIndex)
        EndRow = ValidDefine(DefRow, ValidDefEndRowIndex)
        BranchValue = UCase(Trim(ValidDefine(DefRow, ValidDefValueIndex)))
        sSheetName = Trim(ValidDefine(DefRow, ValidSheetNameIndex))
        BranchRange = ValidDefine(DefRow, ValidDefBranchFieldIndex) + BeginRow + ":" + ValidDefine(DefRow, ValidDefBranchFieldIndex) + EndRow
        
        For Each CurRange In Range(BranchRange)
            If sSheetName = CurSheetName And IsCellInRange(CurRange, Target) Then
                FieldRange = ValidDefine(DefRow, ValidDefFieldIndex) + CStr(CurRange.Row)

                If IsSubStr(UCase(Trim(CurRange.Text)), BranchValue) Then
                    Call SetInValidationMode(FieldRange)
                ElseIf Not isActivateFlag Then
                    Call SetValidationMode(FieldRange)
                End If
       '         If (Trim(CurRange.Text) = "") Then
       '             Call ClearCell(Range(FieldRange))
       '         End If
            End If
        Next CurRange
    Next
End Sub

Public Sub SetFieldValidation(ByVal Target As Range, CurSheetName As String)
    'validdef
    Const ValidSheetNameIndex = 0
    Const ValidDefBranchFieldIndex = 6
    Const ValidDefFieldIndex = 7
    Const ValidDefValueIndex = 4
    Const ValidDefBeginRowIndex = 8
    Const ValidDefEndRowIndex = 9
    Const ValidDefTypeIndex = 10
    Const ValidDefMinIndex = 11
    Const ValidDefMaxIndex = 12
    Const ValidDefListIndex = 13
    Const ValidDefPromptIndex = 14

    Dim CurRange As Range, DefRow As Integer, BeginRow As String, EndRow As String
    Dim BranchRange As String, BranchValue As String, Valid As String, FieldRange As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sPrompt As String
    Dim xType As Excel.XlDVType, sFormula1 As String, sFormula2 As String, sErrPrompt As String, sErrMsg As String
        
    For DefRow = 0 To UBound(RangeDefine) - 1
        BeginRow = RangeDefine(DefRow, ValidDefBeginRowIndex)
        EndRow = RangeDefine(DefRow, ValidDefEndRowIndex)
        BranchValue = UCase(Trim(RangeDefine(DefRow, ValidDefValueIndex)))
        sDataType = Trim(RangeDefine(DefRow, ValidDefTypeIndex))
        sMinVal = Trim(RangeDefine(DefRow, ValidDefMinIndex))
        sMaxVal = Trim(RangeDefine(DefRow, ValidDefMaxIndex))
        sRangeList = Trim(RangeDefine(DefRow, ValidDefListIndex))
        sSheetName = Trim(RangeDefine(DefRow, ValidSheetNameIndex))
        sPrompt = Trim(RangeDefine(DefRow, ValidDefPromptIndex))
        If sMinVal <> "" Then
           sPrompt = "[" + sMinVal + "," + sMaxVal + "]"
        ElseIf sRangeList <> "" Then
           sPrompt = "[" + sRangeList + "]"
        End If
                
        BranchRange = RangeDefine(DefRow, ValidDefBranchFieldIndex) + BeginRow + ":" + RangeDefine(DefRow, ValidDefBranchFieldIndex) + EndRow
        
        For Each CurRange In Range(BranchRange)
            If IsCellInRange(CurRange, Target) And IsSubStr(UCase(Trim(CurRange.Text)), BranchValue) And sSheetName = CurSheetName Then

                If (sDataType = "INT") Then
                    xType = xlValidateWholeNumber
                    sFormula1 = sMinVal
                    sFormula2 = sMaxVal
                End If
                If (sDataType = "STRING") Then
                    xType = xlValidateTextLength
                    sFormula1 = sMinVal
                    sFormula2 = sMaxVal
                End If
                If (sDataType = "LIST") Then
                    xType = xlValidateList
                    sFormula1 = sRangeList
                    sFormula2 = ""
                End If
                FieldRange = RangeDefine(DefRow, ValidDefFieldIndex) + CStr(CurRange.Row)
                Call ClearValidation(FieldRange)
                Call SetValidation(FieldRange, xType, sFormula1, sFormula2, sPrompt)
            End If
        Next CurRange
    Next
End Sub

Public Sub ClearCell(CurCell As Range)
    If (Trim(CurCell.Value) <> "") Then
        CurCell.Value = ""
    End If
End Sub

Public Function IsSystemSheet(CurSheet As Worksheet) As Boolean
    IsSystemSheet = False
    If CurSheet.CodeName = ThisWorkbook.Worksheets("Cover").CodeName _
         Or CurSheet.CodeName = ThisWorkbook.Worksheets("TableDef").CodeName Or CurSheet.CodeName = ThisWorkbook.Worksheets("ValidDef").CodeName _
         Or CurSheet.CodeName = ThisWorkbook.Worksheets("InitTableMap").CodeName Or CurSheet.CodeName = ThisWorkbook.Worksheets("InitFieldMap").CodeName _
         Or CurSheet.CodeName = ThisWorkbook.Worksheets("CMETemplateInfo").CodeName Or CurSheet.CodeName = ThisWorkbook.Worksheets("EnumDef").CodeName Then
        IsSystemSheet = True
    End If
End Function

Public Function CreatePathFileObject(sPath As String, sFileName As String) As Object
  Dim strFileName As String
  If Len(sFileName) = 0 Then
    MsgBox Sheets("TableDef").Range("S3").Text, vbCritical
    Exit Function
  End If
  
  strFileName = sPath + "\" + sFileName
  Dim fs As Object, fstemp As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
 
  If fs.FileExists(strFileName) Then
    Dim iJudge As Integer
    iJudge = MsgBox(Sheets("TableDef").Range("S4").Text + strFileName + Sheets("TableDef").Range("S5").Text, 1)
       If iJudge = 2 Then
         Exit Function
       End If
       Set fstemp = fs.GetFile(strFileName)
       Set CreatePathFileObject = fstemp.OpenAsTextStream(2, -2)
  Else
    Dim txtFile As Object
    Set txtFile = fs.CreateTextFile(strFileName, True)
    Set CreatePathFileObject = txtFile
  End If
End Function

Public Function StringFormat(ByVal strInput As String) As String
    StringFormat = IIf(strInput = "", "''", "'" + strInput + "'")
End Function

Public Function GetAllRange(Category As String) As String
    Dim MaxCol As String
    Dim MaxRow As Integer
    
    MaxCol = "B"
    MaxRow = 8
    
    Dim CurSheet As Worksheet
    Set CurSheet = Sheets("TableDef")
    For iRow = StartTblDataRow To StartTblDataRow + CInt(CurSheet.Cells(5, 7))
        If CurSheet.Cells(iRow, 2) = Category Then
            If CurSheet.Cells(iRow, 10) = "BQ" Then
                CurSheet.Cells(iRow, 10) = CurSheet.Cells(iRow, 10)
            End If
            
            If CurSheet.Cells(iRow, 10) = "AP" Then
                CurSheet.Cells(iRow, 10) = CurSheet.Cells(iRow, 10)
            End If
            
            
            If (Len(MaxCol) < Len(CurSheet.Cells(iRow, 10))) Or _
                (Len(MaxCol) = Len(CurSheet.Cells(iRow, 10)) And MaxCol < CurSheet.Cells(iRow, 10)) Then
                MaxCol = CurSheet.Cells(iRow, 10)
            End If
            If MaxRow < CurSheet.Cells(iRow, 13) Then
                MaxRow = CurSheet.Cells(iRow, 13)
            End If
        End If
    Next
    
    GetAllRange = "B8:" + MaxCol + CStr(MaxRow)
End Function

'**********************************************************
'从列数得到列名：1->A，27->AA
'**********************************************************
Public Function c(iColumn As Integer) As String
  If iColumn >= 257 Or iColumn < 0 Then
    c = ""
    Return
  End If
  
  Dim Result As String
  Dim High, Low As Integer
  
  High = Int((iColumn - 1) / 26)
  Low = iColumn Mod 26
  
  If High > 0 Then
    Result = Chr(High + 64)
  End If

  If Low = 0 Then
    Low = 26
  End If
  
  Result = Result & Chr(Low + 64)
  c = Result
End Function

Public Sub ExpandRow()
    Call ChangeRow(1)
End Sub


Public Sub ChangeRow(iExpandRow As Integer)
  Dim CurSheet As Worksheet
  Dim CurRange As Range
  Dim ObjName As String
  Dim iBasicRow As Integer
  Dim CurRangeRows As Integer
  Dim iRow As Integer
  Dim NextIsMerged As Boolean
  Dim NowStartRow As Integer
  Dim NowEndRow As Integer
  Dim NextStartRow As Integer
  Dim NextEndRow As Integer
  
  Dim sht As Worksheet
  Set sht = ThisWorkbook.Sheets("TableDef")
  
  Set CurSheet = ActiveWindow.ActiveSheet
  Application.DisplayAlerts = False
  If TypeOf Selection Is Range _
        And Selection.Areas.Count = 1 And Selection.MergeCells And Selection.Column = 1 And Selection.Rows.Count > 2 Then  '判断选中了第一列的对象名
    
    Set CurRange = Selection.Areas(1)
    NowStartRow = CurRange.Row
    NowEndRow = CurRange.Row + CurRange.Rows.Count - 1
    
    ObjName = CurRange.Cells(1, 1)
    '插入新行
    Call UnprotectWorkSheet(CurSheet)
    iBasicRow = CurRange.Row + 2
    CurRangeRows = CurRange.Rows.Count
    CurRange.UnMerge
    
    CurSheet.Cells(CurRange.Row + CurRangeRows, 1).Select
    NextIsMerged = Selection.MergeCells
    If NextIsMerged Then
      NextStartRow = Selection.Cells.Row
      NextEndRow = NextStartRow + Selection.Cells.Count - 1
      Selection.UnMerge
    End If
    
    Rows(CStr(CurRange.Row + CurRangeRows) + ":" + CStr(CurRange.Row + CurRangeRows)).Select
    If iExpandRow > 0 Then
      'Selection.Insert Shift:=xlDown
      Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Else
     Selection.Delete Shift:=xlUp
    End If

    If NextIsMerged Then
      CurSheet.Range("A" + CStr(NextStartRow + 1) + ":A" + CStr(NextEndRow + 1)).Select
      Selection.Merge
    End If
    
    CurSheet.Rows(NowEndRow).Select
      Selection.Copy
      CurSheet.Rows(NowEndRow + 1).Select
      Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    CurSheet.Range(c(CurRange.Column) + CStr(CurRange.Row) + ":" + c(CurRange.Column) + CStr(CurRange.Row + CurRangeRows - 1 + iExpandRow)).Merge
    CurSheet.Range("A" + CStr(NowStartRow) + ":A" + CStr(NowStartRow)).Select
    
    Call ProtectWorkSheet(CurSheet)

    '更新TableDef中的定义
    For iRow = StartTblDataRow To StartTblDataRow + CInt(sht.Cells(5, 7)) - 1
      If sht.Cells(iRow, iTitleBeginRow + 1) > iBasicRow And sht.Cells(iRow, 2) = CurSheet.Name Then
        sht.Cells(iRow, iTitleBeginRow + 1) = sht.Cells(iRow, iTitleBeginRow + 1) + iExpandRow
      End If
      If sht.Cells(iRow, iContentEndRow + 1) >= iBasicRow And sht.Cells(iRow, 2) = CurSheet.Name Then
        sht.Cells(iRow, iContentEndRow + 1) = sht.Cells(iRow, iContentEndRow + 1) + iExpandRow
      End If
    Next iRow
    
    '更新ValidDef
    Call RefreshBranchDefRow(CurSheet.Name, NowStartRow + 1)
    Call EnsureRefreshBranchDefine
    
    '更新InitTableDef
    Call GenInitTableMap

  Else
   If bIsEng Then
      MsgBox "Please choose the title of the object,or this object can not be expanded.", vbInformation
    Else
      MsgBox "请定位在对象名上，或此对象不支持扩展行。", vbInformation
   End If
  End If
End Sub



