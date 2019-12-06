Attribute VB_Name = "Utility"
'用以设置颜色
Const SolidColorIdx = 16
Const SolidPattern = xlGray16
Const NullPattern = xlNone

Public Sub UnprotectSheet(curSheet As Worksheet)
    curSheet.Unprotect (GetSheetsPass)
   ' Application.ScreenUpdating = True
End Sub

Public Sub ProtectSheet(curSheet As Worksheet)
    curSheet.Protect Password:=GetSheetsPass, AllowFormattingCells:=True, AllowFormattingColumns:=True
   ' Application.ScreenUpdating = True
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
    Dim curSheet As Worksheet
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
    CoverSht.Select
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
    If (SmallRange.Column >= BigRange.Column) And (SmallRange.Column <= BigRange.Column + BigRange.Columns.Count) _
       And (SmallRange.row >= BigRange.row) And (SmallRange.row <= BigRange.row + BigRange.Rows.Count) Then
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

Public Sub SetInvalidateField(ByVal Target As Range)
    'validdef
    Const ValidDefBranchFieldCol = 1
    Const ValidDefBeginRowCol = 2
    Const ValidDefEndRowCol = 3
    Const ValidDefValueCol = 4
    Const ValidDefFieldCol = 6
    Const ValidDefValidCol = 7
     
    'color
    Const SolidColorIdx = 16
    Const SolidPattern = xlSolid
    Const NullPattern = xlNone
    Dim CurRange As Range, DefRow As Integer, BeginRow As String, EndRow As String
    Dim BranchRange As String, BranchValue As String, Valid As String, FieldRange As String
    Dim PhyLinkTypeRange As String
    Dim TPLinkTypeRange As String
    
    For DefRow = 0 To UBound(ValidDefine) - 1
        BeginRow = ValidDefine(DefRow, ValidDefBeginRowCol)
        EndRow = ValidDefine(DefRow, ValidDefEndRowCol)
        BranchValue = Trim(ValidDefine(DefRow, ValidDefValueCol))
        
        BranchRange = ValidDefine(DefRow, ValidDefBranchFieldCol) + BeginRow + ":" + ValidDefine(DefRow, ValidDefBranchFieldCol) + EndRow
        
        For Each CurRange In Range(BranchRange)
            If IsCellInRange(CurRange, Target) Then
                FieldRange = ValidDefine(DefRow, ValidDefFieldCol) + CStr(CurRange.row)
                
                If IsSubStr(Trim(CurRange.Text), BranchValue) Then
                    Call SetInValidationMode(FieldRange)
                Else
                    Call SetValidationMode(FieldRange)
                End If
                If (Trim(CurRange.Text) = "") Then
                    Call ClearCell(Range(FieldRange))
                End If
            End If
        Next CurRange
    Next
End Sub

Public Sub SetFieldValidation(ByVal Target As Range)
    'validdef
    Const ValidDefBranchFieldCol = 1
    Const ValidDefBeginRowCol = 2
    Const ValidDefEndRowCol = 3
    Const ValidDefValueCol = 4
    Const ValidDefFieldCol = 6
    Const ValidDefTypeCol = 7
    Const ValidDefMinCol = 8
    Const ValidDefMaxCol = 9
    Const ValidDefListCol = 10
    Const ValidDefPromptCol = 11

    Dim CurRange As Range, DefRow As Integer, BeginRow As String, EndRow As String
    Dim BranchRange As String, BranchValue As String, Valid As String, FieldRange As String
    Dim PhyLinkTypeRange As String
    Dim TPLinkTypeRange As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sPrompt As String
    Dim xType As Excel.XlDVType, sFormula1 As String, sFormula2 As String, sErrPrompt As String, sErrMsg As String
        
    For DefRow = 0 To UBound(RangeDefine) - 1
        BeginRow = RangeDefine(DefRow, ValidDefBeginRowCol)
        EndRow = RangeDefine(DefRow, ValidDefEndRowCol)
        BranchValue = Trim(RangeDefine(DefRow, ValidDefValueCol))
        sDataType = Trim(RangeDefine(DefRow, ValidDefTypeCol))
        sMinVal = Trim(RangeDefine(DefRow, ValidDefMinCol))
        sMaxVal = Trim(RangeDefine(DefRow, ValidDefMaxCol))
        sRangeList = Trim(RangeDefine(DefRow, ValidDefListCol))
        sPrompt = Trim(RangeDefine(DefRow, ValidDefPromptCol))
                
        BranchRange = RangeDefine(DefRow, ValidDefBranchFieldCol) + BeginRow + ":" + RangeDefine(DefRow, ValidDefBranchFieldCol) + EndRow
        
        For Each CurRange In Range(BranchRange)
            If IsCellInRange(CurRange, Target) And IsSubStr(Trim(CurRange.Text), BranchValue) Then
                
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
                FieldRange = RangeDefine(DefRow, ValidDefFieldCol) + CStr(CurRange.row)
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

Public Function IsSystemSheet(curSheet As Worksheet) As Boolean
    IsSystemSheet = False
    If curSheet.CodeName = CoverSht.CodeName _
         Or curSheet.CodeName = TableSht.CodeName Or curSheet.CodeName = ValidSht.CodeName _
         Or curSheet.CodeName = InitTableSht.CodeName Or curSheet.CodeName = InitFieldSht.CodeName Then
        IsSystemSheet = True
    End If
End Function

Public Function CreatePathFileObject(sPath As String, sFileName As String) As Object
  Dim strFileName As String
  If Len(sFileName) = 0 Then
    MsgBox "缺少文件名", vbCritical
    Exit Function
  End If
  
  strFileName = sPath + "\" + sFileName
  Dim fs As Object, fstemp As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
 
  If fs.FileExists(strFileName) Then
    Dim iJudge As Integer
    iJudge = MsgBox("目录下已经存在文件" + strFileName + ",需要覆盖吗？", 1)
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

