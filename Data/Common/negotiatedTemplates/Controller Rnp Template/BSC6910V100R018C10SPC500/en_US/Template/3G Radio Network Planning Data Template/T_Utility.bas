Attribute VB_Name = "T_Utility"
'用以设置颜色
Const SolidColorIdx = 16
Const SolidPattern = xlGray16
Const NullPattern = xlNone


'****************************************************************
'设置FieldRange指定区域变为灰色
'****************************************************************
Public Sub SetInValidationMode(FieldRange As String)
    Range(FieldRange).Interior.ColorIndex = SolidColorIdx
    Range(FieldRange).Interior.Pattern = SolidPattern
    Call ClearCell(Range(FieldRange))
End Sub

'****************************************************************
'设置FieldRange指定区域为空白
'****************************************************************
Public Sub SetValidationMode(FieldRange As String)
    Range(FieldRange).Interior.ColorIndex = NullPattern
    Range(FieldRange).Interior.Pattern = NullPattern
End Sub

'****************************************************************
'清除CurCell指定单元格的值
'****************************************************************
Public Sub ClearCell(CurCell As Range)
    If (Trim(CurCell.Value) <> "") Then
        CurCell.Value = ""
    End If
End Sub

'****************************************************************
'弹出不可编辑提示框
'****************************************************************
Public Sub NoValueNeeded(TargetFieldRange As String)
    MsgBox "No input is required.", vbOKOnly, "CME"
    Range(TargetFieldRange).Select
End Sub

'****************************************************************
'判断substr是否包含在str中，是返回True  否返回False
'****************************************************************
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

'****************************************************************
'设置与Target指定单元格相关的其他区域是否为灰色
'****************************************************************
Public Sub SetInvalidateField(ByVal Target As Range, CurSheetName As String)
    Const ValidSheetNameCol = 0
    Const ValidDefBranchFieldCol = 2
    Const ValidDefBeginRowCol = 3
    Const ValidDefEndRowCol = 4
    Const ValidDefValueCol = 5
    Const ValidDefFieldCol = 7
    Const ValidDefValidCol = 8
    'color
    Const SolidColorIdx = 16
    Const SolidPattern = xlSolid
    Const NullPattern = xlNone
    
    Dim CurRange As Range, DefRow As Integer, BeginRow As String, EndRow As String
    Dim BranchValue As String, FieldRange As String, sSheetName As String
    
    For DefRow = 0 To UBound(ValidDefine) - 1
        BeginRow = ValidDefine(DefRow, ValidDefBeginRowCol)
        EndRow = ValidDefine(DefRow, ValidDefEndRowCol)
        BranchValue = Trim(ValidDefine(DefRow, ValidDefValueCol))
        sSheetName = Trim(ValidDefine(DefRow, ValidSheetNameCol))
        BranchRange = ValidDefine(DefRow, ValidDefBranchFieldCol) + BeginRow + ":" + ValidDefine(DefRow, ValidDefBranchFieldCol) + EndRow
        
        If sSheetName = CurSheetName Then
            For Each CurRange In Target
                If CurRange.Column = Range(BranchRange).Column Then
                    FieldRange = ValidDefine(DefRow, ValidDefFieldCol) + CStr(CurRange.Row)
                    If IsSubStr(UCase(Trim(CurRange.Text)), UCase(BranchValue)) Then
                        Call SetInValidationMode(FieldRange)
                    Else
                        Call SetValidationMode(FieldRange)
                    End If
                    If (Trim(CurRange.Text) = "") Then
                        Call ClearCell(Range(FieldRange))
                    End If
                End If
            Next CurRange
        End If
    Next
End Sub

'****************************************************************
'
'****************************************************************
Public Sub SetFieldValidation(ByVal Target As Range, CurSheetName As String)
    Const ValidSheetNameCol = 0
    Const ValidDefBranchFieldCol = 2
    Const ValidDefBeginRowCol = 3
    Const ValidDefEndRowCol = 4
    Const ValidDefValueCol = 5
    Const ValidDefFieldCol = 7
    Const ValidDefTypeCol = 8
    Const ValidDefMinCol = 9
    Const ValidDefMaxCol = 10
    Const ValidDefListCol = 11
    Const ValidDefPromptCol = 12

    Dim CurRange As Range, DefRow As Integer, BeginRow As String, EndRow As String
    Dim BranchRange As String, BranchValue As String, FieldRange As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sPrompt As String, sSheetName As String
    Dim xType As Excel.XlDVType, sFormula1 As String, sFormula2 As String, sErrPrompt As String, sErrMsg As String
    Dim sInputStr As String
        
    For DefRow = 0 To UBound(RangeDefine) - 1
        BeginRow = RangeDefine(DefRow, ValidDefBeginRowCol)
        EndRow = RangeDefine(DefRow, ValidDefEndRowCol)
        BranchValue = Trim(RangeDefine(DefRow, ValidDefValueCol))
        sDataType = Trim(RangeDefine(DefRow, ValidDefTypeCol))
        sMinVal = Trim(RangeDefine(DefRow, ValidDefMinCol))
        sMaxVal = Trim(RangeDefine(DefRow, ValidDefMaxCol))
        sRangeList = Trim(RangeDefine(DefRow, ValidDefListCol))
        sPrompt = Trim(RangeDefine(DefRow, ValidDefPromptCol))
        sSheetName = Trim(RangeDefine(DefRow, ValidSheetNameCol))
        BranchRange = RangeDefine(DefRow, ValidDefBranchFieldCol) + BeginRow + ":" + RangeDefine(DefRow, ValidDefBranchFieldCol) + EndRow
                          
        If sSheetName = CurSheetName Then
            For Each CurRange In Target
                If CurRange.Column = Range(BranchRange).Column Then  '列号相同
                    If IsSubStr(UCase(Trim(CurRange.Text)), UCase(BranchValue)) Then
                        If (sDataType = "INT") Then
                            xType = xlValidateTextLength
                            If Len(sMinVal) <= Len(sMaxVal) Then
                                sFormula1 = Len(sMinVal)
                                sFormula2 = Len(sMaxVal)
                            Else
                                sFormula1 = Len(sMaxVal)
                                sFormula2 = Len(sMinVal)
                            End If
                            CurRange.NumberFormatLocal = "@"
                            CurRange.HorizontalAlignment = xlRight
                            sInputStr = "[" & sMinVal & ".." & sMaxVal & "]"
'                            xType = xlValidateWholeNumber
'                            sFormula1 = sMinVal
'                            sFormula2 = sMaxVal
                        End If
'                        If (sDataType = "STRING") Then
'                            xType = xlValidateTextLength
'                            sFormula1 = sMinVal
'                            sFormula2 = sMaxVal
'                        End If
'                        If (sDataType = "LIST") Then
'                            xType = xlValidateList
'                            sFormula1 = sRangeList
'                            sFormula2 = ""
'                        End If
                        FieldRange = RangeDefine(DefRow, ValidDefFieldCol) + CStr(CurRange.Row)
                        Call ClearValidation(FieldRange)
                        Call SetValidation(FieldRange, xType, sFormula1, sFormula2, sInputStr, sPrompt)
                    End If
                End If
             Next CurRange
        End If
    Next
End Sub

'****************************************************************
'设置FieldRange指定区域的数据有效性规则
'****************************************************************
Public Sub SetValidation(FieldRange As String, AddType As Long, Formula1 As String, Formula2 As String, InputStr As String, RangeStr As String)
    Dim InputTitle As String
    InputTitle = ""
    If InputStr <> "" Then
        InputTitle = "Range "
    End If

    Dim sErrMsg As String
    sErrMsg = RangeStr + GetCellInfo(Range(FieldRange))
    
    With Range(FieldRange).Validation
        .Delete
        .Add Type:=AddType, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=Formula1, Formula2:=Formula2
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = InputTitle
        .ErrorTitle = "Prompt"
        .InputMessage = RangeStr 'InputStr
        .ErrorMessage = sErrMsg
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Public Function GetCellInfo(ByVal rng As Range) As String
    If rng Is Nothing Then
        GetCellInfo = ""
        Exit Function
    End If
    GetCellInfo = vbLf + "Worksheet=" + rng.Worksheet.Name + "; Column=" + CStr(rng.Column) + "; Row=" + CStr(rng.Row)
End Function

'****************************************************************
'清除FieldRange指定区域的数据有效性规则
'****************************************************************
Public Sub ClearValidation(FieldRange As String)
    Call SetValidation(FieldRange, xlValidateTextLength, "0", "0", "", "No input is required.")
End Sub





