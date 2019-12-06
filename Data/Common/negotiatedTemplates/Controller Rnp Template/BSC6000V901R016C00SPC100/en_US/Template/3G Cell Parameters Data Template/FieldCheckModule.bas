Attribute VB_Name = "FieldCheckModule"
Public m_nDefSheetIndex As Long
Public GeneratingFlag As Integer
Public Const TableEndRow = 65535 + 1
Public Const TableBeginRow = 3
Public Const TblCols = 17
Public Const DefTblRows = 500
Public Const TableBeginCol = 1
'Private TableEndCol As Long
Const TableEndCol = 500
Public ValidDefine(150, 9) As String
Public RangeDefine(500, 13) As String
Public SheetDefine(DefTblRows, TblCols) As String

'****************************************************************
'检查Target指定区域输入的值是否符合数据有效性规则
'****************************************************************
Public Sub CheckFieldData(ByVal nDefSheetIndex As Long, ByVal Target As Range)
    Dim nResponse As Integer
    Dim sColType As String, sErrPrompt As String, sErrMsg As String
    Dim FieldRange As Range
    
    If Target.Count = 1 Then
        sErrMsg = CheckOneFieldData(nDefSheetIndex, Target)
        If sErrMsg <> "" Then
            sErrPrompt = "Prompt"
            'sErrMsg = sErrMsg + GetCellInfo(Target)
            nResponse = MsgBox(sErrMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, sErrPrompt)
            If nResponse = vbRetry Then
                Target.Cells(1, 1).Select
            End If
            Target.Cells(1, 1).ClearContents
        End If
    End If
   
    If Target.Count > 1 Then
        For Each FieldRange In Target
            sErrMsg = CheckOneFieldData(nDefSheetIndex, FieldRange)
            If sErrMsg <> "" Then
                sErrPrompt = "Prompt"
                'sErrMsg = sErrMsg + GetCellInfo(Target)
                MsgBox sErrMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, sErrPrompt
                FieldRange.ClearContents
            End If
        Next
    End If
End Sub
'****************************************************************
'检查Cell Sheet页的Target指定区域输入的值是否符合数据有效性规则
'****************************************************************
Public Sub CellCheckFieldData(ByVal nDefSheetIndex As Long, ByVal Target As Range)
    Dim nResponse As Integer
    Dim sColType As String, sErrPrompt As String, sErrMsg As String
    Dim FieldRange As Range
    
    If Target.Count = 1 Then
        sErrMsg = CellCheckOneFieldData(nDefSheetIndex, Target)
        If sErrMsg <> "" Then
            sErrPrompt = "Prompt"
            'sErrMsg = sErrMsg + GetCellInfo(Target)
            nResponse = MsgBox(sErrMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, sErrPrompt)
            If nResponse = vbRetry Then
                Target.Cells(1, 1).Select
            End If
            Target.Cells(1, 1).ClearContents
        End If
    End If
   
    If Target.Count > 1 Then
        For Each FieldRange In Target
            sErrMsg = CellCheckOneFieldData(nDefSheetIndex, FieldRange)
            If sErrMsg <> "" Then
                sErrPrompt = "Prompt"
                'sErrMsg = sErrMsg + GetCellInfo(Target)
                MsgBox sErrMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, sErrPrompt
                FieldRange.ClearContents
            End If
        Next
    End If
End Sub

Public Function CheckOneFieldData(ByVal nDefSheetIndex As Long, ByVal FieldRange As Range) As String
    Dim nValue As Double, nStrLen As Double, nLoop As Integer
    Dim sValue As String, sItem As String
    Dim bFlag As Boolean
    Dim sColType As String, sFeildName As String, sMinVal As String, sMaxVal As String, sErrMsg As String
    Dim nMinVal As Double, nMaxVal As Double
    Dim TableInfoSheet As Worksheet
    Dim CurSheet As Worksheet
    Dim TypeRowNo As Integer
    Dim RowNo As Integer
    
    Dim sRangeStr As String, sRangeUnit As String
    Dim nPos As Integer, nPriPos As Integer, nPosUnit As Integer, j As Integer
    Dim RangeUnit As New Collection
    Dim RangeInst As RangeInfo
    
    Set TableInfoSheet = ThisWorkbook.Sheets("TableInfo")
    Set CurSheet = ActiveSheet
    RowNo = 5
     
    For RowNo = 5 To 3000
       If CurSheet.name = Trim(TableInfoSheet.Cells(RowNo, 1).Value) And CurSheet.Cells(1, FieldRange.Column) = Trim(TableInfoSheet.Cells(RowNo, 2).Value) Then
          TypeRowNo = RowNo
       End If
    Next RowNo
    
    sColType = Trim(TableInfoSheet.Cells(TypeRowNo, 3).Value)
    
    If sColType = "INT" Then
        sFeildName = Trim(TableInfoSheet.Cells(TypeRowNo, 2).Value)
        sRangeStr = Trim(TableInfoSheet.Cells(TypeRowNo, 6).Value)
        sValue = FieldRange.Text
    
        If sValue <> "" Then
            bFlag = True
            nStrLen = Len(sValue)
        
            For nLoop = 1 To nStrLen
                sItem = Right(Left(sValue, nLoop), 1)
                If sItem < "0" Or sItem > "9" Then
                    If nLoop <> 1 Then
                        bFlag = False
                        Exit For
                    Else
                        If sItem <> "-" Then
                            bFlag = False
                            Exit For
                        End If
                    End If
                End If
            Next
    
            If bFlag Then
                nValue = CDbl(sValue)
                
                nPos = InStr(1, sRangeStr, ",", vbTextCompare)
                If nPos = 0 Then
                    nPos = InStr(1, sRangeStr, "..", vbTextCompare)
                    sMinVal = Left(sRangeStr, nPos - 1)
                    sMaxVal = Right(sRangeStr, Len(sRangeStr) - nPos - 1)
                    nMinVal = CDbl(sMinVal)
                    nMaxVal = CDbl(sMaxVal)
                    If nValue < nMinVal Or nValue > nMaxVal Then
                        bFlag = False
                    End If
                Else
                    nPriPos = 1
                    Do While nPos <> 0
                        sRangeUnit = Left(sRangeStr, nPos - 1)
                        sRangeUnit = Right(sRangeUnit, Len(sRangeUnit) - (nPriPos - 1))
                        nPosUnit = InStr(1, sRangeUnit, "..", vbTextCompare)

                        Set RangeInst = New RangeInfo
                        If (nPosUnit = 0) Then
                            RangeInst.sMinVal = sRangeUnit
                            RangeInst.sMaxVal = sRangeUnit
                        Else
                            RangeInst.sMinVal = Left(sRangeUnit, nPosUnit - 1)
                            RangeInst.sMaxVal = Right(sRangeUnit, Len(sRangeUnit) - nPosUnit - 1)
                        End If
                        RangeUnit.Add RangeInst

                        nPriPos = nPos + 1
                        nPos = InStr(nPriPos, sRangeStr, ",", vbTextCompare)
                    Loop
                    sRangeUnit = Right(sRangeStr, Len(sRangeStr) - (nPriPos - 1))
                    nPosUnit = InStr(1, sRangeUnit, "..", vbTextCompare)
                    Set RangeInst = New RangeInfo
                        If (nPosUnit = 0) Then
                            RangeInst.sMinVal = sRangeUnit
                            RangeInst.sMaxVal = sRangeUnit
                        Else
                            RangeInst.sMinVal = Left(sRangeUnit, nPosUnit)
                            RangeInst.sMaxVal = Right(sRangeUnit, Len(sRangeUnit) - nPosUnit - 1)
                        End If
                    RangeUnit.Add RangeInst

                    bFlag = False
                    For j = 1 To RangeUnit.Count
                        Set RangeInst = RangeUnit.Item(j)
                        nMinVal = CDbl(RangeInst.sMinVal)
                        nMaxVal = CDbl(RangeInst.sMaxVal)
                        If nValue >= nMinVal And nValue <= nMaxVal Then
                            bFlag = True
                            Exit For
                        End If
                    Next
                End If
            End If
            
            CheckOneFieldData = ""
            If Not bFlag Then
                sErrMsg = "Range [" & sRangeStr & "]"
                sErrMsg = sErrMsg + vbLf + "Worksheet = " + CurSheet.name + "; FieldName = " + sFeildName + "; Column = " + CStr(FieldRange.Column)
                CheckOneFieldData = sErrMsg
            End If
        End If
        
    End If
        
    If sColType = "BITMAP" Then
            sValue = FieldRange.Text
            If sValue <> "" Then
                bFlag = True
                nStrLen = Len(sValue)
            
                For nLoop = 1 To nStrLen
                    sItem = Right(Left(sValue, nLoop), 1)
                    If sItem < "0" Or sItem > "1" Then
                        bFlag = False
                        Exit For
                    End If
                    If CInt(sItem) < 0 Or CInt(sItem) > 1 Then
                        bFlag = False
                        Exit For
                    End If
                Next
        
                CheckOneFieldData = ""
                If Not bFlag Then
                    sErrMsg = "Input Range [0,1]"
                    sErrMsg = sErrMsg + vbLf + "Worksheet = " + CurSheet.name + "; FieldName = " + sFeildName + "; Column = " + CStr(FieldRange.Column)
                    CheckOneFieldData = sErrMsg
                End If
            End If
        
    End If
End Function
'****************************************************************
'检查Cell Sheet页的单元格输入的值是否符合数据有效性规则
'****************************************************************
Public Function CellCheckOneFieldData(ByVal nDefSheetIndex As Long, ByVal FieldRange As Range) As String
    Dim nValue As Double, nStrLen As Integer, nLoop As Integer
    Dim sValue As String, sItem As String
    Dim bFlag As Boolean
    Dim sColType As String, sFeildName As String, sMinVal As String, sMaxVal As String
    Dim nMinVal As Double, nMaxVal As Double
    Dim sSheetName As String, sBandValue As String
    Dim BranchRange As String, BranchValue As String, BFieldRange As String, sBSheetName As String
    Dim nBandColumn As Integer, sErrMsg As String
    Dim TableInfoSheet As Worksheet
    Dim ValidInfoSheet As Worksheet
    Dim TypeRowNo As Integer
    Dim RowNo As Integer
    
    Const ValidSheetNameCol = 1
    Const ValidDefBranchFieldCol = 3
    Const ValidDefValueCol = 6
    Const ValidDefFieldCol = 8
    Const ValidDefMinCol = 10
    Const ValidDefMaxCol = 11
    
    Set TableInfoSheet = ThisWorkbook.Sheets("TableInfo")
    Set ValidInfoSheet = ThisWorkbook.Sheets("ValidInfo")
    
    Set CurSheet = ActiveSheet
    RowNo = 5
     
    For RowNo = 5 To 3000
       If CurSheet.name = Trim(TableInfoSheet.Cells(RowNo, 1).Value) And CurSheet.Cells(1, FieldRange.Column) = Trim(TableInfoSheet.Cells(RowNo, 2).Value) Then
          TypeRowNo = RowNo
       End If
    Next RowNo
    
    sColType = Trim(TableInfoSheet.Cells(TypeRowNo, 3).Value)
    
    
    If sColType = "INT" Then
        sFeildName = Trim(TableInfoSheet.Cells(TypeRowNo, 2).Value)
        sSheetName = CurSheet.name
        bRangeFlag = False
        sBandValue = ""
        
        If sSheetName = "CELL" And (FieldRange.Column = 7 Or FieldRange.Column = 8) Then
            nBandColumn = 6
            sBandValue = UCase(Cells(FieldRange.Row, nBandColumn).Value)
        End If
        
        If sBandValue <> "" Then
            For nLoop = 2 To 1000
                sBSheetName = Trim(ValidInfoSheet.Cells(nLoop, 1))
                sBranchValue = Trim(ValidInfoSheet.Cells(nLoop, 6))
                BranchRange = ValidInfoSheet.Cells(nLoop, 3).Value & ":" & ValidInfoSheet.Cells(nLoop, 3).Value
                BFieldRange = ValidInfoSheet.Cells(nLoop, 8).Value & ":" & ValidInfoSheet.Cells(nLoop, 8).Value
                
                If sSheetName = sBSheetName And Range(BranchRange).Column = nBandColumn And Range(BFieldRange).Column = FieldRange.Column And sBranchValue = sBandValue Then
                    sMinVal = Trim(ValidInfoSheet.Cells(nLoop, 10))
                    sMaxVal = Trim(ValidInfoSheet.Cells(nLoop, 11))
                    Exit For
                End If
            Next
        Else
            sMinVal = Trim(TableInfoSheet.Cells(TypeRowNo, 4))
            sMaxVal = Trim(TableInfoSheet.Cells(TypeRowNo, 5))
        End If

        nMinVal = CDbl(sMinVal)
        nMaxVal = CDbl(sMaxVal)
        sValue = FieldRange.Text
    
        If sValue <> "" Then
            bFlag = True
            
            If (sFeildName = "URAIDS") Then
                nStrLen = Len(sValue)
                
                For nLoop = 1 To nStrLen
                    sItem = Right(Left(sValue, nLoop), 1)
                    If sItem < "0" Or sItem > "9" Then
                        If sItem <> ";" Then
                            bFlag = False
                            Exit For
                        End If
                    End If
                Next
                
                If bFlag Then
                    Dim sValueOne As String
                    Dim nPos As Integer, nPPos As Integer
                    
                    nPPos = 1
                    nPos = InStr(nPPos, sValue, ";", vbTextCompare)
                    If (nPos = 0) Then
                        nValue = CDbl(sValue)
                        If nValue < nMinVal Or nValue > nMaxVal Then
                            bFlag = False
                        End If
                    Else
                        Do While nPos <> 0
                            sValueOne = Left(sValue, nPos - 1)
                            sValueOne = Right(sValueOne, Len(sValueOne) - (nPPos - 1))
                            nValue = CDbl(sValueOne)
                            If nValue < nMinVal Or nValue > nMaxVal Then
                                bFlag = False
                                Exit Do
                            End If
                            
                            nPPos = nPos + 1
                            nPos = InStr(nPPos, sValue, ";", vbTextCompare)
                        Loop
                        
                        If bFlag Then
                            sValueOne = Right(sValue, Len(sValue) - (nPPos - 1))
                            nValue = CDbl(sValueOne)
                            If nValue < nMinVal Or nValue > nMaxVal Then
                                bFlag = False
                            End If
                        End If
                    End If
                End If
            Else
                nStrLen = Len(sValue)
            
                For nLoop = 1 To nStrLen
                    sItem = Right(Left(sValue, nLoop), 1)
                    If sItem < "0" Or sItem > "9" Then
                        If nLoop <> 1 Then
                            bFlag = False
                            Exit For
                        Else
                            If sItem <> "-" Then
                                bFlag = False
                                Exit For
                            End If
                        End If
                    End If
                Next
        
                If bFlag Then
                    nValue = CDbl(sValue)
                    If nValue < nMinVal Or nValue > nMaxVal Then
                        bFlag = False
                    End If
                    If sFeildName = "LAC" And nValue = (nMaxVal - 1) Then
                        bFlag = False
                    End If
                End If
            End If
            
            CellCheckOneFieldData = ""
            If Not bFlag Then
                If sFeildName = "LAC" Then
                    sErrMsg = "Range [1..65533, 65535]"
                Else
                    sErrMsg = "Range [" & sMinVal & ".." & sMaxVal & "]"
                End If
                sErrMsg = sErrMsg + vbLf + "Worksheet = " + CurSheet.name + "; FieldName = " + sFeildName + "; Column = " + CStr(FieldRange.Column)
                CellCheckOneFieldData = sErrMsg
            End If
        End If
        
    
      ElseIf sColType = "BITMAP" Then
            sValue = FieldRange.Text
            If sValue <> "" Then
                bFlag = True
                nStrLen = Len(sValue)
            
                For nLoop = 1 To nStrLen
                    sItem = Right(Left(sValue, nLoop), 1)
                    If sItem < "0" Or sItem > "1" Then
                        bFlag = False
                        Exit For
                    End If
                    If CInt(sItem) < 0 Or CInt(sItem) > 1 Then
                        bFlag = False
                        Exit For
                    End If
                Next
        
                CellCheckOneFieldData = ""
                If Not bFlag Then
                    sErrMsg = "Input Range [0,1]"
                    sErrMsg = sErrMsg + vbLf + "Worksheet = " + CurSheet.name + "; FieldName = " + sFeildName + "; Column = " + CStr(FieldRange.Column)
                    CellCheckOneFieldData = sErrMsg
                End If
            End If
        
    End If
    
End Function
'****************************************************************
'从ValidDef取所有定义数据
'****************************************************************
Public Sub GetValidDefineData()
    Dim CurSheet As Worksheet
    Dim iRow As Integer, iCol As Integer
    
    If ValidDefine(0, 0) <> "" And RangeDefine(0, 0) <> "" Then
        Exit Sub
    End If
    
    Set CurSheet = Sheets("ValidDef")
    
    For iRow = 0 To ValidRows - 1
        For iCol = 0 To ValidCols - 1
            ValidDefine(iRow, iCol) = CurSheet.Cells(InvalidBeginRow + iRow, InvalidBeginCol + iCol)
        Next
    Next
    
    For iRow = 0 To RangeRows - 1
        For iCol = 0 To RangeCols - 1
            RangeDefine(iRow, iCol) = CurSheet.Cells(RangeBeginRow + iRow, RangeBeginCol + iCol)
        Next
    Next
End Sub

'****************************************************************
'
'****************************************************************
Public Sub SetFieldValidation(ByVal Target As Range, CurSheetName As String)
    Const ValidSheetNameCol = 1
    Const ValidDefBranchFieldCol = 3
    Const ValidDefBeginRowCol = 4
    Const ValidDefEndRowCol = 5
    Const ValidDefValueCol = 6
    Const ValidDefFieldCol = 8
    Const ValidDefTypeCol = 9
    Const ValidDefMinCol = 10
    Const ValidDefMaxCol = 11
    Const ValidDefListCol = 12
    Const ValidDefPromptCol = 13
    
    Dim ValidInfoSheet As Worksheet
    Dim CurRange As Range, DefRow As Integer, BeginRow As String, EndRow As String
    Dim BranchRange As String, BranchValue As String, FieldRange As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sPrompt As String, sSheetName As String
    Dim xType As Excel.XlDVType, sFormula1 As String, sFormula2 As String, sErrPrompt As String, sErrMsg As String
    Dim sInputStr As String

    
    
    Set ValidInfoSheet = ThisWorkbook.Sheets("ValidInfo")
        
    For DefRow = 2 To 23
        BeginRow = ValidInfoSheet.Cells(DefRow, ValidDefBeginRowCol).Value
        EndRow = ValidInfoSheet.Cells(DefRow, ValidDefEndRowCol).Value
        BranchValue = Trim(ValidInfoSheet.Cells(DefRow, ValidDefValueCol).Value)
        sDataType = Trim(ValidInfoSheet.Cells(DefRow, ValidDefTypeCol).Value)
        sMinVal = Trim(ValidInfoSheet.Cells(DefRow, ValidDefMinCol).Value)
        sMaxVal = Trim(ValidInfoSheet.Cells(DefRow, ValidDefMaxCol).Value)
        sRangeList = Trim(ValidInfoSheet.Cells(DefRow, ValidDefListCol).Value)
        sPrompt = Trim(ValidInfoSheet.Cells(DefRow, ValidDefPromptCol).Value)
        sSheetName = Trim(ValidInfoSheet.Cells(DefRow, ValidSheetNameCol).Value)
        BranchRange = ValidInfoSheet.Cells(DefRow, ValidDefBranchFieldCol).Value + BeginRow + ":" + ValidInfoSheet.Cells(DefRow, ValidDefBranchFieldCol).Value + EndRow
                          
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
                            CurRange.HorizontalAlignment = xlCenter
                            sInputStr = "[" & sMinVal & ".." & sMaxVal & "]"
                        End If
                        FieldRange = ValidInfoSheet.Cells(DefRow, ValidDefFieldCol).Value + CStr(CurRange.Row)
                        Call ClearValidation(FieldRange)
                        Call SetValidation(FieldRange, xType, sFormula1, sFormula2, sInputStr, sPrompt)
                    End If
                End If
             Next CurRange
        End If
    Next
End Sub

'****************************************************************
'返回当前Sheet页的有效列数
'****************************************************************
Public Function GetSheetColCount(sRange As String) As Integer
    Dim sRight As String
    
    sRight = Right(sRange, Len(sRange) - InStr(sRange, ":"))
    If Len(sRight) = 1 Then
        GetSheetColCount = InStr(gRangeStr, sRight)
    ElseIf Len(sRight) = 2 Then
        GetSheetColCount = InStr(gRangeStr, Left(sRight, 1)) * Len(gRangeStr) + InStr(gRangeStr, Right(sRight, 1))
    Else
        GetSheetColCount = -1
    End If
End Function
'判断参数Target指定区域的单元格是否为灰色不可用状态,是则清空该单元格输入的值
Public Sub Ensure_NoValue(ByVal Target As Range)
    Const SolidColorIdx = 16
    Const SolidPattern = xlGray16
    Dim CurRange As Range
    
    For Each CurRange In Target
        If CurRange.Cells(1, 1) <> "" And CurRange.Cells(1, 1).Interior.ColorIndex = SolidColorIdx And CurRange.Cells(1, 1).Interior.Pattern = SolidPattern Then
            NoValueNeeded ("B" + CStr(CurRange.Row))
            CurRange.Cells(1, 1).ClearContents
            CurRange.Select
        End If
    Next CurRange
End Sub
'****************************************************************
'设置与Target指定单元格相关的其他区域是否为灰色
'****************************************************************
Public Sub SetInvalidateField(ByVal Target As Range, CurSheetName As String)
    Const ValidSheetNameCol = 1
    Const ValidDefBranchFieldCol = 5
    Const ValidDefBeginRowCol = 3
    Const ValidDefEndRowCol = 4
    Const ValidDefValueCol = 5
    Const ValidDefFieldCol = 7
    Const ValidDefValidCol = 8
    Const ValidContrlFieldRaw = 100
    'color
    Const SolidColorIdx = 16
    Const SolidPattern = xlSolid
    Const NullPattern = xlNone
    
    Dim CurRange As Range, DefRow As Integer, BeginRow As String, EndRow As String
    Dim sSheetName As String
    Dim ValidInfoSheet As Worksheet
    Dim CurSheet As Worksheet
    Dim BranchValue As String
    Dim BranchRange As Integer
    Dim FieldRange As Range
    Dim sBranchLevel As String

    Set ValidInfoSheet = ThisWorkbook.Sheets("ValidInfo")
    
    '先设置主控字段影响的受控字段为有效（空白）
    For DefRow = 25 To 1000
        BranchValue = Trim(ValidInfoSheet.Cells(DefRow, 4).Value)
        sSheetName = Trim(ValidInfoSheet.Cells(DefRow, 1).Value)
        BranchRange = ValidInfoSheet.Cells(DefRow, 5).Value
        sBranchLevel = Trim(ValidInfoSheet.Cells(DefRow, 7).Value)
        If ValidInfoSheet.Cells(DefRow, 5).Value <> "" Then
          If sSheetName = CurSheetName Then
            Set CurSheet = ThisWorkbook.Sheets(CurSheetName)
            For Each CurRange In Target
                If CurRange.Column = BranchRange Then
                   If ValidInfoSheet.Cells(DefRow, 6).Value <> "" And sBranchLevel <> "2" Then
                      Call SetValidationMode(Range(CurSheet.Cells(CurRange.Row, ValidInfoSheet.Cells(DefRow, 6).Value), CurSheet.Cells(CurRange.Row, ValidInfoSheet.Cells(DefRow, 6).Value)))
                    
                      If (Trim(CurRange.Text) = "") Then
                          Call ClearCell(Range(CurSheet.Cells(CurRange.Row, ValidInfoSheet.Cells(DefRow, 6).Value), CurSheet.Cells(CurRange.Row, ValidInfoSheet.Cells(DefRow, 6).Value)))
                      End If
                    End If
                End If
            Next CurRange
          End If
        End If
    Next
    
    '再根据主控字段的数据设置相关的受控字段为无效（灰色）
    For DefRow = 25 To 1000
        BranchValue = Trim(ValidInfoSheet.Cells(DefRow, 4).Value)
        sSheetName = Trim(ValidInfoSheet.Cells(DefRow, 1).Value)
        BranchRange = ValidInfoSheet.Cells(DefRow, 5).Value
        sBranchLevel = Trim(ValidInfoSheet.Cells(DefRow, 7).Value)
        If ValidInfoSheet.Cells(DefRow, 5).Value <> "" Then
          If sSheetName = CurSheetName Then
            Set CurSheet = ThisWorkbook.Sheets(CurSheetName)
            For Each CurRange In Target
                If CurRange.Column = BranchRange Then
                   If ValidInfoSheet.Cells(DefRow, 6).Value <> "" Then
                      If IsSubStr(UCase(Trim(CurRange.Text)), UCase(BranchValue)) Then
                          Call SetInValidationMode(Range(CurSheet.Cells(CurRange.Row, ValidInfoSheet.Cells(DefRow, 6).Value), CurSheet.Cells(CurRange.Row, ValidInfoSheet.Cells(DefRow, 6).Value)))
                      End If
                    
                      If (Trim(CurRange.Text) = "") Then
                          Call ClearCell(Range(CurSheet.Cells(CurRange.Row, ValidInfoSheet.Cells(DefRow, 6).Value), CurSheet.Cells(CurRange.Row, ValidInfoSheet.Cells(DefRow, 6).Value)))
                      End If
                    End If
                End If
            Next CurRange
          End If
        End If
    Next
    
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
'设置FieldRange指定区域变为灰色
'****************************************************************
'Public Sub SetInValidationMode(FieldRange As String)
   ' Range(FieldRange).Interior.ColorIndex = 16 'SolidColorIdx
    'Range(FieldRange).Interior.Pattern = 17 'SolidPattern
    'Call ClearCell(Range(FieldRange))
'End Sub
'****************************************************************
'设置FieldRange指定区域变为灰色
'****************************************************************
Public Sub SetInValidationMode(ByVal FieldRange As Range)
    FieldRange.Interior.ColorIndex = 16 'SolidColorIdx
    FieldRange.Interior.Pattern = 17 'SolidPattern
    Call ClearCell(FieldRange)
End Sub

'****************************************************************
'设置FieldRange指定区域为空白
'****************************************************************
Public Sub SetValidationMode(ByVal FieldRange As Range)
    FieldRange.Interior.ColorIndex = xlNone
    FieldRange.Interior.Pattern = xlNone
End Sub
'****************************************************************
'清除CurCell指定单元格的值
'****************************************************************
Public Sub ClearCell(CurCell As Range)
    If (Trim(CurCell.Value) <> "") Then
        CurCell.Value = ""
    End If
End Sub
Public Function GetCellInfo(ByVal rng As Range) As String
    If rng Is Nothing Then
        GetCellInfo = ""
        Exit Function
    End If
    GetCellInfo = vbLf + "Worksheet=" + rng.Worksheet.name + "; Column=" + CStr(rng.Column) + "; Row=" + CStr(rng.Row)
End Function


'Private Sub SetFieldValidate(CurSheet As Worksheet, iRow As Integer, sTableRange As String)
'****************************************************************
'设置单元格格式和数据有效性
'****************************************************************
Public Sub SetFieldValidate(CurSheet As Worksheet, iCol As Integer)
    Const FieldNameCol = 2
    Const DataTypeCol = 3
    Const MinValCol = 4
    Const MaxValCol = 5
    Const RangeListCol = 6
    Dim sFieldName As String
    Dim TypeRowNo As Integer
    Dim RowNo As Integer
    Dim FieldCol As Integer
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String
    Dim xType As Excel.XlDVType, sFormula1 As String, sFormula2 As String, sErrPrompt As String, sErrMsg As String
    Dim nMinVal As Double, nMaxVal As Double
    Dim TableInfoSheet As Worksheet
    
    Set TableInfoSheet = ThisWorkbook.Sheets("TableInfo")
    'Set CurSheet = ActiveSheet
    RowNo = 5
    FieldCol = iCol
    For RowNo = 5 To 3000
       If CurSheet.name = Trim(TableInfoSheet.Cells(RowNo, 1).Value) And CurSheet.Cells(1, FieldCol) = Trim(TableInfoSheet.Cells(RowNo, 2).Value) Then
          TypeRowNo = RowNo
          Exit For
       End If
    Next RowNo

    sColType = Trim(TableInfoSheet.Cells(TypeRowNo, 3).Value)
    
    sFieldName = Trim(TableInfoSheet.Cells(TypeRowNo, FieldNameCol).Value)
    sDataType = Trim(TableInfoSheet.Cells(TypeRowNo, DataTypeCol).Value)
    sMinVal = Trim(TableInfoSheet.Cells(TypeRowNo, MinValCol).Value)
    sMaxVal = Trim(TableInfoSheet.Cells(TypeRowNo, MaxValCol).Value)
    sRangeList = Trim(TableInfoSheet.Cells(TypeRowNo, RangeListCol).Value)
    
    sErrPrompt = "Prompt"
    If (sDataType = "STRING" Or sDataType = "BITMAP") Then
        xType = xlValidateTextLength
        sFormula1 = sMinVal
        sFormula2 = sMaxVal
        CurSheet.Columns(FieldCol).NumberFormatLocal = "@"
        CurSheet.Columns(FieldCol).HorizontalAlignment = xlCenter
        CurSheet.Cells(4, FieldCol).HorizontalAlignment = xlCenter
        
        If (sDataType = "STRING") Then
            sErrMsg = "Length [" + sMinVal + ".." + sMaxVal + "]"
        Else
            sErrMsg = "Length [" + sMaxVal + "]"
        End If
    End If
    
    If (sDataType = "LIST") Then
        xType = xlValidateList
        sFormula1 = sRangeList
        sFormula2 = ""
        CurSheet.Columns(FieldCol).NumberFormatLocal = "@"
        CurSheet.Columns(FieldCol).HorizontalAlignment = xlCenter
        CurSheet.Cells(4, FieldCol).HorizontalAlignment = xlCenter
        
        sErrMsg = "Range [" + Left(sRangeList, 100) + "...]"
    Else
        CurSheet.Columns(FieldCol).NumberFormatLocal = "@"
        CurSheet.Columns(FieldCol).HorizontalAlignment = xlCenter
        CurSheet.Cells(4, FieldCol).HorizontalAlignment = xlCenter
    End If
    
'    If sFieldName = "LAC" Then
'        xType = xlValidateCustom
'        sFormula1 = "=OR(AND(" + FieldCol + "1<=65533," + FieldCol + "1>0), " + FieldCol + "1 = 65535)"
'        sFormula2 = ""
'    End If
    
    
    'sErrPrompt = GetValidErrTitle(sDataType)
    'sErrMsg = GetRangeInfo(iRow)
    
    CurSheet.Select
    Columns(FieldCol).Select
    sErrMsg = sErrMsg + vbLf + "Worksheet = " + CurSheet.name + "; FieldName = " + sFieldName + "; Column = " + CStr(FieldCol)
    Call SetDataValidate(xType, sFormula1, sFormula2, sErrPrompt, sErrMsg)
       
    '去掉第1，2行的字段有效性校验
    If Cells(1, FieldCol).Text <> "" Then
        Cells(1, FieldCol).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
            :=xlBetween
            .IgnoreBlank = True
            .InCellDropdown = True
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
    
        Cells(2, FieldCol).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
            :=xlBetween
            .IgnoreBlank = True
            .InCellDropdown = True
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
    End If
End Sub

'****************************************************************
'设置指定Range的数据有效性
'****************************************************************
Private Sub SetDataValidate(xType As Excel.XlDVType, sFormula1 As String, sFormula2 As String, sErrPrompt As String, sErrMsg As String)
    With Selection.Validation
        .Delete
        If Trim(sFormula2) = "" Then
            .Add Type:=xType, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=sFormula1
        Else
            .Add Type:=xType, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=sFormula1, Formula2:=sFormula2
        End If
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = sErrPrompt
        .InputMessage = ""
        .ErrorMessage = sErrMsg
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub


'****************************************************************
'获取有效范围提示或者错误信息
'****************************************************************
Public Function GetRangeInfo(iRow As Integer) As String
    Const FieldNameCol = 2
    Const DataTypeCol = 3
    Const MinValCol = 5
    Const MaxValCol = 6
    Const RangeListCol = 7
    Const ValueTypeCol = 15
    
    Dim sFieldName As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sValueType As String
    
    sFieldName = Trim(SheetDefine(iRow, FieldNameCol))
    sDataType = Trim(SheetDefine(iRow, DataTypeCol))
    sMinVal = Trim(SheetDefine(iRow, MinValCol))
    sMaxVal = Trim(SheetDefine(iRow, MaxValCol))
    sRangeList = Trim(SheetDefine(iRow, RangeListCol))
    sValueType = Trim(SheetDefine(iRow, ValueTypeCol))
    
    GetRangeInfo = ""
    If (sDataType = FINT) Or (sDataType = FSTRING) Then
        If sMinVal = sMaxVal Then
            GetRangeInfo = GetValidErrMsg(sDataType) + "[" + sMinVal + "]"
        Else
            GetRangeInfo = GetValidErrMsg(sDataType) + "[" + sMinVal + ".." + sMaxVal + "]"
        End If
    End If
    
    If sDataType = FLIST Then
        GetRangeInfo = GetValidErrMsg(sDataType) + "[" + sRangeList + "]"
    End If
    
    If Trim(sValueType) = "ATM" Then
        GetRangeInfo = GetRangeInfo + vbCrLf + " Note: Must begin with H'. "
    End If
    
    If Trim(sFieldName) = "LAC" Then
        GetRangeInfo = GetValidErrMsg(sDataType) + "[1..65533,65535]"
    End If
End Function

'****************************************************************
'清除FieldRange指定区域的数据有效性规则
'****************************************************************
Public Sub ClearValidation(FieldRange As String)
    Call SetValidation(FieldRange, xlValidateTextLength, "0", "0", "", "No input is required.")
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


