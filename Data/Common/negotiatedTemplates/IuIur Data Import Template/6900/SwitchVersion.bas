Attribute VB_Name = "SwitchVersion"
Option Explicit

Public bIsEng As Boolean  '用于控制设置中英文

Private Sub SwitchVersion()
    Dim iNodebcount As Integer
    Dim tempShtNodeB As Worksheet, tempSetSumShtWithNodeB As Worksheet
    GeneratingFlag = 1
    Call SetCoverSht
    
    Call UnprotectWorkBook
    Call SetSheetProtected(False)
    Call GetSheetDefineData
    
    Call SetNegotiatedData(Sheets("COMMON"))
    Call SetNegotiatedData(Sheets("IUCS"))
    Call SetNegotiatedData(Sheets("IUPS"))
    Call SetNegotiatedData(Sheets("IUR"))
    
    Call SetSheetProtected(True)
    Call ProtectWorkBook
    GeneratingFlag = 0
    
End Sub
Public Sub SwitchEnglishVersion()
    bIsEng = True
    Sheets("TableDef").Range("P10") = CStr(bIsEng)
    
    CommandBars("Operate Bar").Controls.Item(1).Caption = "&English Version&"
    CommandBars("Operate Bar").Controls.Item(1).TooltipText = "English Version"
    
    CommandBars("Operate Bar").Controls.Item(2).Caption = "&Chinese Version&"
    CommandBars("Operate Bar").Controls.Item(2).TooltipText = "Chinese Version"
    
    CommandBars("Operate Bar").Controls.Item(3).Caption = "&Hide Empty Row"
    CommandBars("Operate Bar").Controls.Item(3).TooltipText = "Hide Empty Row"
    
    CommandBars("Operate Bar").Controls.Item(4).Caption = "&Reset Row"
    CommandBars("Operate Bar").Controls.Item(4).TooltipText = "Reset Row"
    
    CommandBars("Operate Bar").Controls.Item(5).Caption = "&Expand Row"
    CommandBars("Operate Bar").Controls.Item(5).TooltipText = "Expand Row"
      
    Call SwitchVersion
End Sub

Public Sub SwitchChineseVersion()
    bIsEng = False
    Sheets("TableDef").Range("P10") = CStr(bIsEng)
    
    CommandBars("Operate Bar").Controls.Item(1).Caption = Sheets("TableDef").Range("P4").Text
    CommandBars("Operate Bar").Controls.Item(1).TooltipText = Sheets("TableDef").Range("P4").Text
    
    CommandBars("Operate Bar").Controls.Item(2).Caption = Sheets("TableDef").Range("P5").Text
    CommandBars("Operate Bar").Controls.Item(2).TooltipText = Sheets("TableDef").Range("P5").Text
    
    CommandBars("Operate Bar").Controls.Item(3).Caption = Sheets("TableDef").Range("P6").Text
    CommandBars("Operate Bar").Controls.Item(3).TooltipText = Sheets("TableDef").Range("P6").Text
    
    CommandBars("Operate Bar").Controls.Item(4).Caption = Sheets("TableDef").Range("P7").Text
    CommandBars("Operate Bar").Controls.Item(4).TooltipText = Sheets("TableDef").Range("P7").Text
    
    CommandBars("Operate Bar").Controls.Item(5).Caption = Sheets("TableDef").Range("R8").Text
    CommandBars("Operate Bar").Controls.Item(5).TooltipText = Sheets("TableDef").Range("R8").Text
                
    Call SwitchVersion
End Sub
    


'******************************************************************************
'设置中英文注释转换
'******************************************************************************
Private Sub SetPostil(CurSheet As Worksheet, iRow As Integer)
    Dim FieldPostil As String, FieldCol As String
    Dim ENGName As String, CHSName As String, RangeName As String, FieldRow As String
    Dim sQuarryType As String
    
    FieldRow = Trim(SheetDefine(iRow, iTitleBeginRow))
    ENGName = Trim(SheetDefine(iRow, iFieldPostil))
    CHSName = Trim(SheetDefine(iRow, iFieldDisplayName_CHS))
    If Left(CHSName, 1) = "*" Then
        CHSName = Right(CHSName, Len(CHSName) - 1)
    End If
    RangeName = GetRangeInfo(iRow)
    sQuarryType = Trim(SheetDefine(iRow, iQuarry))
    If bIsEng Then
        FieldPostil = ENGName + Chr(10) + "(" + RangeName + ")"
    Else
        FieldPostil = CHSName + Chr(10) + "(" + RangeName + ")"
    End If
    If bIsEng And RangeName = "" Then
        FieldPostil = ENGName
    End If
    If Not bIsEng And RangeName = "" Then
        FieldPostil = CHSName
    End If
    If bIsEng Then
        If sQuarryType = "1" Then
            FieldPostil = FieldPostil + Chr(10) + "Quarry:From Internal Planning"
        End If
        If sQuarryType = "2" Then
            FieldPostil = FieldPostil + Chr(10) + "Quarry:Negotiated with the Peer End"
        End If
        If sQuarryType = "3" Then
            FieldPostil = FieldPostil + Chr(10) + "Quarry:From Network Planning"
        End If
    Else
        If sQuarryType = "1" Then
            FieldPostil = FieldPostil + Chr(10) + Sheets("TableDef").Range("Q3").Text
        End If
        If sQuarryType = "2" Then
            FieldPostil = FieldPostil + Chr(10) + Sheets("TableDef").Range("Q4").Text
        End If
        If sQuarryType = "3" Then
            FieldPostil = FieldPostil + Chr(10) + Sheets("TableDef").Range("Q5").Text
        End If
    End If
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    
    Dim BitmapInstruction As String
    If SheetDefine(iRow, iColumnType2) = "BITMAP" Then
        If bIsEng Then
            BitmapInstruction = "Note: This field uses 1 and 0 to indicate ON and OFF for each switch and does not contain delimiters. Example of the format: 111."
        ElseIf Not bIsEng Then
            BitmapInstruction = "注意：此字段使用1，0分别表示每个开关位的开或关，无分隔符，格式形如：111。"
        End If
        FieldPostil = FieldPostil + Chr(10) + BitmapInstruction
    End If
    CurSheet.Range(FieldCol + FieldRow).ClearComments
    CurSheet.Range(FieldCol + FieldRow).AddComment FieldPostil
    CurSheet.Range(FieldCol + FieldRow).Comment.Shape.Height = 160
    CurSheet.Range(FieldCol + FieldRow).Comment.Shape.Width = 120
End Sub

'******************************************************************************
'设置中英文名称转换
'******************************************************************************
Private Sub SetDisplay(CurSheet As Worksheet, iRow As Integer)
    Dim DisplayName As String, FieldCol As String, FieldRow As String
    
    FieldRow = Trim(SheetDefine(iRow, iTitleBeginRow))
    If bIsEng Then
        DisplayName = Trim(SheetDefine(iRow, iFieldDisplayName_ENG))
    Else
        DisplayName = Trim(SheetDefine(iRow, iFieldDisplayName_CHS))
    End If
    
    If CInt(Trim(SheetDefine(iRow, iCheckNull))) = 0 Then
        DisplayName = "*" + DisplayName
    End If
    
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    CurSheet.Range(FieldCol + FieldRow) = DisplayName
    '字体名称和大小
    With CurSheet.Range(FieldCol + FieldRow).Font
      .Name = "Arial"
      .Size = 8
      .Bold = True
    End With
    'Rows(Trim(CInt(FieldRow) + 1) + ":" + Trim(CInt(FieldRow) + 1)).Select
    'Selection.EntireRow.Hidden = True
End Sub

Private Function GetRangeInfo(iRow As Integer) As String
    Dim sFieldName As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sTempInfo As String
    sFieldName = Trim(SheetDefine(iRow, iColumnFieldName))
    sDataType = Trim(SheetDefine(iRow, iColumnType))
    sMinVal = Trim(SheetDefine(iRow, iMin))
    sMaxVal = Trim(SheetDefine(iRow, iMax))
    sRangeList = Trim(SheetDefine(iRow, iListValue))
    
    If sMinVal = "" And sRangeList = "" Then
        GetRangeInfo = ""
        Exit Function
    End If
    sTempInfo = GetValidErrMsg(sDataType)
    If Not bIsEng Then
        If sTempInfo = "Range" Then
            sTempInfo = Sheets("TableDef").Range("Q6").Text
        End If
        If sTempInfo = "Length" Then
            sTempInfo = Sheets("TableDef").Range("Q7").Text
        End If
    End If
    
    GetRangeInfo = ""
    If (sDataType = FINT) Or (sDataType = FSTRING) Then
        'sTempInfo = GetValidErrMsg(sDataType)
        If sMinVal = sMaxVal Then
            GetRangeInfo = sTempInfo + "[" + sMinVal + "]"
        Else
            GetRangeInfo = sTempInfo + "[" + sMinVal + ".." + sMaxVal + "]"
        End If
    End If
    If sDataType = FLIST Then
        GetRangeInfo = sTempInfo + "[" + sRangeList + "]"
    End If
    If sFieldName = "Physical Type" Then
        GetRangeInfo = ""
    End If
End Function



''******************************************************************************
''对于如AC45对象的行列计算
''******************************************************************************
'Public Function CalculateRange(sRange As String, isize As Integer, itype As Integer)
'
'
'End Function

'******************************************************************************
'为传入参数 CurSheet设置名称和注释
'******************************************************************************
Private Sub SetNegotiatedData(CurSheet As Worksheet)
    Dim iRow As Integer, iDefSheet As Integer
    Dim sht As Worksheet
    Set sht = Sheets("TableDef")
    For iDefSheet = 0 To CInt(sht.Cells(5, 7)) - 1
        If Trim(CurSheet.Name) = "IUCS" And Trim(sht.Cells(iDefSheet + StartTblDataRow, SheetNameCol)) = "IUCS" Then
            Call SetDisplay(CurSheet, iDefSheet)
            Call SetPostil(CurSheet, iDefSheet)
        End If
        If Trim(CurSheet.Name) = "IUPS" And Trim(sht.Cells(iDefSheet + StartTblDataRow, SheetNameCol)) = "IUPS" Then
            Call SetDisplay(CurSheet, iDefSheet)
            Call SetPostil(CurSheet, iDefSheet)
        End If
        If Trim(CurSheet.Name) = "IUR" And Trim(sht.Cells(iDefSheet + StartTblDataRow, SheetNameCol)) = "IUR" Then
            Call SetDisplay(CurSheet, iDefSheet)
            Call SetPostil(CurSheet, iDefSheet)
        End If
        If Trim(CurSheet.Name) = "COMMON" And Trim(sht.Cells(iDefSheet + StartTblDataRow, SheetNameCol)) = "COMMON" Then
            Call SetDisplay(CurSheet, iDefSheet)
            Call SetPostil(CurSheet, iDefSheet)
        End If
    Next
End Sub
Private Sub SetCoverSht()
    Dim sCoverInfo As String
    Dim CurSheet As Worksheet
    Set CurSheet = ThisWorkbook.Sheets("Cover")
    
    CurSheet.Unprotect "XCT100"

    If bIsEng Then
        CurSheet.Range("D5:K6") = Sheets("TableDef").Range("Q8").Text + " IuIur Data Template"
    Else
        CurSheet.Range("D5:K6") = Sheets("TableDef").Range("Q8").Text + " IuIur 数据模板"
    End If

    If bIsEng Then
        CurSheet.Range("E11") = "Read me"
    Else
        CurSheet.Range("E11") = Sheets("TableDef").Range("R3").Text
    End If
    
    If bIsEng Then
        sCoverInfo = Sheets("TableDef").Range("R5").Text
    Else
        sCoverInfo = Sheets("TableDef").Range("R4").Text
    End If
    CurSheet.Range("E12:H28") = sCoverInfo
    sCoverInfo = ""

    'CurSheet.Protect "XCT100", DrawingObjects:=True, Contents:=True, Scenarios:=True

End Sub
