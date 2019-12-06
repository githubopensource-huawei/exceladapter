Attribute VB_Name = "T_DefineNegotiatedFile"

Const FINT = "INT", FSTRING = "STRING", FLIST = "LIST"



'****************************************************************
'用于生成协商数据表
'****************************************************************
Public Sub GenNegotiatedFile()
    GeneratingFlag = 1
    iLanguageType = 0
    ThisWorkbook.Sheets("TableDef").Visible = True
    
    If iHideSheetFlg = 0 Then
        Call HideExtendFucSheet
    End If

    'added by z00102652 at 2010-04-20, begin
    Dim sht As Worksheet
    Dim s As String
    If Not SheetExist(ThisWorkbook, SHT_COVER) Then
        Set sht = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(SHT_TABLE_DEF))
        sht.Name = SHT_COVER
        ThisWorkbook.VBProject.VBComponents(sht.CodeName).Name = "T_" + sht.Name
        s = CODE_REMARK + vbLf + vbLf
        s = s + "Private Sub btn_SwitchChs_Click()" + vbLf
        s = s + "   Call SetChineseUI" + vbLf
        s = s + "End Sub" + vbLf + vbLf
        s = s + "Private Sub btn_SwitchEng_Click()" + vbLf
        s = s + "   Call SetEnglishUI" + vbLf
        s = s + "End Sub" + vbLf
        ThisWorkbook.VBProject.VBComponents(sht.CodeName).CodeModule.AddFromString (s)
    End If

    Call RebuildSheet(ThisWorkbook)
    'added by z00102652 at 2010-04-20, end

    '获取每个Sheet中协商数据定义
    Call GetSheetDefineData
    '获取所有配置对象Sheet名称
    Call GetAllSheetName
    '设置Cover Sheet页的版本信息
    Call SetCoverSheet
    '重配置对象Sheet页
    Call GenNegotiatedData
    
    Call SetActiveCell 'added by z00102652 at 2010-04-08

    Call SetSheetInvisible 'added by z00102652 at 2010-04-08

    Call SetGUI_DF 'added by z00102652 at 2010-04-06

    DF_DoubleFrequencyCell.Activate
    ActiveWindow.FreezePanes = False
    DF_DoubleFrequencyCell.Range("A4").Select
    ActiveWindow.FreezePanes = True
    
    Sheets("Cover").Select
    Call HideExtendFucSheet
    GeneratingFlag = 0
    
    MsgBox "Finished to generate RNP tempalte." 'added by z00102652 at 2010-04-08
End Sub

'****************************************************************
'设置Cover Sheet页的版本信息
'****************************************************************
Public Sub SetCoverSheet()
    Dim CurSheet As Worksheet
    Set CurSheet = ThisWorkbook.Sheets("Cover")

    CurSheet.Unprotect "HWCME"
    
    Call SetSheetColor(CurSheet, "Cover")
    
    Call SetSheetDefaultValue(CurSheet)
    Range("A:A").Select
    Selection.ColumnWidth = 8
    
    '设置标题
    Range("D8").Select
    Selection.RowHeight = 40
    Selection.ColumnWidth = 60
    Range("C8:D8").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .size = 22
        .ColorIndex = xlAutomatic
    End With
    Range("C9:D9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    '设置版本信息
    Range("C14:D15").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.RowHeight = 30
    
    Range("C14:C15").Select
    Selection.ColumnWidth = 18
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .size = 14
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 33
        .Pattern = xlSolid
    End With
    
    Range("D14:D15").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    With Selection.Font
        .Name = "宋体"
        .FontStyle = "Normal"
        .size = 14
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 34
        .Pattern = xlSolid
    End With
    
    '设置备注
    Range("C19:D20").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.RowHeight = 30
    
    Range("C19:D19").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .size = 14
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 34
        .Pattern = xlSolid
    End With
    
    Range("C20:D20").Select
    Selection.RowHeight = 270
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With
    With Selection.Font
        .Name = "宋体"
        .FontStyle = "Normal"
        .size = 9
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 34
        .Pattern = xlSolid
    End With
    
    '设置显示内容
    If iLanguageType = 0 Then
        Range("C8:D8") = gEngTemplateName
        'Range("C13") = gEngCMEVersionTitle
        Range("C14") = gEngNEVersion
        Range("C15") = gEngRNPVersion
        Range("C19:D19") = sEngCoverInfoTitle
        Range("C20:D20") = sEngCoverInfo1 + sEngCoverInfo2
    Else
        Range("C8:D8") = gChsTemplateName
        'Range("C13") = gChsCMEVersionTitle
        Range("C14") = gChsNEVersion
        Range("C15") = gChsRNPVersion
        Range("C19:D19") = sChsCoverInfoTitle
        Range("C20:D20") = sChsCoverInfo1 + sChsCoverInfo2
    End If
    'Range("D13") = sCMEVersion
    Range("D14") = sNEVersion
    Range("D15") = sRNPVersion
    
    For i = 0 To 3
        Rows("1:1").Delete
    Next
    Range("A3").Select
    
    CurSheet.Protect "HWCME", DrawingObjects:=True, Contents:=True, Scenarios:=True

End Sub

'****************************************************************
'为Sheet设置字段名、显示名称、批注信息
'****************************************************************
Public Sub GenNegotiatedData()
    Const FieldNameDisplayCol = 11
    
    Dim CurSheet As Worksheet
    Dim iSheet As Integer, iDefSheet As Integer
    Dim SheetID As String, FoundID As String
    Dim BField As Boolean
    Dim FieldDisplayRow As String, sTableRange As String, SheetName As String
    
    For iSheet = 0 To UBound(ArrSheetName) - 1
        BField = False
        SheetID = Trim(ArrSheetName(iSheet, 0))
        SheetName = Trim(ArrSheetName(iSheet, 1))
        sTableRange = Trim(ArrSheetName(iSheet, 6))
        Set CurSheet = ThisWorkbook.Sheets(SheetName)
        
        '设置Sheet页标签颜色
        Call SetSheetColor(CurSheet, SheetName)
        '设置字体等基本属性
        Call SetSheetDefaultValue(CurSheet)
        '设置行高
        Call SetDefRowHeight(CurSheet, iSheet)
        '设置边框，标题行格式
        Call SetFieldBorder(CurSheet, iSheet)
        
        '------------------------------------------------------------
        '设置每一列的属性
        For iDefSheet = 0 To UBound(SheetDefine) - 1
            FoundID = Trim(SheetDefine(iDefSheet, 0))
            If SheetID = FoundID Then
                BField = True
                FieldDisplayRow = Trim(SheetDefine(iDefSheet, FieldNameDisplayCol))
                Exit For
            End If
        Next

        If BField Then
            Do
                '设置列宽
                Call SetFieldColWidth(CurSheet, iDefSheet)
                '设置字段名
                Call SetFieldName(CurSheet, iDefSheet)
                '设置显示名称
                Call SetDisplayName(CurSheet, iDefSheet, FieldDisplayRow)
                '设置批注信息
                Call SetFieldPostil(CurSheet, iDefSheet, FieldDisplayRow)
                '设置数据有效性
                Call SetFieldValidate(CurSheet, iDefSheet, sTableRange)
                
                iDefSheet = iDefSheet + 1
                If iDefSheet >= TblRows Then Exit Do
                FoundID = Trim(SheetDefine(iDefSheet, 0))
            Loop While FoundID = ""
        End If
        
        '------------------------------------------------------------
        '删除多余的行
        CurSheet.Rows("2:3").Select
        Selection.Delete Shift:=xlUp
        CurSheet.Rows("3:3").Select
        Selection.Delete Shift:=xlUp
        
        '将第二行格式刷到第一行
        Rows("2:2").Select
        Selection.Copy
        Rows("1:1").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        
        '将第二行颜色去掉
        Rows("2:2").Select
        Selection.Interior.ColorIndex = xlNone
        Selection.Font.Bold = False
        
        '设置一，二行不进行数据有效性校验
        For iDefSheet = 0 To UBound(SheetDefine) - 1
            FieldDisplayRow = Trim(SheetDefine(iDefSheet, FieldNameDisplayCol))
            Range(FieldDisplayRow + "1" + ":" + FieldDisplayRow + "2").Select
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
        Next
        
        Range("A3").Select
        'CurSheet.Activate
        ActiveWindow.FreezePanes = True
    Next
End Sub

'****************************************************************
'设置Sheet页的标签颜色
'****************************************************************
Private Sub SetSheetColor(CurSheet As Worksheet, SheetName As String)
    If SheetName = "BSCInfo" Or SheetName = "DoubleFrequencyCell" Or SheetName = "WholeNetworkCell" Or SheetName = "DeleteInterNCellRelation" Then
        CurSheet.Tab.ColorIndex = 4
    Else
        CurSheet.Tab.ColorIndex = 33
    End If
End Sub

'****************************************************************
'Sheet缺省行为统一设置
'****************************************************************
Public Function SetSheetDefaultValue(CurSheet As Worksheet) As Integer
    CurSheet.Activate
    With CurSheet
        '缺省行高
        Cells.Select
        Selection.Clear
        Selection.RowHeight = 12
        '字体名称和大小
        With Selection.Font
            .Name = "Arial"
            .size = 9
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        '设置单元格锁定
        Selection.Locked = True
        Selection.FormulaHidden = False
    End With
    '设置零值显示、网格不显示
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayZeros = True
End Function

'****************************************************************
'设置行高
'****************************************************************
Private Sub SetDefRowHeight(CurSheet As Worksheet, iRow As Integer)
    Const iFieldCol = 3, iStartRow = 2
    Dim ArrData() As String, DefRowHeight As String
    Dim i As Integer, iSetRow As String
    
    DefRowHeight = Trim(ArrSheetName(iRow, iFieldCol))
    ArrData = Split(DefRowHeight, ",")
    CurSheet.Activate
    For i = 0 To UBound(ArrData)
        If Trim(ArrData(i)) <> "" Then
            iSetRow = Trim(CStr(iStartRow + i))
            With CurSheet
                Rows(iSetRow + ":" + iSetRow).Select
                Selection.RowHeight = CSng(Trim(ArrData(i)))
            End With
        End If
    Next
End Sub

'****************************************************************
'设置边框、拷贝标题行格式
'****************************************************************
Private Sub SetFieldBorder(CurSheet As Worksheet, iRow As Integer)
    Const iFieldRangeCol = 2, iTitleEndRowCol = 5
    Dim sFieldRangeCol As String, sTitleEndRow As String, sTitleRange As String
    
    sFieldRangeCol = Trim(ArrSheetName(iRow, iFieldRangeCol))
    sTitleEndRow = Trim(ArrSheetName(iRow, iTitleEndRowCol))
    
    CurSheet.Select
    Columns(sFieldRangeCol).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    Rows("1:" + sTitleEndRow).Select
    With Selection
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    '拷贝标题行格式
    sTitleRange = Left(sFieldRangeCol, 1) + sTitleEndRow + Mid(sFieldRangeCol, 2) + sTitleEndRow
    Sheets("TableDef").Range("C17").Copy
    CurSheet.Range(sTitleRange).PasteSpecial xlPasteFormats, xlPasteSpecialOperationNone, False, False
End Sub

'****************************************************************
'设置列宽
'****************************************************************
Private Sub SetFieldColWidth(CurSheet As Worksheet, iRow As Integer)
    Const FieldWidthCol = 10
    Const FieldDefCol = 4
    Dim sFieldWidth As String, FieldCol As String
    
    sFieldWidth = Trim(SheetDefine(iRow, FieldWidthCol))
    FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
    If sFieldWidth = "" Then Exit Sub
    
    CurSheet.Select
    Columns(FieldCol + ":" + FieldCol).Select
    Selection.ColumnWidth = CSng(sFieldWidth)
End Sub

'****************************************************************
'设置字段名称（Sheet页第一行内容）
'****************************************************************
Private Sub SetFieldName(CurSheet As Worksheet, iRow As Integer)
    Const FieldNameCol = 2
    Const FieldDefCol = 4
    Dim FieldName As String, FieldCol As String, FieldRow As String
    
    FieldName = Trim(SheetDefine(iRow, FieldNameCol))
    FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
    FieldRow = 1
    CurSheet.Range(FieldCol + FieldRow) = FieldName
    
    CurSheet.Range(FieldCol + FieldRow).Font.Bold = True
End Sub

'****************************************************************
'设置字段显示名称（Sheet第二行内容）
'****************************************************************
Private Sub SetDisplayName(CurSheet As Worksheet, iRow As Integer, FieldRow As String)
    Const FieldDefCol = 4
    Const FieldNameDisplayCol = 11
    Dim DisplayName As String, FieldCol As String
    
    DisplayName = Trim(SheetDefine(iRow, FieldNameDisplayCol + 1))
    FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
    CurSheet.Range(FieldCol + FieldRow) = DisplayName

    CurSheet.Range(FieldCol + FieldRow).Font.Bold = True

    'CurSheet.Range(FieldCol + Trim(CInt(FieldRow) + 1)) = DisplayName
    Rows(Trim(CInt(FieldRow) + 1) + ":" + Trim(CInt(FieldRow) + 1)).Select
    Selection.EntireRow.Hidden = True
End Sub

'****************************************************************
'设置批注信息
'****************************************************************
Private Sub SetFieldPostil(CurSheet As Worksheet, iRow As Integer, FieldRow As String)
    Const FieldDefCol = 4
    Const FieldNameDisplayCol = 11
    Const IsMustGiveCol = 16
    Dim FieldPostil As String, FieldCol As String
    Dim RangeName As String, DisplayName As String, IsMustGiveFlag As String
    
    DisplayName = Trim(SheetDefine(iRow, FieldNameDisplayCol + 3))
    RangeName = GetRangeInfo(iRow)
    IsMustGiveFlag = Trim(SheetDefine(iRow, IsMustGiveCol))
    
    FieldPostil = DisplayName + "(" + RangeName + ")"
    If UCase(IsMustGiveFlag) = "YES" Then
        FieldPostil = FieldPostil + gEngIsMustGive
    End If
    
    FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
    CurSheet.Range(FieldCol + FieldRow).ClearComments
    CurSheet.Range(FieldCol + FieldRow).AddComment FieldPostil
    CurSheet.Range(FieldCol + FieldRow).Comment.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
    CurSheet.Range(FieldCol + FieldRow).Comment.Shape.ScaleHeight 1, msoFalse, msoScaleFromTopLeft
End Sub

'****************************************************************
'标注必填字段
'****************************************************************
Private Sub SetMustGiveField(CurSheet As Worksheet, iRow As Integer)
    Const MustGiveFlagCol = 16
    Const FieldDefCol = 4
    Dim sMustGiveFlag As String, FieldCol As String
    
    sMustGiveFlag = Trim(SheetDefine(iRow, MustGiveFlagCol))
    FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
    
    If sMustGiveFlag = "YES" Then
        CurSheet.Range(FieldCol + "2").Select
        With Selection.Interior
            .ColorIndex = 36
            .Pattern = xlSolid
        End With
    End If
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
'设置单元格格式和数据有效性
'****************************************************************
Private Sub SetFieldValidate(CurSheet As Worksheet, iRow As Integer, sTableRange As String)
    Const FieldNameCol = 2
    Const DataTypeCol = 3
    Const FieldDefCol = 4
    Const MinValCol = 5
    Const MaxValCol = 6
    Const RangeListCol = 7
    Dim sFieldName As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, FieldCol As String
    Dim xType As Excel.XlDVType, sFormula1 As String, sFormula2 As String, sErrPrompt As String, sErrMsg As String
    Dim nMinVal As Double, nMaxVal As Double
    
    sFieldName = Trim(SheetDefine(iRow, FieldNameCol))
    sDataType = Trim(SheetDefine(iRow, DataTypeCol))
    FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
    sMinVal = Trim(SheetDefine(iRow, MinValCol))
    sMaxVal = Trim(SheetDefine(iRow, MaxValCol))
    sRangeList = Trim(SheetDefine(iRow, RangeListCol))
    
    
    'If (sDataType = FINT) Then
        'xType = xlValidateWholeNumber
        'sFormula1 = sMinVal
        'sFormula2 = sMaxVal
        'xType = xlValidateTextLength
        'If Len(sMinVal) <= Len(sMaxVal) Then
           ' sFormula1 = Len(sMinVal)
           ' sFormula2 = Len(sMaxVal)
       ' Else
           ' sFormula1 = Len(sMaxVal)
           ' sFormula2 = Len(sMinVal)
       ' End If
        'nMinVal = CDbl(sMinVal)
        'nMaxVal = CDbl(sMaxVal)
        'If nMinVal < 0 And nMaxVal > 9 Then
           ' sFormula1 = 1
        'End If
        'CurSheet.Columns(FieldCol).NumberFormatLocal = "@"
        'CurSheet.Columns(FieldCol).HorizontalAlignment = xlRight
        'CurSheet.Cells(4, FieldCol).HorizontalAlignment = xlCenter
    'End If
    
    'If (sDataType = FSTRING) Then
       ' xType = xlValidateTextLength
       ' sFormula1 = sMinVal
       ' 'sFormula2 = sMaxVal
       ' CurSheet.Columns(FieldCol).NumberFormatLocal = "@"
       ' CurSheet.Columns(FieldCol).HorizontalAlignment = xlRight
       ' CurSheet.Cells(4, FieldCol).HorizontalAlignment = xlCenter
    'End If
    
    If sDataType = FLIST Then
        xType = xlValidateList
        sFormula1 = sRangeList
        sFormula2 = ""
        CurSheet.Columns(FieldCol).NumberFormatLocal = "@"
        CurSheet.Columns(FieldCol).HorizontalAlignment = xlRight
        CurSheet.Cells(4, FieldCol).HorizontalAlignment = xlCenter
    End If
    
'    If sFieldName = "LAC" Then
'        xType = xlValidateCustom
'        sFormula1 = "=OR(AND(" + FieldCol + "1<=65533," + FieldCol + "1>0), " + FieldCol + "1 = 65535)"
'        sFormula2 = ""
'    End If
    
    sErrPrompt = GetValidErrTitle(sDataType)
    sErrMsg = GetRangeInfo(iRow)
    
    CurSheet.Select
    Columns(FieldCol + ":" + FieldCol).Select
    sErrMsg = sErrMsg + vbLf + "Worksheet=" + CurSheet.Name + "; Column=" + FieldCol + "; Row=" + str(iRow)
    Call SetDataValidate(xType, sFormula1, sFormula2, sErrPrompt, sErrMsg)
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
            sErrMsg = sErrMsg + GetCellInfo(Target)
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
                sErrMsg = sErrMsg + GetCellInfo(Target)
                MsgBox sErrMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, sErrPrompt
                FieldRange.ClearContents
            End If
        Next
    End If
End Sub

'****************************************************************
'检查单元格输入的值是否符合数据有效性规则
'****************************************************************
Public Function CheckOneFieldData(ByVal nDefSheetIndex As Long, ByVal FieldRange As Range) As String
    Dim nValue As Double, nStrLen As Double, nLoop As Integer
    Dim sValue As String, sItem As String
    Dim bFlag As Boolean
    Dim sColType As String, sFeildName As String, sMinVal As String, sMaxVal As String, sErrMsg As String
    Dim nMinVal As Double, nMaxVal As Double
    
    sColType = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 3))
    If sColType = "INT" Then
        sFeildName = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 2))
        sMinVal = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 5))
        sMaxVal = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 6))
        nMinVal = CDbl(sMinVal)
        nMaxVal = CDbl(sMaxVal)
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
                If nValue < nMinVal Or nValue > nMaxVal Then
                    bFlag = False
                End If
                If sFeildName = "LAC" And nValue = (nMaxVal - 1) Then
                    bFlag = False
                End If
            End If
            
            CheckOneFieldData = ""
            If Not bFlag Then
                If sFeildName = "LAC" Then
                    sErrMsg = "Range [" & sMinVal & ".." & nMaxVal - 2 & ", " & sMaxVal & "]"
                Else
                    sErrMsg = "Range [" & sMinVal & ".." & sMaxVal & "]"
                End If
                CheckOneFieldData = sErrMsg
            End If
        End If
        
    ElseIf sColType = "STRING" Then
        Dim size As Integer
        
        sSheetName = ActiveWindow.ActiveSheet.Name
        sFeildName = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 2))
        sMinVal = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 5))
        sMaxVal = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 6))
        nMinVal = CDbl(sMinVal)
        nMaxVal = CDbl(sMaxVal)
        sValue = FieldRange.Text
        
        If sSheetName = "PhyNBRadio" And sFeildName = "STNAME" Then
            If sValue <> "" Then
                size = getByteLen(sValue)
                If size < nMinVal Or size > nMaxVal Then
                    sErrMsg = "length [" & sMinVal & ".." & sMaxVal & "]"
                    CheckOneFieldData = sErrMsg
                End If
            End If
        Else
            If sValue <> "" Then
                size = Len(sValue)
                If size < nMinVal Or size > nMaxVal Then
                    sErrMsg = "length [" & sMinVal & ".." & sMaxVal & "]"
                    CheckOneFieldData = sErrMsg
                End If
            End If
        End If
        
    End If
End Function

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
            sErrMsg = sErrMsg + GetCellInfo(Target)
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
                sErrMsg = sErrMsg + GetCellInfo(Target)
                MsgBox sErrMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, sErrPrompt
                FieldRange.ClearContents
            End If
        Next
    End If
End Sub

'****************************************************************
'检查Cell Sheet页的单元格输入的值是否符合数据有效性规则
'****************************************************************
Public Function CellCheckOneFieldData(ByVal nDefSheetIndex As Long, ByVal FieldRange As Range) As String
    Dim nValue As Long, nStrLen As Long, nLoop As Integer
    Dim sValue As String, sItem As String
    Dim bFlag As Boolean
    Dim sColType As String, sFeildName As String, sMinVal As String, sMaxVal As String
    Dim nMinVal As Long, nMaxVal As Long
    Dim sSheetName As String, sBandValue As String
    Dim BranchRange As String, BranchValue As String, BFieldRange As String, sBSheetName As String
    Dim nBandColumn As Integer, sErrMsg As String
    
    Const ValidSheetNameCol = 0
    Const ValidDefBranchFieldCol = 2
    Const ValidDefValueCol = 5
    Const ValidDefFieldCol = 7
    Const ValidDefMinCol = 9
    Const ValidDefMaxCol = 10
    
    sColType = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 3))
    If sColType = "INT" Then
        sFeildName = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 2))
        sSheetName = ActiveWindow.ActiveSheet.Name
        bRangeFlag = False
        sBandValue = ""
        
        If sSheetName = "CELL" And (FieldRange.Column = 7 Or FieldRange.Column = 8) Then
            nBandColumn = 6
            sBandValue = UCase(Cells(FieldRange.Row, nBandColumn).Value)
        End If
        If sSheetName = "NRNCCELL" And (FieldRange.Column = 10 Or FieldRange.Column = 11) Then
            nBandColumn = 8
            sBandValue = UCase(Cells(FieldRange.Row, nBandColumn).Value)
        End If
        
        If sBandValue <> "" Then
            For nLoop = 0 To UBound(RangeDefine) - 1
                sBSheetName = Trim(RangeDefine(nLoop, ValidSheetNameCol))
                sBranchValue = Trim(RangeDefine(nLoop, ValidDefValueCol))
                BranchRange = RangeDefine(nLoop, ValidDefBranchFieldCol) & ":" & RangeDefine(nLoop, ValidDefBranchFieldCol)
                BFieldRange = RangeDefine(nLoop, ValidDefFieldCol) & ":" & RangeDefine(nLoop, ValidDefFieldCol)
                
                If sSheetName = sBSheetName And Range(BranchRange).Column = nBandColumn And Range(BFieldRange).Column = FieldRange.Column And sBranchValue = sBandValue Then
                    sMinVal = Trim(RangeDefine(nLoop, ValidDefMinCol))
                    sMaxVal = Trim(RangeDefine(nLoop, ValidDefMaxCol))
                    Exit For
                End If
            Next
        Else
            sMinVal = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 5))
            sMaxVal = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 6))
        End If

        nMinVal = CLng(sMinVal)
        nMaxVal = CLng(sMaxVal)
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
                nValue = CLng(sValue)
                If nValue < nMinVal Or nValue > nMaxVal Then
                    bFlag = False
                End If
                If sFeildName = "LAC" And nValue = (nMaxVal - 1) Then
                    bFlag = False
                End If
            End If
            
            CellCheckOneFieldData = ""
            If Not bFlag Then
                If sFeildName = "LAC" Then
                    sErrMsg = "Range [" & sMinVal & ".." & nMaxVal - 2 & ", " & sMaxVal & "]"
                Else
                    sErrMsg = "Range [" & sMinVal & ".." & sMaxVal & "]"
                End If
                CellCheckOneFieldData = sErrMsg
            End If
        End If
        
    ElseIf sColType = "STRING" Then
        sColType = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 15))
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
        
                CellCheckOneFieldData = ""
                If Not bFlag Then
                    sErrMsg = "Input Range [0,1]"
                    CellCheckOneFieldData = sErrMsg
                End If
            End If
        Else
            sValue = FieldRange.Text
            If sValue <> "" Then
                Dim size As Integer
                sMinVal = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 5))
                sMaxVal = Trim(SheetDefine(nDefSheetIndex + FieldRange.Column - 1, 6))
                nMinVal = CDbl(sMinVal)
                nMaxVal = CDbl(sMaxVal)
                size = Len(sValue)
                If size < nMinVal Or size > nMaxVal Then
                    sErrMsg = "length [" & sMinVal & ".." & sMaxVal & "]"
                    CellCheckOneFieldData = sErrMsg
                End If
            End If
        End If
        
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

Private Sub SetActiveCell()
    Sheets(SHT_COVER).Activate
    Sheets(SHT_COVER).Cells(1, 1).Activate
    Sheets(SHT_CONVERT_TEMPLATE).Activate
    Sheets(SHT_CONVERT_TEMPLATE).Cells(1, 1).Activate
    Sheets(SHT_DOUBLE_FREQ_CELL_SETTING).Activate
    Sheets(SHT_DOUBLE_FREQ_CELL_SETTING).Cells(1, 1).Activate
    Sheets(MOC_DOUBLE_FREQ_CELL).Activate
    Sheets(MOC_DOUBLE_FREQ_CELL).Cells(4, 1).Activate
    Sheets(SHT_TABLE_DEF).Activate
    Sheets(SHT_TABLE_DEF).Cells(1, 1).Activate
    Sheets(SHT_VALID_DEF).Activate
    Sheets(SHT_VALID_DEF).Cells(1, 1).Activate
End Sub

Private Sub SetSheetInvisible()
    Sheets(SHT_TABLE_DEF).Visible = False
    Sheets(SHT_VALID_DEF).Visible = False
    Sheets(SHT_DEBUG).Visible = False
    Sheets(SHT_DOUBLE_FREQ_CELL_SETTING).Visible = False
    Sheets(MOC_DEL_INTERFREQNCELL).Visible = False
    Sheets("TempSheet1").Visible = False
    Sheets("TempSheet2").Visible = False
    Sheets("TempSheet3").Visible = False
    Sheets("TempSheet4").Visible = False
    Sheets("TempSheet5").Visible = False
    Sheets("TempSheet6").Visible = False
    Sheets("TempSheet7").Visible = False
End Sub

Public Function getByteLen(ByVal SStr As String) As Integer
    Dim strlen As Integer, byteLen As Integer, index As Integer
    byteLen = 0
    strlen = Len(SStr)
    For index = 1 To strlen
       If Asc(Mid(SStr, index, 1)) >= 0 And Asc(Mid(SStr, index, 1)) <= 127 Then
          byteLen = byteLen + 1
       Else
          byteLen = byteLen + 2
       End If
    Next index
    getByteLen = byteLen
End Function
