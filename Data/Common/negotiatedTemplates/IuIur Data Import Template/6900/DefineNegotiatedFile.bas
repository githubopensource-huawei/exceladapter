Attribute VB_Name = "DefineNegotiatedFile"
Option Explicit

'定义界面处理类型
Public Const FINT = "INT", FSTRING = "STRING", FLIST = "LIST"
Public Const FBROW = 1, FEROW = 2, FBCOL = 3, FECOL = 4
Public Const DefinedWhite = 2
Public Const DefinedGreen = 35

'TableDef中表定义
Public Const iMocName = 1
Public Const iColumnFieldName = 3
Public Const iColumnType = 4
Public Const iMin = 5
Public Const iMax = 6
Public Const iListValue = 7
Public Const iFieldBeginColumn = 8
Public Const iFieldEndColumn = 9
Public Const iColumnWidth = 10
Public Const iTitleBeginRow = 11
Public Const iContentEndRow = 12
Public Const iUpdateFlagColIndex = 13
Public Const iFieldDisplayName_ENG = 14
Public Const iFieldDisplayName_CHS = 15
Public Const iQuarry = 23
Public Const iFieldPostil = 16
Public Const iMapTableName = 17
Public Const iMapFieldName = 18
Public Const iRealTableName = 19
Public Const iRealFieldName = 20
Public Const iCheckNull = 21
Public Const iColumnType2 = 22


Public Const iSpecialMoName = 25
Public Const iManualValue = 26

Public Const StartTblDataRow = 15

Public Const SheetNameCol = 2

Public Sub AutoGenerate()
    SetTabVisibilityMacro
    Call GenInitTableMap
    Call GenInitFieldMap
    Call RefreshBranchDefPos
    Call EnsureRefreshBranchDefine
    Call DefineNegotiatedFile.GenNegotiatedFile(True)
    Call GenSQLScripts
    SetTabVisibilityMacro
End Sub

Public Sub ClearTableDef()
    Dim CurSheet As Worksheet
    Set CurSheet = Sheets("TableDef")
    CurSheet.Activate
    
    Dim iTableDefEndRow As Integer
    iTableDefEndRow = StartTblDataRow + CInt(Sheets("TableDef").Cells(5, 7)) - 1
    If iTableDefEndRow > 15 Then
        CurSheet.Rows("15:" + CStr(iTableDefEndRow)).ClearContents
    End If
End Sub

'增加新的sheet
Public Sub ReCreateSheet(bAll As Boolean)
  Dim sPreSheetName As String
  Dim sSheetName As String
  Dim sht As Worksheet
  Set sht = Sheets("TableDef")
  Dim CurSheet As Worksheet
  Dim iNewSheet As Integer
  Dim iRow As Integer
  Dim iStartRow As Integer
  Dim iEndRow As Integer
  
  '第一步，删除原有的sheet页
  'Application.DisplayAlerts = False
  'For iRow = StartTblDataRow To StartTblDataRow + CInt(Sht.Cells(5, 7)) - 1
  '    sSheetName = Sht.Cells(iRow, SheetNameCol)
  '    If IsSheetNameExists(sSheetName) Then
  '        ThisWorkbook.Sheets(sSheetName).Delete
  '    End If
  'Next
  'Application.DisplayAlerts = True
  
  '第二步，生成新的sheet页
  '增加第一页
  iNewSheet = 1
  sPreSheetName = "IUCS"
  
 ' If Not IsSheetNameExists(sPreSheetName) Then
   ' ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(iNewSheet)
   ' ThisWorkbook.Sheets(iNewSheet + 1).Name = sPreSheetName
        
   ' iNewSheet = iNewSheet + 1
 ' End If
  
  iStartRow = StartTblDataRow
  Set CurSheet = Sheets(sPreSheetName)
  For iRow = StartTblDataRow To StartTblDataRow + CInt(sht.Cells(5, 7))
      If sPreSheetName <> sht.Cells(iRow, SheetNameCol) Then
          '生成前一页
          iEndRow = iRow - 1
          If bAll Then
            Call ClearSheetData(CurSheet)
          End If
          Call GenNegotiatedData(CurSheet, iStartRow - StartTblDataRow, iEndRow - StartTblDataRow, bAll)
          
          If "" = sht.Cells(iRow, SheetNameCol) Then
            Exit Sub
          End If
          '增加后一页
          iStartRow = iRow
          sPreSheetName = sht.Cells(iRow, SheetNameCol)
          'If Not IsSheetNameExists(sPreSheetName) Then
           ' ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(iNewSheet)
           ' ThisWorkbook.Sheets(iNewSheet + 1).Name = sPreSheetName
        
            'iNewSheet = iNewSheet + 1
          'End If
          Set CurSheet = Sheets(sPreSheetName)
      End If
  Next
End Sub

                                                
Public Sub ClearSheetData(CurSheet As Worksheet)
                                                
  CurSheet.Activate
                                                
  Cells.Select
  Selection.Clear
                                                
  Dim elmShp As Shape
                                                
  For Each elmShp In CurSheet.Shapes
      If elmShp.Name Like "Group *" Then
         'Set elmShp = Nothing
         elmShp.Delete
      End If
  Next
End Sub


'主程序，用于生成协商数据表
Public Sub GenNegotiatedFile(bAll As Boolean)
    GeneratingFlag = 1
    Call UnprotectWorkBook
    ThisWorkbook.Worksheets("TableDef").Visible = True
    '获取协商数据定义
    Call GetSheetDefineData
    Call SetSheetProtected(False)
    '删除sheet页，并创建新的sheet
    Application.ScreenUpdating = False
    Call ReCreateSheet(bAll)
    Application.ScreenUpdating = True
    Call RefreshEnumBiggerThan255
    Call SetSheetProtected(True)
    
  '  ThisWorkbook.Worksheets("Cover").Select
    Call ProtectWorkBook
    GeneratingFlag = 0
End Sub

'为Sheet设置字段名、显示名称、批注信息
Public Sub GenNegotiatedData(CurSheet As Worksheet, iStartRow As Integer, iEndRow As Integer, bAll As Boolean)
    
    Dim iSheet As Integer, iDefSheet As Integer
    Dim FoundID As String
    Dim BField As Boolean
    
    Rows("2:2").Select
    If bAll Then
        Call SetSheetDefaultValue(CurSheet)
    End If
    
    '------------------------------------------------------------
    '设置每一列的属性
    For iDefSheet = iStartRow To iEndRow
        If bAll Or SheetDefine(iDefSheet, iUpdateFlagColIndex) = "1" Then
          '设置边框
          Call SetFieldBorder(CurSheet, iDefSheet)
          '设置列宽、字体
          Call SetFieldColWidth(CurSheet, iDefSheet)
          '设置字段
          Call SetFieldName(CurSheet, iDefSheet)
          '设置显示名称
          Call SetDisplayName(CurSheet, iDefSheet)
          '设置批注信息
          Call SetFieldPostil(CurSheet, iDefSheet)
          '设置数据有效性
          Call SetFieldValidate(CurSheet, iDefSheet)
          '设置单元格锁定
          Call ClearValidateForTitle(CurSheet, iDefSheet)
        End If
    Next
    
    If bAll Then
        Call SetMocName(CurSheet)
    End If
    
    '------------------------------------------------------------
    '设置完成，锁定Sheet页
    Range("CC1").Select
    ActiveCell.FormulaR1C1 = "1"
    Selection.AutoFill Destination:=Range("CC1:CC" + SheetDefine(iEndRow, iContentEndRow)), Type:=xlFillSeries
    Columns("CC:CC").Select
    Selection.EntireColumn.Hidden = True
    Selection.Locked = True
    Selection.FormulaHidden = True
    Rows("1:1").Select
    Selection.EntireRow.Hidden = True
    Selection.Locked = True
    Selection.FormulaHidden = True
    Range("C7").Select
    
    If bAll Then
        Sheets("TableDef").Select
        ActiveSheet.Shapes("Group 820").Select
        Selection.Copy
        CurSheet.Activate
        Range("A3").Select
        CurSheet.Paste
    End If
End Sub

'设置对象名
Private Sub SetMocName(CurSheet As Worksheet)
  Const TableDefMocNameCol = 3
  Const MocNameCol = 1
  Const SheetNameCol = 2
  Const iBeginRow = 12
  Const iEndRow = 13
   
  Dim iStartRow As Integer
  
  Dim sPreMocName As String

  Dim sht As Worksheet
  Set sht = Sheets("TableDef")

  Dim iRow As Integer
  iRow = GetInterfaceBeginRow(CurSheet.Name)
  iStartRow = GetInterfaceBeginRow(CurSheet.Name)
  sPreMocName = sht.Cells(StartTblDataRow, TableDefMocNameCol)
  
  Do Until sht.Cells(iRow, TableDefMocNameCol) = ""
    If (sPreMocName <> sht.Cells(iRow + 1, TableDefMocNameCol)) And (CurSheet.Name = sht.Cells(iRow + 1, SheetNameCol)) Then
      sPreMocName = sht.Cells(iRow, TableDefMocNameCol)
      
      CurSheet.Select
      CurSheet.Cells(CInt(sht.Cells(iStartRow, iBeginRow)), MocNameCol) = sPreMocName
      
      CurSheet.Range("A" + CStr(sht.Cells(iStartRow, iBeginRow)) + ":" + "A" + CStr(sht.Cells(iRow, iEndRow))).Select
      With Selection
        .Merge
        
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Microsoft Sans Serif"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = xlUnderlineStyleNone
        .Font.ColorIndex = 2
        .WrapText = True
        .Interior.ColorIndex = 46
        .Interior.Pattern = xlGray8
        .Interior.PatternColorIndex = xlAutomatic
      End With
      
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
      
      iStartRow = iRow + 1
    End If
    
    iRow = iRow + 1
  Loop
End Sub
        
'清除有效性、设置单元格锁定
Private Sub ClearValidateForTitle(CurSheet As Worksheet, iRow As Integer)
    Dim sFieldRangeCol As String, sTitleBeginRow As String, sTitleRange As String, sFieldRange As String
    Dim sBeginRow As String, sEndRow As String, sBeginCol As String, sEndCol As String
                    
    sBeginRow = Trim(CStr(CInt(SheetDefine(iRow, iTitleBeginRow)) + 1))
    sEndRow = Trim(SheetDefine(iRow, iContentEndRow))
    sBeginCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    sEndCol = Trim(SheetDefine(iRow, iFieldEndColumn))
    sFieldRangeCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    sTitleBeginRow = Trim(SheetDefine(iRow, iTitleBeginRow))
    
    sFieldRange = sBeginCol + sBeginRow + ":" + sEndCol + sEndRow
    sTitleRange = sBeginCol + sTitleBeginRow + ":" + sEndCol + sTitleBeginRow
    CurSheet.Select

    'Field解锁
    Range(sFieldRange).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    'Title加锁
    Range(sTitleRange).Select
    Selection.Locked = True
    Selection.FormulaHidden = False
End Sub

'设置边框,合并单元格
Private Sub SetFieldBorder(CurSheet As Worksheet, iRow As Integer)
    Dim sFieldRangeCol As String, sTitleEndRow As String, sTitleRange As String, sFieldRange As String
    Dim sBeginRow As String, sEndRow As String, sBeginCol As String, sEndCol As String
    Dim iCurRow As Integer
                    
    sBeginRow = Trim(CStr(CInt(SheetDefine(iRow, iTitleBeginRow)) + 1))
    sEndRow = Trim(SheetDefine(iRow, iContentEndRow))
    sBeginCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    sEndCol = Trim(SheetDefine(iRow, iFieldEndColumn))
    
    sFieldRangeCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    sTitleEndRow = Trim(SheetDefine(iRow, iTitleBeginRow))
    
    sFieldRange = sBeginCol + sBeginRow + ":" + sEndCol + sEndRow
    sTitleRange = sBeginCol + sTitleEndRow + ":" + sEndCol + sTitleEndRow
    
    CurSheet.Select
    For iCurRow = CInt(sBeginRow) To CInt(sEndRow)
        Range(sBeginCol + CStr(iCurRow) + ":" + sEndCol + CStr(iCurRow)).Select
        Selection.Merge
    Next
        
    Range(sFieldRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Interior.ColorIndex = DefinedWhite
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
    
    With Selection.Borders(xlInsideHorizontal)
        .Weight = xlThin
        '.ColorIndex = xlAutomatic
        '.LineStyle = xlContinuous
    End With
    
    Sheets("TableDef").Select
    
    Range("C14").Select
    Selection.Copy
    CurSheet.Select
    Range(sTitleRange).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Merge
End Sub

'设置列宽
Private Sub SetFieldColWidth(CurSheet As Worksheet, iRow As Integer)
    Dim iFieldWidth As Integer, FieldCol As String
    
    iFieldWidth = CInt(Trim(SheetDefine(iRow, iColumnWidth)))
    
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    If iFieldWidth = 0 Then Exit Sub
    
    CurSheet.Select
    If Columns(FieldCol + ":" + FieldCol).ColumnWidth < iFieldWidth Then
      Columns(FieldCol + ":" + FieldCol).ColumnWidth = iFieldWidth
    End If
      
    '字体名称和大小
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With

End Sub

'设置数据有效性
Private Sub SetFieldValidate(CurSheet As Worksheet, iRow As Integer)
    Dim sFieldName As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, FieldCol As String
    Dim xType As Excel.XlDVType, sFormula1 As String, sFormula2 As String, sErrPrompt As String, sErrMsg As String
    Dim sFieldRange As String, sBeginRow As String, sEndRow As String
            
    sBeginRow = Trim(CStr(CInt(SheetDefine(iRow, iTitleBeginRow)) + 1))
    sEndRow = Trim(SheetDefine(iRow, iContentEndRow))
           
    sFieldName = Trim(SheetDefine(iRow, iColumnFieldName))
    sDataType = Trim(SheetDefine(iRow, iColumnType))
    sMinVal = Trim(SheetDefine(iRow, iMin))
    sMaxVal = Trim(SheetDefine(iRow, iMax))
    sRangeList = CStr(Trim(SheetDefine(iRow, iListValue)))
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    sFieldRange = FieldCol + sBeginRow + ":" + FieldCol + sEndRow

    If (sMinVal = "" And sRangeList = "") Then
        Exit Sub
    End If
    
    If (sDataType = FINT) Then
        Dim vMinVal, vMaxVal
        vMinVal = Split(sMinVal, ",")
        vMaxVal = Split(sMaxVal, ",")
          
        If UBound(vMinVal) <> 0 Then
          If sMinVal <> sMaxVal Then
            xType = xlValidateCustom
            sFormula1 = "=OR("
            
            Dim Index As Integer
            For Index = 0 To UBound(vMinVal)
              sFormula1 = sFormula1 + "AND(" + FieldCol + sBeginRow + ">=" + vMinVal(Index) + "," + FieldCol + sBeginRow + "<=" + vMaxVal(Index) + "),"
            Next
            sFormula1 = Left(sFormula1, Len(sFormula1) - 1) + ")"
            sFormula2 = ""
          Else
            xType = xlValidateList
            sFormula1 = sMinVal
            sFormula2 = ""
          End If
        Else
          xType = xlValidateWholeNumber
          sFormula1 = sMinVal
          sFormula2 = sMaxVal
        End If
    End If
    If (sDataType = FSTRING) Then
        xType = xlValidateTextLength
        sFormula1 = sMinVal
        sFormula2 = sMaxVal
    End If
    If sDataType = FLIST Then
        xType = xlValidateList
        sFormula1 = sRangeList
        sFormula2 = ""
    End If
    
    sErrPrompt = GetValidErrTitle(sDataType)
    sErrMsg = GetRangeInfo(iRow)
    
    CurSheet.Select
    If sEndRow = "" Then
        Columns(FieldCol + ":" + FieldCol).Select
    Else
        Range(sFieldRange).Select
    End If
    Call SetDataValidate(xType, sFormula1, sFormula2, sErrPrompt, sErrMsg)
        
    '设置单元格数字格式
    If (sDataType = FSTRING) Or (sDataType = FLIST) Then
        Selection.NumberFormatLocal = "@"
    End If
End Sub

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

'刷新某个字段区域的Enum
Private Sub RefreshEnumBiggerThan255_i(SheetName As String, RangeString As String, strListValue As String, FieldName As String, CurrentColumn As Integer)
    '先写入EnumSheet
    Dim EnumSheet As Worksheet
    Set EnumSheet = Sheets("EnumDef")
    EnumSheet.Cells(1, CurrentColumn) = FieldName
    
    Dim EnumEndRow As Integer
    EnumEndRow = 2
    
    Dim ArrData() As String
    Dim i As Integer
    ArrData = Split(strListValue, ",")
    For i = 0 To UBound(ArrData)
        EnumSheet.Cells(EnumEndRow, CurrentColumn) = ArrData(i)
        EnumEndRow = EnumEndRow + 1
    Next
    EnumEndRow = EnumEndRow - 1
    
    '再刷新对应单元格的有效性
    Dim CurSheet As Worksheet
    Set CurSheet = Sheets(SheetName)
    
    Dim FormulaString As String
    FormulaString = "=indirect(""EnumDef!" + c(CurrentColumn) + CStr(2) + ":" + c(CurrentColumn) + CStr(EnumEndRow) + """)"
    CurSheet.Range(RangeString).Validation.Modify Formula1:=FormulaString
End Sub

'针对Enum的长度超过255的，在生成模板后进行这部分字段的Enum刷新
Public Sub RefreshEnumBiggerThan255()
    Call SetSheetProtected(False)
    
    '对TableDef进行遍历，找到Enum长度超过250的
    Dim CurSheet As Worksheet
    Set CurSheet = Sheets("TableDef")
    
    Dim CurrentColumn As Integer
    CurrentColumn = 1
    ThisWorkbook.Worksheets("EnumDef").Cells.ClearContents
    
    Dim RangeString As String
    Dim SheetName As String
    Dim FieldName As String
    Dim strListValue As String
    Dim iRow As Integer
    '对每一行进行遍历
    For iRow = StartTblDataRow To StartTblDataRow + CInt(CurSheet.Cells(5, 7)) - 1
        SheetName = CurSheet.Cells(iRow, 2)
        FieldName = CurSheet.Cells(iRow, 2) + "." + CurSheet.Cells(iRow, 3) + "." + CurSheet.Cells(iRow, 4)
        RangeString = CurSheet.Cells(iRow, 9) + CStr(CurSheet.Cells(iRow, 12) + 1) + ":" + CurSheet.Cells(iRow, 9) + CStr(CurSheet.Cells(iRow, 13))
        
        strListValue = CurSheet.Cells(iRow, iListValue + 1)
        If Len(strListValue) > 250 Then
            Call RefreshEnumBiggerThan255_i(SheetName, RangeString, strListValue, FieldName, CurrentColumn)
            CurrentColumn = CurrentColumn + 1
        End If
    Next
    
    Call SetSheetProtected(True)
End Sub


'设置字段名称
Private Sub SetFieldName(CurSheet As Worksheet, iRow As Integer)
    Dim FieldCol As String, FieldRow As String
    
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    FieldRow = GetSheetFieldDefRow
    CurSheet.Range(FieldCol + FieldRow) = FieldCol
End Sub

'设置字段显示名称
Private Sub SetDisplayName(CurSheet As Worksheet, iRow As Integer)
    Dim DisplayName As String, FieldCol As String, FieldRow As String
    Dim CheckNullValue As Integer
    
    DisplayName = Trim(SheetDefine(iRow, iFieldDisplayName_ENG))
    
    CheckNullValue = CInt(Trim(SheetDefine(iRow, iCheckNull)))
    If CheckNullValue = 0 Then
        DisplayName = "*" + DisplayName
    End If
       
    FieldRow = Trim(SheetDefine(iRow, iTitleBeginRow))
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    CurSheet.Range(FieldCol + FieldRow) = DisplayName
    '字体名称和大小
    With CurSheet.Range(FieldCol + FieldRow).Font
      .Name = "Arial"
      .Size = 9
      .Bold = True
    End With
    'Rows(Trim(CInt(FieldRow) + 1) + ":" + Trim(CInt(FieldRow) + 1)).Select
    'Selection.EntireRow.Hidden = True
End Sub

'设置批注信息
Private Sub SetFieldPostil(CurSheet As Worksheet, iRow As Integer)
    Dim FieldPostil As String, FieldCol As String
    Dim ENGName As String, CHSName As String, RangeName As String, FieldRow As String
    
    FieldRow = Trim(SheetDefine(iRow, iTitleBeginRow))
    ENGName = Trim(SheetDefine(iRow, iFieldPostil))
    CHSName = Trim(SheetDefine(iRow, iFieldDisplayName_CHS))  '此处处理有Bug
    RangeName = GetRangeInfo(iRow)
    If IsEnglishUsed And IsChineseUsed Then
        FieldPostil = ENGName + "(" + CHSName + ")"
    Else
        If IsEnglishUsed And RangeName <> "" Then
            FieldPostil = ENGName + Chr(10) + "(" + RangeName + ")"
        End If
        If IsChineseUsed And RangeName <> "" Then
            FieldPostil = CHSName + Chr(10) + "(" + RangeName + ")"
        End If
        If IsEnglishUsed And RangeName = "" Then
            FieldPostil = ENGName
        End If
        If IsChineseUsed And RangeName = "" Then
            FieldPostil = CHSName
        End If
    End If
    
    Dim BitmapInstruction As String
    If SheetDefine(iRow, iColumnType2) = "BITMAP" Then
        If IsEnglishUsed Then
            BitmapInstruction = "Note: This field uses 1 and 0 to indicate ON and OFF for each switch and does not contain delimiters. Example of the format: 111."
        End If
        If IsChineseUsed Then
            BitmapInstruction = "注意：此字段使用1，0分别表示每个开关位的开或关，无分隔符，格式形如：111。"
        End If
        FieldPostil = FieldPostil + Chr(10) + BitmapInstruction
    End If
    
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    CurSheet.Range(FieldCol + FieldRow).ClearComments
    CurSheet.Range(FieldCol + FieldRow).AddComment FieldPostil
    CurSheet.Range(FieldCol + FieldRow).Comment.Shape.Height = 160
    CurSheet.Range(FieldCol + FieldRow).Comment.Shape.Width = 120
End Sub

'获取有效范围提示或者错误信息
Private Function GetRangeInfo(iRow As Integer) As String
    Dim sFieldName As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String
    Dim sRangeDesp As String
    sFieldName = Trim(SheetDefine(iRow, iColumnFieldName))
    sDataType = Trim(SheetDefine(iRow, iColumnType))
    sMinVal = Trim(SheetDefine(iRow, iMin))
    sMaxVal = Trim(SheetDefine(iRow, iMax))
    sRangeList = Trim(SheetDefine(iRow, iListValue))
    
    If sMinVal = "" And sRangeList = "" Then
        GetRangeInfo = ""
        Exit Function
    End If
    
    GetRangeInfo = ""
    If (sDataType = FINT) Or (sDataType = FSTRING) Then
        If sMinVal = sMaxVal Then
            GetRangeInfo = GetValidErrMsg(sDataType) + "[" + sMinVal + "]"
        Else
            Dim vMinVal, vMaxVal
            vMinVal = Split(sMinVal, ",")
            vMaxVal = Split(sMaxVal, ",")
            
            GetRangeInfo = GetValidErrMsg(sDataType)
               
            Dim Index As Integer
            For Index = 0 To UBound(vMinVal)
                GetRangeInfo = GetRangeInfo + "[" + vMinVal(Index) + ".." + vMaxVal(Index) + "]"
            Next
        End If
    End If
    If sDataType = FLIST Then
        sRangeDesp = sRangeList
        If Len(sRangeList) > 250 Then
            sRangeDesp = Left(sRangeList, 200) + "..."
        End If
        GetRangeInfo = GetValidErrMsg(sDataType) + "[" + sRangeDesp + "]"
    End If
    If sFieldName = "Physical Type" Then
        GetRangeInfo = ""
    End If
End Function

Public Sub SetSheetProtected(fProtected As Boolean)
    Dim CurSheet As Worksheet
    Dim iDefSheet As Integer
    Dim SheetName As String
    
    Dim ShtIdx As Integer
    ShtIdx = 1
    Do While ShtIdx <= ActiveWorkbook.Sheets.Count
        If Not IsSystemSheet(ActiveWorkbook.Sheets(ShtIdx)) Then
            
            Set CurSheet = ActiveWorkbook.Sheets(ShtIdx)
            CurSheet.Activate
            If fProtected Then
            Call ProtectWorkSheet(CurSheet)
                Else
            Call UnprotectWorkSheet(CurSheet)
                End If
        
            CurSheet.Cells(7, 2).Select
        End If
        
        ShtIdx = ShtIdx + 1
    Loop
End Sub

Public Sub CheckFieldData(ByVal fieldValue As Range, CurSheetName As String)
    If IsFieldTypeBitmap(fieldValue, CurSheetName) Then
        Call CheckBitmapFieldData(fieldValue)
    End If
End Sub

Private Function IsFieldTypeBitmap(ByVal fieldValue As Range, CurSheetName As String) As Boolean
  IsFieldTypeBitmap = False
  
  Dim sht As Worksheet
  Set sht = Sheets("TableDef")
  
  Dim iRow As Integer
  For iRow = StartTblDataRow To StartTblDataRow + CInt(sht.Cells(5, 7))
      
      If CurSheetName = sht.Cells(iRow, SheetNameCol) _
          And fieldValue.Row <= CInt(sht.Cells(iRow, iContentEndRow + 1)) _
          And fieldValue.Row > CInt(sht.Cells(iRow, iTitleBeginRow + 1)) _
          And Chr(fieldValue.Column + 64) = sht.Cells(iRow, iFieldBeginColumn + 1) _
          And "BITMAP" = CStr(sht.Cells(iRow, iColumnType2 + 1)) Then
            
          IsFieldTypeBitmap = True
          Exit For
      End If
  Next
End Function

Public Sub CheckBitmapFieldData(ByVal Target As Range)
    Dim nValue As Long, nStrLen As Long, nLoop As Integer
    Dim sValue As String, sItem As String
    Dim bFlag As Boolean, nResponse As Boolean
    
    sValue = Target.Text
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

        If Not bFlag Then
            nResponse = MsgBox("Input Range [0,1]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, "Prompt")
            
            If nResponse = vbRetry Then
                Target.Cells(1, 1).Select
            End If
            
            Target.Cells(1, 1).ClearContents
        End If
    End If
End Sub

Public Sub RefreshBranchDefPos(Optional SheetName As String = "", Optional ObjName As String = "")
    Dim CurSheet As Worksheet
    Set CurSheet = ThisWorkbook.Worksheets("ValidDef")
    
    Dim iStartRow As Integer, iEndRow As Integer, iCurRow As Integer, iCurCol As Integer
    Dim bMasterFind As Boolean, bSubFind As Boolean
    iStartRow = 2
    iEndRow = GetValidDefRows
    
    Dim iTableEndRow As Integer, iTableCurRow As Integer
    
    'InitProgBar
    For iCurRow = iStartRow To iEndRow
        Application.StatusBar = "refresh branch control information..." + CStr((iCurRow - iStartRow) / (iEndRow - iStartRow) * 500)
        'Call SetProgBar("refresh branch control information...", (iCurRow - iStartRow) / (iEndRow - iStartRow) * 500)
        bMasterFind = False
        bSubFind = False
        
        If SheetName = "" Or (SheetName = ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, IDCol) And ObjName = ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, ObjectNameCol)) Then
            For iTableCurRow = StartTblDataRow To StartTblDataRow + CInt(ThisWorkbook.Worksheets("TableDef").Cells(5, 7))
            
                If ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, IDCol) = ThisWorkbook.Worksheets("TableDef").Cells(iTableCurRow, IDCol + 1) And _
                   ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, ObjectNameCol) = ThisWorkbook.Worksheets("TableDef").Cells(iTableCurRow, ObjectNameCol + 1) Then
                    
                    If ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, MasterFieldNameCol) = ThisWorkbook.Worksheets("TableDef").Cells(iTableCurRow, MasterFieldNameCol + 1) Then
                        ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, MasterFieldColCol) = ThisWorkbook.Worksheets("TableDef").Cells(iTableCurRow, iFieldBeginColumn + 1)
                        ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, BeginRowCol) = ThisWorkbook.Worksheets("TableDef").Cells(iTableCurRow, iTitleBeginRow + 1) + 1
                       ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, EndRowCol) = ThisWorkbook.Worksheets("TableDef").Cells(iTableCurRow, iContentEndRow + 1)
                       bMasterFind = True
                    End If
                    If ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, SubFieldNameCol) = ThisWorkbook.Worksheets("TableDef").Cells(iTableCurRow, MasterFieldNameCol + 1) Then
                        ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, SubFieldColColCol) = ThisWorkbook.Worksheets("TableDef").Cells(iTableCurRow, iFieldBeginColumn + 1)
                        bSubFind = True
                    End If
                End If
                If bMasterFind And bSubFind Then
                  Exit For
                End If
            Next iTableCurRow
        End If
    Next iCurRow
    'EndProgBar
    
    Application.StatusBar = "refresh branch define position finished."
End Sub

'针对扩展行之后的行，都增加1的操作
Public Sub RefreshBranchDefRow(SheetName As String, StartRow As Integer)
    Dim CurSheet As Worksheet
    Set CurSheet = ThisWorkbook.Worksheets("ValidDef")
    
    Dim iStartRow As Integer, iEndRow As Integer, iCurRow As Integer, iCurCol As Integer
    Dim bFind As Boolean
    iStartRow = 2
    iEndRow = GetValidDefRows
    
    Application.StatusBar = "refresh branch rows begin..."
    
    For iCurRow = iStartRow To iEndRow
        '只有要刷新的SheetName和当前游标的Valid行相同时，才需要刷新
        If SheetName = ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, IDCol) Then
            If ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, BeginRowCol) > StartRow Then
                ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, BeginRowCol) = ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, BeginRowCol) + 1
            End If
            
            If ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, EndRowCol) > StartRow Then
                ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, EndRowCol) = ThisWorkbook.Worksheets("ValidDef").Cells(iCurRow, EndRowCol) + 1
            End If
        End If
    Next iCurRow
    
    Application.StatusBar = "refresh branch rows finished."
End Sub

Public Sub EnsureRefreshBranchDefine()
    Dim CurSheet As Worksheet
    Set CurSheet = ThisWorkbook.Worksheets("ValidDef")
    
    Dim iStartRow As Integer, iEndRow As Integer, iCurRow As Integer, iCurCol As Integer
    Dim iValidDefineIndex As Integer, iRangeDefineIndex As Integer
    iValidDefineIndex = 0
    iRangeDefineIndex = 0
    
    iStartRow = 2
    iEndRow = GetValidDefRows
    
    For iCurRow = iStartRow To iEndRow
        If CurSheet.Cells(iCurRow, ValidFlagCol) = "NO" Then
            For iCurCol = 1 To ValidDefineColumns
                ValidDefine(iValidDefineIndex, iCurCol - 1) = CurSheet.Cells(iCurRow, iCurCol)
            Next iCurCol
            iValidDefineIndex = iValidDefineIndex + 1
        ElseIf CurSheet.Cells(iCurRow, ValidFlagCol) = "YES" Then
            For iCurCol = 1 To RangeDefineColumns
                RangeDefine(iRangeDefineIndex, iCurCol - 1) = CurSheet.Cells(iCurRow, iCurCol)
            Next iCurCol
            iRangeDefineIndex = iRangeDefineIndex + 1
        End If
    Next iCurRow
    
    Application.StatusBar = ""
End Sub

