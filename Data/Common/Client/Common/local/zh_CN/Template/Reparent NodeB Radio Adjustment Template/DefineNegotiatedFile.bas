Attribute VB_Name = "DefineNegotiatedFile"
Option Explicit

'������洦������
Public Const FINT = "INT", FSTRING = "STRING", FLIST = "LIST"
Public Const FBROW = 1, FEROW = 2, FBCOL = 3, FECOL = 4
Public Const DefinedWhite = 2
Public Const DefinedGreen = 35

'TableDef�б���
Public Const iMocName = 1
Public Const iColumnFieldName = 2
Public Const iColumnType = 3
Public Const iMin = 4
Public Const iMax = 5
Public Const iListValue = 6
Public Const iFieldBeginColumn = 7
Public Const iFieldEndColumn = 8
Public Const iColumnWidth = 9
Public Const iTitleBeginRow = 10
Public Const iContentEndRow = 11
Public Const iRowHeight = 12
Public Const iFieldDisplayName_ENG = 13
Public Const iFieldDisplayName_CHS = 14
Public Const iFieldPostil = 15
Public Const iMapTableName = 16
Public Const iMapFieldName = 17
Public Const iRealTableName = 18
Public Const iRealFieldName = 19
Public Const iCheckNull = 20
Public Const iColumnType2 = 21

Public Const StartTblDataRow = 15
Public Const SheetNameCol = 2

'�����µ�sheet
Public Sub ReCreateSheet()
  Dim sPreSheetName As String
  Dim sSheetName As String
  Dim Sht As Worksheet
  Set Sht = Sheets("TableDef")
  Dim curSheet As Worksheet
  Dim iNewSheet As Integer
  Dim iRow As Integer
  Dim iStartRow As Integer
  Dim iEndRow As Integer
  
  '��һ����ɾ��ԭ�е�sheetҳ
  'Application.DisplayAlerts = False
  'For iRow = StartTblDataRow To StartTblDataRow + CInt(Sht.Cells(5, 7)) - 1
  '    sSheetName = Sht.Cells(iRow, SheetNameCol)
  '    If IsSheetNameExists(sSheetName) Then
  '        ThisWorkbook.Sheets(sSheetName).Delete
  '    End If
  'Next
  'Application.DisplayAlerts = True
  
  '�ڶ����������µ�sheetҳ
  '���ӵ�һҳ
  iNewSheet = 1
  sPreSheetName = "Cell Adjustment"
  
  If Not IsSheetNameExists(sPreSheetName) Then
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(iNewSheet)
    ThisWorkbook.Sheets(iNewSheet + 1).Name = sPreSheetName
        
    iNewSheet = iNewSheet + 1
  End If
  
  iStartRow = StartTblDataRow
  Set curSheet = Sheets(sPreSheetName)
  For iRow = StartTblDataRow To StartTblDataRow + CInt(Sht.Cells(5, 7))
      If sPreSheetName <> Sht.Cells(iRow, SheetNameCol) Then
          '����ǰһҳ
          iEndRow = iRow - 1
          Call ClearSheetData(curSheet)
          Call GenNegotiatedData(curSheet, iStartRow - StartTblDataRow, iEndRow - StartTblDataRow)
          
          If "" = Sht.Cells(iRow, SheetNameCol) Then
            Exit Sub
          End If
          '���Ӻ�һҳ
          iStartRow = iRow
          sPreSheetName = Sht.Cells(iRow, SheetNameCol)
          If Not IsSheetNameExists(sPreSheetName) Then
            ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(iNewSheet)
            ThisWorkbook.Sheets(iNewSheet + 1).Name = sPreSheetName
        
            iNewSheet = iNewSheet + 1
          End If
          Set curSheet = Sheets(sPreSheetName)
      End If
  Next
End Sub

Public Sub ClearCoverSheetData()
  Dim curSheet As Worksheet
  Set curSheet = Sheets("Cover")

  Dim elmShp As Shape

  For Each elmShp In curSheet.Shapes
      If elmShp.Name Like "Group *" Then
         'Set elmShp = Nothing
         elmShp.Delete
      End If
  Next
End Sub

Public Sub ClearSheetData(curSheet As Worksheet)
  curSheet.Activate
  Cells.Select
  Selection.Clear

  Dim elmShp As Shape

  For Each elmShp In curSheet.Shapes
      If elmShp.Name Like "Group *" Then
         'Set elmShp = Nothing
         elmShp.Delete
      End If
  Next
End Sub

'��������������Э�����ݱ�
Public Sub GenNegotiatedFile()
    GeneratingFlag = 1
    Call UnprotectWorkBook
    TableSht.Visible = True
    '��ȡЭ�����ݶ���
    Call GetSheetDefineData
    Call SetSheetProtected(False)
    'ɾ��sheetҳ���������µ�sheet
    Call ReCreateSheet
    Call SetSheetProtected(True)
    
  '  CoverSht.Select
  '  Call ProtectWorkBook
    GeneratingFlag = 0
End Sub

'ΪSheet�����ֶ�������ʾ���ơ���ע��Ϣ
Public Sub GenNegotiatedData(curSheet As Worksheet, iStartRow As Integer, iEndRow As Integer)
    
    Dim iSheet As Integer, iDefSheet As Integer
    Dim FoundID As String
    Dim BField As Boolean
    
    Rows("2:2").Select
    Call SetSheetDefaultValue(curSheet)
    
    '------------------------------------------------------------
    '����ÿһ�е�����
    For iDefSheet = iStartRow To iEndRow
        '���ñ߿�
        Call SetFieldBorder(curSheet, iDefSheet)
        '�����и�
        Call SetDefRowHeight(curSheet, iDefSheet)
        '�����п�����
        Call SetFieldColWidth(curSheet, iDefSheet)
        '�����ֶ�
        Call SetFieldName(curSheet, iDefSheet)
        '������ʾ����
        Call SetDisplayName(curSheet, iDefSheet)
        '������ע��Ϣ
        Call SetFieldPostil(curSheet, iDefSheet)
        '����������Ч��
        Call SetFieldValidate(curSheet, iDefSheet)
        '���õ�Ԫ������
        Call ClearValidateForTitle(curSheet, iDefSheet)
    Next
    
    '------------------------------------------------------------
    '������ɣ�����Sheetҳ
    Range("BB1").Select
    ActiveCell.FormulaR1C1 = "1"
    Selection.AutoFill Destination:=Range("BB1:BB" + SheetDefine(iEndRow, iContentEndRow)), Type:=xlFillSeries
    Columns("BB:BB").Select
    Selection.EntireColumn.Hidden = True
    Selection.Locked = True
    Selection.FormulaHidden = True
    Rows("1:1").Select
    Selection.EntireRow.Hidden = True
    Selection.Locked = True
    Selection.FormulaHidden = True
    Range("B7").Select
    
    Sheets("TableDef").Select
    ActiveSheet.Shapes("Group 236").Select
    Selection.Copy
    curSheet.Activate
    Range("A3").Select
    curSheet.Paste
End Sub

        
'�����Ч�ԡ����õ�Ԫ������
Private Sub ClearValidateForTitle(curSheet As Worksheet, iRow As Integer)
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
    curSheet.Select

    'Field����
    Range(sFieldRange).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    'Title����
    Range(sTitleRange).Select
    Selection.Locked = True
    Selection.FormulaHidden = False
End Sub

'���ñ߿�,�ϲ���Ԫ��
Private Sub SetFieldBorder(curSheet As Worksheet, iRow As Integer)
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
    
    curSheet.Select
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
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    Sheets("TableDef").Select
    
    Range("C" + CStr(StartTblDataRow + iRow)).Select
    Selection.Copy
    curSheet.Select
    Range(sTitleRange).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.Merge
End Sub

'�����и�
Private Sub SetDefRowHeight(curSheet As Worksheet, iRow As Integer)
    Dim DefRowHeight As String
    Dim i As Integer, iSetRow As String
    
    DefRowHeight = Trim(SheetDefine(iRow, iRowHeight))
    curSheet.Activate
    If Trim(DefRowHeight) <> "" Then
        i = Trim(SheetDefine(iRow, iTitleBeginRow))
        iSetRow = Trim(CStr(i))
        With curSheet
            Rows(iSetRow + ":" + iSetRow).Select
            Selection.RowHeight = CSng(DefRowHeight)
        End With
    End If
End Sub

'�����п�
Private Sub SetFieldColWidth(curSheet As Worksheet, iRow As Integer)
    Dim sFieldWidth As String, FieldCol As String
    
    sFieldWidth = Trim(SheetDefine(iRow, iColumnWidth))
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    If sFieldWidth = "" Then Exit Sub
    curSheet.Select
    Columns(FieldCol + ":" + FieldCol).Select
    Selection.ColumnWidth = CSng(sFieldWidth)
    '�������ƺʹ�С
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

'����������Ч��
Private Sub SetFieldValidate(curSheet As Worksheet, iRow As Integer)
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
    sRangeList = Trim(SheetDefine(iRow, iListValue))
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    sFieldRange = FieldCol + sBeginRow + ":" + FieldCol + sEndRow

    If (sMinVal = "" And sRangeList = "") Then
        Exit Sub
    End If
    
    If (sDataType = FINT) Then
        xType = xlValidateWholeNumber
        sFormula1 = sMinVal
        sFormula2 = sMaxVal
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
    
    curSheet.Select
    If sEndRow = "" Then
        Columns(FieldCol + ":" + FieldCol).Select
    Else
        Range(sFieldRange).Select
    End If
    Call SetDataValidate(xType, sFormula1, sFormula2, sErrPrompt, sErrMsg)
        
    '���õ�Ԫ�����ָ�ʽ
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

'�����ֶ�����
Private Sub SetFieldName(curSheet As Worksheet, iRow As Integer)
    Dim FieldCol As String, FieldRow As String
    
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    FieldRow = GetSheetFieldDefRow
    curSheet.Range(FieldCol + FieldRow) = FieldCol
End Sub

'�����ֶ���ʾ����
Private Sub SetDisplayName(curSheet As Worksheet, iRow As Integer)
    Dim DisplayName As String, FieldCol As String, FieldRow As String
    
    FieldRow = Trim(SheetDefine(iRow, iTitleBeginRow))
    DisplayName = Trim(SheetDefine(iRow, iFieldDisplayName_ENG))
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    curSheet.Range(FieldCol + FieldRow) = DisplayName
    '�������ƺʹ�С
    With curSheet.Range(FieldCol + FieldRow).Font
      .Name = "Arial"
      .Size = 9
      .Bold = True
    End With
    'Rows(Trim(CInt(FieldRow) + 1) + ":" + Trim(CInt(FieldRow) + 1)).Select
    'Selection.EntireRow.Hidden = True
End Sub

'������ע��Ϣ
Private Sub SetFieldPostil(curSheet As Worksheet, iRow As Integer)
    Dim FieldPostil As String, FieldCol As String
    Dim ENGName As String, CHSName As String, RangeName As String, FieldRow As String
    
    FieldRow = Trim(SheetDefine(iRow, iTitleBeginRow))
    ENGName = Trim(SheetDefine(iRow, iFieldPostil))
    CHSName = Trim(SheetDefine(iRow, iFieldDisplayName_CHS))  '�˴�������Bug
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
    FieldCol = Trim(SheetDefine(iRow, iFieldBeginColumn))
    curSheet.Range(FieldCol + FieldRow).ClearComments
    curSheet.Range(FieldCol + FieldRow).AddComment FieldPostil
    curSheet.Range(FieldCol + FieldRow).Comment.Shape.Height = 160
    curSheet.Range(FieldCol + FieldRow).Comment.Shape.Width = 120
End Sub

'��ȡ��Ч��Χ��ʾ���ߴ�����Ϣ
Private Function GetRangeInfo(iRow As Integer) As String
    Dim sFieldName As String
    Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String
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
            GetRangeInfo = GetValidErrMsg(sDataType) + "[" + sMinVal + ".." + sMaxVal + "]"
        End If
    End If
    If sDataType = FLIST Then
        GetRangeInfo = GetValidErrMsg(sDataType) + "[" + sRangeList + "]"
    End If
    If sFieldName = "Physical Type" Then
        GetRangeInfo = ""
    End If
End Function

Public Sub SetSheetProtected(fProtected As Boolean)
    Dim curSheet As Worksheet
    Dim iDefSheet As Integer
    Dim SheetName As String
    
    Dim ShtIdx As Integer
    ShtIdx = 1
    Do While ShtIdx <= ActiveWorkbook.Sheets.Count
        If Not IsSystemSheet(ActiveWorkbook.Sheets(ShtIdx)) Then
            Set curSheet = ActiveWorkbook.Sheets(ShtIdx)
            curSheet.Select
            If fProtected Then
            Call ProtectWorkSheet(curSheet)
                Else
            Call UnprotectWorkSheet(curSheet)
                End If
        
            curSheet.Cells(7, 2).Select
        End If
        
        ShtIdx = ShtIdx + 1
    Loop
End Sub

