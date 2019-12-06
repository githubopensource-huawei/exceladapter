Attribute VB_Name = "DefineNegotiatedFile"
'������洦������
Const FINT = "INT", FSTRING = "STRING", FLIST = "LIST"
Const FBROW = 1, FEROW = 2, FBCOL = 3, FECOL = 4

'��������������Э�����ݱ�
Public Sub GenNegotiatedFile()
'ThisWorkbook.Unprotect

ThisWorkbook.Sheets("TableDef").Visible = True
'��ȡÿ��Sheet��Э�����ݶ���
Call GetSheetDefineData
'��ȡ�������ö���Sheet����
Call GetAllSheetName
'ɾ������Sheet
Call DeleteAllSheet

'�������ö���Sheetҳ
Call CreateSheet
Call SetSheetProtected(False)
Call GenNegotiatedData
Call SetSheetUnite
Call SetSheetProtected(True)

ThisWorkbook.Sheets("TableDef").Visible = False
ThisWorkbook.Sheets("BSC").Visible = False
Sheets("Cover").Visible = True
Sheets("Cover").Select

'ThisWorkbook.Protect

End Sub

'ΪSheet�����ֶ�������ʾ���ơ���ע��Ϣ
Public Sub GenNegotiatedData()
'Field Display Name Row  Field Display Name(ENG) Field Display Name(CHS) Field Postil
Const FieldNameDisplayCol = 11  '11 12 13 14

Dim CurSheet As Worksheet
Dim iSheet As Integer, iDefSheet As Integer
Dim SheetID As String, FoundID As String
Dim BField As Boolean
Dim FieldDisplayRow As String, sTableRange As String

For iSheet = 0 To UBound(ArrSheetName) - 1
  BField = False
  SheetID = Trim(ArrSheetName(iSheet, 0))
  SheetName = Trim(ArrSheetName(iSheet, 1))
  sTableRange = Trim(ArrSheetName(iSheet, 6))
  Set CurSheet = ThisWorkbook.Sheets(SheetName)
  '�����и�
  Call SetDefRowHeight(CurSheet, iSheet)
  '���ñ߿�
  Call SetFieldBorder(CurSheet, iSheet)
  '------------------------------------------------------------
  '����ÿһ�е�����
  For iDefSheet = 0 To UBound(SheetDefine) - 1
    FoundID = Trim(SheetDefine(iDefSheet, 0))
    If SheetID = FoundID Then
      BField = True
      FieldDisplayRow = Trim(SheetDefine(iDefSheet, FieldNameDisplayCol))
    End If
    If FoundID <> "" And FoundID <> SheetID Then BField = False
    
    If BField Then
      '�����п�����
      Call SetFieldColWidth(CurSheet, iDefSheet)
      '�����ֶ�
      Call SetFieldName(CurSheet, iDefSheet)
      '������ʾ����
      Call SetDisplayName(CurSheet, iDefSheet, FieldDisplayRow)
      '������ע��Ϣ
      Call SetFieldPostil(CurSheet, iDefSheet, FieldDisplayRow)
      '����������Ч��
      Call SetFieldValidate(CurSheet, iDefSheet, sTableRange)
    End If
  Next
  '------------------------------------------------------------
  '������ɣ�����Sheetҳ
  Call ClearValidateForTitle(CurSheet, iSheet)
Next

End Sub

'�����Ч�ԡ����õ�Ԫ������
Private Sub ClearValidateForTitle(CurSheet As Worksheet, iRow As Integer)
Const iFieldRangeCol = 2, iTitleEndRowCol = 5, TableRangeCol = 6
Dim sFieldRangeCol As String, sTitleEndRow As String, sTitleRange As String, sTableRange As String

sTableRange = Trim(ArrSheetName(iRow, TableRangeCol))
sFieldRangeCol = Trim(ArrSheetName(iRow, iFieldRangeCol))
sTitleEndRow = Trim(ArrSheetName(iRow, iTitleEndRowCol))
CurSheet.Select
If Trim(sTableRange) = "" Then
    '�����Ч��
    Rows("1:" + sTitleEndRow).Select
    Call SetDataValidate(xlValidateInputOnly, "", "", "", "")
    '����
    Columns(sFieldRangeCol).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    '����
    Rows("1:" + sTitleEndRow).Select
    Selection.Locked = True
    Selection.FormulaHidden = False
Else
    '����
    Range(sTableRange).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
End If

End Sub

'���ñ߿򡢿���������
Private Sub SetFieldBorder(CurSheet As Worksheet, iRow As Integer)
Const iFieldRangeCol = 2, iTitleEndRowCol = 5, TableRangeCol = 6
Dim sFieldRangeCol As String, sTitleEndRow As String, sTitleRange As String, sTableRange As String

sFieldRangeCol = Trim(ArrSheetName(iRow, iFieldRangeCol))
sTitleEndRow = Trim(ArrSheetName(iRow, iTitleEndRowCol))
sTableRange = Trim(ArrSheetName(iRow, TableRangeCol))

CurSheet.Select
If Trim(sTableRange) = "" Then
  Columns(sFieldRangeCol).Select
Else
  Range(sTableRange).Select
End If
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

If Trim(sTableRange) = "" Then
    Rows("1:" + sTitleEndRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End If

sTitleRange = Left(sFieldRangeCol, 1) + sTitleEndRow + Mid(sFieldRangeCol, 2) + sTitleEndRow
Sheets("TableDef").Select
Range("C14").Select
Selection.Copy
CurSheet.Select
Range(sTitleRange).Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False

End Sub

'�����и�
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

'�����п�
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
'�������ƺʹ�С
With Selection.Font
  .Name = "Arial"
  .Size = 9
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

sFieldName = Trim(SheetDefine(iRow, FieldNameCol))
sDataType = Trim(SheetDefine(iRow, DataTypeCol))
FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
sMinVal = Trim(SheetDefine(iRow, MinValCol))
sMaxVal = Trim(SheetDefine(iRow, MaxValCol))
sRangeList = Trim(SheetDefine(iRow, RangeListCol))

If (sDataType = FINT) Then
  xType = xlValidateWholeNumber
  sFormula1 = sMinVal
  sFormula2 = sMaxVal
End If
If (sDataType = FSTRING) Then
  xType = xlValidateTextLength
  sFormula1 = sMinVal
  sFormula2 = sMaxVal
  CurSheet.Columns(FieldCol).NumberFormatLocal = "@"
  CurSheet.Columns(FieldCol).HorizontalAlignment = xlLeft
  CurSheet.Cells(4, FieldCol).HorizontalAlignment = xlCenter
End If
If sDataType = FLIST Then
  xType = xlValidateList
  sFormula1 = sRangeList
  sFormula2 = ""
  CurSheet.Columns(FieldCol).NumberFormatLocal = "@"
  CurSheet.Columns(FieldCol).HorizontalAlignment = xlLeft
  CurSheet.Cells(4, FieldCol).HorizontalAlignment = xlCenter
End If
If sFieldName = "LAC" Then
  xType = xlValidateCustom
  sFormula1 = "=OR(AND(" + FieldCol + "1<=65533," + FieldCol + "1>0), " + FieldCol + "1 = 65535)"
  sFormula2 = ""
End If

sErrPrompt = GetValidErrTitle(sDataType)
sErrMsg = GetRangeInfo(iRow)

CurSheet.Select
If sTableRange = "" Then
  Columns(FieldCol + ":" + FieldCol).Select
Else
  Range(FieldCol + GetCellRowOrCol(sTableRange, FBROW) + ":" + FieldCol + GetCellRowOrCol(sTableRange, FEROW)).Select
End If
Call SetDataValidate(xType, sFormula1, sFormula2, sErrPrompt, sErrMsg)

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
Private Sub SetFieldName(CurSheet As Worksheet, iRow As Integer)
Const FieldNameCol = 2
Const FieldDefCol = 4
Dim FieldName As String, FieldCol As String, FieldRow As String

FieldName = Trim(SheetDefine(iRow, FieldNameCol))
FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
FieldRow = GetSheetFieldDefRow
CurSheet.Range(FieldCol + FieldRow) = FieldName
CurSheet.Cells(4, FieldCol).Interior.ColorIndex = SheetDefine(iRow, 26)
End Sub

'�����ֶ���ʾ����
Private Sub SetDisplayName(CurSheet As Worksheet, iRow As Integer, FieldRow As String)
Const FieldDefCol = 4
Const FieldNameDisplayCol = 11
Const FieldNameCol = 2
Dim FieldName As String, DisplayName As String, FieldCol As String

DisplayName = Trim(SheetDefine(iRow, FieldNameDisplayCol + 1))
FieldName = Trim(SheetDefine(iRow, FieldNameCol))
FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
CurSheet.Range(FieldCol + FieldRow) = DisplayName
'�������ƺʹ�С
With CurSheet.Range(FieldCol + FieldRow).Font
  .Name = "Arial"
  .Size = 9
  .Bold = True
End With
CurSheet.Range(FieldCol + Trim(CInt(FieldRow) + 1)) = FieldName
Rows(Trim(CInt(FieldRow) + 1) + ":" + Trim(CInt(FieldRow) + 1)).Select
Selection.EntireRow.Hidden = True

End Sub

'������ע��Ϣ
Private Sub SetFieldPostil(CurSheet As Worksheet, iRow As Integer, FieldRow As String)
Const FieldDefCol = 4
Const FieldNameDisplayCol = 11
Dim FieldPostil As String, FieldCol As String
Dim ENGName As String, CHSName As String, RangeName As String

ENGName = Trim(SheetDefine(iRow, FieldNameDisplayCol + 3))
CHSName = Trim(SheetDefine(iRow, FieldNameDisplayCol + 2))
RangeName = GetRangeInfo(iRow)
If IsEnglishUsed And IsChineseUsed Then
  FieldPostil = ENGName + "(" + CHSName + ")"
Else
  If IsEnglishUsed Then
    FieldPostil = ENGName + "(" + RangeName + ")"
  End If
  If IsChineseUsed Then
    FieldPostil = CHSName + "(" + RangeName + ")"
  End If
End If
FieldCol = Trim(SheetDefine(iRow, FieldDefCol))
CurSheet.Range(FieldCol + FieldRow).ClearComments
CurSheet.Range(FieldCol + FieldRow).AddComment FieldPostil
 
End Sub

'��ȡ��Ч��Χ��ʾ���ߴ�����Ϣ
Private Function GetRangeInfo(iRow As Integer) As String
Const FieldNameCol = 2
Const DataTypeCol = 3
Const MinValCol = 5
Const MaxValCol = 6
Const RangeListCol = 7
Const ValueTypeCol = 24

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

Public Sub SetSheetUnite()
Dim CurSheet As Worksheet
Dim iDefSheet As Integer
Dim SheetName As String

On Error Resume Next

For iDefSheet = 0 To UBound(ArrSheetName) - 1
  SheetName = Trim(ArrSheetName(iDefSheet, 1))
  Set CurSheet = ThisWorkbook.Sheets(SheetName)
  CurSheet.Select
  '���úϲ�����
  Call SetCellUnite(CurSheet)
Next

End Sub

Public Sub SetSheetProtected(fProtected As Boolean)
Dim CurSheet As Worksheet
Dim iDefSheet As Integer
Dim SheetName As String

'���Sheet��Ŷ������,����WorkSheet���ܵ�ҳ��,���ܵ�������Sheet����
On Error Resume Next

For iDefSheet = 0 To UBound(ArrSheetName) - 1
  SheetName = Trim(ArrSheetName(iDefSheet, 1))
  Set CurSheet = ThisWorkbook.Sheets(SheetName)
  CurSheet.Select
'  If fProtected Then
'    Call ProtectWorkSheet(CurSheet)
'  Else
'    Call UnprotectWorkSheet(CurSheet)
'  End If
  CurSheet.Cells(6, 2).Select
Next

End Sub

Public Sub CreateSheet()
Dim CurSheet As Worksheet
Dim iSheet As Integer, iDefSheet As Integer, iNewSheet As Integer
Dim bFound As Boolean, SheetName As String, bTableDef  As Boolean

'���Sheet��Ŷ������,����WorkSheet���ܵ�ҳ��,���ܵ�������Sheet����
On Error Resume Next

For iDefSheet = 0 To UBound(ArrSheetName) - 1
  iNewSheet = CInt(ArrSheetName(iDefSheet, 0))
  bTableDef = (Trim(ArrSheetName(iDefSheet, 6)) = "")   '�Ƿ���Ҫ�����
  If iNewSheet > 1 Then iNewSheet = iNewSheet - 1
  SheetName = Trim(ArrSheetName(iDefSheet, 1))
  
  bFound = False
  For iSheet = 1 To ThisWorkbook.Sheets.count
    Set CurSheet = ThisWorkbook.Sheets(iSheet)
    If Trim(CurSheet.Name) = SheetName Then
      bFound = True
      Exit For
    End If
  Next
  
  If bFound Then ThisWorkbook.Sheets(SheetName).Delete
  
  If bTableDef And Not bFound Then
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(iNewSheet)
    ThisWorkbook.Sheets(iNewSheet + 1).Name = SheetName
    Set CurSheet = ThisWorkbook.Sheets(iNewSheet + 1)
    Call SetSheetDefaultValue(CurSheet)
  End If
  
Next

End Sub

Public Sub DeleteAllSheet()
Dim SheetName As String, iDefSheet  As Integer

On Error Resume Next

Application.DisplayAlerts = False
For iDefSheet = 0 To UBound(ArrSheetName) - 1
  SheetName = Trim(ArrSheetName(iDefSheet, 1))
  ThisWorkbook.Sheets(SheetName).Delete
Next
Application.DisplayAlerts = True
Application.Wait 500

End Sub
Private Function GetCellRowOrCol(sRange As String, sCellType As Integer) As String
Const SepChar = ":"
Dim BCell As String, ECell As String

GetCellRowOrCol = ""
BCell = Trim(Left(sRange, InStr(sRange, SepChar) - 1))
ECell = Trim(Mid(sRange, InStr(sRange, SepChar) + 1))

If sCellType = FBROW Then
  GetCellRowOrCol = Mid(BCell, 2)
End If
If sCellType = FEROW Then
  GetCellRowOrCol = Mid(ECell, 2)
End If
If sCellType = FBCOL Then
  GetCellRowOrCol = Left(BCell, 1)
End If
If sCellType = FECOL Then
  GetCellRowOrCol = Left(ECell, 1)
End If

End Function


