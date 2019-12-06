Attribute VB_Name = "DataDefine"
Const SheetNums = 29  'Sheet数量
Const TblRows = 383
Const TblCols = 33
Public SheetDefine(TblRows, TblCols) As String
Public ArrSheetName(SheetNums, 16) As String
Public ArrCellUnite(21, 3) As String
'******************************************************************************************
'对外接口函数：
'1、GetSheetDefineData
'2、IsEnglishUsed
'3、IsChineseUsed
'4、GetValidErrTitle
'5、GetValidErrMsg
'6、GetAllSheetName
'7、GetSheetFieldDefRow
'8、SetSheetDefaultValue

'******************************************************************************************
Const startRow = 3   'Define data start row
Const startCol = 1    'Define data start col
Const DefineSheetName = "TableDef"  '协商数据格式定义数据Sheet页
Const StartErrDataRow = startRow + 2
Const StartErrDataCol = startCol + 1
Const StartTblDataRow = StartErrDataRow + 10
Const StartTblDataCol = startCol
Const ErrRows = 6
Const ErrCols = 5

Dim ErrDefine(ErrRows, ErrCols) As String

'取所有定义数据
Public Sub GetSheetDefineData()
Dim CurSheet As Worksheet
Dim iRow As Integer, iCol As Integer

Set CurSheet = ThisWorkbook.Sheets(DefineSheetName)

For iRow = 0 To ErrRows - 1
  For iCol = 0 To ErrCols - 1
    ErrDefine(iRow, iCol) = CurSheet.Cells(StartErrDataRow + iRow, StartErrDataCol + iCol)
  Next
Next

For iRow = 0 To TblRows - 1
  For iCol = 0 To TblCols - 1
    SheetDefine(iRow, iCol) = CurSheet.Cells(StartTblDataRow + iRow, StartTblDataCol + iCol)
  Next
Next

For iRow = 0 To TblRows - 1
  SheetDefine(iRow, 26) = CurSheet.Cells(StartTblDataRow + iRow, StartTblDataCol + 2).Interior.ColorIndex
Next

For iRow = 0 To 20
  For iCol = 0 To 2
    ArrCellUnite(iRow, iCol) = CurSheet.Cells(StartErrDataRow + iCol, 11 + iRow)
  Next
Next

End Sub

'合并单元格处理
Public Sub SetCellUnite(CurSheet As Worksheet)
Dim iRow As Integer
Dim BRange As String

CurSheet.Select
For iRow = 0 To UBound(ArrCellUnite)
  If Trim(ArrCellUnite(iRow, 0)) = Trim(CurSheet.Name) Then
    Sheets("TableDef").Select
    range("C14").Select
    Selection.Copy
    CurSheet.Select
    range(ArrCellUnite(iRow, 1)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    BRange = Left(ArrCellUnite(iRow, 1), InStr(ArrCellUnite(iRow, 1), ":") - 1)
    range(BRange).value = Trim(ArrCellUnite(iRow, 2))
    range(BRange).HorizontalAlignment = xlLeft
    With range(BRange).Font
      .Name = "Arial"
      .Size = 9
      .Bold = True
    End With
    range(ArrCellUnite(iRow, 1)).Select
    Selection.Merge
  End If
Next
End Sub


'取所有协商数据Sheet页的语言版本信息
Public Function IsEnglishUsed() As Boolean
Const iTitleRow = 5
Const iTitleCol = 1
Const EngLng = "ENG"

IsEnglishUsed = (ErrDefine(iTitleRow, iTitleCol) = EngLng)
End Function

Public Function IsChineseUsed() As Boolean
Const iTitleRow = 5
Const iTitleCol = 2
Const ChsLng = "CHS"

IsChineseUsed = (ErrDefine(iTitleRow, iTitleCol) = ChsLng)
End Function

'取所有协商数据Sheet页的字段定义行
Public Function GetSheetFieldDefRow() As String
Const iFieldRow = 5
Const iFieldCol = 4

GetSheetFieldDefRow = ErrDefine(iFieldRow, iFieldCol)
End Function

'取所有协商数据Sheet页的首列列宽定义
Private Function GetSheetFisrtColWidth() As Single
Const iFieldRow = 5
Const iFieldCol = 3

GetSheetFisrtColWidth = CSng(ErrDefine(iFieldRow, iFieldCol))
End Function

'Sheet缺省行为统一设置
Public Function SetSheetDefaultValue(CurSheet As Worksheet) As Integer

CurSheet.Activate
With CurSheet
  '缺省行高
  Cells.Select
  Selection.RowHeight = 12
  '字体名称和大小
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
  '设置单元格锁定
  Selection.Locked = True
  Selection.FormulaHidden = False
  '隐藏字段行(缺省为第一行)
  Rows(GetSheetFieldDefRow + ":" + GetSheetFieldDefRow).Select
  Selection.EntireRow.Hidden = True
  '设置列宽（具体行列根据定义设置）
  Columns("A:A").Select
  Selection.ColumnWidth = GetSheetFisrtColWidth
End With
'设置零值显示、网格不显示
ActiveWindow.DisplayGridlines = False
ActiveWindow.DisplayZeros = True

End Function

'取错误描述信息标题
Public Function GetValidErrTitle(DataType As String) As String
Const iTitleCol = 3
Dim iRow As Integer

GetValidErrTitle = ""
iRow = GetValidErrData(DataType)
If iRow >= 0 Then GetValidErrTitle = ErrDefine(iRow, iTitleCol)
End Function

'取错误描述信息
Public Function GetValidErrMsg(DataType As String) As String
Const iMsgCol = 4
Dim iRow As Integer

GetValidErrMsg = ""
iRow = GetValidErrData(DataType)
If iRow >= 0 Then GetValidErrMsg = ErrDefine(iRow, iMsgCol)
End Function

Private Function GetValidErrData(DataType As String) As Integer
Const iDataType = 0
Dim iRow As Integer

GetValidErrData = -1
For iRow = 0 To ErrRows - 1
  If Trim(ErrDefine(iRow, iDataTypeCol)) = Trim(DataType) Then
    GetValidErrData = iRow
    Exit Function
  End If
Next
End Function

'获取所有Sheet页名称
Public Sub GetAllSheetName()
Const iSheetNumCol = 0
Const iSheetNameCol = 1
Const iDefRowHeightCol = 8
Const iTitleEndCol = 9
Const iDisplayTitleCol = 11

Dim iRow As Integer, iSheetNum As Integer, iCol As Integer

iSheetNum = 0
For iRow = 0 To TblRows - 1
  If Trim(SheetDefine(iRow, iSheetNumCol)) <> "" Then
    ArrSheetName(iSheetNum, 0) = Trim(SheetDefine(iRow, iSheetNumCol))
    ArrSheetName(iSheetNum, 1) = Trim(SheetDefine(iRow, iSheetNameCol))
    ArrSheetName(iSheetNum, 2) = Trim(SheetDefine(iRow + 1, iSheetNameCol))
    ArrSheetName(iSheetNum, 3) = Trim(SheetDefine(iRow, iDefRowHeightCol))
    ArrSheetName(iSheetNum, 4) = Trim(SheetDefine(iRow, iTitleEndCol))
    ArrSheetName(iSheetNum, 5) = Trim(SheetDefine(iRow, iDisplayTitleCol))
    ArrSheetName(iSheetNum, 6) = Trim(SheetDefine(iRow + 1, iDisplayTitleCol))
    iSheetNum = iSheetNum + 1
  End If
Next
End Sub
