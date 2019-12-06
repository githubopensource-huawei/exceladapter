Attribute VB_Name = "DataDefine"
Const SheetNums = 1  'Sheet数量
'Const TblRows = 67  '常修改
Const TblCols = 26
'Const ValidRows = 19 '常修改
Const ValidCols = 9
'Const RangeBeginRow = 29 '常修改
Const RangeBeginCol = 2
'Const RangeRows = 9 '常修改
Const RangeCols = 12
Const InvalidBeginRow = 4
Const InvalidBeginCol = 2
Const DefinedGray = 15

'ValidDef中的常量
Public Const ValidDefineColumns = 10
Public Const RangeDefineColumns = 15
Public Const ValidFlagCol = 6
Public Const MasterFieldColCol = 7
Public Const SubFieldColColCol = 8
Public Const BeginRowCol = 9
Public Const EndRowCol = 10
Public Const IDCol = 1
Public Const ObjectNameCol = 2
Public Const MasterFieldNameCol = 3
Public Const SubFieldNameCol = 4

Public SheetDefine(1200, 52) As String '此处仍有Bug
Public ValidDefine(600, ValidDefineColumns) As String
Public RangeDefine(100, RangeDefineColumns) As String

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
Const StartRow = 3   'Define data start row
Const StartCol = 1    'Define data start col
Const DefineSheetName = "TableDef"  '协商数据格式定义数据Sheet页
Const StartErrDataRow = StartRow + 2
Const StartErrDataCol = StartCol + 1
Const StartTblDataRow = StartErrDataRow + 10
Const StartTblDataCol = StartCol
Const ErrRows = 6
Const ErrCols = 5

Public GeneratingFlag As Integer  '0表示正在生成
Dim ErrDefine(ErrRows, ErrCols) As String

Public Function TblRows() As Integer
  TblRows = CInt(Sheets("TableDef").Cells(5, 7))
End Function

Public Function ValidRows() As Integer
  ValidRows = CInt(Sheets("ValidDef").Cells(1, 3))
End Function

Public Function RangeBeginRow() As Integer
  RangeBeginRow = CInt(Sheets("ValidDef").Cells(1, 5))
End Function

Public Function RangeRows() As Integer
  RangeRows = CInt(Sheets("ValidDef").Cells(1, 7))
End Function

Public Function GetInterfaceBeginRow(InterfaceType As String) As Integer
  Set TableDefSht = Sheets("TableDef")
  TableDefSht.Activate

  TableDefSht.UsedRange.Range("B" + CStr(StartTblDataRow) + ":B" + CStr(TblRows + TableDefSht_BeginRow)).Select
  GetInterfaceBeginRow = Selection.Find(What:=InterfaceType, After:=ActiveCell).Row
End Function
  
  
'取所有定义数据
Public Sub GetSheetDefineData()
    Dim CurSheet As Worksheet
    Dim iRow As Integer, iCol As Integer
    
    Set CurSheet = ThisWorkbook.Worksheets("TableDef")
    
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
      Selection.Delete Shift:=xlUp
      Selection.RowHeight = 11.25
      Selection.Interior.ColorIndex = DefinedGray
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

'取ValidDef行数
Public Function GetValidDefRows() As Integer
    Dim iRow As Integer
    
    For iRow = 1 To 1000
        If ThisWorkbook.Worksheets("ValidDef").Cells(iRow, 2) = "" Then
            GetValidDefRows = iRow - 1
            Exit Function
        End If
    Next
End Function

'取所有定义数据
Public Sub GetValidDefineData()
    Dim CurSheet As Worksheet
    Dim iRow As Integer, iCol As Integer, iAllRow As Integer, iAllCol As Integer
    iAllRow = GetValidDefRows()
    
    If ValidDefine(0, 0) <> "" And RangeDefine(0, 0) <> "" Then
      Exit Sub
    End If
    
    Call EnsureRefreshBranchDefine
End Sub

