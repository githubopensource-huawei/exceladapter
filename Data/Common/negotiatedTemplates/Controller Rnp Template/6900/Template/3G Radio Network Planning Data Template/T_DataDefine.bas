Attribute VB_Name = "T_DataDefine"

Public Const gRangeStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'---------------------------------------------------------------------

'2010-03-31 z00102652增加4页 BSCInfo DoubleFrequencyCell WholeNetworkCell DeleteInterNCellRelation
'2010-09-30 z00221168删除3页 BSCInfo DoubleFrequencyCell WholeNetworkCell
Const SheetNums = 11 + 4 - 3 'Sheet数量 9+4-3

Const TblCols = 17
Const DefTblRows = 500
Public SheetDefine(DefTblRows, TblCols) As String
Public ArrSheetName(SheetNums, 6) As String

Public GeneratingFlag As Integer '1表示正在保存 0表示没有在保存的状态

Const StartRow = 3   'Define data start row
Const StartCol = 1    'Define data start col
Const DefineSheetName = "TableDef"  '协商数据格式定义数据Sheet页
Const StartErrDataRow = StartRow + 2
Const StartErrDataCol = StartCol + 1
Const StartTblDataRow = StartRow + 15
Const StartTblDataCol = StartCol
Const ErrRows = 4
Const ErrCols = 5

Public ValidDefine(150, 9) As String
Public RangeDefine(500, 13) As String
Const ValidCols = 9
Const RangeCols = 13
Const RangeBeginCol = 2
Const InvalidBeginRow = 4
Const InvalidBeginCol = 2

Dim ErrDefine(ErrRows, ErrCols) As String

'版本信息
Public sCMEVersion As String
Public sNEVersion As String
Public sRNPVersion As String

Public sChsCoverInfoTitle, sChsCoverInfo1, sChsCoverInfo2, sEngCoverInfoTitle, sEngCoverInfo1, sEngCoverInfo2 As String

Public gEngNEVersion, gChsNEVersion As String
Public gChsRNPVersion, gEngRNPVersion As String
Public gChsTemplateName, gEngTemplateName As String

Public gChsIsMustGive, gEngIsMustGive As String

Public iLanguageType As Integer   '语言设置  0 - 英文  1 - 中文  ......

Public iHideSheetFlg As Integer   '扩展功能Sheet页隐藏标识   0 - 隐藏   1 - 显示




'****************************************************************
'返回各表格定义参数的总行数
'****************************************************************
Function TblRows() As Integer
    TblRows = CInt(ThisWorkbook.Sheets(DefineSheetName).Cells(9, 7))
End Function

'****************************************************************
'从TableDef表中读取所有Sheet的定义参数数据
'****************************************************************
Public Sub GetSheetDefineData()
    Dim CurSheet As Worksheet
    Dim iRow As Integer, iCol As Integer
    
    Set CurSheet = ThisWorkbook.Sheets(DefineSheetName)
    
    iLanguageType = CInt(CurSheet.Range("H9").Text)
    If iLanguageType = 1 Then
        sChsCoverInfoTitle = CurSheet.Range("I5").Text
        sChsCoverInfo1 = CurSheet.Range("J5").Text
        sChsCoverInfo2 = CurSheet.Range("K5").Text
        gChsTemplateName = CurSheet.Range("I6").Text
        gChsNEVersion = CurSheet.Range("I7").Text
        gChsRNPVersion = CurSheet.Range("I8").Text
        gChsIsMustGive = CurSheet.Range("I9").Text
    Else
        sEngCoverInfoTitle = CurSheet.Range("I4").Text
        sEngCoverInfo1 = CurSheet.Range("J4").Text
        sEngCoverInfo2 = CurSheet.Range("K4").Text
        gEngTemplateName = CurSheet.Range("J6").Text
        gEngNEVersion = CurSheet.Range("J7").Text
        gEngRNPVersion = CurSheet.Range("J8").Text
        gEngIsMustGive = CurSheet.Range("J9").Text
        
    End If
    sCMEVersion = CurSheet.Range("H4").Text
    sNEVersion = CurSheet.Range("H5").Text
    sRNPVersion = CurSheet.Range("H6").Text
    
    For iRow = 0 To ErrRows - 1
        For iCol = 0 To ErrCols - 1
            ErrDefine(iRow, iCol) = CurSheet.Cells(StartErrDataRow + iRow, StartErrDataCol + iCol)
        Next
    Next
    
    
    For iRow = 0 To TblRows - 1
        For iCol = 0 To TblCols - 1
            SheetDefine(iRow, iCol) = CurSheet.Cells(StartTblDataRow + iRow, StartTblDataCol + iCol)
            'Debug.Print "[" & iRow & "," & iCol & "]" & SheetDefine(iRow, iCol)
        Next
        'Debug.Print "[" & iRow & "," & 0 & "]" & SheetDefine(iRow, 0)
        
    Next
    
End Sub

'****************************************************************
'获取所有Sheet页名称
'****************************************************************
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
            
            iSheetNum = iSheetNum + 1
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

'****************************************************************
'取错误描述信息标题
'****************************************************************
Public Function GetValidErrTitle(DataType As String) As String
    Const iTitleCol = 3
    Dim iRow As Integer
    
    GetValidErrTitle = ""
    iRow = GetValidErrData(DataType)
    If iRow >= 0 Then GetValidErrTitle = ErrDefine(iRow, iTitleCol)
End Function

'****************************************************************
'取错误描述信息
'****************************************************************
Public Function GetValidErrMsg(DataType As String) As String
    Const iMsgCol = 4
    Dim iRow As Integer
    
    GetValidErrMsg = ""
    iRow = GetValidErrData(DataType)
    If iRow >= 0 Then GetValidErrMsg = ErrDefine(iRow, iMsgCol)
End Function

'****************************************************************
'取得DataType指定类型在ErrDefine中的行索引
'****************************************************************
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

Public Function ValidRows() As Integer
  ValidRows = CInt(Sheets("ValidDef").Cells(1, 3))
End Function

Public Function RangeRows() As Integer
  RangeRows = CInt(Sheets("ValidDef").Cells(1, 7))
End Function

Public Function RangeBeginRow() As Integer
  RangeBeginRow = CInt(Sheets("ValidDef").Cells(1, 5))
End Function


