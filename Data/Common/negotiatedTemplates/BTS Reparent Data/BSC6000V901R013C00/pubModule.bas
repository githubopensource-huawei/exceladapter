Attribute VB_Name = "pubModule"

Const CheckFlagCol As Byte = 7
Const MaxSheetPara As Byte = 5           'SheetName DataSheetName StartCol EndCol MaxRows
Const MaxImportSheets As Integer = 10

Dim StartTableRow As Byte                '数据起始行

Dim ErrLogLines As Long                  '日志行数
Dim ArrSheetPara(MaxImportSheets, MaxSheetPara) As String
Dim WorkBookMaxLines As Long


Private Function GetWorkBookPass() As String
  GetWorkBookPass = Sheets("Cover").Cells(1, 2)
End Function


Private Function IsInList(ByVal iValue As String, ByVal iList As String) As Boolean
Dim ArrData() As String
Dim i As Integer

ArrData = Split(iList, ",")
For i = 0 To UBound(ArrData)
  If Trim(ArrData(i)) = Trim(iValue) Then
    IsInList = True
    Exit Function
  End If
Next
IsInList = False
End Function


'****************************************************************
'判断行数据是否有空值,若有1个单元格无数据（非空行）校验通过
'****************************************************************
Function RowDataHasNull(CurSheet As Worksheet, intRow As Long, startCol As Long, EndCol As Long) As Boolean
Dim i As Long

RowDataHasNull = False
For i = startCol To EndCol Step 1
  CurSheet.Cells(intRow, i).value = Trim(CurSheet.Cells(intRow, i).value)
  If (isEmpty(CurSheet.Cells(intRow, i)) Or (Trim(CurSheet.Cells(intRow, i)) = "")) Then
    RowDataHasNull = True
    Exit For
  End If
Next
End Function

'****************************************************************
'判断是否为空行,若连续20个单元格无数据则认为是空行
'****************************************************************
Function RowIsNull(CurSheet As Worksheet, intRow As Long) As Boolean
Const ROWCOUNT As Byte = 20
Dim i As Long

RowIsNull = True
For i = 1 To ROWCOUNT Step 1
  If (CurSheet.Cells(intRow, i) <> Empty) And (Trim(CurSheet.Cells(intRow, i)) <> "") Then
    RowIsNull = False
    Exit For
  End If
Next
End Function

'****************************************************************
'进度条指示
'****************************************************************
Private Sub StatusProcess(iStep As Long)
Dim iBase As Long
Dim iValue As Long

iBase = 100
If iStep > 0 Then
  iValue = StatusForm.ProgressBar1.value + iStep * 100 / iBase
  If iValue > 100 Then
    iValue = 100
  End If
  StatusForm.ProgressBar1.value = iValue
  StatusForm.Repaint
End If
End Sub

Private Sub WriteResult(fResult As Boolean, fPrompt As Boolean)
If fResult Then
  If fPrompt Then MsgBox "Validity check successfully."
  Sheets("Work").Cells(3, CheckFlagCol) = "Success"
  Sheets("Cover").Activate
Else
  Sheets("Work").Cells(3, CheckFlagCol) = "Failure"
  If fPrompt Then
    'ActiveWorkbook.Unprotect Password:=GetSheetsPass
    Sheets("CheckResult").Visible = True
    Sheets("CheckResult").Activate
    'ActiveWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
  End If
End If
End Sub

Public Sub ClearResult()
On Error Resume Next
Sheets("Work").Cells(3, CheckFlagCol) = ""
End Sub
'****************************************************************
'记录错误信息
'****************************************************************
Private Function WriteErrLog(CurSheet As Worksheet, ErrMsg As String, iRow As Long) As Boolean
  ErrLog Space(2) + "Error: " + ErrMsg
  CurSheet.Cells(iRow, 1).value = "X"
  WriteErrLog = False
End Function
'****************************************************************
'记录警告信息
'****************************************************************
Private Function WriteWarning(CurSheet As Worksheet, ErrMsg As String, iRow As Long) As Boolean
  ErrLog Space(2) + "Warning: " + ErrMsg
  CurSheet.Cells(iRow, 1).value = "?"
  WriteWarning = False
End Function

Private Sub ClearErrOrWarning(CurSheet As Worksheet, iRow As Long)
  CurSheet.Cells(iRow, 1).value = ""
End Sub

'****************************************************************
'登记错误日志文件
'****************************************************************
Private Sub ErrLog(ErrMsg As String)
  ErrLogLines = ErrLogLines + 1
  Sheets("CheckResult").Cells(ErrLogLines, 2).value = ErrMsg
End Sub

'****************************************************************
'获取上下频点和频段指示对应关系表
'****************************************************************
Private Sub GetBandAndFreq(CurSheet As Worksheet, startRow As Integer, startCol As Integer)
Dim i As Integer, j As Integer

For i = 0 To MaxBANDINDs - 1
  For j = 0 To MaxBANDINDCol - 1
    
  Next
Next
End Sub

'****************************************************************
'获取 Sheet 参数值（sheet名称 数据起始行 数据起始列 数据结束列）
'****************************************************************
Private Sub GetSheetPara(CurSheet As Worksheet, startRow As Integer, startCol As Integer)
Dim i As Integer, j As Integer

For i = 0 To MaxImportSheets - 1
  For j = 0 To MaxSheetPara - 2
    ArrSheetPara(i, j) = Trim(CurSheet.Cells(startRow + i, startCol + j))
  Next
  'ArrSheetPara(i, j) = Trim(CStr(GetWorkSheetMaxLines(Sheets(ArrSheetPara(i, 0)))))
Next

End Sub

'****************************************************************
'获取随机数
'****************************************************************
Private Function GetRandomValue(iPercent As Integer) As Integer
  Randomize
  GetRandomValue = CInt((iPercent * Rnd()) + 1)
  If GetRandomValue < (iPercent / 2) Then GetRandomValue = iPercent / 2 + GetRandomValue
End Function

'****************************************************************
'清除错误日志
'****************************************************************
Private Sub ClearErrLog()
  'ActiveWorkbook.Unprotect Password:=GetSheetsPass
  Sheets("CheckResult").Visible = False
  Sheets("CheckResult").Activate
  Range("B2:B65536").Select
  Selection.ClearContents
  'ActiveWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
End Sub

'****************************************************************
'判断是否为表尾,若连续5行为空行则认为是表尾
'****************************************************************
Function EndOfSheet(CurSheet As Worksheet, intRow As Long) As Boolean
Dim i As Long
Const ROWCOUNT As Byte = 5

EndOfSheet = True
For i = intRow To intRow + ROWCOUNT Step 1
  If (Not RowIsNull(CurSheet, i)) Then
    EndOfSheet = False
    Exit For
  End If
Next

End Function

Sub InsertUserToolBar()
    Dim cmbNewBar As CommandBar
    Dim ctlBtn As CommandBarButton
    
    '中英文按钮栏
    Dim sCHSBarName As String
    Dim sENGBarName As String
    Dim iLanguage  As Integer
    Dim sheetTableDef As Worksheet
    Set sheetTableDef = ThisWorkbook.Sheets("TableDef")
    iLanguage = sheetTableDef.Cells(5, 8).value
    If iLanguage = 0 Then
       sENGBarName = sheetTableDef.Cells(5, 9).value
       sCHSBarName = sheetTableDef.Cells(6, 9).value
      End If
         
    If iLanguage = 1 Then
       sENGBarName = sheetTableDef.Cells(5, 10).value
       sCHSBarName = sheetTableDef.Cells(6, 10).value
      End If
      
    On Error Resume Next
    Set cmbNewBar = CommandBars.Add(Name:="Operate Bar")
    
    'English
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = sENGBarName
            .TooltipText = sENGBarName
            .OnAction = "SwitchEng"
            .FaceId = 28
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
    'Chinese
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = sCHSBarName
            .TooltipText = sCHSBarName
            .OnAction = "SwitchChs"
            .FaceId = 28
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
    
End Sub

Sub DeleteUserToolBar()
  On Error Resume Next
  CommandBars("Operate Bar").Delete
End Sub

Sub UsrExportXML()

End Sub

Sub UsrImportXML()

End Sub

Public Sub UnprotectWorkSheet(CurSheet As Worksheet)
  On Error Resume Next
  CurSheet.Unprotect (GetSheetsPass)
  Application.ScreenUpdating = True
End Sub

Public Sub ProtectWorkSheet(CurSheet As Worksheet)
  On Error Resume Next
  CurSheet.Protect Password:=GetSheetsPass, AllowFormattingCells:=True, AllowFormattingColumns:=True
  Application.ScreenUpdating = True
End Sub

Public Sub ProtectWorkBook()
  On Error Resume Next
  ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
  ThisWorkbook.Save
End Sub

Public Sub UnprotectWorkBook()
  On Error Resume Next
  ThisWorkbook.Unprotect Password:=GetSheetsPass
  ThisWorkbook.Save
End Sub

Sub SetSysOption()
  With ActiveWindow
    If .DisplayZeros Then Exit Sub
    .DisplayGridlines = False
    .DisplayZeros = True
  End With
End Sub


Public Sub SwitchChs()

  ThisWorkbook.Unprotect
  
  Dim nRowIndex As Integer
  Dim sheetTableDef As Worksheet
  Dim sheetCurrent As Worksheet
  Dim strObjId As String
  Dim nObjId As Integer
  Dim strObjName As String
  Dim TableName As String
  
  SetBarLanguage ("CHS")
     
  Set sheetTableDef = ThisWorkbook.Sheets("TableDef")
  
  '设置批注信息
  Const FieldDefCol = 4
  Const FieldNameDisplayCol = 11
  Dim CHSName As String
  Dim RangeName As String
  Dim FieldPostil As String
  Dim nObjIndex As Integer
    
  For nRowIndex = 15 To 15000
  
    If Trim(sheetTableDef.Cells(nRowIndex, 3).value) = "" Then
      Exit For
    End If
    
    strObjId = Trim(sheetTableDef.Cells(nRowIndex, 1).value)
    If strObjId <> "" Then
      TableName = Trim(sheetTableDef.Cells(nRowIndex, 2).value)
      Set sheetCurrent = ThisWorkbook.Sheets(TableName)
      nObjIndex = 2
    End If
        
    ' 设置批注信息
    CHSName = Trim(sheetTableDef.Cells(nRowIndex, 14))
    RangeName = GetRangeInfoCHS(nRowIndex, sheetTableDef)
    FieldPostil = CHSName + "(" + RangeName + ")"
    
    FieldPostil = FieldPostil + Chr(13) + Chr(10) + sheetTableDef.Cells(nRowIndex, 19)
    
    sheetCurrent.Cells(4, nObjIndex).ClearComments
    sheetCurrent.Cells(4, nObjIndex).AddComment FieldPostil
    sheetCurrent.Cells(4, nObjIndex).Comment.Shape.Width = 200
    sheetCurrent.Cells(4, nObjIndex).Comment.Shape.Height = 200
    sheetCurrent.Cells(4, nObjIndex) = CHSName
    
    nObjIndex = nObjIndex + 1
    
  Next nRowIndex
 
  Dim sheetCoverEng As Worksheet
  Dim sheetCoverChs As Worksheet
  
  Set sheetCoverEng = ThisWorkbook.Sheets("Cover")
  Set sheetCoverChs = ThisWorkbook.Sheets("封面")
  
  sheetCoverEng.Visible = False
  sheetCoverChs.Visible = True
  
  MsgBox "完成."
  
  ThisWorkbook.Protect
  
End Sub


Public Sub SwitchEng()
  ThisWorkbook.Unprotect
  
  Dim nRowIndex As Integer
  Dim sheetTableDef As Worksheet
  Dim sheetCurrent As Worksheet
  Dim strObjId As String
  Dim nObjId As Integer
  Dim strObjName As String
  Dim TableName As String
  
  SetBarLanguage ("ENG")
     
  Set sheetTableDef = ThisWorkbook.Sheets("TableDef")
  
  '设置批注信息
  Const FieldDefCol = 4
  Const FieldNameDisplayCol = 11
  Dim CHSName As String
  Dim RangeName As String
  Dim FieldPostil As String
  Dim nObjIndex As Integer
    
  For nRowIndex = 15 To 15000
  
    If Trim(sheetTableDef.Cells(nRowIndex, 3).value) = "" Then
      Exit For
    End If
    
    strObjId = Trim(sheetTableDef.Cells(nRowIndex, 1).value)
    If strObjId <> "" Then
      TableName = Trim(sheetTableDef.Cells(nRowIndex, 2).value)
      Set sheetCurrent = ThisWorkbook.Sheets(TableName)
      nObjIndex = 2
    End If
        
    ' 设置批注信息
    CHSName = Trim(sheetTableDef.Cells(nRowIndex, 13))
    RangeName = GetRangeInfoENG(nRowIndex, sheetTableDef)
    FieldPostil = CHSName + "(" + RangeName + ")"
    
    FieldPostil = FieldPostil + Chr(10) + sheetTableDef.Cells(nRowIndex, 17)
    
    sheetCurrent.Cells(4, nObjIndex).ClearComments
    sheetCurrent.Cells(4, nObjIndex).AddComment FieldPostil
    sheetCurrent.Cells(4, nObjIndex).Comment.Shape.Width = 200
    sheetCurrent.Cells(4, nObjIndex).Comment.Shape.Height = 200
    sheetCurrent.Cells(4, nObjIndex) = CHSName
    
    nObjIndex = nObjIndex + 1
    
  Next nRowIndex
  
  Dim sheetCoverEng As Worksheet
  Dim sheetCoverChs As Worksheet
  
  Set sheetCoverEng = ThisWorkbook.Sheets("Cover")
  Set sheetCoverChs = ThisWorkbook.Sheets("封面")
  
  sheetCoverEng.Visible = True
  sheetCoverChs.Visible = False
  
  MsgBox "OK."

  ThisWorkbook.Protect
  
End Sub


'获取有效范围提示或者错误信息
Private Function GetRangeInfoCHS(nRowIndex As Integer, sheetTableDef As Worksheet) As String
  
  Const FieldNameCol = 2
  Const DataTypeCol = 3
  Const MinValCol = 5
  Const MaxValCol = 6
  Const RangeListCol = 7
  Const ValueTypeCol = 24
  Const FINT = "INT", FSTRING = "STRING", FLIST = "LIST"

  Dim sFieldName As String
  Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sValueType As String
  
  sFieldName = Trim(sheetTableDef.Cells(nRowIndex, 3).value)
  sDataType = Trim(sheetTableDef.Cells(nRowIndex, 4).value)
  sMinVal = Trim(sheetTableDef.Cells(nRowIndex, 6).value)
  sMaxVal = Trim(sheetTableDef.Cells(nRowIndex, 7).value)
  sRangeList = Trim(sheetTableDef.Cells(nRowIndex, 8).value)
  
  GetRangeInfoCHS = ""
  If (sDataType = FSTRING) Then
    If sMinVal = sMaxVal Then
      GetRangeInfoCHS = "长度范围" + "[" + sMinVal + "]"
    Else
      GetRangeInfoCHS = "长度范围" + "[" + sMinVal + ".." + sMaxVal + "]"
    End If
  End If
  
  If (sDataType = FINT) Then
    If sMinVal = sMaxVal Then
      GetRangeInfoCHS = "取值范围" + "[" + sMinVal + "]"
    Else
      GetRangeInfoCHS = "取值范围" + "[" + sMinVal + ".." + sMaxVal + "]"
    End If
  End If
  
  If sDataType = FLIST Then
      GetRangeInfoCHS = "取值范围" + "[" + sRangeList + "]"
  End If
  
  If Trim(sValueType) = "ATM" Then
    GetRangeInfoCHS = GetRangeInfoCHS + vbCrLf + " 注意: 需要加前缀 H'. "
  End If
  
  If Trim(sFieldName) = "LAC" Then
    GetRangeInfoCHS = "取值范围" + "[1..65533,65535]"
  End If

End Function


'获取有效范围提示或者错误信息
Private Function GetRangeInfoENG(nRowIndex As Integer, sheetTableDef As Worksheet) As String
  Const FieldNameCol = 2
  Const DataTypeCol = 3
  Const MinValCol = 5
  Const MaxValCol = 6
  Const RangeListCol = 7
  Const ValueTypeCol = 24
  Const FINT = "INT", FSTRING = "STRING", FLIST = "LIST"

  Dim sFieldName As String
  Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sValueType As String
  
  sFieldName = Trim(sheetTableDef.Cells(nRowIndex, 3).value)
  sDataType = Trim(sheetTableDef.Cells(nRowIndex, 4).value)
  sMinVal = Trim(sheetTableDef.Cells(nRowIndex, 6).value)
  sMaxVal = Trim(sheetTableDef.Cells(nRowIndex, 7).value)
  sRangeList = Trim(sheetTableDef.Cells(nRowIndex, 8).value)
  
  GetRangeInfoENG = ""
  If (sDataType = FSTRING) Then
    If sMinVal = sMaxVal Then
      GetRangeInfoENG = "Length" + "[" + sMinVal + "]"
    Else
      GetRangeInfoENG = "Length" + "[" + sMinVal + ".." + sMaxVal + "]"
    End If
  End If
  
  If (sDataType = FINT) Then
    If sMinVal = sMaxVal Then
      GetRangeInfoENG = "Range" + "[" + sMinVal + "]"
    Else
      GetRangeInfoENG = "Range" + "[" + sMinVal + ".." + sMaxVal + "]"
    End If
  End If
  
  If sDataType = FLIST Then
      GetRangeInfoENG = "Range" + "[" + sRangeList + "]"
  End If
  If Trim(sValueType) = "ATM" Then
    GetRangeInfoENG = GetRangeInfoENG + vbCrLf + " Note: Must begin with H'."
  End If
  If Trim(sFieldName) = "LAC" Then
    GetRangeInfoENG = "Range" + "[1..65533,65535]"
  End If

End Function
Public Sub SetBarLanguage(sLanguageType As String)
    '中英文按钮栏
    Dim sCHSBarName As String
    Dim sENGBarName As String
    Dim iLanguage  As Integer
    Dim sheetTableDef As Worksheet
    Set sheetTableDef = ThisWorkbook.Sheets("TableDef")
    If sLanguageType = "ENG" Then
        iLanguage = 0
        sheetTableDef.Cells(5, 8).value = iLanguage
       End If
    If sLanguageType = "CHS" Then
        iLanguage = 1
        sheetTableDef.Cells(5, 8).value = iLanguage
       End If
       
    If iLanguage = 0 Then
       sENGBarName = sheetTableDef.Cells(5, 9).value
       sCHSBarName = sheetTableDef.Cells(6, 9).value
      End If
         
    If iLanguage = 1 Then
       sENGBarName = sheetTableDef.Cells(5, 10).value
       sCHSBarName = sheetTableDef.Cells(6, 10).value
      End If
      
    If iLanguage = 1 Or iLanguage = 0 Then
        CommandBars("Operate Bar").Controls.Item(1).Caption = sENGBarName
        CommandBars("Operate Bar").Controls.Item(1).TooltipText = sENGBarName
        CommandBars("Operate Bar").Controls.Item(2).Caption = sCHSBarName
        CommandBars("Operate Bar").Controls.Item(2).TooltipText = sCHSBarName
      End If
 Call ReHOME(sLanguageType)
End Sub
Private Sub ReHOME(sType As String)
  Dim sheetHOME As Worksheet
  Dim sheetTableList As Worksheet
  Dim sheetCurrent As Worksheet
  Dim sMOName As String
  Dim sTableName As String
  Dim nRowIndex As Integer
  Dim nRow As Integer
  Dim nObjIndex As Integer
  
  If (sType <> "CHS") And (sType <> "ENG") Then
      Exit Sub
    End If
    
  If sType = "CHS" Then
    nRow = 4
   End If
   
  If sType = "ENG" Then
    nRow = 3
   End If
   
  Set sheetHOME = ThisWorkbook.Sheets("HOME")
  Set sheetTableList = ThisWorkbook.Sheets("TableList")
  
  sMOName = Trim(sheetTableList.Cells(1, nRow).value)
  sheetHOME.Cells(1, 2).value = sMOName
    
  For nRowIndex = 2 To 1000
  
    If Trim(sheetTableList.Cells(nRowIndex, 2).value) = "" Then
      Exit For
    End If
    
    For nObjIndex = 2 To 1000
  
    If Trim(sheetHOME.Cells(nObjIndex, 1).value) = "" Then
      Exit For
    End If
    
    sTableName = Trim(sheetTableList.Cells(nRowIndex, 2).value)
    
    If Trim(sheetHOME.Cells(nObjIndex, 1).value) = sTableName Then
      sMOName = Trim(sheetTableList.Cells(nRowIndex, nRow).value)
      sheetHOME.Cells(nObjIndex, 2).value = sMOName
    End If
    Next nObjIndex
  Next nRowIndex
  
  If sType = "CHS" Then
    sheetHOME.Cells(1, 1).value = "MO名称"
   End If
   
  If sType = "ENG" Then
    sheetHOME.Cells(1, 1).value = "MO Name"
   End If
End Sub
