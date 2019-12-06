Attribute VB_Name = "pubModule"
Option Explicit

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
  CurSheet.Cells(intRow, i).Value = Trim(CurSheet.Cells(intRow, i).Value)
  If (IsEmpty(CurSheet.Cells(intRow, i)) Or (Trim(CurSheet.Cells(intRow, i)) = "")) Then
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
  iValue = StatusForm.ProgressBar1.Value + iStep * 100 / iBase
  If iValue > 100 Then
    iValue = 100
  End If
  StatusForm.ProgressBar1.Value = iValue
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
  CurSheet.Cells(iRow, 1).Value = "X"
  WriteErrLog = False
End Function
'****************************************************************
'记录警告信息
'****************************************************************
Private Function WriteWarning(CurSheet As Worksheet, ErrMsg As String, iRow As Long) As Boolean
  ErrLog Space(2) + "Warning: " + ErrMsg
  CurSheet.Cells(iRow, 1).Value = "?"
  WriteWarning = False
End Function

Private Sub ClearErrOrWarning(CurSheet As Worksheet, iRow As Long)
  CurSheet.Cells(iRow, 1).Value = ""
End Sub

'****************************************************************
'登记错误日志文件
'****************************************************************
Private Sub ErrLog(ErrMsg As String)
  ErrLogLines = ErrLogLines + 1
  Sheets("CheckResult").Cells(ErrLogLines, 2).Value = ErrMsg
End Sub

'****************************************************************
'获取上下频点和频段指示对应关系表
'****************************************************************
Private Sub GetBandAndFreq(CurSheet As Worksheet, startRow As Integer, startCol As Integer)
Dim i As Integer, j As Integer

For i = 0 To MaxBANDINDs - 1
  For j = 0 To MaxBANDINDCol - 1
    BandAndFreq(i, j) = CurSheet.Cells(startRow + i, startCol + j)
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
    Dim sTemplateBarName As String
    Dim sGTRXBarName As String
    Dim iLanguage  As Integer
    Dim sheetTableDef As Worksheet
    Set sheetTableDef = ThisWorkbook.Sheets("TableDef")
    iLanguage = sheetTableDef.Cells(5, 8).Value
    If iLanguage = 0 Then
       sENGBarName = sheetTableDef.Cells(5, 9).Value
       sCHSBarName = sheetTableDef.Cells(6, 9).Value
       sTemplateBarName = sheetTableDef.Cells(7, 9).Value
       sGTRXBarName = sheetTableDef.Cells(8, 9).Value
      End If
         
    If iLanguage = 1 Then
       sENGBarName = sheetTableDef.Cells(5, 10).Value
       sCHSBarName = sheetTableDef.Cells(6, 10).Value
       sTemplateBarName = sheetTableDef.Cells(7, 10).Value
       sGTRXBarName = sheetTableDef.Cells(8, 10).Value
      End If
    
    On Error Resume Next
    
    Set cmbNewBar = CommandBars.Add(Name:="Operate Bar")
    
'    'reset sheets
'    With cmbNewBar
'        Set ctlBtn = .Controls.Add
'        With ctlBtn
'            .Style = msoButtonIconAndCaption
'            .BeginGroup = True
'            .Caption = "&Reset Sheets"
'            .TooltipText = "Reset all sheets"
'            .OnAction = "GenNegotiatedFile"
'            .FaceId = 28
'        End With
'        .Protection = msoBarNoCustomize
'        .Position = msoBarTop
'        .Visible = True
'    End With
'
'    Export XML
'    With cmbNewBar
'        Set ctlBtn = .Controls.Add
'        With ctlBtn
'            .Style = msoButtonIconAndCaption
'            .BeginGroup = True
'            .Caption = "&Export XML..."
'            .TooltipText = "Export XML..."
'            .OnAction = "UsrExportXML"
'            .FaceId = 23
'        End With
'        .Protection = msoBarNoCustomize
'        .Position = msoBarTop
'        .Visible = True
'    End With
    
    'customized template
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = sTemplateBarName
            .TooltipText = sTemplateBarName
            .OnAction = "OpenCfgForm"
            .FaceId = 50
        End With
        .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
    
    'transfer
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = sGTRXBarName
            .TooltipText = sGTRXBarName
            .OnAction = "ConvertGTRX"
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
Op
Sub SetSysOption()
  With ActiveWindow
    If .DisplayZeros Then Exit Sub
    .DisplayGridlines = False
    .DisplayZeros = True
  End With
End Sub

Public Sub OpenCfgForm()
    Dim iLanguage As Integer
    iLanguage = ThisWorkbook.Sheets("TableDef").Cells(5, 8).Value
    If iLanguage = 0 Then
       TemplateCMForm.Caption = "Customize Template"
       TemplateCMForm.ToolFrame.Caption = "Summary"
    ElseIf iLanguage = 1 Then
       TemplateCMForm.Caption = "定制模板"
       TemplateCMForm.ToolFrame.Caption = "汇总"
    End If
    TemplateCMForm.Show vbModeless
End Sub

Public Sub ConvertGTRX()
  Call CvtTRXMAPtoGTRX
End Sub

Public Sub SwitchChs()
  ThisWorkbook.Unprotect
  
  Dim nRowIndex As Integer
  Dim sheetTableDef As Worksheet
  Dim sheetCurrent As Worksheet
  Dim strObjId As String
  Dim nObjId As Integer
  Dim strObjName As String
  Dim FieldCol As String
  Dim sDataType As String
  Dim sRangeList As String
  Dim sFormula1 As String
  Dim sFormula2 As String
  Dim xType As Excel.XlDVType
  Dim sFieldName As String
  Dim sTableName As String
  
  SetBarLanguage ("CHS")
      
  TemplateCMForm.Caption = "定制模板"
       
  Set sheetTableDef = ThisWorkbook.Sheets("TableDef")
  
  '设置批注信息
  Const FieldDefCol = 4
  Const FieldNameDisplayCol = 11
  Dim CHSName As String
  Dim RangeName As String
  Dim FieldPostil As String
  Dim nObjIndex As Integer
    
  For nRowIndex = 15 To 15000
  
    If Trim(sheetTableDef.Cells(nRowIndex, 3).Value) = "" Then
      Exit For
    End If
    
    strObjId = Trim(sheetTableDef.Cells(nRowIndex, 1).Value)
    If strObjId <> "" Then
      sTableName = Trim(sheetTableDef.Cells(nRowIndex, 2).Value)
      Set sheetCurrent = ThisWorkbook.Sheets(sTableName)
      nObjIndex = 2
    End If
        
    ' 设置批注信息
    CHSName = Trim(sheetTableDef.Cells(nRowIndex, 14))
    RangeName = GetRangeInfoCHS(nRowIndex, sheetTableDef)
    FieldPostil = CHSName + "(" + RangeName + ")"
    sFieldName = Trim(sheetTableDef.Cells(nRowIndex, 3))
    
        '数据有效性变量
    sDataType = Trim(sheetTableDef.Cells(nRowIndex, 4))
    sRangeList = Trim(sheetTableDef.Cells(nRowIndex, 8))
    
    FieldPostil = FieldPostil + Chr(13) + Chr(10) + sheetTableDef.Cells(nRowIndex, 19)
    
    sheetCurrent.Cells(4, nObjIndex).ClearComments
    sheetCurrent.Cells(4, nObjIndex).AddComment FieldPostil
    sheetCurrent.Cells(4, nObjIndex).Comment.Shape.Width = 200
    sheetCurrent.Cells(4, nObjIndex).Comment.Shape.Height = 200
    sheetCurrent.Cells(4, nObjIndex) = CHSName
    sheetCurrent.Cells(5, nObjIndex) = sFieldName
    sheetCurrent.Cells(1, nObjIndex) = sFieldName
    
    FieldCol = GetWidthName(nObjIndex)
      
    ' 设置表空框的属性值，包括数据有效性。
    '暂时删除下拉列表测试
    If sDataType = "LIST" Then
       sheetCurrent.Activate
       
       Range(FieldCol + "6" + ":" + FieldCol + "2000").Select
       
       xType = xlValidateList
       sFormula1 = sRangeList
       sFormula2 = ""
       sheetCurrent.Columns(FieldCol).NumberFormatLocal = "@"
       sheetCurrent.Columns(FieldCol).HorizontalAlignment = xlLeft
       
       
       ' 调用有效数据

       
       Call SetDataValidate(xType, sFormula1, sFormula2, "LIST Prompt", "Range")
     Else
       sheetCurrent.Activate
       Range(FieldCol + "6" + ":" + FieldCol + "2000").Select
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
    sheetCurrent.Cells(6, 2).Select
    nObjIndex = nObjIndex + 1
    
  Next nRowIndex
 
  Dim sheetFreqTool As Worksheet
  Set sheetFreqTool = ThisWorkbook.Sheets("Frequency Tool")
  sheetFreqTool.Rows(3).EntireRow.Hidden = False
  sheetFreqTool.Rows(4).EntireRow.Hidden = True

  
  Dim sheetCoverChs As Worksheet
  Dim sheetCoverEng As Worksheet
  
  Set sheetCoverEng = ThisWorkbook.Sheets("Cover")
  Set sheetCoverChs = ThisWorkbook.Sheets("封面")
  
  sheetCoverEng.Visible = False
  sheetCoverChs.Visible = True
  
  ThisWorkbook.Protect
  
  MsgBox "完成."
End Sub
Public Sub AutoSet()
  ThisWorkbook.Unprotect
  
Const CONST_BasicInfoROW = 6

   Dim iRow As Integer
   Dim kRow As Integer
   Dim TRXId As String
   Dim ChannelHopIndex As String
   Dim ChannelMaio As String
   Dim ChilSht As Worksheet
   Set ChilSht = Sheets("GTRXCHANHOP")
   iRow = CONST_BasicInfoROW
   
   Do Until (ChilSht.Cells(iRow, 2) = "")
        TRXId = ChilSht.Cells(iRow, 2).Text
        ChannelHopIndex = ChilSht.Cells(iRow, 4).Text
        ChannelMaio = ChilSht.Cells(iRow, 5).Text
        kRow = iRow + 1
      Do Until (ChilSht.Cells(kRow, 2) = "")
        If ChilSht.Cells(kRow, 2).Text = TRXId Then
          ChilSht.Cells(kRow, 4) = ChannelHopIndex
          ChilSht.Cells(kRow, 5) = ChannelMaio
        End If
    
        kRow = kRow + 1

      Loop

        iRow = iRow + 1
   Loop
End Sub
Public Sub Highlight()
  ThisWorkbook.Unprotect
  
Const CONST_BasicInfoROW = 6

   Dim iRow As Integer
   Dim kRow As Integer
   Dim Flag As Integer
   Dim TRXId As String
   Dim ChilSht As Worksheet
   Set ChilSht = Sheets("GTRXCHANHOP")
   iRow = CONST_BasicInfoROW
   
   Do Until (ChilSht.Cells(iRow, 2) = "")
        TRXId = ChilSht.Cells(iRow, 2).Text
        kRow = iRow - 1
        Flag = 0
     If kRow > 5 Then
        
       Do Until (kRow = 5)
         If ChilSht.Cells(kRow, 2).Text = TRXId Then
           Flag = 1
         End If
 
         kRow = kRow - 1

       Loop
     End If

     If Flag <> 1 Then
       With ChilSht.Cells(iRow, 2).Interior
          .ColorIndex = 6
          .Pattern = xlSolid
       End With
       With ChilSht.Cells(iRow, 3).Interior
          .ColorIndex = 6
          .Pattern = xlSolid
       End With
       With ChilSht.Cells(iRow, 4).Interior
          .ColorIndex = 6
          .Pattern = xlSolid
       End With
       With ChilSht.Cells(iRow, 5).Interior
          .ColorIndex = 6
          .Pattern = xlSolid
       End With
     End If

       iRow = iRow + 1
   Loop

End Sub



Public Sub SwitchEng()
  
  ThisWorkbook.Unprotect
  
  Dim nRowIndex As Integer
  Dim sheetTableDef As Worksheet
  Dim sheetCurrent As Worksheet
  Dim strObjId As String
  Dim nObjId As Integer
  Dim strObjName As String
  Dim FieldCol As String
  Dim sDataType As String
  Dim sRangeList As String
  Dim sFormula1 As String
  Dim sFormula2 As String
  Dim xType As Excel.XlDVType
  Dim sFieldName As String
  Dim sTableName As String
  
  SetBarLanguage ("ENG")
      
  TemplateCMForm.Caption = "Customize Template"
   
  Set sheetTableDef = ThisWorkbook.Sheets("TableDef")
  
  '设置批注信息
  Const FieldDefCol = 4
  Const FieldNameDisplayCol = 11
  Dim CHSName As String
  Dim RangeName As String
  Dim FieldPostil As String
  Dim nObjIndex As Integer
    
  For nRowIndex = 15 To 15000
  
    If Trim(sheetTableDef.Cells(nRowIndex, 3).Value) = "" Then
      Exit For
    End If
    
    strObjId = Trim(sheetTableDef.Cells(nRowIndex, 1).Value)
    If strObjId <> "" Then
      sTableName = Trim(sheetTableDef.Cells(nRowIndex, 2).Value)
      Set sheetCurrent = ThisWorkbook.Sheets(sTableName)
      nObjIndex = 2
    End If
        
    ' 设置批注信息
    CHSName = Trim(sheetTableDef.Cells(nRowIndex, 13))
    RangeName = GetRangeInfoENG(nRowIndex, sheetTableDef)
    FieldPostil = CHSName + "(" + RangeName + ")"
    sFieldName = Trim(sheetTableDef.Cells(nRowIndex, 3))
    
    '数据有效性变量
    sDataType = Trim(sheetTableDef.Cells(nRowIndex, 4))
    sRangeList = Trim(sheetTableDef.Cells(nRowIndex, 8))
    
    FieldPostil = FieldPostil + Chr(10) + sheetTableDef.Cells(nRowIndex, 17)
    
    sheetCurrent.Cells(4, nObjIndex).ClearComments
    sheetCurrent.Cells(4, nObjIndex).AddComment FieldPostil
    sheetCurrent.Cells(4, nObjIndex).Comment.Shape.Width = 200
    sheetCurrent.Cells(4, nObjIndex).Comment.Shape.Height = 200
    sheetCurrent.Cells(4, nObjIndex) = CHSName
    sheetCurrent.Cells(5, nObjIndex) = sFieldName
    sheetCurrent.Cells(1, nObjIndex) = sFieldName
    
    FieldCol = GetWidthName(nObjIndex)
      
    ' 设置表空框的属性值，包括数据有效性。
    '暂时删除下拉列表测试
    If sDataType = "LIST" Then
       sheetCurrent.Activate
       
       Range(FieldCol + "5" + ":" + FieldCol + "2000").Select
       
       xType = xlValidateList
       sFormula1 = sRangeList
       sFormula2 = ""
       sheetCurrent.Columns(FieldCol).NumberFormatLocal = "@"
       sheetCurrent.Columns(FieldCol).HorizontalAlignment = xlLeft
       
       
       ' 调用有效数据

       
       Call SetDataValidate(xType, sFormula1, sFormula2, "LIST Prompt", "Range")
     Else
       sheetCurrent.Activate
       Range(FieldCol + "6" + ":" + FieldCol + "2000").Select
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
    sheetCurrent.Cells(6, 2).Select
    nObjIndex = nObjIndex + 1
    
  Next nRowIndex
  
  Dim sheetFreqTool As Worksheet
  Set sheetFreqTool = ThisWorkbook.Sheets("Frequency Tool")
  sheetFreqTool.Rows(3).EntireRow.Hidden = True
  sheetFreqTool.Rows(4).EntireRow.Hidden = False
  
  Dim sheetCoverChs As Worksheet
  Dim sheetCoverEng As Worksheet
  
  Set sheetCoverEng = ThisWorkbook.Sheets("Cover")
  Set sheetCoverChs = ThisWorkbook.Sheets("封面")
  
  sheetCoverEng.Visible = True
  sheetCoverChs.Visible = False
  
  ThisWorkbook.Protect
  MsgBox "OK."

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
  
  sFieldName = Trim(sheetTableDef.Cells(nRowIndex, 3).Value)
  sDataType = Trim(sheetTableDef.Cells(nRowIndex, 4).Value)
  sMinVal = Trim(sheetTableDef.Cells(nRowIndex, 6).Value)
  sMaxVal = Trim(sheetTableDef.Cells(nRowIndex, 7).Value)
  sRangeList = Trim(sheetTableDef.Cells(nRowIndex, 8).Value)
  
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
  
  sFieldName = Trim(sheetTableDef.Cells(nRowIndex, 3).Value)
  sDataType = Trim(sheetTableDef.Cells(nRowIndex, 4).Value)
  sMinVal = Trim(sheetTableDef.Cells(nRowIndex, 6).Value)
  sMaxVal = Trim(sheetTableDef.Cells(nRowIndex, 7).Value)
  sRangeList = Trim(sheetTableDef.Cells(nRowIndex, 8).Value)
  
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
Public Sub SetDataValidate(xType As Excel.XlDVType, sFormula1 As String, sFormula2 As String, sErrPrompt As String, sErrMsg As String)
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
Public Function GetWidthName(iRow As Integer) As String
Dim iTemRow As Integer
Dim sWidthName(1 To 26) As String
sWidthName(1) = "A"
sWidthName(2) = "B"
sWidthName(3) = "C"
sWidthName(4) = "D"
sWidthName(5) = "E"
sWidthName(6) = "F"
sWidthName(7) = "G"
sWidthName(8) = "H"
sWidthName(9) = "I"
sWidthName(10) = "J"
sWidthName(11) = "K"
sWidthName(12) = "L"
sWidthName(13) = "M"
sWidthName(14) = "N"
sWidthName(15) = "O"
sWidthName(16) = "P"
sWidthName(17) = "Q"
sWidthName(18) = "R"
sWidthName(19) = "S"
sWidthName(20) = "T"
sWidthName(21) = "U"
sWidthName(22) = "V"
sWidthName(23) = "W"
sWidthName(24) = "X"
sWidthName(25) = "Y"
sWidthName(26) = "Z"


 iTemRow = iRow
  If iTemRow < 27 Then
   GetWidthName = Trim(sWidthName(iTemRow))
  End If
 If iTemRow > 26 And iTemRow < 53 Then
   iTemRow = iTemRow - 26
   GetWidthName = "A" + Trim(sWidthName(iTemRow))
  End If
 If iTemRow > 52 And iTemRow < 79 Then
   iTemRow = iTemRow - 52
   GetWidthName = "B" + Trim(sWidthName(iTemRow))
  End If

 If iTemRow > 78 And iTemRow < 105 Then
   iTemRow = iTemRow - 78
   GetWidthName = "C" + Trim(sWidthName(iTemRow))
  End If
  
 If iTemRow > 104 And iTemRow < 131 Then
   iTemRow = iTemRow - 104
   GetWidthName = "D" + Trim(sWidthName(iTemRow))
  End If
  
 If iTemRow > 130 And iTemRow < 157 Then
   iTemRow = iTemRow - 130
   GetWidthName = "E" + Trim(sWidthName(iTemRow))
  End If
  
 If iTemRow > 156 And iTemRow < 183 Then
   iTemRow = iTemRow - 156
   GetWidthName = "F" + Trim(sWidthName(iTemRow))
  End If
End Function
Public Sub SetBarLanguage(sLanguageType As String)
    '中英文按钮栏
    Dim sCHSBarName As String
    Dim sENGBarName As String
    Dim sTemplateBarName As String
    Dim sGTRXBarName As String
    Dim iLanguage  As Integer
    Dim sheetTableDef As Worksheet
    Set sheetTableDef = ThisWorkbook.Sheets("TableDef")
    If sLanguageType = "ENG" Then
        iLanguage = 0
        sheetTableDef.Cells(5, 8).Value = iLanguage
       End If
    If sLanguageType = "CHS" Then
        iLanguage = 1
        sheetTableDef.Cells(5, 8).Value = iLanguage
       End If
       
    If iLanguage = 0 Then
       sENGBarName = sheetTableDef.Cells(5, 9).Value
       sCHSBarName = sheetTableDef.Cells(6, 9).Value
       sTemplateBarName = sheetTableDef.Cells(7, 9).Value
       sGTRXBarName = sheetTableDef.Cells(8, 9).Value
      End If
         
    If iLanguage = 1 Then
       sENGBarName = sheetTableDef.Cells(5, 10).Value
       sCHSBarName = sheetTableDef.Cells(6, 10).Value
       sTemplateBarName = sheetTableDef.Cells(7, 10).Value
       sGTRXBarName = sheetTableDef.Cells(8, 10).Value
      End If
      
    If iLanguage = 1 Or iLanguage = 0 Then
        CommandBars("Operate Bar").Controls.Item(1).Caption = sTemplateBarName
        CommandBars("Operate Bar").Controls.Item(1).TooltipText = sTemplateBarName
        CommandBars("Operate Bar").Controls.Item(2).Caption = sGTRXBarName
        CommandBars("Operate Bar").Controls.Item(2).TooltipText = sGTRXBarName
        CommandBars("Operate Bar").Controls.Item(3).Caption = sENGBarName
        CommandBars("Operate Bar").Controls.Item(3).TooltipText = sENGBarName
        CommandBars("Operate Bar").Controls.Item(4).Caption = sCHSBarName
        CommandBars("Operate Bar").Controls.Item(4).TooltipText = sCHSBarName
      End If
  Call ReHOME(sLanguageType)
End Sub
Public Sub ReHOME(sType As String)
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
  
  sMOName = Trim(sheetTableList.Cells(1, nRow).Value)
  sheetHOME.Cells(1, 2).Value = sMOName
  
  If sType = "CHS" Then
    sheetHOME.Cells(1, 1).Value = "MO 名称"
   End If
   
   If sType = "ENG" Then
    sheetHOME.Cells(1, 1).Value = "MO Name"
   End If
    
  For nRowIndex = 2 To 1000
  
    If Trim(sheetTableList.Cells(nRowIndex, 2).Value) = "" Then
      Exit For
    End If
    
    For nObjIndex = 2 To 1000
  
    If Trim(sheetHOME.Cells(nObjIndex, 1).Value) = "" Then
      Exit For
    End If
    
    sTableName = Trim(sheetTableList.Cells(nRowIndex, 2).Value)
    
    If Trim(sheetHOME.Cells(nObjIndex, 1).Value) = sTableName Then
      sMOName = Trim(sheetTableList.Cells(nRowIndex, nRow).Value)
      sheetHOME.Cells(nObjIndex, 2).Value = sMOName
    End If
    Next nObjIndex
  Next nRowIndex
End Sub
