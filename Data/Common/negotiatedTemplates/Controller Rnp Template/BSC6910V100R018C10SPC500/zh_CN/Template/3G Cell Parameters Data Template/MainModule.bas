Attribute VB_Name = "MainModule"

Option Explicit
Public ParamMocList As New Collection
'版本信息
Public sCMEVersion As String
Public sNEVersion As String
Public sRNPVersion As String
Public RNC_VERSION As String

Public Const RSC_STR_FINISHED = "Finished."
Public Const INTTYPE = 0
Public Const BIGINTTYPE = 11
Public Const BITMAPTYPE = 5
Public Const ENUMTYPE = 101
Public Const STRINGTYPE = 2
Public Const FieldRow = 40
Public Const FieldColmn = 5



Type struRangeDef
   sMaxVal As String        '参数信息
   sMinVal As String        '显示名称
End Type


Sub DeleteUserToolBar()
  On Error Resume Next
  CommandBars("Operate Bar").Delete
End Sub

'****************************************************************
'从TableDef表中读取所有Sheet的定义参数数据
'****************************************************************
Public Sub GetSheetDefineData()
    Dim CurSheet As Worksheet
    Dim iRow As Integer, iCol As Integer
    
    Set CurSheet = ThisWorkbook.Sheets("Refresh")
    
    sCMEVersion = CurSheet.Range("B2").Text
    sNEVersion = CurSheet.Range("B3").Text
    sRNPVersion = CurSheet.Range("B4").Text
    RNC_VERSION = CurSheet.Range("B1").Text
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
            .name = "Arial"
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

        '设置列宽（具体行列根据定义设置）
        Columns("A:A").Select
       ' Selection.ColumnWidth = GetSheetFisrtColWidth
         Selection.ColumnWidth = 20
    End With
    '设置零值显示、网格不显示
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayZeros = True
End Function

Public Sub SetCoverSheet()
    Dim CurSheet As Worksheet
    Set CurSheet = ThisWorkbook.Sheets("Cover")
    
   ' CurSheet.Unprotect Password:=GetSheetsPass

    
    Call SetSheetDefaultValue(CurSheet)
    Range("A:A").Select
    Selection.ColumnWidth = 8
    
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
        .name = "Arial"
        .FontStyle = "加粗"
        .Size = 24
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
    Range("C8:D8") = "3G Radio Algorithm Parameters Template"
    
    Range("C15:D17").Select
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
    
    Range("C15:C17").Select
    Selection.ColumnWidth = 18
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Selection.Font
        .name = "Arial"
        .FontStyle = "加粗"
        .Size = 14
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 33
        .Pattern = xlSolid
    End With
    Range("C14") = "NE  Version"
    
    Range("D15:D17").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    With Selection.Font
        .name = "宋体"
        .FontStyle = "常规"
        .Size = 14
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 34
        .Pattern = xlSolid
    End With

    Range("D14") = sNEVersion

    Range("A3").Select
    
    'CurSheet.Protect Password:=GetSheetsPass, DrawingObjects:=True, Contents:=True, Scenarios:=True

End Sub

Private Sub ClearParamSheets()
  Dim CurSheet As Worksheet, HomeSheet As Worksheet
  Dim i As Integer
  Dim sCellVal As String
    
  UnprotectWorkBook
   
  Set HomeSheet = ActiveWorkbook.Sheets("Home")
  i = 2
  sCellVal = HomeSheet.Cells(i, 1).Value
  Do While (Trim(sCellVal) <> "")
    HomeSheet.Cells(i, 1).Clear
    
    i = i + 1
    sCellVal = HomeSheet.Cells(i, 1).Value
  Loop
  
  Application.DisplayAlerts = False
  For i = ActiveWorkbook.Sheets.Count To 1 Step -1
    Set CurSheet = ActiveWorkbook.Sheets(i)
    If StrComp(CurSheet.name, "Refresh", vbTextCompare) <> 0 And StrComp(CurSheet.name, "TableInfo", vbTextCompare) <> 0 And StrComp(CurSheet.name, "Cover", vbTextCompare) <> 0 _
      And StrComp(CurSheet.name, "CMETemplateInfo", vbTextCompare) <> 0 _
      And StrComp(CurSheet.name, "ValidInfo", vbTextCompare) <> 0 _
      And StrComp(CurSheet.name, "Home", vbTextCompare) _
      And StrComp(CurSheet.name, "CHS", vbTextCompare) _
      And StrComp(CurSheet.name, "ENG", vbTextCompare) Then
      CurSheet.Activate
      CurSheet.Unprotect Password:=GetSheetsPass
      CurSheet.Cells.Select
      Selection.Clear
      Range("A1").Select
    End If
  Next
  Application.DisplayAlerts = True
    
End Sub
Private Sub GenerateMOSheet(ByRef MocInst As MocInfo)
  Call DeleteParamSheets
  sOldMocName = "Home"
  For iIndex = 1 To ParamMocList.Count
    Set MocInst = ParamMocList.Item(iIndex)
    Call AddParamSheets(MocInst, iIndex, sOldMocName)
    sOldMocName = MocInst.MOCName
  Next

End Sub
Private Sub SetFieldProperty(ByRef MocInst As MocInfo, ByVal sFiledName As String, ByVal sDspName As String, ByVal iFieldType As Integer, ByVal iMinVal As Double, ByVal iMaxVal As Double)
  Dim ParamInst As ParamInfo
  Dim RangeInst As RangeInfo

  Set ParamInst = New ParamInfo

  ParamInst.FieldName = sFiledName
  ParamInst.DisplayName = sDspName
  ParamInst.FieldType = iFieldType
  
  If (iFieldType = INTTYPE Or iFieldType = BIGINTTYPE) Then
    ParamInst.FieldTypeName = "INT"
  ElseIf (iFieldType = STRINGTYPE) Then
    ParamInst.FieldTypeName = "STRING"
  End If
  
  ParamInst.FieldMinValue = iMinVal
  ParamInst.FieldMaxValue = iMaxVal
  
  Set RangeInst = New RangeInfo
  RangeInst.sMinVal = CStr(iMinVal)
  RangeInst.sMaxVal = CStr(iMaxVal)
  ParamInst.RangeDef.Add RangeInst
  
  MocInst.Params.Add ParamInst
End Sub

Private Sub ProcessCELLMO(ByRef MocInst As MocInfo, ByRef EntyNum As Integer)
Dim sFiledName As String

  If EntyNum = 0 Then
   EntyNum = 1
  ElseIf EntyNum <> 0 Then
    Exit Sub
  End If
  
  Call AddPrimKey(MocInst, "CELL")

  sFiledName = "LOGICRNCID"
  Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "NODEBNAME"
  Call GetCELLFieldProperty(MocInst, sFiledName)

  sFiledName = "CELLID"
   Call GetCELLFieldProperty(MocInst, sFiledName)

  sFiledName = "CELLNAME"
   Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "CNOPGRPINDEX"
  Call GetCELLFieldProperty(MocInst, sFiledName)

  sFiledName = "BANDIND"
   Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "UARFCNUPLINK"
   Call GetCELLFieldProperty(MocInst, sFiledName)
  
   sFiledName = "UARFCNDOWNLINK"
   Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "PSCRAMBCODE"
   Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "LAC"
   Call GetCELLFieldProperty(MocInst, sFiledName)
 
  sFiledName = "SAC"
   Call GetCELLFieldProperty(MocInst, sFiledName)
 
  sFiledName = "RAC"
  Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "SPGID"
   Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "URAIDS"
  Call SetFieldProperty(MocInst, sFiledName, "URA ID", INTTYPE, 0, 65535)
  
  sFiledName = "CIO"
   Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "TCELL"
   Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "LOCELL"
  Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "MAXTXPOWER"
   Call GetCELLFieldProperty(MocInst, sFiledName)
  
  sFiledName = "PCPICHPOWER"
  Call SetFieldProperty(MocInst, sFiledName, "PCPICH transmit power[0.1dBm]", INTTYPE, -100, 500)
  
  
 sFiledName = "SRN"
 Call GetCELLFieldProperty(MocInst, sFiledName)
  
 sFiledName = "SN"
 Call GetCELLFieldProperty(MocInst, sFiledName)
  
 sFiledName = "SSN"
 Call GetCELLFieldProperty(MocInst, sFiledName)
 
 
  sFiledName = "TEMPLATENAME"
  Call SetFieldProperty(MocInst, sFiledName, "Template Name", STRINGTYPE, 0, 255)
              
End Sub

Private Sub AddPrimKey(ByRef aMoc As MocInfo, ByVal MOCName As String)
  Dim ParamInst As ParamInfo
  Set ParamInst = New ParamInfo
  ParamInst.FieldName = "BSCName"
  ParamInst.DisplayName = "BSC Name"
  ParamInst.CHSDisplayName = "BSC名称"
  ParamInst.FieldISMustGive = "YES"
  ParamInst.FieldMaxValue = 255
  ParamInst.FieldMinValue = 1
  ParamInst.FieldTypeName = "STRING"
  ParamInst.FieldType = 2
  aMoc.Params.Add ParamInst
  
  If MOCName <> "CELL" Then
    Set ParamInst = New ParamInfo
    ParamInst.FieldName = "CELLNAME"
    ParamInst.DisplayName = "Cell Name"
    ParamInst.CHSDisplayName = " 小区名称"
    ParamInst.FieldISMustGive = "YES"
    ParamInst.FieldMaxValue = 64
    ParamInst.FieldMinValue = 1
    ParamInst.FieldTypeName = "STRING"
    ParamInst.FieldType = 2
    aMoc.Params.Add ParamInst
  End If
End Sub
Private Sub GenerateMOInfo(ByRef sheetName As String, ByRef MOInfo As MocInfo, ByRef Counter As Integer)
  Dim MOName, FieldName, FieldType, FieldLocation, EnumList, FieldDspName, CHSFieldName, BitmapField, StrRange As String
  Dim MinValue, MaxValue, RowNo, ColumnNo, i, j As Integer
  Dim aParam As ParamInfo
  Dim EnumInst As EnumInfo
  Dim RangeInst As RangeInfo
  Dim CurSheet As Worksheet
  
  Set CurSheet = ThisWorkbook.Sheets(sheetName)
  CurSheet.Activate
  RowNo = 5
  ColumnNo = 1
  
  '填充对象的信息，列名如下：
  'MONAME, Field Name,  Field Type,  Min Value , Max Value, List Value , Field Display Name(ENG), Field Display Name(CHS),  IsMustGive
  
  For i = 1 To MOInfo.Params.Count
    Set aParam = MOInfo.Params.Item(i)
    
    CurSheet.Cells(Counter, ColumnNo) = MOInfo.MOCName
    ColumnNo = ColumnNo + 1
    CurSheet.Cells(Counter, ColumnNo) = aParam.FieldName
    ColumnNo = ColumnNo + 1
    CurSheet.Cells(Counter, ColumnNo) = aParam.FieldTypeName
    ColumnNo = ColumnNo + 1
    CurSheet.Cells(Counter, ColumnNo) = aParam.FieldMinValue
    ColumnNo = ColumnNo + 1
    CurSheet.Cells(Counter, ColumnNo) = aParam.FieldMaxValue
    
    ColumnNo = ColumnNo + 1
    If aParam.FieldType = ENUMTYPE Then
        StrRange = ""
        For j = 1 To aParam.EnumRangeDef.Count
          Set EnumInst = aParam.EnumRangeDef.Item(j)
          StrRange = StrRange & EnumInst.sEnumStr + ","
        Next
        StrRange = Left(StrRange, Len(StrRange) - Len(","))
        CurSheet.Cells(Counter, ColumnNo) = StrRange
    ElseIf (aParam.FieldType = INTTYPE) Or (aParam.FieldType = BIGINTTYPE) Then
        StrRange = ""
        For j = 1 To aParam.RangeDef.Count
            Set RangeInst = aParam.RangeDef.Item(j)
            If (RangeInst.sMaxVal = RangeInst.sMinVal) Then
                StrRange = StrRange + RangeInst.sMaxVal + ","
            Else
                StrRange = StrRange + RangeInst.sMinVal + ".." + RangeInst.sMaxVal + ","
            End If
        Next
        StrRange = Left(StrRange, Len(StrRange) - Len(","))
        CurSheet.Cells(Counter, ColumnNo) = StrRange
    End If
    
    ColumnNo = ColumnNo + 1
    CurSheet.Cells(Counter, ColumnNo) = aParam.DisplayName
    ColumnNo = ColumnNo + 1
    CurSheet.Cells(Counter, ColumnNo) = aParam.CHSDisplayName
    ColumnNo = ColumnNo + 1
    CurSheet.Cells(Counter, ColumnNo) = aParam.FieldISMustGive
    ColumnNo = 1
    Counter = Counter + 1
  Next

End Sub
Private Sub DeleteParamSheets()
  Dim CurSheet As Worksheet, HomeSheet As Worksheet
  Dim i As Integer
  Dim sCellVal As String
    
  UnprotectWorkBook
   
  Set HomeSheet = ActiveWorkbook.Sheets("Home")
  i = 2
  sCellVal = HomeSheet.Cells(i, 1).Value
  Do While (Trim(sCellVal) <> "")
    HomeSheet.Cells(i, 1).Clear
    
    i = i + 1
    sCellVal = HomeSheet.Cells(i, 1).Value
  Loop
  
  Application.DisplayAlerts = False
  For i = ActiveWorkbook.Sheets.Count To 1 Step -1
    Set CurSheet = ActiveWorkbook.Sheets(i)
    If StrComp(CurSheet.name, "Refresh", vbTextCompare) <> 0 And StrComp(CurSheet.name, "TableInfo", vbTextCompare) <> 0 And StrComp(CurSheet.name, "Cover", vbTextCompare) <> 0 And StrComp(CurSheet.name, "ValidInfo", vbTextCompare) <> 0 _
      And StrComp(CurSheet.name, "CMETemplateInfo", vbTextCompare) _
      And StrComp(CurSheet.name, "Home", vbTextCompare) _
      And StrComp(CurSheet.name, "CHS", vbTextCompare) _
      And StrComp(CurSheet.name, "ENG", vbTextCompare) Then
      CurSheet.Activate
      ActiveSheet.Delete
    End If
  Next
  Application.DisplayAlerts = True
    
End Sub

Private Sub AddParamSheets(ByRef aMoc As MocInfo, ByVal Index As Integer, ByVal PreSheetName As String)
  Dim CurSheet As Worksheet, HomeSheet As Worksheet
  Dim i As Integer, j As Integer
  Dim StrRange As String, hyperStr As String
  Dim RangeInst As RangeInfo
  Dim EnumInst As EnumInfo
  Dim AA As String
  Dim SheetIsNotExist As Integer

  
  
SheetIsNotExist = 1
For i = ActiveWorkbook.Sheets.Count To 1 Step -1
   If aMoc.MOCName = ActiveWorkbook.Sheets(i).name Then
      SheetIsNotExist = 0
   End If
Next

If SheetIsNotExist Then
  ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(PreSheetName)
  Set CurSheet = ActiveSheet
  CurSheet.Activate
  CurSheet.name = aMoc.MOCName
Else
  Set CurSheet = ActiveWorkbook.Sheets(aMoc.MOCName)
  CurSheet.Activate
  CurSheet.name = aMoc.MOCName
End If

  
  Dim aParam As ParamInfo
  
  For i = 1 To aMoc.Params.Count
    Set aParam = aMoc.Params.Item(i)
    
    CurSheet.Cells(1, i) = aParam.FieldName
    CurSheet.Cells(2, i) = aParam.DisplayName
    
    If aParam.FieldType = ENUMTYPE Then
      Call FormatColumn(CurSheet, i, aParam)
    End If
    
    StrRange = ""

    If aParam.FieldType = ENUMTYPE Then
      StrRange = aParam.DisplayName + "(Range["
      For j = 1 To aParam.EnumRangeDef.Count
        Set EnumInst = aParam.EnumRangeDef.Item(j)
        StrRange = StrRange & EnumInst.sEnumStr + ","
      Next
      StrRange = Left(StrRange, Len(StrRange) - Len(","))
      StrRange = StrRange + "])"
    ElseIf aParam.FieldType = BITMAPTYPE Then
      If aParam.RangeDef.Count > 0 Then
        StrRange = aParam.DisplayName + "(Length[%iMaxVal%])"
        Set RangeInst = aParam.RangeDef.Item(1)
        StrRange = Replace(StrRange, "%iMaxVal%", RangeInst.sMaxVal)
      End If
    ElseIf aParam.FieldType = INTTYPE Or aParam.FieldType = BIGINTTYPE Then
      If aParam.RangeDef.Count > 0 Then
        StrRange = aParam.DisplayName + "(Range["
        For j = 1 To aParam.RangeDef.Count
          Set RangeInst = aParam.RangeDef.Item(j)
          If (RangeInst.sMinVal = RangeInst.sMaxVal) Then
            StrRange = StrRange + RangeInst.sMaxVal + ","
          Else
            StrRange = StrRange + RangeInst.sMinVal + ".." + RangeInst.sMaxVal + ","
          End If
        Next
        StrRange = Left(StrRange, Len(StrRange) - Len(","))
        StrRange = StrRange + "])"
      
'        StrRange = aParam.DisplayName + "(Range[%iMinVal%..%iMaxVal%])"
'        Set RangeInst = aParam.RangeDef.Item(1)
'        StrRange = Replace(StrRange, "%iMinVal%", RangeInst.sMinVal)
'        StrRange = Replace(StrRange, "%iMaxVal%", RangeInst.sMaxVal)
      End If
    Else
        StrRange = aParam.DisplayName + "(Length[%iMinVal%..%iMaxVal%])"
        StrRange = Replace(StrRange, "%iMinVal%", aParam.FieldMinValue)
        StrRange = Replace(StrRange, "%iMaxVal%", aParam.FieldMaxValue)
    End If
    
    If Len(StrRange) > 0 Then
      If aParam.FieldName = "LAC" Then
        StrRange = aParam.DisplayName + "(Range[1..65533,65535])"
        CurSheet.Cells(2, i).AddComment (StrRange)
      ElseIf aParam.FieldName = "PCPICHPOWER" Then
        StrRange = aParam.DisplayName + "(Range[-100..500])"
        CurSheet.Cells(2, i).AddComment (StrRange)
      Else
        CurSheet.Cells(2, i).AddComment (StrRange)
      End If
      CurSheet.Cells(2, i).Comment.Visible = True
      CurSheet.Cells(2, i).Comment.Shape.Select
      Selection.ShapeRange.ScaleWidth 3.1, msoFalse, msoScaleFromTopLeft
      Selection.ShapeRange.ScaleHeight 1.6, msoFalse, msoScaleFromTopLeft
      CurSheet.Cells(2, i).Comment.Visible = False
    End If
    Call SetFieldValidate(CurSheet, i)
    
  Next
  
  Call FormatSheetFont(CurSheet, aMoc.Params.Count)
    
    CurSheet.Activate
    Cells.Select
    Selection.Locked = False
    Rows("1:2").Select
    Selection.Locked = True
    Range("A3").Select
    ActiveWindow.FreezePanes = True
  '  CurSheet.Protect Password:=GetSheetsPass, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True
        
  Set HomeSheet = ActiveWorkbook.Sheets("Home")
  HomeSheet.Cells(Index + 1, 1) = aMoc.MOCName
 
  '增加超链接
  HomeSheet.Activate
  hyperStr = "A" & Trim(str(Index + 1))
  HomeSheet.Range(hyperStr).Select
  hyperStr = aMoc.MOCName & "!A1"
  HomeSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        hyperStr, TextToDisplay:=aMoc.MOCName

End Sub

'增加枚举类型的有效性
Private Sub FormatColumn(ByRef CurSheet As Worksheet, ColNo As Integer, ByRef ParaInst As ParamInfo)
  Dim EnumInst As EnumInfo
  Dim EnumList As String
  Dim i As Integer
  
  For i = 1 To ParaInst.EnumRangeDef.Count
    Set EnumInst = ParaInst.EnumRangeDef.Item(i)
    EnumList = EnumList + EnumInst.sEnumStr + ","
  Next
  EnumList = Left(EnumList, Len(EnumList) - Len(","))
  
  CurSheet.Columns(ColNo).Select
  With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=EnumList
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    Range(Cells(1, ColNo), Cells(2, ColNo)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub
'增加整数的有效性
Private Sub IntegerValid(ByRef CurSheet As Worksheet, ColNo As Integer, RowNo As Integer, ByRef ParaInst As ParamInfo)
  Dim TableInfoSheet As Worksheet
  Dim i As Integer
  Dim s As String, isAdded As Boolean
  
  Set TableInfoSheet = ThisWorkbook.Worksheets("TableInfo")
  TableInfoSheet.Activate
  
  
  CurSheet.Activate
  
  If ParaInst.FieldTypeName <> "INT" Then
   Return
   
  CurSheet.Columns(ColNo).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=ParaInst.FieldMinValue, Formula2:=ParaInst.FieldMaxValue
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Prompt"
        .InputMessage = ""
        .ErrorMessage = "WorkSheet:" + CurSheet.name + "; FieldName:" + ParaInst.FieldName + "; " + vbLf _
                       + "Range:[" + ParaInst.FieldMinValue + ".." + ParaInst.FieldMaxValue + "]"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
  
End Sub

Private Sub StringLengthValid(ByRef CurSheet As Worksheet, ColNo As Integer, RowNo As Integer, ByRef ParaInst As ParamInfo)

  If ParaInst.FieldTypeName <> "STRING" Then
   Return
      
   CurSheet.Columns(ColNo).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="1", Formula2:="20"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "AA"
        .InputMessage = ""
        .ErrorMessage = "BB"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub

    

Private Sub FormatSheetFont(ByRef CurSheet As Worksheet, TotalCol As Integer)
  Range(CurSheet.Columns(1), CurSheet.Columns(TotalCol)).Columns.Select
  
  With Selection.Font
        .name = "Arial"
        .FontStyle = "常规"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
  End With
  Selection.NumberFormatLocal = "@"
  
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
  
  CurSheet.Cells.EntireColumn.AutoFit
  CurSheet.Cells(1, 1).Select
    
  Range(Cells(1, 1), Cells(1, TotalCol)).Select
  With Selection
      .Interior.ColorIndex = 34
      .Interior.Pattern = xlSolid
      .Font.FontStyle = "加粗"
  End With
End Sub


Sub InsertUserToolBar()
    Dim cmbNewBar As CommandBar
    Dim ctlBtn As CommandBarButton
    
    '中英文按钮栏
    Dim sCHSBarName As String
    Dim sENGBarName As String
    Dim sHideBarName As String
    Dim iHide  As Integer
    Dim iLanguage  As Integer
    Dim TableInfoSheet As Worksheet
    
    Set TableInfoSheet = ThisWorkbook.Sheets("TableInfo")
    
    iLanguage = TableInfoSheet.Cells(2, 1).Value
    
    If iLanguage = 0 Then
       sENGBarName = TableInfoSheet.Cells(2, 2).Value
       sCHSBarName = TableInfoSheet.Cells(3, 2).Value
       iHide = TableInfoSheet.Cells(2, 4).Value
       If iHide = 0 Then
          sHideBarName = TableInfoSheet.Cells(2, 5).Value
        End If
       If iHide = 1 Then
          sHideBarName = TableInfoSheet.Cells(3, 5).Value
        End If
      End If
         
    If iLanguage = 1 Then
       sENGBarName = TableInfoSheet.Cells(2, 3).Value
       sCHSBarName = TableInfoSheet.Cells(3, 3).Value
       iHide = TableInfoSheet.Cells(2, 4).Value
       If iHide = 0 Then
          sHideBarName = TableInfoSheet.Cells(2, 6).Value
        End If
       If iHide = 1 Then
          sHideBarName = TableInfoSheet.Cells(3, 6).Value
        End If
      End If
     
    
    On Error Resume Next
    
    Set cmbNewBar = CommandBars.Add(name:="Operate Bar")
    'Translate to English
    'With cmbNewBar
        'Set ctlBtn = .Controls.Add
        'With ctlBtn
            '.Style = msoButtonIconAndCaption
            '.BeginGroup = True
            '.Caption = sENGBarName
            '.TooltipText = sENGBarName
            '.OnAction = "TranslatetoEnglish"
            '.FaceId = 50
        'End With
       '' .Protection = msoBarNoCustomize
        '.Position = msoBarTop
        '.Visible = True
    'End With
    
     'Translate to Chinese
    'With cmbNewBar
        'Set ctlBtn = .Controls.Add
        'With ctlBtn
            '.Style = msoButtonIconAndCaption
            '.BeginGroup = True
            '.Caption = sCHSBarName
            '.TooltipText = sCHSBarName
            '.OnAction = "TranslatetoChinese"
            '.FaceId = 50
        'End With
        ''.Protection = msoBarNoCustomize
        '.Position = msoBarTop
        '.Visible = True
    'End With

    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = sHideBarName
            .TooltipText = sHideBarName
            .OnAction = "HideSheet"
            .FaceId = 50
        End With
      '  .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
End Sub


Private Sub TranslatetoChinese()
ThisWorkbook.Unprotect Password:=GetSheetsPass

  Dim nRowIndex As Integer
  Dim TableInfoSheet As Worksheet
  Dim sheetTableList As Worksheet
  Dim sheetCurrent As Worksheet
  Dim strObjId As String
  Dim nObjId As Integer
  Dim strObjName As String
  Dim oldObjName As String
  Dim FieldCol As String
  Dim sTableName As String
  Dim nObjIndex As Integer

  SetBarLanguage ("CHS")

  Set TableInfoSheet = ThisWorkbook.Sheets("TableInfo")


  Dim CHSName As String
  Dim RangeName As String
  Dim FieldPostil As String
  Dim sFieldName As String
  Dim sDataType As String
  Dim sRangeList As String


  oldObjName = ""
  For nRowIndex = 5 To 30000

    If Trim(TableInfoSheet.Cells(nRowIndex, 2).Value) = "" Then
      Exit For
    End If

    strObjId = Trim(TableInfoSheet.Cells(nRowIndex, 1).Value)
    If strObjId <> "" Then
       ' nObjId = CInt(strObjId)
      Set sheetCurrent = ThisWorkbook.Sheets(strObjId)
      
      sheetCurrent.Unprotect Password:=GetSheetsPass
      
      If oldObjName <> strObjId Then
           nObjIndex = 1
           oldObjName = strObjId
           sheetCurrent.Rows(2).ClearContents
        End If
    End If

    ' 设置批注信息
    CHSName = Trim(TableInfoSheet.Cells(nRowIndex, 8))
    sFieldName = Trim(TableInfoSheet.Cells(nRowIndex, 2))
    RangeName = GetRangeInfoCHS(nRowIndex, TableInfoSheet)
    FieldPostil = CHSName + "(" + RangeName + ")"

    sDataType = Trim(TableInfoSheet.Cells(nRowIndex, 3))
    sRangeList = Trim(TableInfoSheet.Cells(nRowIndex, 6))

    sheetCurrent.Activate
    sheetCurrent.Cells(2, nObjIndex).Select
    sheetCurrent.Cells(2, nObjIndex).ClearComments
    sheetCurrent.Cells(2, nObjIndex).AddComment FieldPostil
    sheetCurrent.Cells(2, nObjIndex).Comment.Shape.Width = 260
    sheetCurrent.Cells(2, nObjIndex).Comment.Shape.Height = 100
    sheetCurrent.Cells(2, nObjIndex) = CHSName

    sheetCurrent.Cells(3, 1).Select

    nObjIndex = nObjIndex + 1
    
    
    Cells.Select
    Selection.Locked = False
    Rows("1:2").Select
    Selection.Locked = True
    Range("A3").Select
    'sheetCurrent.Protect Password:=GetSheetsPass, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True

  Next nRowIndex


  '设置表头格式
  'SetGrdiMaster

  'ThisWorkbook.Protect Password:=GetSheetsPass

  MsgBox "完成."
End Sub
Private Sub TranslatetoEnglish()
ThisWorkbook.Unprotect Password:=GetSheetsPass
  Dim nRowIndex As Integer
  Dim TableInfoSheet As Worksheet
  Dim sheetTableList As Worksheet
  Dim sheetCurrent As Worksheet
  Dim strObjId As String
  Dim nObjId As Integer
  Dim strObjName As String
  Dim oldObjName As String
  Dim FieldCol As String
  Dim sTableName As String
  Dim nObjIndex As Integer

  SetBarLanguage ("ENG")

  Set TableInfoSheet = ThisWorkbook.Sheets("TableInfo")


  Dim ENGName As String
  Dim RangeName As String
  Dim FieldPostil As String
  Dim sFieldName As String
  Dim sDataType As String
  Dim sRangeList As String



  oldObjName = ""
  For nRowIndex = 5 To 30000

    If Trim(TableInfoSheet.Cells(nRowIndex, 2).Value) = "" Then
      Exit For
    End If

    strObjId = Trim(TableInfoSheet.Cells(nRowIndex, 1).Value)
    If strObjId <> "" Then
       ' nObjId = CInt(strObjId)
      Set sheetCurrent = ThisWorkbook.Sheets(strObjId)
      
      sheetCurrent.Unprotect Password:=GetSheetsPass
      
      If oldObjName <> strObjId Then
           nObjIndex = 1
           oldObjName = strObjId
           sheetCurrent.Rows(2).ClearContents
        End If
    End If

    ' 设置批注信息
    ENGName = Trim(TableInfoSheet.Cells(nRowIndex, 7))
    sFieldName = Trim(TableInfoSheet.Cells(nRowIndex, 2))
    RangeName = GetRangeInfoENG(nRowIndex, TableInfoSheet)
    FieldPostil = ENGName + "(" + RangeName + ")"

    sDataType = Trim(TableInfoSheet.Cells(nRowIndex, 3))
    sRangeList = Trim(TableInfoSheet.Cells(nRowIndex, 6))

     sheetCurrent.Activate
    sheetCurrent.Cells(2, nObjIndex).Select
    sheetCurrent.Cells(2, nObjIndex).ClearComments
    sheetCurrent.Cells(2, nObjIndex).AddComment FieldPostil
    sheetCurrent.Cells(2, nObjIndex).Comment.Shape.Width = 260
    sheetCurrent.Cells(2, nObjIndex).Comment.Shape.Height = 100
    sheetCurrent.Cells(2, nObjIndex) = ENGName

    sheetCurrent.Cells(3, 1).Select

    nObjIndex = nObjIndex + 1
    
    Cells.Select
    Selection.Locked = False
    Rows("1:2").Select
    Selection.Locked = True
    Range("A3").Select
   ' sheetCurrent.Protect Password:=GetSheetsPass, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True

  Next nRowIndex

  '  ThisWorkbook.Protect Password:=GetSheetsPass
  '设置表头格式
  'SetGrdiMaster

  MsgBox "完成."
End Sub
Public Sub HideSheet()
    '隐藏空数据表
 ThisWorkbook.Unprotect Password:=GetSheetsPass
    Dim sheetCurrent As Worksheet
    Dim sheetTableInfo As Worksheet
    Dim sheetHOME As Worksheet
    Dim sTableName  As String
    Dim sHideBarName  As String
    Dim iLanguage  As Integer
    Dim iHide  As Integer
    Dim i As Integer
    Dim iHome  As Integer
    
    Set sheetTableInfo = ThisWorkbook.Sheets("TableInfo")
    Set sheetHOME = ThisWorkbook.Sheets("HOME")
    
    iLanguage = sheetTableInfo.Cells(2, 1).Value
    iHide = sheetTableInfo.Cells(2, 4).Value
    
    For i = 1 To ThisWorkbook.Sheets.Count
        sTableName = ThisWorkbook.Sheets(i).name
        If sTableName = "" Then
           Exit For
         End If
        
        Set sheetCurrent = ThisWorkbook.Sheets(sTableName)
        
        If iHide = 1 And sTableName <> "CMETemplateInfo" And sTableName <> "Refresh" And sTableName <> "FileIdentification" And sTableName <> "ValidInfo" And sTableName <> "TableInfo" And sTableName <> "UserSelectMoc" And sTableName <> "EnumDef" Then
            sheetCurrent.Visible = True
'            For iHome = 2 To 30000
'               If sheetHOME.Cells(iHome, 1).Value = "" Then
'                  Exit For
'                 End If
'               If sheetHOME.Cells(iHome, 1).Value = sTableName Then
'                   sheetHOME.Rows(iHome).Hidden = False
'                  Exit For
'                 End If
'            Next iHome
         End If

        If (sheetCurrent.Cells(3, 1).Value = "") And (iHide = 0) And sTableName <> "Cover" And sTableName <> "Home" And sTableName <> "UserSelectMoc" Then
          sheetCurrent.Visible = False
'          For iHome = 2 To 30000
'               If sheetHOME.Cells(iHome, 1).Value = "" Then
'                  Exit For
'                 End If
'               If sheetHOME.Cells(iHome, 1).Value = sTableName Then
'                   sheetHOME.Rows(iHome).Hidden = True
'                  Exit For
'                 End If
'            Next iHome
        End If

    Next i
    
    If iHide = 0 Then
      sheetTableInfo.Cells(2, 4).Value = 1
    End If
    If iHide = 1 Then
      sheetTableInfo.Cells(2, 4).Value = 0
    End If
    
    iHide = sheetTableInfo.Cells(2, 4).Value
    
    If iLanguage = 0 Then
       If iHide = 0 Then
          sHideBarName = sheetTableInfo.Cells(2, 5).Value
        End If
       If iHide = 1 Then
          sHideBarName = sheetTableInfo.Cells(3, 5).Value
        End If
      End If
            
    If iLanguage = 1 Then
       iHide = sheetTableInfo.Cells(2, 4).Value
       If iHide = 0 Then
          sHideBarName = sheetTableInfo.Cells(2, 6).Value
        End If
       If iHide = 1 Then
          sHideBarName = sheetTableInfo.Cells(3, 6).Value
        End If
      End If
      
    If iLanguage = 1 Or iLanguage = 0 Then
        CommandBars("Operate Bar").Controls.Item(1).Caption = sHideBarName
        CommandBars("Operate Bar").Controls.Item(1).TooltipText = sHideBarName
      End If
      
' ThisWorkbook.Protect Password:=GetSheetsPass
End Sub


Public Sub SetBarLanguage(sLanguageType As String)
    '中英文按钮栏
    Dim sCHSBarName As String
    Dim sENGBarName As String
    Dim iLanguage  As Integer
    Dim iHide  As Integer
    Dim sHideBarName  As String
    Dim TableInfoSheet As Worksheet
    Set TableInfoSheet = ThisWorkbook.Sheets("TableInfo")
           
    If sLanguageType = "ENG" Then
        iLanguage = 0
        TableInfoSheet.Cells(2, 1).Value = iLanguage
       End If
    If sLanguageType = "CHS" Then
        iLanguage = 1
        TableInfoSheet.Cells(2, 1).Value = iLanguage
       End If
       
    If iLanguage = 0 Then
       sENGBarName = TableInfoSheet.Cells(2, 2).Value
       sCHSBarName = TableInfoSheet.Cells(3, 2).Value
       iHide = TableInfoSheet.Cells(2, 4).Value
       If iHide = 0 Then
          sHideBarName = TableInfoSheet.Cells(2, 5).Value
        End If
       If iHide = 1 Then
          sHideBarName = TableInfoSheet.Cells(3, 5).Value
        End If
      End If
         
    If iLanguage = 1 Then
       sENGBarName = TableInfoSheet.Cells(2, 3).Value
       sCHSBarName = TableInfoSheet.Cells(3, 3).Value
       iHide = TableInfoSheet.Cells(2, 4).Value
       If iHide = 0 Then
          sHideBarName = TableInfoSheet.Cells(2, 6).Value
        End If
       If iHide = 1 Then
          sHideBarName = TableInfoSheet.Cells(3, 6).Value
        End If
      End If
      
    If iLanguage = 1 Or iLanguage = 0 Then
        CommandBars("Operate Bar").Controls.Item(1).Caption = sENGBarName
        CommandBars("Operate Bar").Controls.Item(1).TooltipText = sENGBarName
        CommandBars("Operate Bar").Controls.Item(2).Caption = sCHSBarName
        CommandBars("Operate Bar").Controls.Item(2).TooltipText = sCHSBarName
        CommandBars("Operate Bar").Controls.Item(3).Caption = sHideBarName
        CommandBars("Operate Bar").Controls.Item(3).TooltipText = sHideBarName
      End If
'  Call ReHOME(sLanguageType)
End Sub

'获取有效范围提示或者错误信息
Private Function GetRangeInfoCHS(nRowIndex As Integer, TableInfoSheet As Worksheet) As String
  Const FieldNameCol = 2
  Const DataTypeCol = 3
  Const MinValCol = 5
  Const MaxValCol = 6
  Const RangeListCol = 7
  Const ValueTypeCol = 24
  Const FINT = "INT", FSTRING = "STRING", FLIST = "LIST", FBITMAP = "BITMAP"

  Dim sFieldName As String
  Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sValueType As String
  
  sFieldName = Trim(TableInfoSheet.Cells(nRowIndex, 2).Value)
  sDataType = Trim(TableInfoSheet.Cells(nRowIndex, 3).Value)
  sMinVal = Trim(TableInfoSheet.Cells(nRowIndex, 4).Value)
  sMaxVal = Trim(TableInfoSheet.Cells(nRowIndex, 5).Value)
  sRangeList = Trim(TableInfoSheet.Cells(nRowIndex, 6).Value)
  
  GetRangeInfoCHS = ""
  If (sDataType = FSTRING) Then
    If sMinVal = sMaxVal Then
      GetRangeInfoCHS = "长度范围" + "[" + sMinVal + "]"
    Else
      GetRangeInfoCHS = "长度范围" + "[" + sMinVal + ".." + sMaxVal + "]"
    End If
  End If
  
  If (sDataType = FBITMAP) Then
      GetRangeInfoCHS = "长度 = " + sMaxVal
  End If
  
  If (sDataType = FINT) Then
    GetRangeInfoCHS = "取值范围" + "[" + sRangeList + "]"
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
Private Function GetRangeInfoENG(nRowIndex As Integer, TableInfoSheet As Worksheet) As String
  Const FieldNameCol = 2
  Const DataTypeCol = 3
  Const MinValCol = 5
  Const MaxValCol = 6
  Const RangeListCol = 7
  Const ValueTypeCol = 24
  Const FINT = "INT", FSTRING = "STRING", FLIST = "LIST", FBITMAP = "BITMAP"

  Dim sFieldName As String
  Dim sDataType As String, sMinVal As String, sMaxVal As String, sRangeList As String, sValueType As String
  
  sFieldName = Trim(TableInfoSheet.Cells(nRowIndex, 2).Value)
  sDataType = Trim(TableInfoSheet.Cells(nRowIndex, 3).Value)
  sMinVal = Trim(TableInfoSheet.Cells(nRowIndex, 4).Value)
  sMaxVal = Trim(TableInfoSheet.Cells(nRowIndex, 5).Value)
  sRangeList = Trim(TableInfoSheet.Cells(nRowIndex, 6).Value)
  
  GetRangeInfoENG = ""
  If (sDataType = FSTRING) Then
    If sMinVal = sMaxVal Then
      GetRangeInfoENG = "Length" + "[" + sMinVal + "]"
    Else
      GetRangeInfoENG = "Length" + "[" + sMinVal + ".." + sMaxVal + "]"
    End If
  End If
  
  If (sDataType = FBITMAP) Then

     GetRangeInfoENG = "Length = " + sMaxVal

  End If
  
  
  If (sDataType = FINT) Then
      GetRangeInfoENG = "Range" + "[" + sRangeList + "]"
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
  
  sMOName = Trim(sheetTableList.Cells(1, nRow).Value)
  sheetHOME.Cells(1, 2).Value = sMOName
    
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


Public Sub UnprotectWorkBook()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=GetSheetsPass
    'ThisWorkbook.Save
End Sub

Public Function GetSheetsPass() As String
    GetSheetsPass = "HWCME"
End Function

'获取表格分支参数位置信息
Public Sub GetBranchParameterLocalRange()
Dim CurSheetCoulumn As Integer
Dim ValidInfoSheet As Worksheet
Dim CurSheet As Worksheet
Dim ValidInfoRow As Integer
Dim sFieldName As String
Dim sTableName As String
Dim sBranchFieldName As String
Dim i As Integer


 Set ValidInfoSheet = ThisWorkbook.Sheets("ValidInfo")
 
 For ValidInfoRow = 25 To 1000
    If ValidInfoSheet.Cells(ValidInfoRow, 1).Value = "" Then
        Exit For
    End If
    sTableName = ValidInfoSheet.Cells(ValidInfoRow, 1).Value
    For i = ActiveWorkbook.Sheets.Count To 1 Step -1
        Set CurSheet = ActiveWorkbook.Sheets(i)
        If sTableName = CurSheet.name Then
            sBranchFieldName = ValidInfoSheet.Cells(ValidInfoRow, 2).Value
            sFieldName = ValidInfoSheet.Cells(ValidInfoRow, 3).Value
            For CurSheetCoulumn = 1 To 1000
                If CurSheet.Cells(1, CurSheetCoulumn).Value = "" Then
                    Exit For
               End If
               If CurSheet.Cells(1, CurSheetCoulumn).Value = sBranchFieldName Then
                  ValidInfoSheet.Cells(ValidInfoRow, 5) = CurSheetCoulumn
                  Exit For
               End If
             Next CurSheetCoulumn
             For CurSheetCoulumn = 1 To 1000
                If CurSheet.Cells(1, CurSheetCoulumn).Value = "" Then
                    Exit For
               End If
               If CurSheet.Cells(1, CurSheetCoulumn).Value = sFieldName Then
                  ValidInfoSheet.Cells(ValidInfoRow, 6) = CurSheetCoulumn
                  Exit For
               End If
             Next CurSheetCoulumn
             Exit For
        End If
    Next i
Next ValidInfoRow

End Sub


