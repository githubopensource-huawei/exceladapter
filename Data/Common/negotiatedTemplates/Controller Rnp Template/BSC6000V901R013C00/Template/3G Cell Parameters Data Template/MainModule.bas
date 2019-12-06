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

Sub ShowUserForm()
  UserForm1.Show vbModeless
End Sub
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
        .Name = "Arial"
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
        .Name = "Arial"
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
        .Name = "宋体"
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

Public Sub Refresh()
  Dim cmd As New ADODB.Command
  Dim rs As New ADODB.Recordset
  Dim sMocName As String, sOldMocName As String, sFiledName As String, sOldFieldName As String
  Dim sDspName As String, sMaxVal As String, sMinVal As String, sOldMaxVal As String, sOldMinVal As String, sMustGive As String

  Dim SQLCmd As New ADODB.Command
  Dim SQLRs As New ADODB.Recordset
  
  If conn.State = adStateClosed Then
    MsgBox "请先连接数据库。"
    Exit Sub
  End If
  
  Set cmd.ActiveConnection = conn
  
  Dim iFieldType As Integer
  Dim MocInst As MocInfo
  Dim ParamInst As ParamInfo
  Dim RangeInst As RangeInfo
  Dim EnumInst As EnumInfo
  Dim iIndex As Integer
  Dim iEntryNum As Integer
  
  iEntryNum = 0
  
  Call GetSheetDefineData
  
  cmd.CommandText = " if exists (select * from sysobjects where name = 'v_RNO_MOCNameInfo') drop view v_RNO_MOCNameInfo "
  rs.CursorLocation = adUseClient
  rs.Open cmd
  
  cmd.CommandText = " create view v_RNO_MOCNameInfo " _
   & " as " _
   & " select sVersion, TableName from Utils_CellWPARAM_TableMap where ClassID in (9, 11) and sVersion = '" + RNC_VERSION + "'"

  rs.CursorLocation = adUseClient
  rs.Open cmd
  
  cmd.CommandText = " if exists (select * from sysobjects where name = 'v_RNO_FieldInfo') drop view v_RNO_FieldInfo "
  rs.CursorLocation = adUseClient
  rs.Open cmd
  
   cmd.CommandText = " create view v_RNO_FieldInfo " _
  & " as  " _
  & " select a.sTableName, b.sFieldName, b.sDspName, b.iFieldType, b.iSortField, c.iValue, c.sInput, c.sDisplay, b.iMustGive,b.iSortIndex" _
  & "  from v_RNO_MOCNameInfo o, Utils_TableDef a, Utils_FieldDef b, Utils_FieldEnumDef c " _
  & "  where a.iMode = b.iMode And b.iMode = c.iMode And a.iMode = 2" _
  & "  and a.iTableId = b.iTableId and b.iTableId = c.iTableId" _
  & "  and b.iFieldId = c.iFieldId and b.iFieldType in (101, 5)" _
  & "  and (b.iVisible=1 or b.iVisible is null)" _
  & "  and a.sVersion = b.sVersion and b.sVersion = c.sVersion and a.sVersion = '" + RNC_VERSION + "'" _
  & "  and o.TableName = a.sTableName and o.sVersion = a.sVersion " _
  & " Union" _
  & " select a.sTableName, a.sFieldName, a.sDspName, a.iFieldType, a.iSortField, c.iMinValue, convert(varchar(255), c.iMaxValue)," + " '' As sDisplay, a.iMustGive, a.iSortIndex" _
  & "  from view_FieldAllInfo a " _
  & "  join v_RNO_MOCNameInfo o " _
  & "  on o.TableName = a.sTableName and o.sVersion = a.sVersion " _
  & "  left join Utils_FieldRangeDef c " _
  & "  on a.iMode = c.iMode and a.sVersion = c.sVersion and a.iTableId = c.iTableId and a.iFieldId = c.iFieldId " _
  & "  where a.sVersion ='" + RNC_VERSION + "' and a.iMode = 2 and a.iFieldType not in (101, 5) " _
  & "  and (a.iVisible=1 or a.iVisible is null)  "

  rs.CursorLocation = adUseClient
  rs.Open cmd
  
  cmd.CommandText = "select sFieldName, iFieldType, sTableName, sDspName, iValue, sInput, sDisplay, iMustGive from v_RNO_FieldInfo  order by sTableName, iSortField, iSortIndex"

  
  Dim sFieldTypeName As String, sCHSDspName As String, sBitmapStr As String      '类型名， 中文显示名，bitmap标识
  Dim iFieldLocation As Integer                                                  '字段位置
  Dim SQLStr, PTableName As String
  
  iFieldLocation = 1
  
  rs.CursorLocation = adUseClient
  rs.Open cmd
  
  For iIndex = ParamMocList.Count To 1 Step -1
    ParamMocList.Remove (iIndex)
  Next
  
  Do While Not rs.EOF
    
    sMocName = GetFieldValue(rs, "sTableName")
    sFiledName = GetFieldValue(rs, "sFieldName")
    iFieldType = GetFieldValue(rs, "iFieldType")
    sDspName = GetFieldValue(rs, "sDspName")
    
    sMustGive = GetFieldValue(rs, "iMustGive")
    If sMustGive = "1" Then
     sMustGive = "YES"
    Else
     sMustGive = "NO"
    End If
    
    If (iFieldType = INTTYPE Or iFieldType = BIGINTTYPE) Then
      sMinVal = GetFieldValue(rs, "iValue")
      sMaxVal = GetFieldValue(rs, "sInput")
      If sMinVal = "" Then
       sMinVal = 0
      End If
      If sMaxVal = "" Then
        sMaxVal = 65535
      End If
    ElseIf (iFieldType = ENUMTYPE) Then
      sMinVal = GetFieldValue(rs, "iValue")
      sMaxVal = GetFieldValue(rs, "sInput")
    End If
    
    PTableName = "t_P_" + sMocName + "_" + RNC_VERSION
        
    If iFieldType <> INTTYPE And iFieldType <> BIGINTTYPE And iFieldType <> ENUMTYPE And iFieldType <> BITMAPTYPE Then
        Set SQLCmd.ActiveConnection = conn
        SQLStr = "select length from syscolumns where name = '" + sFiledName + "' and id = object_id('" + PTableName + "') "
        SQLCmd.CommandText = SQLStr
        
        SQLRs.CursorLocation = adUseClient
        SQLRs.Open SQLCmd
        sMinVal = 1
        sMaxVal = GetFieldValue(SQLRs, "length")
        SQLRs.Close
    End If
    
    If iFieldType = BITMAPTYPE Then
       Set SQLCmd.ActiveConnection = conn
       SQLStr = "select count(iValue) as length from v_RNO_FieldInfo where iFieldType = 5 and sFieldName = '" + sFiledName + "' "
       SQLCmd.CommandText = SQLStr
    
       SQLRs.CursorLocation = adUseClient
       SQLRs.Open SQLCmd
       sMinVal = 1
       sMaxVal = GetFieldValue(SQLRs, "length")
       SQLRs.Close
    End If
            
    If (iFieldType = INTTYPE) Or (iFieldType = BIGINTTYPE) Then
       sFieldTypeName = "INT"
    ElseIf iFieldType = ENUMTYPE Then
       sFieldTypeName = "LIST"
    ElseIf iFieldType = BITMAPTYPE Then
       sFieldTypeName = "BITMAP"
    Else
       sFieldTypeName = "STRING"
    End If
   
    If StrComp(sOldMocName, sMocName, vbTextCompare) <> 0 Then
      sOldMocName = sMocName
      sOldFieldName = ""
      Set MocInst = New MocInfo
      ParamMocList.Add MocInst
      MocInst.MOCName = sMocName
      
      If sMocName <> "CELL" Then
        Call AddPrimKey(MocInst, sMocName)
      Else
        Call ProcessCELLMO(MocInst, iEntryNum)
      End If
      
    End If
   
   If sMocName <> "CELL" Then
    If Not (StrComp(sFiledName, sOldFieldName, vbTextCompare) = 0 And StrComp(sOldMocName, sMocName, vbTextCompare) = 0) Then
      Set ParamInst = New ParamInfo
      ParamInst.FieldName = sFiledName
      ParamInst.DisplayName = sDspName
      ParamInst.FieldType = iFieldType
      
      If (iFieldType <> ENUMTYPE) Then
        ParamInst.FieldMinValue = sMinVal
        ParamInst.FieldMaxValue = sMaxVal
      End If
      
      ParamInst.FieldISMustGive = sMustGive
      ParamInst.FieldTypeName = sFieldTypeName
      MocInst.Params.Add ParamInst
    End If
    
    If (StrComp(sFiledName, sOldFieldName, vbTextCompare) = 0 And StrComp(sOldMocName, sMocName, vbTextCompare) = 0) Then
        If (iFieldType = INTTYPE) Or (iFieldType = BIGINTTYPE) Then
            ParamInst.FieldMaxValue = sMaxVal
        End If
    End If
    
    If ((StrComp(sFiledName, sOldFieldName, vbTextCompare) <> 0) Or (StrComp(sMaxVal, sOldMaxVal, vbTextCompare) <> 0) Or (StrComp(sMinVal, sOldMinVal, vbTextCompare) <> 0)) Then
      sOldMaxVal = sMaxVal
      sOldMinVal = sMinVal

      If (sMaxVal <> "" And sOldMinVal <> "") Then
        If iFieldType = ENUMTYPE Then
          Set EnumInst = New EnumInfo
          EnumInst.sEnumStr = sMaxVal
          EnumInst.sEnumInt = sMinVal
          ParamInst.EnumRangeDef.Add EnumInst
        Else
          Set RangeInst = New RangeInfo
          RangeInst.sMaxVal = sMaxVal
          RangeInst.sMinVal = sMinVal
          ParamInst.RangeDef.Add RangeInst
        End If
      End If
    End If
    
    sOldFieldName = sFiledName
    
    End If
    rs.MoveNext
  Loop
  rs.Close

  Call DeleteParamSheets
  sOldMocName = "Home"
 
  Dim RowCnt As Integer
  RowCnt = 5
  For iIndex = 1 To ParamMocList.Count
    Set MocInst = ParamMocList.Item(iIndex)
    'If MocInst.MOCName = "CELLAMRC" Then
     Call GenerateMOInfo("TableInfo", MocInst, RowCnt)
     Call AddParamSheets(MocInst, iIndex, sOldMocName)
     sOldMocName = MocInst.MOCName
    'End If
  Next
  
  Call GetBranchParameterFieldRange
  Call GetBranchParameterLocalRange
  Call RebuildSheet
  
  MsgBox "OK."
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
    If StrComp(CurSheet.Name, "Refresh", vbTextCompare) <> 0 And StrComp(CurSheet.Name, "TableInfo", vbTextCompare) <> 0 And StrComp(CurSheet.Name, "Cover", vbTextCompare) <> 0 _
      And StrComp(CurSheet.Name, "ValidInfo", vbTextCompare) <> 0 _
      And StrComp(CurSheet.Name, "Home", vbTextCompare) _
      And StrComp(CurSheet.Name, "CHS", vbTextCompare) _
      And StrComp(CurSheet.Name, "ENG", vbTextCompare) Then
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

Private Sub GetCELLFieldProperty(ByRef MocInst As MocInfo, ByVal sFiledName As String)
  Dim cmd As New ADODB.Command
  Dim rs As New ADODB.Recordset
  Dim ParamInst As ParamInfo
  Dim sDspName As String, sMaxVal As String, sMinVal As String, sOldMaxVal As String
  Dim sOldMinVal As String, sOldFieldName As String, sFieldTypeName As String, sMustGive, PTableName, SQLStr As String
  Dim iFieldType As Integer
  Dim RangeInst As RangeInfo
  Dim EnumInst As EnumInfo
  Dim SQLCmd As New ADODB.Command
  Dim SQLRs As New ADODB.Recordset
  
  Set cmd.ActiveConnection = conn
  
 cmd.CommandText = " if exists (select * from sysobjects where name = 'v_RNO_CELLFieldInfo') drop view v_RNO_CELLFieldInfo "
  rs.CursorLocation = adUseClient
  rs.Open cmd
  
Set ParamInst = New ParamInfo
sOldFieldName = ""

  cmd.CommandText = " create view v_RNO_CELLFieldInfo " _
  & " as  " _
  & " select a.sTableName, b.sFieldName, b.sDspName, b.iFieldType, b.iSortField, c.iValue, c.sInput, b.iMustGive " _
  & "  from v_RNO_MOCNameInfo o, Utils_TableDef a, Utils_FieldDef b, Utils_FieldEnumDef c " _
  & "  where a.iMode = b.iMode And b.iMode = c.iMode And a.iMode = 2" _
  & "  and a.iTableId = b.iTableId and b.iTableId = c.iTableId" _
  & "  and b.iFieldId = c.iFieldId and b.iFieldType in (101, 5)" _
  & "  and (b.iVisible=1 or b.iVisible is null)" _
  & "  and a.sVersion = b.sVersion and b.sVersion = c.sVersion and a.sVersion ='" + RNC_VERSION + "' " _
  & "  and o.TableName = a.sTableName and o.sVersion = a.sVersion " + " and a.sTableName = 'CELL' and b.sFieldName = '" + sFiledName + "'" _
  & " Union" _
  & " select a.sTableName, a.sFieldName, a.sDspName, a.iFieldType, a.iSortField, c.iMinValue, convert(varchar(255), c.iMaxValue), a.iMustGive" _
  & "  from view_FieldAllInfo a " _
  & "  join v_RNO_MOCNameInfo o " _
  & "  on o.TableName = a.sTableName and o.sVersion = a.sVersion " _
  & "  left join Utils_FieldRangeDef c " _
  & "  on a.iMode = c.iMode and a.sVersion = c.sVersion and a.iTableId = c.iTableId and a.iFieldId = c.iFieldId " _
  & "  where a.sVersion ='" + RNC_VERSION + "' and a.iMode = 2 and a.iFieldType not in (101, 5) " _
  & "  and (a.iVisible=1 or a.iVisible is null) " + " and a.sTableName = 'CELL' and a.sFieldName ='" + sFiledName + "'" _
  
  rs.CursorLocation = adUseClient
  rs.Open cmd
  
  cmd.CommandText = "select sFieldName, iFieldType,  sDspName, iValue, sInput, iMustGive from v_RNO_CELLFieldInfo order by iSortField"
  rs.CursorLocation = adUseClient
  rs.Open cmd
  
  Do While Not rs.EOF
    iFieldType = GetFieldValue(rs, "iFieldType")
    sDspName = GetFieldValue(rs, "sDspName")
    sMinVal = GetFieldValue(rs, "iValue")
    sMaxVal = GetFieldValue(rs, "sInput")
    
    sMustGive = GetFieldValue(rs, "iMustGive")
    If sMustGive = "1" Then
      sMustGive = "YES"
    Else
      sMustGive = "NO"
    End If
    
    PTableName = "t_P_CELL_" + RNC_VERSION
    If iFieldType <> INTTYPE And iFieldType <> BIGINTTYPE And iFieldType <> ENUMTYPE And iFieldType <> BITMAPTYPE Then
        Set SQLCmd.ActiveConnection = conn
        SQLStr = "select length from syscolumns where name = '" + sFiledName + "' and id = object_id('" + PTableName + "') "
        SQLCmd.CommandText = SQLStr
        
        SQLRs.CursorLocation = adUseClient
        SQLRs.Open SQLCmd
        sMinVal = 1
        sMaxVal = GetFieldValue(SQLRs, "length")
        SQLRs.Close
    End If
      
    If Not StrComp(sFiledName, sOldFieldName, vbTextCompare) = 0 Then
      ParamInst.FieldName = sFiledName
      ParamInst.DisplayName = sDspName
      ParamInst.FieldType = iFieldType
      If (iFieldType = INTTYPE) Or (iFieldType = BIGINTTYPE) Then
        sFieldTypeName = "INT"
      ElseIf iFieldType = ENUMTYPE Then
        sFieldTypeName = "LIST"
      ElseIf iFieldType = BITMAPTYPE Then
        sFieldTypeName = "BITMAP"
      Else
        sFieldTypeName = "STRING"
      End If
      ParamInst.FieldTypeName = sFieldTypeName
      ParamInst.FieldISMustGive = sMustGive
      
      If iFieldType <> ENUMTYPE Then
        ParamInst.FieldMaxValue = sMaxVal
        ParamInst.FieldMinValue = sMinVal
      End If
      MocInst.Params.Add ParamInst
    End If
    
    If StrComp(sFiledName, sOldFieldName, vbTextCompare) = 0 Then
        If (iFieldType = INTTYPE) Or (iFieldType = BIGINTTYPE) Then
            ParamInst.FieldMaxValue = sMaxVal
        End If
    End If
    
    If ((StrComp(sFiledName, sOldFieldName, vbTextCompare) <> 0) Or (StrComp(sMaxVal, sOldMaxVal, vbTextCompare) <> 0) Or (StrComp(sMinVal, sOldMinVal, vbTextCompare) <> 0)) Then
      sOldMaxVal = sMaxVal
      sOldMinVal = sMinVal

      If (sMaxVal <> "" And sOldMinVal <> "") Then
        If iFieldType = ENUMTYPE Then
          Set EnumInst = New EnumInfo
          EnumInst.sEnumStr = sMaxVal
          EnumInst.sEnumInt = sMinVal
          ParamInst.EnumRangeDef.Add EnumInst
        Else
          Set RangeInst = New RangeInfo
          RangeInst.sMaxVal = sMaxVal
          RangeInst.sMinVal = sMinVal
          ParamInst.RangeDef.Add RangeInst
        End If
      End If
    End If
   
    sOldFieldName = sFiledName
    rs.MoveNext
  Loop
  rs.Close
  
  
End Sub

Private Sub ProcessCELLMO(ByRef MocInst As MocInfo, ByRef EntyNum As Integer)
Dim sFiledName As String

  If EntyNum = 0 Then
   EntyNum = 1
  ElseIf EntyNum <> 0 Then
    Exit Sub
  End If
  
  Call AddPrimKey(MocInst, "CELL")
   
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
  
 sFiledName = "PEERISVALID"
 Call GetCELLFieldProperty(MocInst, sFiledName)

 sFiledName = "PEERCELLID"
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
Private Sub GenerateMOInfo(ByRef SheetName As String, ByRef MOInfo As MocInfo, ByRef Counter As Integer)
  Dim MOName, FieldName, FieldType, FieldLocation, EnumList, FieldDspName, CHSFieldName, BitmapField, StrRange As String
  Dim MinValue, MaxValue, RowNo, ColumnNo, i, j As Integer
  Dim aParam As ParamInfo
  Dim EnumInst As EnumInfo
  Dim RangeInst As RangeInfo
  Dim CurSheet As Worksheet
  
  Set CurSheet = ThisWorkbook.Sheets(SheetName)
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
    If StrComp(CurSheet.Name, "Refresh", vbTextCompare) <> 0 And StrComp(CurSheet.Name, "TableInfo", vbTextCompare) <> 0 And StrComp(CurSheet.Name, "Cover", vbTextCompare) <> 0 And StrComp(CurSheet.Name, "ValidInfo", vbTextCompare) <> 0 _
      And StrComp(CurSheet.Name, "Home", vbTextCompare) _
      And StrComp(CurSheet.Name, "CHS", vbTextCompare) _
      And StrComp(CurSheet.Name, "ENG", vbTextCompare) Then
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
   If aMoc.MOCName = ActiveWorkbook.Sheets(i).Name Then
      SheetIsNotExist = 0
   End If
Next

If SheetIsNotExist Then
  ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(PreSheetName)
  Set CurSheet = ActiveSheet
  CurSheet.Activate
  CurSheet.Name = aMoc.MOCName
Else
  Set CurSheet = ActiveWorkbook.Sheets(aMoc.MOCName)
  CurSheet.Activate
  CurSheet.Name = aMoc.MOCName
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
        .ErrorMessage = "WorkSheet:" + CurSheet.Name + "; FieldName:" + ParaInst.FieldName + "; " + vbLf _
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
        .Name = "Arial"
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
    
    Set cmbNewBar = CommandBars.Add(Name:="Operate Bar")
    'Translate to English
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = sENGBarName
            .TooltipText = sENGBarName
            .OnAction = "TranslatetoEnglish"
            .FaceId = 50
        End With
       ' .Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With
    
     'Translate to Chinese
    With cmbNewBar
        Set ctlBtn = .Controls.Add
        With ctlBtn
            .Style = msoButtonIconAndCaption
            .BeginGroup = True
            .Caption = sCHSBarName
            .TooltipText = sCHSBarName
            .OnAction = "TranslatetoChinese"
            .FaceId = 50
        End With
        '.Protection = msoBarNoCustomize
        .Position = msoBarTop
        .Visible = True
    End With

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
    Dim iRow  As Integer
    Dim iHome  As Integer
    
    Set sheetTableInfo = ThisWorkbook.Sheets("TableInfo")
    Set sheetHOME = ThisWorkbook.Sheets("HOME")
    
    iLanguage = sheetTableInfo.Cells(2, 1).Value
    iHide = sheetTableInfo.Cells(2, 4).Value
    
    For iRow = 5 To 30000
        sTableName = sheetTableInfo.Cells(iRow, 1).Value
        If sTableName = "" Then
           Exit For
         End If
        
        Set sheetCurrent = ThisWorkbook.Sheets(sTableName)
        
        If iHide = 1 Then
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

        If (sheetCurrent.Cells(3, 1).Value = "") And (iHide = 0) Then
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

    Next iRow
    
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
        CommandBars("Operate Bar").Controls.Item(3).Caption = sHideBarName
        CommandBars("Operate Bar").Controls.Item(3).TooltipText = sHideBarName
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

'获取表格中文名称,参数为CHS中文，ENG英文
Public Sub ReChinese(sTableType As String)
'Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim TableInfoSheet As Worksheet
Dim sheetRefresh As Worksheet
Dim sFiledName As String
Dim sDspName As String
Dim sTableName As String
Dim sheetTableName As String
Dim sFieldName As String
Dim iTemRows As Integer
Dim sPostil1 As String, sPostil2 As String, sPostil As String, sPostilSign As String
Dim iRows As Integer, icount As Integer, iSign  As Integer, iSheetNum  As Integer, iRsCount  As Integer
Dim sVersion As String

icount = 0
iSign = 0

On Error GoTo ErrHandler
 ' If conn.State = adStateOpen Then
 '   conn.Close
 ' End If
  
'conn.Open "Provider=Sybase OLEDB Provider;Persist Security Info=False;Initial Catalog = cmedb;Data Source= 127.0.0.1,5000 ;User Id= sa ;Password= emsems ;"
 'ConnectDatabase = True
 If conn.State = adStateClosed Then
    MsgBox "请先连接数据库。"
     Exit Sub
  End If
Set cmd.ActiveConnection = conn
Set TableInfoSheet = ThisWorkbook.Sheets("TableInfo")

Set sheetRefresh = ThisWorkbook.Sheets("Refresh")

sVersion = Trim(sheetRefresh.Cells(1, 2).Value)
     If sVersion = "" Then
        MsgBox "网元版本编号不能为空"
            Exit Sub
       End If

iRows = 4
' 设置行标

For iSheetNum = 5 To 2000
''''''''对每个Sheet进行编辑
' ThisWorkbook.Sheets(EachSheet(iSheetNum)).Rows(1).Clear
' ThisWorkbook.Sheets(EachSheet(iSheetNum)).Rows(2).Clear
     If Trim(TableInfoSheet.Cells(iSheetNum, 1).Value) = "" Then
                Exit For
       End If
       
sTableName = Trim(TableInfoSheet.Cells(iSheetNum, 1).Value)
sFieldName = Trim(TableInfoSheet.Cells(iSheetNum, 2).Value)

If sTableName = "CELL" And sFieldName = "URAIDS" Then
TableInfoSheet.Cells(iSheetNum, 8) = "URA标识"
End If

If sTableName = "CELL" And sFieldName = "PCPICHPOWER" Then
TableInfoSheet.Cells(iSheetNum, 8) = "PCPICH发射功率"
End If

If sTableName = "CELL" And sFieldName = "TEMPLATENAME" Then
TableInfoSheet.Cells(iSheetNum, 8) = "模板名称"
End If

If sFieldName = "BSCName" Then
TableInfoSheet.Cells(iSheetNum, 8) = "网元名称"
End If

If sFieldName = "CELLNAME" Then
TableInfoSheet.Cells(iSheetNum, 8) = "小区名称"
End If


Application.StatusBar = "正在取得表" + sTableName + "的数据，请稍候……"
cmd.CommandText = "select sDspName,sFieldName from view_FieldAllInfo where sVersion = '" + sVersion + "'and iMode = 2 and iVisible  = 1 and sTableName = '" + sTableName + "'and sFieldName = '" + sFieldName + "'"


rs.CursorLocation = adUseClient
rs.Open cmd
Application.StatusBar = ""
iTemRows = iRows
Do While Not rs.EOF
    iRows = iRows + 1
 '   sheetTableDef.Rows(iRows).Clear
 ' 不清楚当前编辑行
   For icount = iRows To 5000
        If Trim(TableInfoSheet.Cells(icount, 2).Value) = "" Then
                Exit For
           End If
    sFiledName = Trim(TableInfoSheet.Cells(icount, 2).Value)
    sheetTableName = Trim(TableInfoSheet.Cells(icount, 1).Value)
    sDspName = rs("sDspName")
    If sFiledName = rs("sFieldName") And sTableName = sheetTableName Then
       If sTableType = "CHS" Then
          TableInfoSheet.Cells(icount, 8) = sDspName
         End If
       If sTableType = "ENG" Then
          TableInfoSheet.Cells(icount, 7) = sDspName
         End If
     End If
    Next icount
    rs.MoveNext
    Loop
    rs.Close
''''''''对每个Sheet进行编辑
Next iSheetNum

'InsertUserToolBar

MsgBox "OK"
 
 Exit Sub

ErrHandler:
  'ConnectDatabase = False
End Sub

Public Sub UnprotectWorkBook()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=GetSheetsPass
    'ThisWorkbook.Save
End Sub

Public Function GetSheetsPass() As String
    GetSheetsPass = "HWCME"
End Function
'获取表格分支参数信息
Public Sub GetBranchParameterFieldRange()
'Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim BranchCmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim BranchRs As New ADODB.Recordset
Dim HomeSheet As Worksheet
Dim ValidInfoSheet As Worksheet
Dim sheetRefresh As Worksheet
Dim sFiledName As String
Dim sTableName As String
Dim sheetTableName As String
Dim sFieldName As String
Dim iRows As Integer
Dim sVersion As String
Dim icount As Integer
Dim iSheetNum As Integer
Dim sEnumValue As String
Dim sBranchFieldName As String


Set cmd.ActiveConnection = conn
Set BranchCmd.ActiveConnection = conn

Set ValidInfoSheet = ThisWorkbook.Sheets("ValidInfo")
Set HomeSheet = ThisWorkbook.Sheets("Home")
Set sheetRefresh = ThisWorkbook.Sheets("Refresh")
sVersion = Trim(sheetRefresh.Cells(1, 2).Value)

If sVersion = "" Then
     MsgBox "网元版本编号不能为空"
     Exit Sub
End If
Application.StatusBar = "正在取得表" + "sBranchName" + "的数据，请稍候……"

rs.CursorLocation = adUseClient
BranchRs.CursorLocation = adUseClient

iRows = 25
' 设置行标
For iSheetNum = 2 To 5000
If Trim(HomeSheet.Cells(iSheetNum, 1).Value) = "" Then
    Exit For
End If
sTableName = Trim(HomeSheet.Cells(iSheetNum, 1).Value)
cmd.CommandText = "select distinct a.sTableName,a.sBranchFieldName,a.sFieldName ,b.sInput from view_BranchFields a,view_FieldEnum b where a.iMode = b.iMode and a.iMode = 2 and a.iValidFlag = 0 and b.sVersion = a.sVersion and a.sVersion = '" + sVersion + "'and a.sBranchFieldName = b.sFieldName and a.sBranchMinValue = b.iValue and a.sTableName = b.sTableName and a.sTableName = '" + sTableName + "'" + "and not (a.sFieldName = 'REDIRUARFCNUPLINK'or a.sFieldName = 'UARFCNUPLINK') order by a.sTableName,a.sBranchFieldName,a.sFieldName ,b.sInput"
rs.Open cmd

Do While Not rs.EOF
    sBranchFieldName = rs("sBranchFieldName")
    sFiledName = rs("sFieldName")
    
    '查询二级分支参数的字段，这些字段也必须被控制
    BranchCmd.CommandText = "select distinct sFieldName from view_BranchFields  where sTableName= '" + sTableName + "'and sVersion = '" _
     + sVersion + "' and iMode = 2 and iValidFlag = 0  and sBranchFieldName = '" + sFiledName + "'"
    
    ValidInfoSheet.Cells(iRows, 1) = sTableName
    ValidInfoSheet.Cells(iRows, 2) = rs("sBranchFieldName")
    ValidInfoSheet.Cells(iRows, 3) = rs("sFieldName")
    ValidInfoSheet.Cells(iRows, 4) = rs("sInput")
    ValidInfoSheet.Cells(iRows, 7) = "1"
    iRows = iRows + 1
    
    BranchRs.Open BranchCmd
    Do While Not BranchRs.EOF
       sFiledName = BranchRs("sFieldName")
       If sFiledName <> "REDIRUARFCNUPLINK" And sFiledName <> "UARFCNUPLINK" Then
            ValidInfoSheet.Cells(iRows, 1) = sTableName
            ValidInfoSheet.Cells(iRows, 2) = rs("sBranchFieldName")
            ValidInfoSheet.Cells(iRows, 3) = BranchRs("sFieldName")
            ValidInfoSheet.Cells(iRows, 4) = rs("sInput")
            ValidInfoSheet.Cells(iRows, 7) = "2"
            iRows = iRows + 1
         End If
        BranchRs.MoveNext
    Loop
    BranchRs.Close
    rs.MoveNext
    Loop
rs.Close
Next iSheetNum

Application.StatusBar = ""

End Sub
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
        If sTableName = CurSheet.Name Then
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


