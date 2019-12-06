Attribute VB_Name = "UpdateGenSummaryData"
Const RangeRow = 3
Const TitleRow = 2
Const DataBegin_Row = 4
Dim ErrMsg As String
Dim DefReShetName As String
Dim RefeShetCount As String

Function GetReferenceColumn() As Long
  Dim Summary_Shet As Worksheet
  Set Summary_Shet = Sheets(GetMainSheetName)
  
  Dim iColumn As Long
  iColumn = 1
  
  Do While (Summary_Shet.Cells(TitleRow, iColumn) <> "")
    If (Summary_Shet.Cells(TitleRow, iColumn) = getResByKey("Referenced_Site")) Then
      GoTo Mark
    End If
    iColumn = iColumn + 1
  Loop
  
Mark:
  GetReferenceColumn = iColumn
End Function

Function IsExistsSheet(sheetName As String) As Boolean
  Dim ShtIdx As Long
  Dim OpSht As Worksheet
  
  ShtIdx = 1
  Do While (ShtIdx <= ActiveWorkbook.Sheets.count)
      Set OpSht = ActiveWorkbook.Sheets(ShtIdx)
      If OpSht.name = sheetName Then
        IsExistsSheet = True
        Exit Function
      End If
      ShtIdx = ShtIdx + 1
  Loop
  IsExistsSheet = False
End Function

Sub UpdateOneSiteData(Summary_Shet As Worksheet, siteName As String, CurRow As Long)
  Dim Site_Shet As Worksheet
  Dim iColumn As Long
  Dim curRange As range
  Dim mapCell As range
  Dim mapRangeStr As String
  Dim mapRangeArr() As String
  Dim ErrorCells As String
  Dim index As Long
  Dim tmpMsg As String
  
  If IsExistsSheet(siteName) = False Then
    Exit Sub
  End If
  
  Set Site_Shet = Sheets(siteName)
  iColumn = 1

  Dim firstCellValue As String
  Dim firstCellRange As range
  
  Do While (Summary_Shet.Cells(TitleRow, iColumn) <> "")
    If (Summary_Shet.Cells(RangeRow, iColumn) <> "") Then
    
        mapRangeStr = Summary_Shet.Cells(RangeRow, iColumn)
        mapRangeArr = Split(mapRangeStr, ",")
        ErrorCells = ""
        
        'Two cells is not equa, error
        For index = 0 To UBound(mapRangeArr)
            '设置上第一个有值的单元格和该单元格值
            If Site_Shet.range(mapRangeArr(index)).value <> "" And firstCellRange Is Nothing Then
                Set firstCellRange = Site_Shet.range(mapRangeArr(index))
                firstCellValue = firstCellRange.value
            End If
            For Each mapCell In Site_Shet.range(mapRangeArr(index))
'              If CStr(mapCell.value) <> "" And Site_Shet.Range(mapRangeArr(index - 1)).Cells(1, 1) <> "" _
'              And CStr(mapCell.value) <> Site_Shet.Range(mapRangeArr(index - 1)).Cells(1, 1) Then
'                    ErrorCells = ErrorCells + C(mapCell.column) + CStr(mapCell.row) + ","
'              End If
                '之前代码无法处理3,"",4,4这种情况，因此改为如下判断
                If mapCell.value <> "" And mapCell.value <> firstCellValue Then
                    ErrorCells = ErrorCells + C(mapCell.column) + CStr(mapCell.row) + ","
                End If
            Next mapCell
        Next
        
        If ErrorCells <> "" Then
          tmpMsg = Replace(getResByKey("VALUE_NOT_EQUA"), "(%0)", Left(ErrorCells, Len(ErrorCells) - 1))
          tmpMsg = Replace(tmpMsg, "(%1)", firstCellRange.address(False, False))
          tmpMsg = Replace(tmpMsg, "(%3)", siteName)
          ErrMsg = ErrMsg + tmpMsg + vbCrLf
        Else
          Dim Range_array() As String
          Dim Indx As Long
          Dim cellRange As range
          Range_array = Split(Summary_Shet.Cells(RangeRow, iColumn), ",")
          For Indx = LBound(Range_array) To UBound(Range_array)
            'Refresh data from Site
            Set cellRange = Site_Shet.range(Range_array(Indx))
            '增加新的判断，基站页签值不为空，才刷新到传输页
            With cellRange
                If .value <> "" Then Summary_Shet.Cells(CurRow, iColumn).value = .value
            End With
'            For Each CurRange In Site_Shet.Range(Range_array(Indx))
'              Summary_Shet.Cells(CurRow, iColumn) = CurRange.Cells(1, 1)
'              Exit For
'            Next CurRange
          Next
        End If
    End If
    iColumn = iColumn + 1
    Set firstCellRange = Nothing
    firstCellValue = ""
  Loop
  
End Sub

'Refresh Summary Data from Site
Sub UpdateSummaryFromSiteSheet()
  Dim iRow As Long
  Dim notExistSites As Long
  Dim notExistSitesName As String
  
  Dim iSheet As Worksheet
  
  ErrMsg = ""
  Call refreshStart
  Dim Summary_Shet As Worksheet
  Set Summary_Shet = Sheets(GetMainSheetName)
  Dim isFound As Boolean

  For Each iSheet In ThisWorkbook.Sheets
        If iSheet.Tab.colorIndex = BluePrintSheetColor Then
                isFound = False
                 iRow = DataBegin_Row
                Do While (Trim(Summary_Shet.Cells(iRow, 1).value) <> "")
                  If Len(Trim(Summary_Shet.Cells(iRow, 1).value)) <= 31 Then
                    If Trim(Summary_Shet.Cells(iRow, 1).value) = iSheet.name Then
                        isFound = True
                        Exit Do
                    End If
                  Else
                      If Trim(Summary_Shet.Cells(iRow, 2).value) = iSheet.name Then
                        isFound = True
                        Exit Do
                      End If
                  End If
                    iRow = iRow + 1
                Loop
                If isFound = fasle Then
                    If notExistSites = 0 Then
                        notExistSitesName = iSheet.name
                    ElseIf notExistSites = 1 Then
                        notExistSitesName = notExistSitesName + "," + iSheet.name
                    ElseIf notExistSites = 2 Then
                        notExistSitesName = notExistSitesName + "..."
                    Else
                        Exit For
                    End If
                    notExistSites = notExistSites + 1
                End If
        End If
  Next
  
  If notExistSitesName <> "" Then
     MsgBox Replace(getResByKey("SITES_NOT_EXIST_IN_TRANSPORT"), "(0%)", "[" + notExistSitesName + "]"), vbInformation, getResByKey("Information")
  End If
  
  iRow = DataBegin_Row
  Do While (Trim(Summary_Shet.Cells(iRow, 1)) <> "")
    If Len(Trim(Summary_Shet.Cells(iRow, 1).value)) <= 31 Then
    Call UpdateOneSiteData(Summary_Shet, Trim(Summary_Shet.Cells(iRow, 1)), iRow)
    Else
      Call UpdateOneSiteData(Summary_Shet, Trim(Summary_Shet.Cells(iRow, 2)), iRow)
    End If
    iRow = iRow + 1
  Loop
  
  If ErrMsg = "" Then
    ErrMsg = getResByKey("REFRESH_SUMMARY_DATA_FINISH")
  End If
  
  Summary_Shet.Activate
  Dim msgArray() As String
  msgArray = Split(ErrMsg, vbCrLf)
  If UBound(msgArray) > 3 Then
     ErrMsg = msgArray(0) + vbCrLf + msgArray(1) + vbCrLf + msgArray(1) + vbCrLf + getResByKey("MSG_TOO_LONG")
  End If
  MsgBox ErrMsg, vbInformation, getResByKey("Information")
  Call refreshEnd

End Sub
Function IsValidSheetName(siteName As String) As Boolean
  IsValidSheetName = True
  
If Len(siteName) > 64 _
     Or InStr(siteName, "\") > 0 _
     Or InStr(siteName, "/") > 0 _
     Or InStr(siteName, ":") > 0 _
     Or InStr(siteName, "*") > 0 _
     Or InStr(siteName, "?") > 0 _
     Or InStr(siteName, Chr(34)) > 0 _
     Or InStr(siteName, "<") > 0 _
     Or InStr(siteName, ">") > 0 _
     Or InStr(siteName, "|") > 0 _
     Or InStr(siteName, ",") > 0 _
     Or InStr(siteName, ";") > 0 _
     Or InStr(siteName, "=") > 0 _
     Or InStr(siteName, "!") > 0 _
     Or InStr(siteName, "^") > 0 _
     Or InStr(siteName, "[") > 0 _
     Or InStr(siteName, "]") > 0 _
     Or InStr(siteName, "  ") > 0 _
     Or InStr(siteName, "+++") > 0 _
     Or Trim(siteName) = "" Then
     ErrMsg = ErrMsg + Replace(getResByKey("SITE_NAME_INVALID"), "(%0)", siteName) + vbCrLf
     IsValidSheetName = False
     
  End If

End Function


Function CreateNewSheet(siteName As String) As Boolean
  CreateNewSheet = True
  
  If IsValidSheetName(siteName) = False Then
    CreateNewSheet = False
    Exit Function
  End If
  
  If IsSystemSheet(ThisWorkbook.ActiveSheet) Then
    ErrMsg = ErrMsg + Replace(getResByKey("SHET_CAN_NOT_COPY"), "(%0)", ThisWorkbook.ActiveSheet.name) + vbCrLf
    CreateNewSheet = False
    Exit Function
  End If
  
  ThisWorkbook.ActiveSheet.Copy after:=ThisWorkbook.ActiveSheet
  ThisWorkbook.ActiveSheet.name = siteName
  
End Function

Function ForNewSiteName(siteName As String, BeginRow As Long) As String
  Dim OldSiteName As String
  Dim OriSiteName As String
  Dim NameNum As Integer
  Dim NameEndString As String
  
  Dim BaseStationShet As Worksheet
  Set BaseStationShet = Sheets(getResByKey("BaseTransPort"))
  
  SiteNameMoreThan = True
  
  Dim index As Long
  Dim endRowNum As Long
  'endRowNum = BaseStationShet.UsedRange.rows.count
  endRowNum = getUsedRowCount(BaseStationShet)
  MaxNameNum = 0
  MinNameNum = 9999
  
  If Len(siteName) > 31 Then
    For index = 3 To endRowNum
      OriSiteName = GetCell(BaseStationShet, index, 1)
      NameEndString = Right(OriSiteName, 5)
    
      If Left(NameEndString, 1) = "~" Then
        NameNum = CInt(Right(NameEndString, 4))
        If NameNum > MaxNameNum Then
          MaxNameNum = NameNum
        End If
        If NameNum < MinNameNum Then
          MinNameNum = NameNum
        End If
      End If
    Next index
    
    If GetCell(BaseStationShet, 2, 2) = getResByKey("SheetNameForSite") Then
      For index = 3 To endRowNum
        OriSiteName = GetCell(BaseStationShet, index, 2)
        NameEndString = Right(OriSiteName, 5)
    
        If Left(NameEndString, 1) = "~" Then
          NameNum = CInt(Right(NameEndString, 4))
          If NameNum > MaxNameNum Then
            MaxNameNum = NameNum
          End If
          If NameNum < MinNameNum Then
            MinNameNum = NameNum
          End If
        End If
      Next index
    End If
    
  
    If MaxNameNum < 9999 Then
      NameEndString = Trim(str(MaxNameNum + 1))
    ElseIf MinNameNum > 1 Then
      NameEndString = Trim(str(MinNameNum - 1))
    End If

    For index = 1 To (4 - Len(Trim(str(NameEndString))))
      NameEndString = "0" + NameEndString
    Next index
  
    NameEndString = "~" + NameEndString
    ForNewSiteName = Left(siteName, 26) + NameEndString
    
    If GetCell(BaseStationShet, 2, 2) <> getResByKey("SheetNameForSite") Then
      Call ForNewSiteNameColumn
      Call ForInsertNewSiteNameRow(GetCell(BaseStationShet, 1, 1))
    End If
    BaseStationShet.Cells(BeginRow, 2) = ForNewSiteName
  End If
End Function
Function ForNewSiteNameColumn()
    Dim sheet As Worksheet
    Set sheet = Sheets(getResByKey("BaseTransPort"))
    sheet.Activate

    With sheet
        range("B2").Select
        Selection.EntireColumn.Insert
        Cells(2, 2) = getResByKey("SheetNameForSite")
    End With
End Function
Function ForInsertNewSiteNameRow(groupName As String)
    Dim sheet As Worksheet
    Set sheet = Sheets("MAPPING DEF")
    
    sheet.Activate
    
    With sheet
        rows("2:2").Select
        'Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.EntireRow.Insert
        Cells(2, 1) = getResByKey("BaseTransPort")
        Cells(2, 2) = groupName
        Cells(2, 3) = getResByKey("SheetNameForSite")
        Cells(2, 4) = "NODE"
        Cells(2, 5) = "SHEETNAMEFORSITE"
        Cells(2, 12) = Cells(3, 12)
        Cells(2, 13) = Cells(3, 13)
        Cells(2, readOnlyColNumInMappingDef) = True
    End With
End Function

'**********************************************************
'Copy data from summry
'**********************************************************
Public Function CopyCell(curRange As range, Content As String, curSht As Worksheet) As Boolean
  Dim ShtCell As range
  Dim row As Long
  Dim cnt As Long
  Dim ChkCell As range

  CopyCell = True

  For Each ShtCell In curRange   ' Not Refresh Empty Line.
    row = ShtCell.row
    
    If curSht.Cells(row, 1) <> "" Then 'Title Row
       CopyCell = False
       ErrMsg = ErrMsg + Replace(getResByKey("ROW_IS_TITLE"), "(%0)", CStr(row)) + vbCrLf
       Exit Function
    End If
    
    Dim titleColumnNumber As Long, groupStartRow As Long, groupEndRow As Long
    'noneEmptyCellsNumber = Application.WorksheetFunction.CountA(curSht.rows(row))

    Call getVerticalGroupName(curSht, row, "A", groupStartRow, groupEndRow)
    titleColumnNumber = Application.WorksheetFunction.CountA(curSht.rows(groupStartRow))
    
    Dim rowIsBlankFlag As Boolean
    rowIsBlankFlag = rowIsBlank(curSht, row)
    
    'The Row is not Empty 或者 当前只有一个参数列 或者 是支持操作符的表格且为ADD 时需要刷新
    If rowIsBlankFlag = False Or titleColumnNumber = 2 Or isOperationAdd(curSht, groupStartRow, row) Then
      ShtCell.value = Content
    End If
  Next
  
End Function

'Public Function isOperator(ByRef curSht As Worksheet, ByVal groupStartRow As Long, ByRef curRange As range) As Boolean
'    isOperator = False
'
'    If curSht.Cells(groupStartRow, 2) = getResByKey("OPERATION") And curRange.column = 2 Then
'        isOperator = True
'        Exit Function
'    End If
'End Function

Public Function isOperationAdd(ByRef curSht As Worksheet, ByVal groupStartRow As Long, ByVal row As Integer) As Boolean
    isOperationAdd = False
    
    'workbook without operation
    If curSht.Cells(groupStartRow, 2) <> getResByKey("OPERATION") Then
        Exit Function
    End If
    
    'workbook with operation, but operation has not been referenced
    If curSht.range("B" & row).Interior.colorIndex <> HyperLinkColorIndex Then
        If curSht.Cells(row, 2) = "ADD" Then
            isOperationAdd = True
            Exit Function
        End If
        Exit Function 'isOperationAdd = False
    End If
    
    'workbook with operation and operation has been referenced
    Dim sheetName As String
    Dim sheetRow As Integer
    sheetName = curSht.name
    
    Dim TransportShet As Worksheet
    Set TransportShet = Sheets(GetMainSheetName)
    
    Dim maxRow, maxCol, idx As Integer
    maxRow = TransportShet.range("a1048576").End(xlUp).row
    maxCol = TransportShet.range("XFD2").End(xlToLeft).column

    'find the row of current sitesheet in transportsheet
    'it can be done before for performance's sake
    For idx = RangeRow + 1 To maxRow
        If TransportShet.Cells(idx, 1).value = sheetName Then
            sheetRow = idx
            Exit For
        End If
    Next
    
    If sheetRow <= RangeRow Or sheetRow > maxRow Then
        Exit Function
    End If

    For idx = 1 To maxCol
        Dim refRange As String
        refRange = TransportShet.Cells(RangeRow, idx).value
        If refRange <> "" Then
            If InStr(refRange, "B" & row) <> 0 And TransportShet.Cells(sheetRow, idx).value = "ADD" Then
                isOperationAdd = True
                Exit Function
            End If
        End If
    Next
End Function

Sub GenOneSiteData(Summary_Shet As Worksheet, CurSiteName As String, ReferenceSiteName As String, CurRow As Long)

  'Refresh Data from Summary???ìo?if The site not exits, create it.
  If IsExistsSheet(CurSiteName) = False Then
    If IsExistsSheet(ReferenceSiteName) Then
      ActiveWorkbook.Sheets(ReferenceSiteName).Select
      If CreateNewSheet(CurSiteName) = False Then
        If GetCell(Summary_Shet, 2, 2) = getResByKey("SheetNameForSite") Then
            Summary_Shet.Cells(CurRow, 2) = ""
        End If
        Exit Sub
      End If
    ElseIf ReferenceSiteName <> "" Then
      ErrMsg = ErrMsg + Replace(getResByKey("REFERENCE_SHET_NOT_FOUND"), "(%0)", ReferenceSiteName) + vbCrLf
      If GetCell(Summary_Shet, 2, 2) = getResByKey("SheetNameForSite") Then
         Summary_Shet.Cells(CurRow, 2) = ""
      End If
      Exit Sub
    End If
    
  End If
  
  If IsExistsSheet(CurSiteName) Then
    'Refresh data from Summary
    Dim SiteShet As Worksheet
    Set SiteShet = Sheets(CurSiteName)
    Dim iColumn As Long
  
    iColumn = 1
  
    Do While (Summary_Shet.Cells(TitleRow, iColumn) <> "")
      If (Summary_Shet.Cells(RangeRow, iColumn) <> "") Then
        
        Dim Range_array() As String
        Dim index As Long
        Range_array = Split(Summary_Shet.Cells(RangeRow, iColumn), ",")
        For index = LBound(Range_array) To UBound(Range_array)
          If CopyCell(SiteShet.range(Range_array(index)), Summary_Shet.Cells(CurRow, iColumn), SiteShet) = False Then
            Exit Sub
          End If
        Next
      End If
      iColumn = iColumn + 1
    Loop
  End If

End Sub
Sub Process_DefSite()
  Dim flag As Long
  Dim ShtIdx As Long
  Dim OpSht As Worksheet
  
  flag = 0
  RefeShetCount = 0
  ShtIdx = 1
  DefReShetName = ""
  
  Do While (ShtIdx <= ActiveWorkbook.Sheets.count)
      Set OpSht = ActiveWorkbook.Sheets(ShtIdx)
      If IsBluePrintSheetName(OpSht.name) Then
         If flag = 0 Then
           DefReShetName = OpSht.name
           flag = 1
         End If
         RefeShetCount = RefeShetCount + 1
      End If
      ShtIdx = ShtIdx + 1
  Loop
End Sub
Function IsReferenceDef(Summary_Shet As Worksheet, ReferenceColumn As Long) As Boolean
  IsReferenceDef = False
  Dim NoReferSiteNames As String
  Dim iRow As Long
  iRow = DataBegin_Row
  NoReferSiteNames = ""
  
  If Trim(Summary_Shet.Cells(iRow, 1).value) <> "" Then
     Do While (Trim(Summary_Shet.Cells(iRow, 1).value) <> "")
       If IsExistsSheet(Trim(Summary_Shet.Cells(iRow, 1).value)) = False And Summary_Shet.Cells(iRow, ReferenceColumn) = "" Then
         NoReferSiteNames = NoReferSiteNames + Trim(Summary_Shet.Cells(iRow, 1)) + ","
       End If
       iRow = iRow + 1
     Loop
  End If
  
  If NoReferSiteNames <> "" Then
    NoReferSiteNames = Left(NoReferSiteNames, Len(NoReferSiteNames) - 1)
    ErrMsg = ErrMsg + Replace(getResByKey("REFERENCE_SHET_EMPTY"), "(%0)", NoReferSiteNames)
    ErrMsg = Replace(ErrMsg, "(%1)", DefReShetName)
  
    If MsgBox(ErrMsg, vbYesNo) = vbYes Then
      IsReferenceDef = True
    End If
    
    ErrMsg = ""
  End If
  
End Function

Private Function checkSameSitesName() As Boolean
    checkSameSitesName = True
    Call InitTemplateVersion
    
    Dim NodeBBeginRow As Long, NodeBEndRow As Long
    NodeBBeginRow = g_NodeBBeginRow + 1
    
    Dim baseStationSheet As Worksheet
    Set baseStationSheet = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    NodeBEndRow = baseStationSheet.range("A1048576").End(xlUp).row
    
    Dim siteName As String, ucaseSiteName As String
    
    Dim sitesNameCol As New Collection
    For rowNumber = NodeBBeginRow To NodeBEndRow
        siteName = Trim(baseStationSheet.range("A" & rowNumber).value)
        ucaseSiteName = UCase(siteName)
        If siteName = "" Then GoTo NextLoop
        If Not Contains(sitesNameCol, ucaseSiteName) Then
            sitesNameCol.Add Item:=ucaseSiteName, key:=ucaseSiteName
        Else
            MsgBox getResByKey("EXIST_[") + siteName + getResByKey("]_SAME_BS_NAME_PLEASE_MOD"), vbCritical
            checkSameSitesName = False
            Exit Function
        End If
NextLoop:
    Next rowNumber
End Function

'Generate Site data from Summary
'Generate Site data from Summary
Sub GenSiteSheetFromSummary()
  Dim iRow As Long
  Dim ReferenceSiteName As String
  Dim ReferenceColumn As Long
  Dim GenDefFlag As Boolean
  Dim NewSiteName As String
  
  If checkSameSitesName = False Then Exit Sub '检查名称是否相同，大小写不敏感
  
  Call refreshStart
  iRow = DataBegin_Row
  ErrMsg = ""
  GenDefFlag = True
  
  Dim Summary_Shet As Worksheet
  Set Summary_Shet = Sheets(GetMainSheetName)
  ReferenceColumn = GetReferenceColumn()
  
  Call Process_DefSite
  
  If RefeShetCount > 1 Then
    GenDefFlag = IsReferenceDef(Summary_Shet, ReferenceColumn)
  End If
  
  If DefReShetName = "" Then
    ErrMsg = getResByKey("NO_REFERENCE_SITE")
  
  End If

  Do While (Trim(Summary_Shet.Cells(iRow, 1)) <> "")
    index = DataBegin_Row
    ReferenceSiteName = Trim(Summary_Shet.Cells(iRow, ReferenceColumn))
    If Len(ReferenceSiteName) > 31 Then
      Do While (Trim(Summary_Shet.Cells(index, 1)) <> "")
        If Trim(Summary_Shet.Cells(index, 1)) = ReferenceSiteName And Trim(Summary_Shet.Cells(index, 2)) <> "" Then
          If GetCell(Summary_Shet, 2, 2) = getResByKey("SheetNameForSite") Then
            ReferenceSiteName = Trim(Summary_Shet.Cells(index, 2))
            Exit Do
          End If
        End If
        index = index + 1
      Loop
    
    End If
    
    If ReferenceSiteName = "" And GenDefFlag = True Then
      ReferenceSiteName = DefReShetName
    End If
    
    NewSiteName = Trim(Summary_Shet.Cells(iRow, 1))
    If Len(NewSiteName) > 31 Then
      If GetCell(Summary_Shet, 2, 2) = getResByKey("SheetNameForSite") Then
        If Summary_Shet.Cells(iRow, 2) <> "" Then
          NewSiteName = Trim(Summary_Shet.Cells(iRow, 2))
          If Len(NewSiteName) > 31 Then
                ErrMsg = getResByKey("SheetNameLenGrateThan31")
                MsgBox ErrMsg
                Call refreshEnd
                Exit Sub
          End If
          If IsValidSheetName(NewSiteName) = False Then
            MsgBox ErrMsg
            Call refreshEnd
            Exit Sub
          End If
        Else
          NewSiteName = GenNewName(NewSiteName, iRow)
        End If
      Else
        If IsExistsSheet(NewSiteName) = False Then
          NewSiteName = GenNewName(NewSiteName, iRow)
        End If
      End If
    End If
    
    Call GenOneSiteData(Summary_Shet, NewSiteName, ReferenceSiteName, iRow)
    iRow = iRow + 1
  Loop
  
  If ErrMsg = "" Then
    ErrMsg = getResByKey("GENERATE_SITE_DATA_FINISH")
  End If
  
  Summary_Shet.Activate
  MsgBox ErrMsg
  Call refreshEnd
  
End Sub
Function GenNewName(siteName As String, BeginRow As Long) As String
  Dim OldSiteName As String
  Dim OriSiteName As String
  Dim NameNum As Integer
  Dim index As Integer
  Dim NameEndString As String
  
  Dim BaseStationShet As Worksheet
  Set BaseStationShet = Sheets(getResByKey("BaseTransPort"))
  
  MaxNameNum = 0
  MinNameNum = 9999
  
  Dim ShtIdx As Long
  Dim OpSht As Worksheet
  
  ShtIdx = 1
  If Len(siteName) > 31 Then
    Do While (ShtIdx <= ActiveWorkbook.Sheets.count)
      Set OpSht = ActiveWorkbook.Sheets(ShtIdx)
      NameEndString = Right(OpSht.name, 5)
      If Left(NameEndString, 1) = "~" Then
        NameNum = CInt(Right(NameEndString, 4))
        If NameNum > MaxNameNum Then
          MaxNameNum = NameNum
        End If
        If NameNum < MinNameNum Then
          MinNameNum = NameNum
        End If
      End If
      ShtIdx = ShtIdx + 1
    Loop
  
    If MaxNameNum < 9999 Then
      NameEndString = Trim(str(MaxNameNum + 1))
    ElseIf MinNameNum > 1 Then
      NameEndString = Trim(str(MinNameNum - 1))
    End If

    For index = 1 To (4 - Len(Trim(str(NameEndString))))
      NameEndString = "0" + NameEndString
    Next index
  
    NameEndString = "~" + NameEndString
    GenNewName = Left(siteName, 26) + NameEndString
    
    If GetCell(BaseStationShet, 2, 2) <> getResByKey("SheetNameForSite") Then
      Call ForNewSiteNameColumn
    End If
    BaseStationShet.Cells(BeginRow, 2) = GenNewName
  End If
End Function
