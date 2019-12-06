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
  Dim CurRange As Range
  Dim mapCell As Range
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
  
  Do While (Summary_Shet.Cells(TitleRow, iColumn) <> "")
    If (Summary_Shet.Cells(RangeRow, iColumn) <> "") Then
    
        mapRangeStr = Summary_Shet.Cells(RangeRow, iColumn)
        mapRangeArr = Split(mapRangeStr, ",")
        ErrorCells = ""
        
        'Two cells is not equa, error
        For index = 1 To UBound(mapRangeArr)
            For Each mapCell In Site_Shet.Range(mapRangeArr(index))
              If CStr(mapCell.value) <> "" And Site_Shet.Range(mapRangeArr(index - 1)).Cells(1, 1) <> "" _
              And CStr(mapCell.value) <> Site_Shet.Range(mapRangeArr(index - 1)).Cells(1, 1) Then
                    ErrorCells = ErrorCells + C(mapCell.column) + CStr(mapCell.row) + ","
              End If
            Next mapCell
        Next
        
        If ErrorCells <> "" Then
          tmpMsg = Replace(getResByKey("VALUE_NOT_EQUA"), "(%0)", Left(ErrorCells, Len(ErrorCells) - 1))
          tmpMsg = Replace(tmpMsg, "(%1)", Split(mapRangeArr(0), ":")(0))
          tmpMsg = Replace(tmpMsg, "(%3)", siteName)
          ErrMsg = ErrMsg + tmpMsg + vbCrLf
        Else
          Dim Range_array() As String
          Dim Indx As Long
          Range_array = Split(Summary_Shet.Cells(RangeRow, iColumn), ",")
          For Indx = LBound(Range_array) To UBound(Range_array)
            'Refresh data from Site
            For Each CurRange In Site_Shet.Range(Range_array(Indx))
              Summary_Shet.Cells(CurRow, iColumn) = CurRange.Cells(1, 1)
              Exit For
            Next CurRange
          Next
        End If
    End If
    iColumn = iColumn + 1
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
  Dim isfound As Boolean

  For Each iSheet In ThisWorkbook.Sheets
        If iSheet.Tab.colorIndex = BluePrintSheetColor Then
                isfound = False
                 iRow = DataBegin_Row
                Do While (Trim(Summary_Shet.Cells(iRow, 1).value) <> "")
                    If Trim(Summary_Shet.Cells(iRow, 1).value) = iSheet.name Then
                        isfound = True
                        Exit Do
                    End If
                    iRow = iRow + 1
                Loop
                If isfound = fasle Then
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
    Call UpdateOneSiteData(Summary_Shet, Trim(Summary_Shet.Cells(iRow, 1)), iRow)
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
  
If Len(siteName) > 31 _
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
     Or InStr(siteName, "@") > 0 _
     Or InStr(siteName, "#") > 0 _
     Or InStr(siteName, "&") > 0 _
     Or InStr(siteName, "%") > 0 _
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

'**********************************************************
'Copy data from summry
'**********************************************************
Public Function CopyCell(CurRange As Range, Content As String, curSht As Worksheet) As Boolean
  Dim ShtCell As Range
  Dim row As Long
  Dim cnt As Long
  Dim ChkCell As Range

  CopyCell = True

  For Each ShtCell In CurRange   ' Not Refresh Empty Line.
    row = ShtCell.row
    
    If curSht.Cells(row, 1) <> "" Then 'Title Row
       CopyCell = False
       ErrMsg = ErrMsg + Replace(getResByKey("ROW_IS_TITLE"), "(%0)", CStr(row)) + vbCrLf
       Exit Function
    End If
    
    Dim noneEmptyCellsNumber As Long, titleColumnNumber As Long, groupStartRow As Long, groupEndRow As Long
    noneEmptyCellsNumber = Application.WorksheetFunction.CountA(curSht.rows(row))
    
    Call getVerticalGroupName(curSht, row, "A", groupStartRow, groupEndRow)
    titleColumnNumber = Application.WorksheetFunction.CountA(curSht.rows(groupStartRow))
    
    If noneEmptyCellsNumber <> 0 Or (noneEmptyCellsNumber = 0 And titleColumnNumber = 2) Then 'The Row is not Empty 或者当前只有一个参数列时需要刷新
      ShtCell.value = Content
    End If
  Next
  
End Function

Sub GenOneSiteData(Summary_Shet As Worksheet, CurSiteName As String, ReferenceSiteName As String, CurRow As Long)

  'Refresh Data from Summary???ìo?if The site not exits, create it.
  If IsExistsSheet(CurSiteName) = False Then
    If IsExistsSheet(ReferenceSiteName) Then
      ActiveWorkbook.Sheets(ReferenceSiteName).Select
      If CreateNewSheet(CurSiteName) = False Then
        Exit Sub
      End If
    ElseIf ReferenceSiteName <> "" Then
      ErrMsg = ErrMsg + Replace(getResByKey("REFERENCE_SHET_NOT_FOUND"), "(%0)", ReferenceSiteName) + vbCrLf
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
          If CopyCell(SiteShet.Range(Range_array(index)), Summary_Shet.Cells(CurRow, iColumn), SiteShet) = False Then
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
    NodeBEndRow = baseStationSheet.Range("A65535").End(xlUp).row
    
    Dim siteName As String, ucaseSiteName As String
    
    Dim sitesNameCol As New Collection
    For rowNumber = NodeBBeginRow To NodeBEndRow
        siteName = Trim(baseStationSheet.Range("A" & rowNumber).value)
        ucaseSiteName = UCase(siteName)
        If Not Contains(sitesNameCol, ucaseSiteName) Then
            sitesNameCol.Add Item:=ucaseSiteName, key:=ucaseSiteName
        Else
            If bIsEng Then
                MsgBox "The name of [" + siteName + "] is the same as another NE name, please modify the name.", vbCritical
            Else
                MsgBox "存在与[" + siteName + "]同名的网元名称，请修改名称。", vbCritical
            End If
            checkSameSitesName = False
            Exit Function
        End If
    Next rowNumber
End Function

'Generate Site data from Summary
'Generate Site data from Summary
Sub GenSiteSheetFromSummary()
  Dim iRow As Long
  Dim ReferenceSiteName As String
  Dim ReferenceColumn As Long
  Dim GenDefFlag As Boolean
  
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
    ReferenceSiteName = Trim(Summary_Shet.Cells(iRow, ReferenceColumn))
    If ReferenceSiteName = "" And GenDefFlag = True Then
      ReferenceSiteName = DefReShetName
    End If
    Call GenOneSiteData(Summary_Shet, Trim(Summary_Shet.Cells(iRow, 1)), ReferenceSiteName, iRow)
    iRow = iRow + 1
  Loop
  
  If ErrMsg = "" Then
    ErrMsg = getResByKey("GENERATE_SITE_DATA_FINISH")
  End If
  
  Summary_Shet.Activate
  MsgBox ErrMsg
  Call refreshEnd
  
End Sub


