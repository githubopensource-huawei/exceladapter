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
        If (Summary_Shet.Cells(TitleRow, iColumn) = getResByKey("Referenced_Site")) Then GoTo Mark
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

Public Sub UpdateOneSiteData(Summary_Shet As Worksheet, siteName As String, CurRow As Long)
    On Error GoTo ErrorHandler
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
                Next
            End If
        End If
        iColumn = iColumn + 1
        Set firstCellRange = Nothing
        firstCellValue = ""
    Loop
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in UpdateOneSiteData, " & Err.Description
End Sub

'Refresh Summary Data from Site
Public Sub UpdateSummaryFromSiteSheet()
    Dim iRow As Long
    Dim notExistSites As Long
    Dim notExistSitesName As String
    
    Dim iSheet As Worksheet
    
    ErrMsg = ""
    Call refreshStart
    Dim Summary_Shet As Worksheet
    Set Summary_Shet = Sheets(GetMainSheetName)
    Dim isFound As Boolean
    
    Dim siteNameCol As Long
    siteNameCol = siteNameColNum(Summary_Shet)
    For Each iSheet In ThisWorkbook.Sheets
        If iSheet.Tab.colorIndex = BluePrintSheetColor Then
            isFound = False
            iRow = DataBegin_Row
            Do While (Trim(Summary_Shet.Cells(iRow, siteNameCol).value) <> "")
                If Len(Trim(Summary_Shet.Cells(iRow, siteNameCol).value)) <= 31 Then
                    If Trim(Summary_Shet.Cells(iRow, siteNameCol).value) = iSheet.name Then
                        isFound = True
                        Exit Do
                    End If
                Else
                    If Trim(Summary_Shet.Cells(iRow, siteNameCol + 1).value) = iSheet.name Then
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
    
    If notExistSitesName <> "" Then MsgBox Replace(getResByKey("SITES_NOT_EXIST_IN_TRANSPORT"), "(0%)", "[" + notExistSitesName + "]"), vbInformation, getResByKey("Information")
    
    iRow = DataBegin_Row
    Do While (Trim(Summary_Shet.Cells(iRow, siteNameCol)) <> "")
        If Len(Trim(Summary_Shet.Cells(iRow, siteNameCol).value)) <= 31 Then
            Call UpdateOneSiteData(Summary_Shet, Trim(Summary_Shet.Cells(iRow, siteNameCol)), iRow)
        Else
            Call UpdateOneSiteData(Summary_Shet, Trim(Summary_Shet.Cells(iRow, siteNameCol + 1)), iRow)
        End If
        iRow = iRow + 1
    Loop
    
    If ErrMsg = "" Then ErrMsg = getResByKey("REFRESH_SUMMARY_DATA_FINISH")
    
    Summary_Shet.Activate
    Dim msgArray() As String
    msgArray = Split(ErrMsg, vbCrLf)
    If UBound(msgArray) > 3 Then ErrMsg = msgArray(0) + vbCrLf + msgArray(1) + vbCrLf + msgArray(1) + vbCrLf + getResByKey("MSG_TOO_LONG")
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

Function CreateNewSheet(siteName As String, ByRef siteSheet As Worksheet) As Boolean
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
    Set siteSheet = ThisWorkbook.ActiveSheet
    siteSheet.name = siteName
End Function

Function ForNewSiteName(siteName As String, BeginRow As Long) As String
    Dim OldSiteName As String
    Dim OriSiteName As String
    Dim NameNum As Integer
    Dim NameEndString As String
    
    Dim BaseStationShet As Worksheet
    Set BaseStationShet = Sheets(getResByKey("BaseTransPort"))
  
    Dim siteNameCol As Long
    siteNameCol = siteNameColNum(BaseStationShet)
    SiteNameMoreThan = True
  
    Dim index As Long
    Dim endRowNum As Long
    endRowNum = getUsedRowCount(BaseStationShet)
    
    MaxNameNum = 0
    MinNameNum = 9999
  
    If Len(siteName) > 31 Then
        For index = 3 To endRowNum
            OriSiteName = GetCell(BaseStationShet, index, siteNameCol)
            NameEndString = Right(OriSiteName, 5)
        
            If Left(NameEndString, 1) = "~" Then
                NameNum = CInt(Right(NameEndString, 4))
                If NameNum > MaxNameNum Then
                    MaxNameNum = NameNum
                ElseIf NameNum < MinNameNum Then
                    MinNameNum = NameNum
                End If
            End If
        Next index
    
        If GetCell(BaseStationShet, 2, siteNameCol + 1) = getResByKey("SheetNameForSite") Then
            For index = 3 To endRowNum
                OriSiteName = GetCell(BaseStationShet, index, siteNameCol + 1)
                NameEndString = Right(OriSiteName, 5)
        
                If Left(NameEndString, 1) = "~" Then
                    NameNum = CInt(Right(NameEndString, 4))
                    If NameNum > MaxNameNum Then
                        MaxNameNum = NameNum
                    ElseIf NameNum < MinNameNum Then
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
        BaseStationShet.Cells(BeginRow, siteNameCol + 1) = ForNewSiteName
    End If
End Function

Function ForNewSiteNameColumn()
    Dim sheet As Worksheet
    Set sheet = Sheets(getResByKey("BaseTransPort"))
    sheet.Activate
    
    Dim siteNameCol As Long
    siteNameCol = siteNameColNum(sheet)

    With sheet
        range("B" & (siteNameCol + 1)).Select
        Selection.EntireColumn.INSERT
        Cells(2, siteNameCol + 1) = getResByKey("SheetNameForSite")
    End With
End Function

Function ForInsertNewSiteNameRow(groupName As String)
    Dim sheet As Worksheet
    Set sheet = Sheets("MAPPING DEF")
    
    sheet.Activate
    
    With sheet
        rows("2:2").Select
        Selection.EntireRow.INSERT
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
    maxRow = getSheetUsedRows(TransportShet)
    maxCol = TransportShet.range("IV2").End(xlToLeft).column
    
    Dim siteNameCol As Long
    siteNameCol = siteNameColNum(TransportShet)

    'find the row of current sitesheet in transportsheet
    'it can be done before for performance's sake
    For idx = RangeRow + 1 To maxRow
        If TransportShet.Cells(idx, siteNameCol).value = sheetName Then
            sheetRow = idx
            Exit For
        End If
    Next
    
    If sheetRow <= RangeRow Or sheetRow > maxRow Then Exit Function

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

Private Sub clearRows(ByRef iubsheet As Worksheet, ByRef startRow As Long, ByRef endRow As Long)
    Dim maxColumnNumber As Long
    '获取该MOC的最大列数
    maxColumnNumber = iubsheet.range("IV" & startRow).End(xlToLeft).column
    
    Dim dataRange As range, cellRange As range
    Set dataRange = iubsheet.range(iubsheet.Cells(startRow + 1, 2), iubsheet.Cells(endRow, maxColumnNumber))
    
    '遍历所有单元格，清空数据范围
    For Each cellRange In dataRange
        cellRange.ClearContents
        If cellIsGray(cellRange) Then
            Call resetBackgroundStyle(cellRange)
        End If
    Next cellRange
End Sub

Private Sub copyRows(ByRef iubsheet As Worksheet, ByRef originalIubSheet As Worksheet, ByRef startRow As Long, ByRef endRow As Long)
    Dim maxColumnNumber As Long
    
    '获取该MOC的最大列数
    maxColumnNumber = originalIubSheet.range("IV" & startRow).End(xlToLeft).column
    
    Dim dataRange As range, cellRange As range
    Set dataRange = originalIubSheet.range(originalIubSheet.Cells(startRow + 1, 2), originalIubSheet.Cells(endRow, maxColumnNumber))
    
    dataRange.Copy
    iubsheet.Paste Destination:=iubsheet.Cells(startRow + 1, 2)
End Sub

'清空IUB页签的未选定MOC数据行
Private Sub clearIubSheetCertainRowsData(ByRef mocCollection As Collection, ByRef iubsheet As Worksheet)
    Dim startRow As Long, endRow As Long
    Dim mocName As Variant
    For Each mocName In mocCollection
        '获取该MOC的起始行和结束行
        Call iubMocManager.getMocRows(CStr(mocName), startRow, endRow)
        
        '清空iub页签里的指定起始行和结束行
        Call clearRows(iubsheet, startRow, endRow)
    Next mocName
End Sub

'拷贝IUB页签的未选定MOC数据行
Private Sub copyIubSheetCertainRowsData(ByRef mocCollection As Collection, ByRef iubsheet As Worksheet, ByRef originalIubSheet As Worksheet)
    Dim startRow As Long, endRow As Long
    Dim mocName As Variant
    For Each mocName In mocCollection
        '获取该MOC的起始行和结束行
        Call iubMocManager.getMocRows(CStr(mocName), startRow, endRow)
        
        '清空iub页签里的指定起始行和结束行
        'Call clearRows(iubsheet, startRow, endRow)
        
        Call copyRows(iubsheet, originalIubSheet, startRow, endRow)
    Next mocName
End Sub

Private Function CreateNewSiteSheet(ByRef curSiteName As String, ByRef referenceSiteName As String, ByRef unselectedMocCol As Collection) As Boolean
    CreateNewSiteSheet = False
    Dim referencedSiteSheet As Worksheet, CurSheet As Worksheet
    If containsASheet(ThisWorkbook, referenceSiteName, referencedSiteSheet) Then
        referencedSiteSheet.Select
        If CreateNewSheet(curSiteName, CurSheet) = False Then
            Exit Function
        End If
        
        '需要把未选定的MOC数据行清空
        Call clearIubSheetCertainRowsData(unselectedMocCol, CurSheet)
    ElseIf referenceSiteName <> "" Then
        ErrMsg = ErrMsg + Replace(getResByKey("REFERENCE_SHET_NOT_FOUND"), "(%0)", referenceSiteName) + vbCrLf
        Exit Function
    ElseIf referenceSiteName = "" Then
        '此分支能进来是在未填写引用站，并且未选定所提示的站来生成蓝本页签时
        Exit Function
    End If
    
    CreateNewSiteSheet = True
End Function

Private Function createExistingSiteSheet(ByRef siteShet As Worksheet, ByRef curSiteName As String, ByRef referenceSiteName As String, ByRef unselectedMocCol As Collection) As Boolean
    createExistingSiteSheet = False
    Dim referenceSiteSheet As Worksheet, CurSheet As Worksheet
    If referenceSiteName <> "" And curSiteName <> referenceSiteName Then
        '有引用站，并且引用站名称和自己名称不一致，需要先删除当前站，再重新生成当前该站
        If Not containsASheet(ThisWorkbook, referenceSiteName, referenceSiteSheet) Then
            ErrMsg = ErrMsg + Replace(getResByKey("REFERENCE_SHET_NOT_FOUND"), "(%0)", referenceSiteName) + vbCrLf
            Exit Function '没有该引用站页签，则退出，在已经有基站页签，并且填写的引用站不存在时会进来
        End If
        
        'siteShet.name = siteShet.name & "_ori"
        siteShet.name = "7177180_ori"
        
        referenceSiteSheet.Select
        If CreateNewSheet(curSiteName, CurSheet) = False Then
            Exit Function
        End If
        
        '需要把未选定的MOC数据行清空
        'Call clearIubSheetCertainRowsData(unselectedMocCol, curSheet)
        Call copyIubSheetCertainRowsData(unselectedMocCol, CurSheet, siteShet)
        
        Application.DisplayAlerts = False
        siteShet.Delete
    End If
    '其余场景，没有引用站，或引用站名称和自己名称不一致，则不需要重新生成
    
    createExistingSiteSheet = True
End Function

'if the site sheet does not exist, create it.
Sub GenOneSiteData(Summary_Shet As Worksheet, curSiteName As String, referenceSiteName As String, CurRow As Long, ByRef unselectedMocCol As Collection)
    Dim siteNameCol As Long
    siteNameCol = siteNameColNum(Summary_Shet)
    
    Dim siteShet As Worksheet
    If containsASheet(ThisWorkbook, curSiteName, siteShet) = False Then
        '没有基站页，则需要新生成一个基站页签
        If CreateNewSiteSheet(curSiteName, referenceSiteName, unselectedMocCol) = False Then
            If GetCell(Summary_Shet, 2, siteNameCol + 1) = getResByKey("SheetNameForSite") Then Summary_Shet.Cells(CurRow, siteNameCol + 1).value = ""
            Exit Sub
        End If
    Else
        '新加的功能，已经有了基站页签，需要判断这个基站的引用站是否为空，如果为空，则不做操作，如果不为空，则需要以引用站重新生成该基站页
        If createExistingSiteSheet(siteShet, curSiteName, referenceSiteName, unselectedMocCol) = False Then
            If GetCell(Summary_Shet, 2, siteNameCol + 1) = getResByKey("SheetNameForSite") Then Summary_Shet.Cells(CurRow, siteNameCol + 1).value = ""
            Exit Sub
        End If
    End If
    
    If containsASheet(ThisWorkbook, curSiteName, siteShet) Then
        Dim iColumn As Long
        
        iColumn = 1
        Do While (Summary_Shet.Cells(TitleRow, iColumn) <> "")
            If (Summary_Shet.Cells(RangeRow, iColumn) <> "") Then
                Dim Range_array() As String
                Dim index As Long
                Range_array = Split(Summary_Shet.Cells(RangeRow, iColumn), ",")
                For index = LBound(Range_array) To UBound(Range_array)
                    If CopyCell(siteShet.range(Range_array(index)), Summary_Shet.Cells(CurRow, iColumn), siteShet) = False Then Exit Sub
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

Private Function existsIUBSht(Summary_Shet As Worksheet, ByVal rowIdx As Integer, siteName As String) As Boolean
On Error GoTo ErrorHandler
    existsIUBSht = False
    Dim shtName As String
    shtName = siteName
    If Summary_Shet.Cells(TitleRow, 2) = getResByKey("SheetNameForSite") Then
        If Summary_Shet.Cells(rowIdx, 2) <> "" Then shtName = Trim(Summary_Shet.Cells(rowIdx, 2).value)
    End If
    
    existsIUBSht = existsASheet(shtName)
    
    Exit Function
ErrorHandler:
    existsIUBSht = False
    Debug.Print "some exception in existsIUBSht, " & Err.Description
End Function

Function IsReferenceDef(Summary_Shet As Worksheet, referenceColumn As Long) As Boolean
    IsReferenceDef = False
    Dim NoReferSiteNames As String
    Dim iRow As Long
    iRow = DataBegin_Row
    NoReferSiteNames = ""
  
    Dim siteNameCol As Long
    siteNameCol = siteNameColNum(Summary_Shet)
    
    If Trim(Summary_Shet.Cells(iRow, siteNameCol).value) <> "" Then
        Do While (Trim(Summary_Shet.Cells(iRow, siteNameCol).value) <> "")
            If existsIUBSht(Summary_Shet, iRow, Trim(Summary_Shet.Cells(iRow, siteNameCol).value)) = False And Summary_Shet.Cells(iRow, referenceColumn) = "" Then
                NoReferSiteNames = NoReferSiteNames + Trim(Summary_Shet.Cells(iRow, siteNameCol)) + ","
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
    NodeBBeginRow = DataBegin_Row
    
    Dim baseStationSheet As Worksheet
    Set baseStationSheet = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    NodeBEndRow = getSheetUsedRows(baseStationSheet)
    
    Dim siteName As String, ucaseSiteName As String
    Dim NameCol As Long
    NameCol = siteNameColNum(baseStationSheet)
    
    Dim sitesNameCol As New Collection
    For rowNumber = NodeBBeginRow To NodeBEndRow
        siteName = Trim(baseStationSheet.Cells(rowNumber, NameCol).value)
        ucaseSiteName = UCase(siteName)
        If siteName = "" Then GoTo NextLoop
        If Not Contains(sitesNameCol, ucaseSiteName) Then
            sitesNameCol.Add Item:=ucaseSiteName, key:=ucaseSiteName
        Else
            MsgBox getResByKey("Exists") + siteName + getResByKey("SameNameSite"), vbCritical
            checkSameSitesName = False
            Exit Function
        End If
NextLoop:
    Next rowNumber
End Function

Public Function isNeedfilter() As Boolean
    On Error GoTo ErrorHandler
    Dim referenceColumn As Long
    
    referenceColumn = GetReferenceColumn
    If referenceColumn <= 1 Then
        isNeedfilter = False
        Exit Function
    End If
    Dim usedColCount As Long, index As Long
    usedColCount = Sheets(GetMainSheetName()).range("a65536").End(xlUp).row
    
    For index = 4 To usedColCount
        If Len(Trim(Sheets(GetMainSheetName()).Cells(index, referenceColumn).value)) <> 0 Or IsSheetExist(Sheets(GetMainSheetName()).Cells(index, referenceColumn - 1).value) = False Then
            isNeedfilter = True
            Exit Function
        End If
    Next
    isNeedfilter = False
    Exit Function
ErrorHandler:
    isNeedfilter = True
End Function


'Generate Site data from Summary
Sub GenSiteSheetFromSummary()
    On Error GoTo ErrorHandler
    Dim iRow As Long
    Dim referenceSiteName As String
    Dim referenceColumn As Long
    Dim GenDefFlag As Boolean
    Dim NewSiteName As String

    If checkSameSitesName = False Then Exit Sub '检查名称是否相同，大小写不敏感
    
    '获取变量unselectedMocCollection
    Dim unselectedMocCol As Collection
    If isNeedfilter = True Then
        If getIubUnselectMocCollection(unselectedMocCol) = False Then Exit Sub '如果在弹出窗体中选择了取消，则退出，不进行转换
    Else
        Set unselectedMocCol = New Collection
    End If
    
    Call refreshStart
    iRow = DataBegin_Row
    ErrMsg = ""
    GenDefFlag = True
    
    Dim Summary_Shet As Worksheet
    Set Summary_Shet = Sheets(GetMainSheetName)
    referenceColumn = GetReferenceColumn()
    
    Call Process_DefSite
    
    If RefeShetCount > 1 Then GenDefFlag = IsReferenceDef(Summary_Shet, referenceColumn)
    
    If DefReShetName = "" Then ErrMsg = getResByKey("NO_REFERENCE_SITE")
    
    Dim curSiteName As String
    Dim siteNameCol As Long
    siteNameCol = siteNameColNum(Summary_Shet)
    
    Do While (Trim(Summary_Shet.Cells(iRow, siteNameCol)) <> "")
        index = DataBegin_Row
        referenceSiteName = Trim(Summary_Shet.Cells(iRow, referenceColumn))
        If Len(referenceSiteName) > 31 Then
            Do While (Trim(Summary_Shet.Cells(index, siteNameCol)) <> "")
                If Trim(Summary_Shet.Cells(index, siteNameCol)) = referenceSiteName And Trim(Summary_Shet.Cells(index, siteNameCol + 1)) <> "" Then
                    If GetCell(Summary_Shet, 2, siteNameCol + 1) = getResByKey("SheetNameForSite") Then
                        referenceSiteName = Trim(Summary_Shet.Cells(index, siteNameCol + 1))
                        Exit Do
                    End If
                End If
                index = index + 1
            Loop
            
        End If
        
        If referenceSiteName = "" And GenDefFlag = True Then referenceSiteName = DefReShetName
        
        curSiteName = Trim(Summary_Shet.Cells(iRow, siteNameCol).value)
        If Len(curSiteName) > 31 Then
            If GetCell(Summary_Shet, 2, siteNameCol + 1) = getResByKey("SheetNameForSite") Then
                If Summary_Shet.Cells(iRow, siteNameCol + 1) <> "" Then
                    curSiteName = Trim(Summary_Shet.Cells(iRow, siteNameCol + 1))
                    If Len(curSiteName) > 31 Or IsValidSheetName(curSiteName) = False Then
                        ErrMsg = getResByKey("SheetNameLenGrateThan31")
                        MsgBox ErrMsg
                        Call refreshEnd
                        Exit Sub
                    End If
                Else
                    curSiteName = GenNewName(curSiteName, iRow)
                End If
            Else
                If IsExistsSheet(curSiteName) = False Then curSiteName = GenNewName(curSiteName, iRow)
            End If
        End If
        
        Call GenOneSiteData(Summary_Shet, curSiteName, referenceSiteName, iRow, unselectedMocCol)
        iRow = iRow + 1
    Loop
    
    If ErrMsg = "" Then ErrMsg = getResByKey("GENERATE_SITE_DATA_FINISH")
    
    Summary_Shet.Activate
    MsgBox ErrMsg
    Call refreshEnd
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in GenSiteSheetFromSummary " & Err.Description
End Sub

Function GenNewName(siteName As String, BeginRow As Long) As String
    On Error GoTo ErrorHandler
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
  
    Dim siteNameCol As Long
    siteNameCol = siteNameColNum(BaseStationShet)
    ShtIdx = 1
    If Len(siteName) > 31 Then
        Do While (ShtIdx <= ActiveWorkbook.Sheets.count)
            Set OpSht = ActiveWorkbook.Sheets(ShtIdx)
            NameEndString = Right(OpSht.name, 5)
            If Left(NameEndString, 1) = "~" Then
                NameNum = CInt(Right(NameEndString, 4))
                If NameNum > MaxNameNum Then
                  MaxNameNum = NameNum
                ElseIf NameNum < MinNameNum Then
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
        
        If GetCell(BaseStationShet, 2, siteNameCol + 1) <> getResByKey("SheetNameForSite") Then Call ForNewSiteNameColumn
        BaseStationShet.Cells(BeginRow, siteNameCol + 1) = GenNewName
    End If
    Exit Function
ErrorHandler:
    Debug.Print "some exception in GenNewName " & Err.Description
End Function
