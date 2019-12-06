VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMocViewManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const m_maxSiteCount = 200
Private Const m_listShtGrpRow = 1
Private Const m_listShtAttrRow = 2
Private Const m_maxSiteNameLen = 63
Private Const m_maxShtNameLen = 31
Private Const m_mocStartRowColName = "StartRow"
Private Const m_mocEndRowColName = "EndRow"
Private Const m_OprationCol = 2

Private m_siteNameShtNameMap As CMap  'key=sitename, value=shtname

Private m_neType As String
Private m_baseSht As Worksheet
Private m_longNameIdx As Integer
Private m_shtName4SiteCol As Integer


Private Sub Class_Initialize()
    Application.ScreenUpdating = False
    m_longNameIdx = 1
End Sub

Private Sub Class_Terminate()
    Set m_siteNameShtNameMap = Nothing
    Application.ScreenUpdating = True
End Sub


''=============================================================================================================================================
''interface called by MocViewManager instance
''baseSht : Basestation Transport Data
''siteNames : sites to be transferred, for MRAT is node name
''targetMocNames : acctually is a sheet name collection, which will be transferred to IUB sheet
''ignoreShtNames : sheets those will not transferred to IUB sheet, eg. summary cell sheets, sheets without site(controller) names
''prefix/posfix : append to site names(IUB sheet name), eg. sitename="abc", IUB sheetName="abc(NE Managerment)"
''fileType : Summary, Bulk, Expand ....
''neType : NeType_MRAT, NeType_GSM, NeType_UMTS, NeType_LTE, NeType_USU, NeType_ICS
''=============================================================================================================================================
Public Function genMocView( _
    ByRef baseSht As Worksheet, _
    ByRef siteNames As Collection, _
    ByRef targetMocNames As Collection, _
    Optional neType As String = NeType_MRAT, _
    Optional prefixOfIUBShtName As String = "", _
    Optional postfixOfIUBShtName As String = "", _
    Optional ignoreShtNames As Collection) As Boolean
    
On Error GoTo ErrorHandler
    
    genMocView = True
    
    If Not checkInput(baseSht, siteNames, targetMocNames, prefixOfIUBShtName, postfixOfIUBShtName) Then Exit Function
    
    Call preProcessCtrlMocShts(targetMocNames, siteNames)
    
    Call initClassResources(baseSht, neType)
    
    Dim templateShtName As String
    templateShtName = createTemplateSiteSht(siteNames.Item(1), prefixOfIUBShtName, postfixOfIUBShtName)
    
    Dim mocMaxCountPerSite As New CMap 'key=mocname, value=maxCount
    Dim mocNameBeginRowMap As New CMap 'key=mocName, value=begin row in IUB sheet
    Call setMocPositionsInTemplateSht(templateShtName, targetMocNames, ignoreShtNames, mocMaxCountPerSite, mocNameBeginRowMap)
    
    Call setMocPositionsInShtDef(targetMocNames, mocMaxCountPerSite)

    Call createAllSiteSht(templateShtName, siteNames, prefixOfIUBShtName, postfixOfIUBShtName)
    
    If Not copyData(targetMocNames, mocNameBeginRowMap) Then
        genMocView = False
        Exit Function
    End If
    
    Call fixAction
    
    Call releaseClassResources
    
    Exit Function
    
ErrorHandler:
    Debug.Print "some exceptio in genMocView, " & Err.Description
    Call releaseClassResources
    genMocView = False
End Function

Private Function checkInput(ByRef baseSht As Worksheet, ByRef siteNames As Collection, ByRef targetMocs As Collection, _
    prefixOfIUBShtName As String, postfixOfIUBShtName As String) As Boolean
    
On Error GoTo ErrorHandle
    checkInput = True
    '(prefix)xxxx~0001(postfix)
    If baseSht Is Nothing _
       Or siteNames.count = 0 _
       Or targetMocs.count = 0 _
       Or Len(prefixOfIUBShtName) + Len(postfixOfIUBShtName) > 26 _
       Or Not validateSiteName(siteNames, prefixOfIUBShtName, postfixOfIUBShtName) _
       Then GoTo ErrorHandle
       
     Exit Function
ErrorHandle:
    checkInput = False
End Function

Private Sub initClassResources(baseSht As Worksheet, neType As String)
    Set m_siteNameShtNameMap = Nothing
    Set m_siteNameShtNameMap = New CMap

    m_neType = neType
    Set m_baseSht = baseSht
End Sub

Private Sub releaseClassResources()
    If Not m_siteNameShtNameMap Is Nothing Then Set m_siteNameShtNameMap = Nothing
End Sub

Private Function createTemplateSiteSht(siteName As String, prefixOfIUBShtName As String, postfixOfIUBShtName As String) As String
    On Error GoTo ErrorHandler
    createTemplateSiteSht = prefixOfIUBShtName & siteName & postfixOfIUBShtName
    
    If Len(createTemplateSiteSht) > 31 Then
        createTemplateSiteSht = shtName4LongSiteName(siteName, prefixOfIUBShtName, postfixOfIUBShtName)
        m_shtName4SiteCol = genShtNameForSiteCol
        Call setShtNameForSite(siteName, createTemplateSiteSht, m_shtName4SiteCol)
    End If
    
    m_baseSht.Activate
    Sheets.Add after:=ThisWorkbook.ActiveSheet
    With ActiveSheet
        .name = createTemplateSiteSht
        .Tab.colorIndex = BluePrintSheetColor
        With .Cells.Interior
            .colorIndex = 15
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
        End With
    End With

    Call m_siteNameShtNameMap.SetAt(siteName, createTemplateSiteSht)
    Exit Function
ErrorHandler:
    Debug.Print "some exception in createTemplateSiteSht, " & Err.Description
End Function

'mocMaxCountPerSite : key=mocname, value=maxCount
'mocNameBeginRowMap : key=mocName, value=begin row in IUB sheet
Private Sub setMocPositionsInTemplateSht(templateShtName As String, targetMocs As Collection, ByRef ignoreShtNames As Collection, _
    ByRef mocMaxCountPerSite As CMap, ByRef mocNameBeginRowMap As CMap)
    On Error GoTo ErrorHandler
    Dim moc
    For Each moc In targetMocs
        Dim mocSht As Worksheet
        Set mocSht = Sheets(moc)
        
        If Not ignoreShtNames Is Nothing Then
            If Contains(ignoreShtNames, CStr(moc)) Then GoTo NextLoop
        End If
        'If Not needTransfer(mocSht) Then GoTo NextLoop
        
        Dim maxCount As Integer
        maxCount = getMaxCountPerSite(mocSht)
        
        'mocMaxCountPerSite.Add Item:=maxCount, key:=moc
        Call mocMaxCountPerSite.SetAt(moc, maxCount)
NextLoop:
    Next
    
    Dim templateSht As Worksheet
    Set templateSht = Sheets(templateShtName)
    
    Call formatTemplateSht(templateSht, targetMocs, mocMaxCountPerSite, mocNameBeginRowMap)
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setMocPositionsInTemplateSht, " & Err.Description
End Sub

Private Sub setMocPositionsInShtDef(ByRef targetMocNames As Collection, ByRef mocMaxCountPerSite As CMap)
    On Error GoTo ErrorHandler
    If mocMaxCountPerSite Is Nothing Then Exit Sub
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim shtDef As Worksheet
    Set shtDef = Sheets("SHEET DEF")
    
    Dim shtNameCol As Integer
    shtNameCol = innerPositionMgr.sheetDef_shtNameColNo
    
    Dim startRowCol As Integer
    Dim endRowCol As Integer
    startRowCol = insertSheeDefTitle(shtDef)
    endRowCol = startRowCol + 1
    
    Dim emptyRowAdded As Boolean  '基站和控制器MOC之间增加一行空行
    Dim ctrlMocProcessed As Boolean
    Dim btsMocProcessed As Boolean
    emptyRowAdded = False
    ctrlMocProcessed = False
    btsMocProcessed = False
    
    Dim mocStartPos As Integer
    Dim mocEndPos As Integer
    mocStartPos = 1
    
    Dim mocName
    For Each mocName In targetMocNames
        Dim mocRowIdxInShtDef As Integer
        mocRowIdxInShtDef = shtRowInShtDef(CStr(mocName))
        
        If emptyRowAdded = False Then
            If isCtrlMoc(CStr(mocName)) Then
                If btsMocProcessed Then
                    mocStartPos = mocStartPos + 1
                    emptyRowAdded = True
                End If
                ctrlMocProcessed = True
            Else
                If ctrlMocProcessed Then
                    mocStartPos = mocStartPos + 1
                    emptyRowAdded = True
                End If
                btsMocProcessed = True
            End If
        End If
        
        mocEndPos = mocStartPos + mocMaxCountPerSite.GetAt(mocName) + 1
        
        shtDef.Cells(mocRowIdxInShtDef, startRowCol) = mocStartPos
        shtDef.Cells(mocRowIdxInShtDef, endRowCol) = mocEndPos
        
        mocStartPos = mocEndPos + 1
    Next
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setMocPositionsInShtDef, " & Err.Description
End Sub

'key=siteName, value=shtName
Private Sub createAllSiteSht(templateShtName As String, siteNames As Collection, prefixOfIUBShtName As String, postfixOfIUBShtName As String)
    On Error GoTo ErrorHandler
    Dim templateSht As Worksheet
    Set templateSht = Sheets(templateShtName)
    templateSht.Activate
    
    Dim tmpSht As Worksheet
    Set tmpSht = templateSht
    
    Dim idx As Integer
    For idx = 2 To siteNames.count 'siteNames.item(1) is templateShtName
        Dim siteShtName As String
        siteShtName = prefixOfIUBShtName & siteNames.Item(idx) & postfixOfIUBShtName
        If Len(siteShtName) > 31 Then
            siteShtName = shtName4LongSiteName(siteNames.Item(idx), prefixOfIUBShtName, postfixOfIUBShtName)
            m_shtName4SiteCol = genShtNameForSiteCol
            Call setShtNameForSite(siteNames.Item(idx), siteShtName, m_shtName4SiteCol)
        End If
        
        tmpSht.Copy after:=tmpSht
        Set tmpSht = ActiveSheet
        tmpSht.name = siteShtName
        
        Call m_siteNameShtNameMap.SetAt(siteNames.Item(idx), siteShtName)
    Next
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in createAllSiteSht, " & Err.Description
End Sub

'MOC list 页签中的名称是rat名称，一个rat名称有可能对应多个Node名称
'mocNameBeginRowMap : key=mocName, value=moc begin row in iub sheet
Private Function copyData(targetMocs As Collection, mocNameBeginRowMap As CMap) As Boolean
    On Error GoTo ErrorHandler
    Dim moc
    For Each moc In targetMocs
        Dim mocSht As Worksheet
        Set mocSht = Sheets(moc)
        
        If Not copyDataByMoc(mocSht, mocNameBeginRowMap) Then
            copyData = False
            Exit Function
        End If
        
        Application.DisplayAlerts = False
        mocSht.Delete
        Application.DisplayAlerts = True
    Next
    copyData = True
    Exit Function
ErrorHandler:
    Debug.Print "some exception in copyData, " & Err.Description
End Function

Private Sub fixAction()
    On Error GoTo ErrorHandler
    If innerPositionMgr Is Nothing Then loadInnerPositions
    innerPositionMgr.updateSheetDefPositions
    
    m_baseSht.Activate
    
    Dim SiteMergeCount As Long
    With m_baseSht
        rows((listShtAttrRow + 1) & ":" & (listShtAttrRow + 1)).Insert Shift:=xlDown
        rows((listShtAttrRow + 1) & ":" & (listShtAttrRow + 1)).Hidden = True
        
        Dim siteNameCol As Integer
        siteNameCol = siteNameColNum(m_baseSht)
        
        If m_baseSht.Cells(listShtAttrRow, siteNameCol + 1) <> getResByKey("SheetNameForSite") Then
            Cells(listShtAttrRow, siteNameCol + 1).EntireColumn.Insert
            Cells(listShtAttrRow, siteNameCol + 1) = getResByKey("Referenced_Site")
        Else
            Cells(listShtAttrRow, siteNameCol + 2).EntireColumn.Insert
            Cells(listShtAttrRow, siteNameCol + 2) = getResByKey("Referenced_Site")
        End If
               
        Range("A1").Select
        
        If Selection.count = 1 Then
            Range("A1:B1").Merge
        End If
        
        Call setBoard(m_baseSht.name)
    End With
    
    If isMultiVersionWb Then Call delModelDiffReferences
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in fixAction, " & Err.Description
End Sub

Private Function validateSiteName(ByRef siteNames As Collection, prefixOfIUBShtName As String, postfixOfIUBShtName As String) As Boolean
    On Error GoTo ErrorHandler
    validateSiteName = True
    
    Dim prefixLen As Integer
    Dim postfixLen As Integer
    prefixLen = Len(prefixOfIUBShtName)
    postfixLen = Len(postfixOfIUBShtName)
    
    Dim siteName
    For Each siteName In siteNames
        If Len(siteName) > m_maxSiteNameLen _
            Or InStr(siteName, "\") > 0 _
            Or InStr(siteName, "/") > 0 _
            Or InStr(siteName, ":") > 0 _
            Or InStr(siteName, "*") > 0 _
            Or InStr(siteName, "?") > 0 _
            Or InStr(siteName, Chr(34)) > 0 _
            Or InStr(siteName, "[") > 0 _
            Or InStr(siteName, "]") > 0 _
            Or Trim(siteName) = "" Then
            MsgBox Replace(getResByKey("SITE_NAME_INVALID"), "(%0)", siteName) + vbCrLf
            validateSiteName = False
            Exit Function
        End If
    Next
    Exit Function
ErrorHandler:
    Debug.Print "some exception in validateSiteName, " & Err.Description
End Function

Private Function shtName4LongSiteName(siteName As String, prefixOfIUBShtName As String, postfixOfIUBShtName As String) As String
    On Error GoTo ErrorHandler
    Dim prefixLen As Integer
    Dim postfixLen As Integer
    prefixLen = Len(prefixOfIUBShtName)
    postfixLen = Len(postfixOfIUBShtName)
    
    Dim nameLen As Integer
    nameLen = m_maxShtNameLen - prefixLen - postfixLen
    
    Dim siteNameLeft As String
    siteNameLeft = Left(siteName, nameLen - 5) 'xxxx~0001
    siteNameLeft = siteNameLeft & "~"
    
    Dim strLongNameIdx As String
    strLongNameIdx = CStr(m_longNameIdx)
    
    Dim pos As Integer ' to test
    For pos = 1 To 4 - Len(strLongNameIdx)
        siteNameLeft = siteNameLeft & "0"
    Next
    siteNameLeft = siteNameLeft & strLongNameIdx
    
    shtName4LongSiteName = prefixOfIUBShtName & siteNameLeft & postfixOfIUBShtName

    m_longNameIdx = m_longNameIdx + 1
    Exit Function
ErrorHandler:
    Debug.Print "some exception in shtName4LongSiteName, " & Err.Description
End Function

'return sheet name for site column
Private Function genShtNameForSiteCol() As Integer
    On Error GoTo ErrorHandler
    genShtNameForSiteCol = colNumByAttr(m_baseSht, getResByKey("SheetNameForSite"))
    If genShtNameForSiteCol <> -1 Then Exit Function
    
    Dim siteNameCol As Integer
    siteNameCol = siteNameColNum(m_baseSht)
    genShtNameForSiteCol = siteNameCol + 1

    m_baseSht.Cells(listShtAttrRow, genShtNameForSiteCol).EntireColumn.Insert
    m_baseSht.Cells(listShtAttrRow, genShtNameForSiteCol) = getResByKey("SheetNameForSite")
    
    Call insertShtNameForSiteColInMappingDef
    Exit Function
ErrorHandler:
    Debug.Print "some exception in genShtNameForSiteCol, " & Err.Description
End Function

Private Sub insertShtNameForSiteColInMappingDef()
    On Error GoTo ErrorHandler
    Dim mappingDef As Worksheet
    Set mappingDef = Sheets("MAPPING DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim shtNameCol As Integer
    Dim colNameCol As Integer
    shtNameCol = innerPositionMgr.mappingDef_shtNameColNo
    colNameCol = innerPositionMgr.mappingDef_colNameColNo
    
    mappingDef.Activate
    
    Dim insertRowPos As Integer
    For insertRowPos = functionShtTitleRow + 1 To getSheetUsedRows(mappingDef)
        If Cells(insertRowPos, shtNameCol) = m_baseSht.name And is_Site(Cells(insertRowPos, colNameCol)) Then Exit For
    Next
    insertRowPos = insertRowPos + 1

    rows(insertRowPos & ":" & insertRowPos).Insert Shift:=xlDown
    Cells(insertRowPos, innerPositionMgr.mappingDef_shtNameColNo) = m_baseSht.name
    Cells(insertRowPos, innerPositionMgr.mappingDef_grpNameColNo) = m_baseSht.Cells(1, 1)
    Cells(insertRowPos, innerPositionMgr.mappingDef_colNameColNo) = getResByKey("SheetNameForSite")
    Cells(insertRowPos, innerPositionMgr.mappingDef_mocNameColNo) = "NODE"
    Cells(insertRowPos, innerPositionMgr.mappingDef_attrNameColNo) = "SHEETNAMEFORSITE"
    Cells(insertRowPos, innerPositionMgr.mappingDef_neTypeColNo) = Cells(insertRowPos + 1, innerPositionMgr.mappingDef_neTypeColNo)
    Cells(insertRowPos, innerPositionMgr.mappingDef_neVersionColNo) = Cells(insertRowPos + 1, innerPositionMgr.mappingDef_neVersionColNo)
    Cells(insertRowPos, innerPositionMgr.mappingDef_readOnlyColNo) = "TRUE"
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in insertShtNameForSiteColInMappingDef, " & Err.Description
End Sub

Private Sub setShtNameForSite(siteName As String, shtName As String, ByVal shtName4SiteCol As Integer)
    On Error GoTo ErrorHandler
    Dim siteNameCol As Integer
    siteNameCol = shtName4SiteCol - 1
    
    Dim rowIdx As Integer
    For rowIdx = listShtAttrRow + 1 To getSheetUsedRows(m_baseSht)
        If m_baseSht.Cells(rowIdx, siteNameCol) = siteName Then
            m_baseSht.Cells(rowIdx, shtName4SiteCol) = shtName
            Exit Sub
        End If
    Next
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in setShtNameForSite, " & Err.Description
End Sub

Private Function getMaxCountPerSite(ByRef mocSht As Worksheet) As Integer
    On Error GoTo ErrorHandler
    getMaxCountPerSite = 0

    Dim NameCol As Integer
    NameCol = siteNameColNum(mocSht)
    If NameCol = -1 Then NameCol = controllerNameColNum(mocSht)
    If NameCol = -1 Then Exit Function
    
    Dim siteNameMoiCountMap As New Collection
    
    Dim maxRow As Integer
    maxRow = getSheetUsedRows(mocSht)
    
    Dim siteName As String
    Dim rowIdx As Integer
    For rowIdx = listShtAttrRow + 1 To maxRow
        siteName = mocSht.Cells(rowIdx, NameCol)
        If siteName = "" Then GoTo NextLoop
        Call increaseKeyCounter(siteNameMoiCountMap, siteName)
NextLoop:
    Next
    
    Dim iter
    For Each iter In siteNameMoiCountMap
        If getMaxCountPerSite < iter Then getMaxCountPerSite = iter
    Next
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getMaxCountPerSite, " & Err.Description
End Function

Private Sub increaseKeyCounter(ByRef coll As Collection, k As String)
    On Error GoTo ErrorHandler
    If coll Is Nothing Then Exit Sub
    
    Dim v As Integer
    If Contains(coll, k) Then
        v = coll(k)
        coll.Remove k
    Else
        v = 0
    End If

    v = v + 1
    coll.Add Item:=v, key:=k
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in increaseKeyCounter, " & Err.Description
End Sub

Private Function insertSheeDefTitle(ByRef shtDef As Worksheet) As Integer
    On Error GoTo ErrorHandler
    insertSheeDefTitle = 0
    Dim startColumn As Integer
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    insertSheeDefTitle = innerPositionMgr.sheetDef_startRowColNo
    If insertSheeDefTitle <> -1 Then Exit Function
    
    startColumn = getSheetUsedColumnsByRow(shtDef, functionShtTitleRow)
    shtDef.Cells(functionShtTitleRow, startColumn + 1) = m_mocStartRowColName
    shtDef.Cells(functionShtTitleRow, startColumn + 2) = m_mocEndRowColName
    insertSheeDefTitle = startColumn + 1
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in insertSheeDefTitle, " & Err.Description
End Function

Private Function isCtrlMoc(mocName As String, Optional ctrlType As String) As Boolean
    On Error GoTo ErrorHandler
    isCtrlMoc = False
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim targetRange As Range
    Dim firstAddr As String
    Dim targetOffset As Integer
    targetOffset = innerPositionMgr.mappingDef_neTypeColNo - innerPositionMgr.mappingDef_shtNameColNo
    With mappingDef.columns(innerPositionMgr.mappingDef_shtNameColNo)
        Set targetRange = .Find(mocName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                ctrlType = targetRange.Offset(0, targetOffset).value
                If ctrlType = getResByKey("umtsController") Or ctrlType = getResByKey("gsmController") Then
                    isCtrlMoc = True
                    Exit Function
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    Exit Function
ErrorHandler:
    Debug.Print "some exception in isCtrlMoc, " & Err.Description
End Function

'mocNameBeginRowMap is an out parameter
Private Sub formatTemplateSht(templateSht As Worksheet, targetMocs As Collection, mocMaxCountPerSite As CMap, mocNameBeginRowMap As CMap)
    On Error GoTo ErrorHandler
    If templateSht Is Nothing Or mocMaxCountPerSite.count = 0 Then Exit Sub

    Dim controllerMocs As New Collection
    Dim siteMocs As New Collection
    Call classifyMocs(targetMocs, controllerMocs, siteMocs)
    
    Dim formatBeginRow As Integer
    Dim formatEndRow As Integer
    Dim mocStartRow As Integer
    Dim mocEndRow As Integer
    formatBeginRow = 1
    mocEndRow = 0
    
    If controllerMocs.count > 0 Then
        Dim ctrlMoc
        For Each ctrlMoc In controllerMocs
            mocStartRow = mocEndRow + 1
            mocEndRow = formatMoc(templateSht, CStr(ctrlMoc), mocStartRow, mocMaxCountPerSite, mocNameBeginRowMap)
        Next

        formatEndRow = mocEndRow
        Call formatMocNames(templateSht, formatBeginRow, formatEndRow)
        
        mocEndRow = mocEndRow + 1
        formatBeginRow = mocEndRow + 1
    End If
    
    If siteMocs.count > 0 Then
        Dim siteMoc
        For Each siteMoc In siteMocs
            mocStartRow = mocEndRow + 1
            mocEndRow = formatMoc(templateSht, CStr(siteMoc), mocStartRow, mocMaxCountPerSite, mocNameBeginRowMap)
        Next
    
        formatEndRow = mocEndRow
        Call formatMocNames(templateSht, formatBeginRow, formatEndRow)
    End If
    
    Call formatTemplateSht_i(templateSht)
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in formatTemplateSht, " & Err.Description
End Sub

Private Function formatMoc(templateSht As Worksheet, mocName As String, mocStartRow As Integer, mocMaxCountPerSite As CMap, mocNameBeginRowMap As CMap) As Integer
    On Error GoTo ErrorHandler
    Dim mocSht As Worksheet
    Set mocSht = Sheets(mocName)
    
    Dim maxCol As Long
    maxCol = getCopySrcMaxCol(mocSht)
    
    Dim mocEndRow As Integer
    mocEndRow = mocStartRow + mocMaxCountPerSite.GetAt(mocName) + 1
    
    Dim mocRange As Range
    With templateSht
        .Activate
        Set mocRange = .Range("A" & CStr(mocStartRow) + ":" & C(maxCol) + CStr(mocEndRow))
    End With
    
    Call setRangeBoard(mocRange)
    mocRange.Interior.colorIndex = xlColorIndexNone
    Range("A" + CStr(mocStartRow) + ":" + "A" + CStr(mocEndRow)).Merge
    
    Dim nameColInListSht As Integer
    nameColInListSht = getNeNameColInListSht(mocSht)
    Call copyTitle(templateSht, mocStartRow, mocSht, nameColInListSht, maxCol)
    templateSht.Cells(mocStartRow, 1).value = mocName
    'Call writeMocName(templateSht, mocStartRow, mocName)

    Call mocNameBeginRowMap.SetAt(mocName, mocStartRow)
    
    formatMoc = mocEndRow
    Exit Function
ErrorHandler:
    Debug.Print "some exception in formatMoc, " & Err.Description
End Function

Private Sub classifyMocs(targetMocs As Collection, controllerMocs As Collection, siteMocs As Collection)
    On Error GoTo ErrorHandler
    Dim moc
    For Each moc In targetMocs
        Dim mocName As String
        mocName = CStr(moc)
        If isCtrlMoc(mocName) Then
            controllerMocs.Add Item:=mocName, key:=mocName
        Else
            siteMocs.Add Item:=mocName, key:=mocName
        End If
    Next
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in classifyMocs, " & Err.Description
End Sub

Private Sub copyTitle(templateSht As Worksheet, ByVal mocStartRow As Integer, mocSht As Worksheet, ByVal NameCol As Integer, ByVal maxCol As Integer)
    On Error GoTo ErrorHandler
    Dim copyStartColInTemplateSht As Long
    copyStartColInTemplateSht = 2
    
    'work sheet with operation, column 2 is operation
    If isOperationWs(mocSht) Then
        copyStartColInTemplateSht = 3
    End If

    Call copyDataByRow(mocSht, listShtAttrRow, NameCol + 1, maxCol, templateSht, mocStartRow, copyStartColInTemplateSht)
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in copyTitle, " & Err.Description
End Sub

Private Sub formatMocNames(templateSht As Worksheet, ByVal startRow As Integer, ByVal endRow As Integer)
    On Error GoTo ErrorHandler
     With templateSht.Range("A" & startRow & ":A" & endRow)
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.name = "Microsoft Sans Serif"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = xlUnderlineStyleNone
        .Font.colorIndex = 2
        .WrapText = True
        .Interior.colorIndex = 9
        .Interior.Pattern = xlGray8
        .Interior.PatternColorIndex = xlAutomatic
    End With
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in formatMocNames, " & Err.Description
End Sub

Private Sub formatTemplateSht_i(templateSht As Worksheet)
    On Error GoTo ErrorHandler
    With templateSht.Cells
        .ColumnWidth = 20
        .WrapText = True
    End With
    
    templateSht.Select
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 0
    End With
    ActiveWindow.FreezePanes = True
    
    Dim maxRow As Integer
    maxRow = getSheetUsedRows(templateSht)
    Dim rowIdx As Integer
    For rowIdx = 1 To maxRow
        If Application.WorksheetFunction.CountA(templateSht.rows(rowIdx)) <> 0 Then templateSht.rows.WrapText = False
    Next
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in formatTemplateSht_i, " & Err.Description
End Sub

Private Sub copyDataByRow(srcSht As Worksheet, ByVal srcRow As Integer, ByVal srcStartCol As Long, _
    ByVal srcEndCol As Long, dstSht As Worksheet, ByVal dstRow As Integer, ByVal dstStartCol)
    On Error GoTo ErrorHandler
    dstSht.Activate
    dstSht.Cells(dstRow, dstStartCol).Select
    srcSht.Range(C(srcStartCol) + CStr(srcRow) + ":" + C(srcEndCol) + CStr(srcRow)).Copy Destination:=dstSht.Cells(dstRow, dstStartCol)
    
    'copy operation
    If isOperationWs(srcSht) Then
        Dim opCol As Long
        opCol = operationColNum(srcSht)
        If opCol <> -1 Then
            srcSht.Cells(srcRow, opCol).Copy Destination:=dstSht.Cells(dstRow, m_OprationCol)
        End If
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in copyDataByRow, " & Err.Description
End Sub

Private Sub copyDataByRows(srcSht As Worksheet, ByVal srcStartRow As Integer, ByVal srcEndRow As Integer, _
    ByVal srcStartCol As Long, ByVal srcEndCol As Long, dstSht As Worksheet, ByVal dstRow As Integer, ByVal dstStartCol)
    On Error GoTo ErrorHandler
    With dstSht
        .Activate
        .Cells(dstRow, dstStartCol).Select
        srcSht.Range(C(srcStartCol) + CStr(srcStartRow) + ":" + C(srcEndCol) + CStr(srcEndRow)).Copy Destination:=.Cells(dstRow, dstStartCol)
    End With
    
    'copy operation
    If isOperationWs(srcSht) Then
        Dim opCol As Long
        opCol = operationColNum(srcSht)
        If opCol <> -1 Then
            srcSht.Range(C(opCol) + CStr(srcStartRow) + ":" + C(opCol) + CStr(srcEndRow)).Copy Destination:=dstSht.Cells(dstRow, m_OprationCol)
        End If
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in copyDataByRows, " & Err.Description
End Sub

Private Function getNeNameColInListSht(listSht As Worksheet) As Integer
    On Error GoTo ErrorHandler
    getNeNameColInListSht = 1

    getNeNameColInListSht = siteNameColNum(listSht)
    If getNeNameColInListSht = -1 Then
        getNeNameColInListSht = controllerNameColNum(listSht)
    End If
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getNeNameColInListSht, " & Err.Description
End Function

Private Function getListShtUsedColumns(sheet As Worksheet) As Long
    On Error GoTo ErrorHandler
    getListShtUsedColumns = 0
    
    Do While Trim(sheet.Cells(listShtAttrRow, getListShtUsedColumns + 1)) <> ""
      getListShtUsedColumns = getListShtUsedColumns + 1
    Loop
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getListShtUsedColumns, " & Err.Description
End Function

Private Function isSharedCtrlName(ws As Worksheet, ByVal maxCol As Integer) As Boolean
    On Error GoTo ErrorHandler
    isSharedCtrlName = False
    Dim colName As String
    colName = ws.Cells(listShtAttrRow, maxCol)
    If colName = getResByKey("NodeBName") Or colName = getResByKey("BTSName") Then isSharedCtrlName = True
    Exit Function
ErrorHandler:
    Debug.Print "some exception in isSharedCtrlName, " & Err.Description
End Function

Private Function copyDataByMoc(mocSht As Worksheet, mocNameBeginRowMap As CMap) As Boolean
    On Error GoTo ErrorHandler
    copyDataByMoc = True
    Call deleteEmptyRows(mocSht)
    If siteNameColNum(mocSht) <> -1 Then
        copyDataByMoc = copySiteMocData(mocSht, mocNameBeginRowMap)
    ElseIf controllerNameColNum(mocSht) <> -1 Then
        copyDataByMoc = copyCtrlMocData(mocSht, mocNameBeginRowMap)
    End If
    Exit Function
ErrorHandler:
    Debug.Print "some exception in copyDataByMoc, " & Err.Description
End Function

Private Sub deleteEmptyRows(mocSht As Worksheet)
    On Error GoTo ErrorHandler
    Dim maxRow As Integer
    maxRow = getSheetUsedRows(mocSht)
    If maxRow > listShtAttrRow Then
        Dim rowIdx As Long
        With mocSht
            For rowIdx = maxRow To listShtAttrRow Step -1
                If rowIsBlank(mocSht, rowIdx) Then .rows(rowIdx).Delete
            Next
        End With
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in deleteEmptyRows, " & Err.Description
End Sub

'siteShtNames : key=siteName, value=shtName
'mocNameBeginRowMap : key=mocName, value=moc begin row in iub sheet
Private Function copySiteMocData(mocSht As Worksheet, mocNameBeginRowMap As CMap) As Boolean
    On Error GoTo ErrorHandler
    copySiteMocData = True
    Dim dstRow As Integer
    Dim dstCol As Integer

    dstCol = 2
    If isOperationWs(mocSht) Then dstCol = 3
    
    Dim nameColInListSht As Integer
    nameColInListSht = getNeNameColInListSht(mocSht)
        
    Dim copySiteDstRowMap As New CMap 'key=siteshtname, value=copy dest row
    
    Dim maxRow As Integer
    maxRow = getSheetUsedRows(mocSht)
    
    Dim cpBeginRow As Integer
    Dim cpEndRow As Integer
    Dim preName As String
    Dim curName As String
    
    With mocSht
        If maxRow > listShtAttrRow Then
            preName = .Cells(listShtAttrRow + 1, nameColInListSht)
            cpBeginRow = listShtAttrRow + 1
            cpEndRow = cpBeginRow
        End If
        
        'Dim ratName As String
        Dim rowIdx As Integer
        For rowIdx = listShtAttrRow + 1 To maxRow + 1
            curName = .Cells(rowIdx, nameColInListSht)
            If curName = preName Then
                GoTo NextLoop
            End If
            
            Dim dstSiteShtNames As Collection
            Set dstSiteShtNames = getDstSiteShtNames(mocSht, nameColInListSht, preName, .Cells(listShtAttrRow, nameColInListSht)) 'DTS2016091909724
            
            If dstSiteShtNames.count = 0 Then
                MsgBox getResByKey("SiteNameNotExist") & .name & "[" & (rowIdx - 1) & ":" & nameColInListSht & "]" & getResByKey("CheckAndTryAgain")
                copySiteMocData = False
                Exit Function
            End If
            
            Dim dstSiteShtName
            For Each dstSiteShtName In dstSiteShtNames
                If copySiteDstRowMap.hasKey(CStr(dstSiteShtName)) Then
                    dstRow = copySiteDstRowMap.GetAt(CStr(dstSiteShtName))
                Else
                    dstRow = mocNameBeginRowMap.GetAt(.name) + 1
                End If
        
                Call copyDataByRows(mocSht, cpBeginRow, cpEndRow, nameColInListSht + 1, getCopySrcMaxCol(mocSht), Sheets(CStr(dstSiteShtName)), dstRow, dstCol)
                
                dstRow = dstRow + 1 + cpEndRow - cpBeginRow
                Call copySiteDstRowMap.SetAt(CStr(dstSiteShtName), dstRow)
            Next
            
            cpBeginRow = rowIdx
            preName = curName
NextLoop:
            cpEndRow = rowIdx
        Next
    End With
    Exit Function
ErrorHandler:
    Debug.Print "some exception in copySiteMocData, " & Err.Description
End Function

'siteShtNames : key=siteName, value=shtName
'mocNameBeginRowMap : key=mocName, value=moc begin row in iub sheet
Private Function copyCtrlMocData(mocSht As Worksheet, mocNameBeginRowMap As CMap) As Boolean
    On Error GoTo ErrorHandler
    copyCtrlMocData = True
    Dim ctrlType As String
    If Not isCtrlMoc(mocSht.name, ctrlType) Then
        Debug.Print "not controller moc!"
        GoTo ErrorHandler
    End If
    
    Dim dstRow As Integer
    Dim dstCol As Integer

    dstCol = 2
    If isOperationWs(mocSht) Then dstCol = 3
    
    Dim nameColInListSht As Integer
    nameColInListSht = getNeNameColInListSht(mocSht)
    
    Dim copySiteDstRowMap As New CMap 'key=siteshtname, value=copy dest row
    
    Dim ratNameColName As String
    ratNameColName = getResByKey("NodeBName")
    If ctrlType = getResByKey("gsmController") Then ratNameColName = getResByKey("BTSName")
    
    Dim maxCol As Long
    maxCol = getListShtUsedColumns(mocSht)
    With mocSht
        If .Cells(listShtAttrRow, maxCol) <> ratNameColName Then Exit Function
        
        Dim maxRow As Integer
        maxRow = getSheetUsedRows(mocSht)
    
        Dim sharedSiteRatNames As Variant
        
        'Dim ratName As String
        Dim rowIdx As Integer
        For rowIdx = listShtAttrRow + 1 To maxRow
            sharedSiteRatNames = Split(.Cells(rowIdx, maxCol), ",")
            
            Dim ratName
            For Each ratName In sharedSiteRatNames
                Dim dstSiteShtNames As Collection
                Set dstSiteShtNames = getDstSiteShtNames(mocSht, nameColInListSht, CStr(ratName), "*" & ratNameColName) 'DTS2016091909724
                
                If dstSiteShtNames.count = 0 Then
                    MsgBox getResByKey("SiteNameNotExist") & mocSht.name & "[" & (rowIdx - 1) & ":" & nameColInListSht & "]" & getResByKey("CheckAndTryAgain")
                    copyCtrlMocData = False
                    Exit Function
                End If
                
                Dim dstSiteShtName
                For Each dstSiteShtName In dstSiteShtNames
                    If copySiteDstRowMap.hasKey(CStr(dstSiteShtName)) Then
                        dstRow = copySiteDstRowMap.GetAt(CStr(dstSiteShtName))
                    Else
                        dstRow = mocNameBeginRowMap.GetAt(mocSht.name) + 1
                    End If
            
                    Call copyDataByRows(mocSht, rowIdx, rowIdx, nameColInListSht + 1, maxCol - 1, Sheets(CStr(dstSiteShtName)), dstRow, dstCol)
                    
                    dstRow = dstRow + 1
                    Call copySiteDstRowMap.SetAt(CStr(dstSiteShtName), dstRow)
                Next
            Next
        Next
    End With
    Exit Function
    
ErrorHandler:
    Debug.Print "some exception in copyCtrlMocData, " & Err.Description
End Function

Private Function getCopySrcMaxCol(mocSht As Worksheet) As Integer
    On Error GoTo ErrorHandler
    getCopySrcMaxCol = getListShtUsedColumns(mocSht)
    
    If isSharedCtrlName(mocSht, getCopySrcMaxCol) Then
        getCopySrcMaxCol = getCopySrcMaxCol - 1
    End If
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getCopySrcMaxCol, " & Err.Description
End Function

Private Function getDstSiteShtNames(mocSht As Worksheet, ByVal nameColInListSht As Integer, ratName As String, ratNameTitle As String) As Collection
    On Error GoTo ErrorHandler
    Set getDstSiteShtNames = New Collection
    
    Dim shtName As String
    
    If m_siteNameShtNameMap.hasKey(ratName) Then
        shtName = m_siteNameShtNameMap.GetAt(ratName)
        getDstSiteShtNames.Add Item:=shtName, key:=shtName
        Exit Function
    End If
    
    Dim nodeNames As Collection
    Set nodeNames = getNodeNamesByRatName(mocSht, ratName, ratNameTitle)
    
    Dim nodeName
    For Each nodeName In nodeNames
        If m_siteNameShtNameMap.hasKey(CStr(nodeName)) Then
            shtName = m_siteNameShtNameMap.GetAt(nodeName)
            getDstSiteShtNames.Add Item:=shtName, key:=shtName
        End If
    Next
    If getDstSiteShtNames.count >= 1 Then Exit Function
    
    Dim maxCol As Integer
    maxCol = getListShtUsedColumns(mocSht)
    If isSharedCtrlName(mocSht, maxCol) Then
        Dim rowIdx As Integer
        With mocSht
            For rowIdx = listShtAttrRow + 1 To getSheetUsedRows(mocSht)
                If .Cells(rowIdx, nameColInListSht).value = ratName Then
                    shtName = .Cells(rowIdx, maxCol).value
                    getDstSiteShtNames.Add Item:=shtName, key:=shtName
                End If
            Next
        End With
    End If
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getDstSiteShtNames, " & Err.Description
End Function

Private Function getNodeNamesByRatName(mocSht As Worksheet, ratName As String, ratNameTitle As String) As Collection
    On Error GoTo ErrorHandler
    Set getNodeNamesByRatName = New Collection
    
    If m_neType = NeType_MRAT Then
        Set getNodeNamesByRatName = nodeNames(ratNameTitle, ratName)
    Else
        getNodeNamesByRatName.Add Item:=ratName, key:=ratName
    End If
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getNodeNamesByRatName, " & Err.Description
End Function

' 传输页中的function名称列可能带*号，也可能不带*号
Private Function nodeNames(ratNameColName As String, ratName As String) As Collection
    On Error GoTo ErrorHandler
    Set nodeNames = New Collection
    
    Dim ratNameCol As Integer
    ratNameCol = colNumByAttr(m_baseSht, ratNameColName)
    
    If ratNameCol = -1 Then
        If Left(ratNameColName, 1) = "*" Then
            Dim functionColName As String
            functionColName = Right(ratNameColName, Len(ratNameColName) - 1)
            ratNameCol = colNumByAttr(m_baseSht, functionColName)
        End If
    End If
    
    If ratNameCol = -1 Then
        nodeNames.Add Item:=ratName, key:=ratName
        Exit Function
    End If
    
    Dim rowIdx As Integer
    Dim nodeName As String
    For rowIdx = listShtAttrRow + 1 To getSheetUsedRows(m_baseSht)
        If m_baseSht.Cells(rowIdx, ratNameCol) = ratName Then
            nodeName = m_baseSht.Cells(rowIdx, siteNameColNum(m_baseSht))
            nodeNames.Add Item:=nodeName, key:=nodeName
        End If
    Next
    Exit Function
ErrorHandler:
    Debug.Print "some exception in nodeNames, " & Err.Description
End Function

Private Sub delModelDiffReferences()
    On Error GoTo ErrorHandler
    Dim modelDiffSht As Worksheet
    Set modelDiffSht = ThisWorkbook.Worksheets(getResByKey("ModelDiffSht"))
    
    Dim targetShtName As String
    Dim targetGrpName As String
    Dim targetColName As String
    Dim refArray As Variant
    Dim rowIdx As Integer
    
    With modelDiffSht
        For rowIdx = 2 To .Range("a65535").End(xlUp).row
            If Not isValidReference(.Cells(rowIdx, refColInModDiffSht), refArray, "\") Then GoTo NextLoop
            
            targetShtName = refArray(0)
            targetGrpName = refArray(1)
            targetColName = refArray(2)
            
            Dim shtType As String
            If getSheetType(targetShtName) = SheetType_List Then
                With .Cells(rowIdx, refColInModDiffSht)
                    .Hyperlinks.Delete
                    .value = targetShtName & "\" & targetColName
                End With
            End If
NextLoop:
        Next
        Call setBorders(.UsedRange)
    End With
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in delModelDiffReferences, " & Err.Description
End Sub

Private Sub preProcessCtrlMocShts(targetMocNames As Collection, siteNames As Collection)
    On Error GoTo ErrorHandler
    Dim neType As String
    Dim moc As Variant
    For Each moc In targetMocNames
        If isCtrlMoc(CStr(moc), neType) Then
            Dim sht As Worksheet
            Set sht = ThisWorkbook.Worksheets(CStr(moc))
            Call preProcessCtrlMocBySht(sht, siteNames, neType)
        End If
    Next
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in preProcessCtrlMocShts, " & Err.Description
End Sub

'for customize templates, there is no shared site names in the last hidden column in pure controller moc sheets, these mocs should be transferred to all site sheets
'while pure controller moc sheets with shared sitenames in the last hidden column, but without sites, these mocs should be transferred to all site sheets
Private Sub preProcessCtrlMocBySht(ws As Worksheet, siteNames As Collection, neType As String)
    On Error GoTo ErrorHandler

    Dim allSiteNames As String
    allSiteNames = collectionJoin(siteNames)
    
    Dim maxCol As Integer, maxRow As Integer
    maxCol = getListShtUsedColumns(ws)
    With ws
        maxRow = .Range("a65535").End(xlUp).row

        If isSharedCtrlName(ws, maxCol) Then
            If maxRow > listShtAttrRow Then
                Dim rowIdx As Integer
                For rowIdx = listShtAttrRow + 1 To maxRow
                    If .Cells(rowIdx, maxCol).value = "" Then .Cells(rowIdx, maxCol).value = allSiteNames
                Next
            End If
        Else
            Dim insertColStr As String
            insertColStr = getColStr(maxCol + 1)
            If neType = getResByKey("gsmController") Then
                .Cells(listShtGrpRow, maxCol + 1).value = "Shared by BTS"
                .Cells(listShtAttrRow, maxCol + 1).value = getResByKey("BTSName")
            ElseIf neType = getResByKey("umtsController") Then
                .Cells(listShtGrpRow, maxCol + 1).value = "Shared by NodeB"
                .Cells(listShtAttrRow, maxCol + 1).value = getResByKey("NodeBName")
            End If
            If maxRow > listShtAttrRow Then
                .Range(insertColStr & (listShtAttrRow + 1) & ":" & insertColStr & maxRow).value = allSiteNames
            End If
            .columns(maxCol + 1).Hidden = True
        End If
    End With
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in preProcessCtrlMocBySht, " & Err.Description
End Sub
