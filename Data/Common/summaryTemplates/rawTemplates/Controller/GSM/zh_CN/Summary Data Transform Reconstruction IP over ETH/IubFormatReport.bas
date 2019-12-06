Attribute VB_Name = "IubFormatReport"
Option Explicit

Public bIsEng As Boolean  '用于控制设置中英文
'==================================================================================================
'生成Iub形式代码
'==================================================================================================
Public Const HyperLinkColorIndex = 6
Public Const BluePrintSheetColor = 5
Public Const MaxChosenSiteNum = 1000

Public Const g_NodeBBeginRow = 3
Public Const g_TitleRowNo As Long = 2
Public NodeNameArray(MaxChosenSiteNum) As String
Public RatNameArray(MaxChosenSiteNum) As String
Public g_NodeIndex As Long

Public ChosenSiteArray(MaxChosenSiteNum) As String
Public ChosenSiteNames As New Collection

Public FILE_TYPE As String '0- Summary 1-Bulk

Sub setFileType(fileType As String)
        FILE_TYPE = fileType
End Sub
Function is_SiteorController(columnName As String) As Boolean
    is_SiteorController = False
    If (is_Site(columnName) Or is_Controller(columnName)) Then
        is_SiteorController = True
    End If
End Function

'[Summary转类似原IUB表格功能]将List Sheet数据转到NodeB Sheet页
Public Sub ConvertList(ByVal strSheetName As String, ByVal startRow As Long, ByVal endRow As Long, ByVal IsEmptyDataSheet As Boolean)
    Dim vListSht As Worksheet
    Dim vName As String
    Dim vMocEndRowNum As Long
    Dim vNameEnd As String
    Dim vListColumnNum As Long

    If (Not IsSheetExist(strSheetName)) Then
        Exit Sub
    End If
    
    Set vListSht = Sheets(strSheetName)
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    
    'name colum may not be the first colum
    Dim siteNameCol As Long
    Dim ctrlNameCol As Long
    siteNameCol = siteNameColNum(vListSht)
    ctrlNameCol = controllerNameColNum(vListSht)
    
    Dim nameCol As Long
    nameCol = siteNameCol
    If nameCol = -1 Then
        nameCol = ctrlNameCol
    End If
    
    '只处理NodeB&BTS对象
    If (nameCol = -1) Or (isTrasnPortSheet(strSheetName) And FILE_TYPE = "0") Then
        Exit Sub
    End If
    
    vName = GetCell(vListSht, g_TitleRowNo, nameCol)
    vNameEnd = GetCell(vListSht, g_TitleRowNo, vListColumnNum)

    Call FormatFirst(vListSht, strSheetName, startRow, endRow)
    
    vMocEndRowNum = GetSheetUsedRows(vListSht)
    If (vMocEndRowNum <= g_TitleRowNo) Then
        vMocEndRowNum = g_TitleRowNo + 1
    End If
    
    'copy continuous columns from next to nameCol to the end(but not include the end when is shared controller name)
    Call copyTitle(vListSht, strSheetName, startRow, nameCol, vListColumnNum)

'    If is_Controller(vName) Then
'        '如果最后一列是共享名称，则少拷贝一列'最后一列不是共享名称，则正常拷贝所有列
'        If isSharedControllerName(vNameEnd) Then
'            Call copyControllerTitle_1(vListSht, strSheetName, startRow)
'        Else
'            Call copyControllerTitle(vListSht, strSheetName, startRow)
'        End If
'    Else
'        Call CopySiteTitle(vListSht, strSheetName, startRow)
''        Call CopySiteData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet)
'    End If
    
    Call setCopiedRangeFont(vListSht, startRow, endRow)
End Sub

Public Sub copyTitle(ByRef vListSht As Worksheet, ByRef strSheetName As String, ByVal startRow As Long, ByVal nameCol As Long, ByVal endCol As Long)
    Dim cpEndCol As Long
    cpEndCol = endCol
    If isSharedControllerName(GetCell(vListSht, g_TitleRowNo, endCol)) Then
        cpEndCol = endCol - 1
    End If
    
    Dim vSiteIndex As Long
    Dim vListColumnNum As Long
    Dim vSiteName As String
    Dim vSiteSheet As Worksheet
    Dim myRange As Range

    vSiteName = ChosenSiteArray(0)
    If ("" = vSiteName) Then
        Exit Sub
    End If
    Call WriteDatabyRow(vListSht, vSiteName, g_TitleRowNo, nameCol + 1, cpEndCol, startRow)

    Set vSiteSheet = Sheets(vSiteName)
    Call WriteSheetName(vSiteSheet, startRow, strSheetName)
    '关闭备注大小调整，提示刷表格效率
    Set myRange = vSiteSheet.Range("B" + CStr(startRow) + ":" + C(cpEndCol) + CStr(startRow))
    Call refreshComment(myRange)
End Sub

Private Function isSharedControllerName(ByRef vName As String) As Boolean
    If (vName = "NodeB Name" Or vName = getResByKey("A158")) _
        Or (vName = "BTS Name" Or vName = getResByKey("A159")) Then
        isSharedControllerName = True
    Else
        isSharedControllerName = False
    End If
End Function

Private Sub setCopiedRangeFont(ByRef vListSht As Worksheet, ByRef startRow As Long, ByRef endRow As Long)
    Dim vSiteName As String
    vSiteName = ChosenSiteArray(0)
    If ("" = vSiteName) Then
        Exit Sub
    End If
    
    Dim vSiteSheet As Worksheet
    Dim vListColumnNum As Long
    Set vSiteSheet = ThisWorkbook.Worksheets(vSiteName)
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    
    Dim myRange As Range
    Set myRange = vSiteSheet.Range("B" + CStr(startRow) + ":" + C(vListColumnNum) + CStr(endRow))
    Call setRangeFont(myRange, "Arial")
    Call setRangeFontSize(myRange, 10)
End Sub

Public Sub setRangeFont(ByRef certainRange As Range, ByRef fontName As String)
    certainRange.Font.name = fontName
End Sub

Public Sub setRangeWrap(ByRef certainRange As Range, ByRef flag As Boolean)
    certainRange.WrapText = flag
End Sub

Public Sub setRangeFontSize(ByRef certainRange As Range, ByRef fontSize As Long)
    certainRange.Font.Size = fontSize
End Sub

Public Sub FormatFirst(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long, ByVal siteEndRow As Long)
    Dim vSiteIndex As Long
    Dim vSiteName As String
    Dim vSiteSheet As Worksheet
    Dim vRowIndex As Long
    Dim vListColumnNum As Long
    Dim myRange As Range
    Dim vName As String
    Dim vNameEnd As String
    
    vSiteName = ChosenSiteArray(0)
    If ("" = vSiteName) Then
        Exit Sub
    End If
    
    'name colum may not be the first colum
    Dim siteNameCol As Long
    Dim ctrlNameCol As Long
    siteNameCol = siteNameColNum(vListSht)
    ctrlNameCol = controllerNameColNum(vListSht)
    
    Dim nameCol As Long
    nameCol = siteNameCol
    If nameCol = -1 Then
        nameCol = ctrlNameCol
    End If
    
    Set vSiteSheet = Sheets(vSiteName)
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    vName = GetCell(vListSht, g_TitleRowNo, nameCol)
    vNameEnd = GetCell(vListSht, g_TitleRowNo, vListColumnNum)
    
    '如果是控制器则少生成一列,如果最后一列是共享名称，则少拷贝一列'最后一列不是共享名称，正常拷贝所有列
    If is_Controller(vName) And isSharedControllerName(vNameEnd) Then
        vListColumnNum = vListColumnNum - 1
    End If
    
    vSiteSheet.Activate
    Set myRange = vSiteSheet.Range("A" + CStr(SiteStartRow) + ":" + C(vListColumnNum) + CStr(siteEndRow))
    Call AddListRefSub.setRangeBoard(myRange)

    myRange.Select
    With Selection.Interior
        .colorIndex = xlColorIndexNone
    End With
    vSiteSheet.Range("A" + CStr(SiteStartRow) + ":" + "A" + CStr(siteEndRow)).Merge
End Sub

'Public Sub copyControllerTitle_1(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long)
'    Dim vSiteIndex As Long
'    Dim vListColumnNum As Long
'    Dim vSiteName As String
'    Dim vSiteSheet As Worksheet
'    Dim myRange As Range
'
'    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
'    vSiteName = ChosenSiteArray(0)
'    If ("" = vSiteName) Then
'        Exit Sub
'    End If
'    Call WriteDatabyRow(vListSht, vSiteName, g_TitleRowNo, vListColumnNum - 1, SiteStartRow)
'    Set vSiteSheet = Sheets(vSiteName)
'    Call WriteSheetName(vSiteSheet, SiteStartRow, strSheetName)
'    '关闭备注大小调整，提示刷表格效率
'    Set myRange = vSiteSheet.Range("B" + CStr(SiteStartRow) + ":" + C(vListColumnNum - 1) + CStr(SiteStartRow))
'    Call refreshComment(myRange)
'End Sub

'Public Sub copyControllerTitle(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long)
'    Dim vSiteIndex As Long
'    Dim vListColumnNum As Long
'    Dim vSiteName As String
'    Dim vSiteSheet As Worksheet
'    Dim myRange As Range
'
'    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
'    vSiteName = ChosenSiteArray(0)
'    If ("" = vSiteName) Then
'        Exit Sub
'    End If
'    Call WriteDatabyRow(vListSht, vSiteName, g_TitleRowNo, vListColumnNum, SiteStartRow)
'    Set vSiteSheet = Sheets(vSiteName)
'    Call WriteSheetName(vSiteSheet, SiteStartRow, strSheetName)
'    '关闭备注大小调整，提示刷表格效率
'    Set myRange = vSiteSheet.Range("B" + CStr(SiteStartRow) + ":" + C(vListColumnNum) + CStr(SiteStartRow))
'    Call refreshComment(myRange)
'End Sub

Public Sub CopyControllerData(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal vMocEndRowNum As Long, ByVal SiteStartRow As Long, ByVal IsEmptyDataSheet As Boolean)
    Dim vSiteIndex As Long
    Dim vSiteName As String
    Dim vSiteSheet As Worksheet
    Dim vSiteLastRow As Long
    Dim vListColumnNum As Long
    Dim vRowIndex As Long
    Dim vSiteList As String
    Dim myRange As Range
    
    Dim ctrlNameCol As Long
    ctrlNameCol = controllerNameColNum(vListSht)
    
    '最后一列是Site List,不需要拷贝
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    For vSiteIndex = 0 To UBound(ChosenSiteArray)
        vSiteName = ChosenSiteArray(vSiteIndex)
        If ("" = vSiteName) Then
            Exit Sub
        End If
        
        Set vSiteSheet = Sheets(vSiteName)
        vSiteLastRow = SiteStartRow
        For vRowIndex = g_TitleRowNo To vMocEndRowNum
            vSiteList = vListSht.Cells(vRowIndex, vListColumnNum)
            If (vRowIndex = g_TitleRowNo) Then
                vSiteLastRow = vSiteLastRow + 1
            ElseIf IsSharedbySite(vSiteName, vSiteList, strSheetName, vListColumnNum) Then
                Call WriteDatabyRow(vListSht, vSiteName, vRowIndex, ctrlNameCol + 1, vListColumnNum - 1, vSiteLastRow)
                vSiteLastRow = vSiteLastRow + 1
            End If
            If (IsEmptyDataSheet) Then
                vRowIndex = vMocEndRowNum + 1
            End If
        Next vRowIndex
        
    Next vSiteIndex
End Sub

Private Function IsSharedbySite(ByVal vSiteName As String, ByVal vSiteList As String, ByVal vMocSheetName As String, ByVal mocColumnNo As String)
    IsSharedbySite = False
    If ("" = vSiteList) Then
        IsSharedbySite = True
        Exit Function
    End If
    
    Dim vSiteArray() As String
    Dim vSiteIndex As Long
    Dim vTmpSite As String
    
    vSiteArray() = Split(vSiteList, ",")
    For vSiteIndex = 0 To UBound(vSiteArray)
        vTmpSite = Trim(vSiteArray(vSiteIndex))
        If ("" = vTmpSite) Then
            Exit Function
        End If
        
        Dim nodeNames As Collection
        Set nodeNames = getSiteSheetNames(vTmpSite, ThisWorkbook.Worksheets(vMocSheetName), mocColumnNo)
        Dim nodeName
        For Each nodeName In nodeNames
            If vSiteName = nodeName Then
                IsSharedbySite = True
                Exit Function
            End If
        Next
    Next vSiteIndex
End Function

Public Sub CopySiteData(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal vMocEndRowNum As Long, ByVal SiteStartRow As Long, ByVal IsEmptyDataSheet As Boolean)
    Dim vRowIndex As Long
    Dim vNodeBName As String
    Dim vNodeBSheet As Worksheet
    Dim vSiteSheetName As String
    Dim vSiteLastRow As Long
    Dim vLastNodeName As String
    Dim vListColumnNum As Long
    Dim siteLastRowMap As New CMap '以每个基站为Key值，以最后一行空行为索引
    'Copy title first
    'Call CopySiteTitle(vListSht, strSheetName, SiteStartRow)
    
    If (IsEmptyDataSheet) Then
        Exit Sub
    End If
    
    'name colum may not be the first colum
    Dim siteNameCol As Long
    siteNameCol = siteNameColNum(vListSht)
    
    vLastNodeName = ""
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    vSiteLastRow = SiteStartRow + 1
    For vRowIndex = g_TitleRowNo + 1 To vMocEndRowNum
        vNodeBName = Trim(GetCell(vListSht, vRowIndex, siteNameCol))
        If ("" <> vNodeBName) Then
            Dim siteSheetNames As Collection
            Set siteSheetNames = getSiteSheetNames(vNodeBName, Sheets(strSheetName), siteNameCol)
            Dim shtName
            For Each shtName In siteSheetNames
                If ("" <> shtName And IsChosenSite(CStr(shtName))) Then
                    Set vNodeBSheet = Sheets(shtName)
                    
                    If Not siteLastRowMap.hasKey(shtName) Then
                        vSiteLastRow = SiteStartRow + 1
                        Call siteLastRowMap.SetAt(shtName, CStr(vSiteLastRow))
                    Else
                        vSiteLastRow = CLng(siteLastRowMap.GetAt(shtName))
                    End If
    
                    '拷贝数据
                    Call WriteDatabyRow(vListSht, shtName, vRowIndex, siteNameCol + 1, vListColumnNum, vSiteLastRow)
                    vSiteLastRow = vSiteLastRow + 1
                    
                    Call siteLastRowMap.SetAt(shtName, CStr(vSiteLastRow)) '更新最后一行空行的记录
                End If
            Next
        End If
    Next vRowIndex
End Sub

'Public Sub CopySiteTitle(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long)
'    Dim vSiteIndex As Long
'    Dim vListColumnNum As Long
'    Dim vSiteName As String
'    Dim vSiteSheet As Worksheet
'    Dim myRange As Range
'
'    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
'    'For vSiteIndex = 0 To UBound(ChosenSiteArray)
'        vSiteName = ChosenSiteArray(0)
'        If ("" = vSiteName) Then
'            Exit Sub
'        End If
'
'        Call WriteDatabyRow(vListSht, vSiteName, g_TitleRowNo, vListColumnNum, SiteStartRow)
'
'        Set vSiteSheet = Sheets(vSiteName)
'        Call WriteSheetName(vSiteSheet, SiteStartRow, strSheetName)
'
'        '关闭备注大小调整，提示刷表格效率
'        Set myRange = vSiteSheet.Range("B" + CStr(SiteStartRow) + ":" + C(vListColumnNum) + CStr(SiteStartRow))
'        Call refreshComment(myRange)
'    'Next vSiteIndex
'End Sub

Private Sub refreshComment(ByVal myRange As Range)
    On Error Resume Next
    Dim cell As Range
    For Each cell In myRange
        cell.comment.Shape.TextFrame.AutoSize = False
    Next
End Sub

'根据function名称查找Node的页签名称，单制式的时候，页签名就是Function名称，共主控的时候，页签名是Node的名称
Public Function getSiteSheetNames(ByVal functionName As String, ByRef mocSheet As Worksheet, ByVal functionNameColNo As Long) As Collection
    Set getSiteSheetNames = New Collection

    If getNeType = NeType_MRAT Then
        Set getSiteSheetNames = getNodeNamesByRatName(functionName)
        If (getSiteSheetNames.count <> 0) Then
            Exit Function
        End If
    ElseIf IsSheetExist(functionName) Then
        getSiteSheetNames.Add Item:=functionName, key:=functionName
        Exit Function
    End If

    Dim functionNameColName As String
    Dim transportSht As Worksheet
    Dim functionNameColNumInTransportSht As Long
    
    functionNameColName = GetCell(mocSheet, 2, functionNameColNo)
    If (0 = InStr(functionNameColName, "*")) Then
        functionNameColName = "*" + functionNameColName
    End If
    
    Set transportSht = Sheets(getResByKey("BaseTransPort"))
    
    functionNameColNumInTransportSht = colNumByAttr(transportSht, functionNameColName)
    If functionNameColNumInTransportSht = -1 Then
        Exit Function
    End If
    
    Dim nodeNameColNo As Long
    nodeNameColNo = siteNameColNum(transportSht)
    If nodeNameColNo = -1 Then Exit Function
    
    Dim tmpFunctionName As String
    Dim nodeName As String
    Dim rowIdx As Integer
    For rowIdx = 3 To GetSheetUsedRows(transportSht)
        tmpFunctionName = Trim(GetCell(transportSht, rowIdx, functionNameColNumInTransportSht))
        If tmpFunctionName = functionName Then
            nodeName = Trim(GetCell(transportSht, rowIdx, nodeNameColNo))
            If Len(nodeName) > 31 Then
                Dim siteNameRefCol As Long
                siteNameRefCol = colNumByAttr(transportSht, "Sheet Name for Site")
                If siteNameRefCol = -1 Then
                    siteNameRefCol = colNumByAttr(transportSht, getResByKey("A160"))
                End If
                
                If siteNameRefCol <> -1 Then
                   nodeName = Trim(GetCell(transportSht, rowIdx, siteNameRefCol))
                End If
            End If
            getSiteSheetNames.Add Item:=nodeName, key:=nodeName
            
            NodeNameArray(g_NodeIndex) = Trim(nodeName)
            RatNameArray(g_NodeIndex) = Trim(functionName)
            g_NodeIndex = g_NodeIndex + 1
        End If
    Next

'    Dim siteNameCol As Long
'    siteNameCol = siteNameColNum(transportSht)
'    If siteNameCol = -1 Then Exit Function
'
'    getSiteSheetName = Trim(GetCell(transportSht, findRowNo, siteNameCol))
'    If Len(getSiteSheetName) > 31 Then
'        Dim siteNameRefCol As Long
'        siteNameRefCol = colNumByAttr(transportSht, "Sheet Name for Site")
'        If siteNameRefCol = -1 Then
'            siteNameRefCol = colNumByAttr(transportSht, getResByKey("A160"))
'        End If
'
'        If siteNameRefCol <> -1 Then
'           getSiteSheetName = Trim(GetCell(transportSht, findRowNo, siteNameRefCol))
'        End If
'    End If
'    NodeNameArray(g_NodeIndex) = Trim(getSiteSheetName)
'    RatNameArray(g_NodeIndex) = Trim(siteName)
'    g_NodeIndex = g_NodeIndex + 1
End Function

'Public Function Get_RefColbyColumnName(sheetName As String, recordRow As Long, ColValue As String) As Long
'    Dim vSiteSheet As Worksheet
'    Dim vColIndex As Long
'    Dim vListColumnNum As Long
'    Dim vCell As String
'
'    Get_RefColbyColumnName = 0
'    Set vSiteSheet = Sheets(sheetName)
'    vListColumnNum = GetColNumByRowIndex(vSiteSheet)
'
'    For vColIndex = 1 To vListColumnNum
'        vCell = vSiteSheet.Cells(recordRow, vColIndex)
'        If (vCell = ColValue) Then
'            Get_RefColbyColumnName = vColIndex
'            Exit Function
'        End If
'    Next vColIndex
'End Function

'there may be different Co-MPT sites with the same function name(rat name)
Public Function getNodeNamesByRatName(ByVal ratName As String) As Collection
    Set getNodeNamesByRatName = New Collection
    
    Dim ratIndex As Long
    Dim tmpRatName As String
    Dim nodeName As String
    
    For ratIndex = 0 To UBound(ChosenSiteArray)
        tmpRatName = RatNameArray(ratIndex)
        If ("" = tmpRatName) Then
            Exit Function
        End If
        
        If (ratName = tmpRatName) Then
            nodeName = NodeNameArray(ratIndex)
            getNodeNamesByRatName.Add Item:=nodeName, key:=nodeName
        End If
    Next
End Function

'Public Function GetRowNumbyValue(ByVal value As String, ByVal columnNum As Long) As Long
'    Dim baseTransSheetName As String
'    baseTransSheetName = getResByKey("BaseTransPort")
'
'    Dim baseTransSheet As Worksheet
'    Set baseTransSheet = Sheets(baseTransSheetName)
'
'    Dim index As Long
'    Dim endRowNum As Long
'    endRowNum = baseTransSheet.UsedRange.rows.count
'    For index = 3 To endRowNum
'        If (value = GetCell(baseTransSheet, index, columnNum)) Then
'            GetRowNumbyValue = index
'            Exit Function
'        End If
'    Next index
'    GetRowNumbyValue = 0
'End Function

'将MOC(Sheet页名称)写入到NodeB表
Public Sub WriteSheetName(ByVal vNodeBSheet As Worksheet, ByVal vNodeBLastRow As Long, ByVal strSheetName As String)
     vNodeBSheet.Cells(vNodeBLastRow, 1) = strSheetName
     With vNodeBSheet.Cells(vNodeBLastRow, 1)
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
End Sub

'将数据按行拷贝到NodeB表
Public Sub WriteDatabyRow(ByVal vListSht As Worksheet, ByVal vSiteName As String, ByVal vRowIndex As Long, ByVal cpStartCol As Long, ByVal cpEndCol As Long, ByVal vSiteWrtRow As Long)
    Dim vSiteSheet As Worksheet
    Set vSiteSheet = Sheets(vSiteName)
    
    Dim dstStartCol As Long
    dstStartCol = 2
    
    'work sheet with operation, column 2 is operation
    If isOperationWs(vListSht) Then
        dstStartCol = 3
    End If
    
    vSiteSheet.Activate
    vSiteSheet.Cells(vSiteWrtRow, dstStartCol).Select
    vListSht.Range(C(cpStartCol) + CStr(vRowIndex) + ":" + C(cpEndCol) + CStr(vRowIndex)).Copy Destination:=vSiteSheet.Cells(vSiteWrtRow, dstStartCol)
    
    'copy operation
    If dstStartCol = 3 Then
        Dim opCol As Long
        opCol = operationColNum(vListSht)
        If opCol <> -1 Then
            vListSht.Cells(vRowIndex, opCol).Copy Destination:=vSiteSheet.Cells(vSiteWrtRow, 2)
        End If
    End If
End Sub


'获取Sheet页某行的列表,多了最后一个空列，使用的时候需要-1
Public Function GetColNumByRowIndex(ByVal vWorkSht As Worksheet) As Long
    Dim vColumnIndex As Long
    vColumnIndex = 1
    
    While ("" <> Trim(GetCell(vWorkSht, g_TitleRowNo, vColumnIndex)))
        vColumnIndex = vColumnIndex + 1
    Wend
    
    GetColNumByRowIndex = vColumnIndex
End Function
'chenjun end

Public Function IsSheetExist(sheetName As String) As Boolean
    Dim SheetNum, SheetCount As Long 'SheetCount每个原始数据文件的Sheet页总数
    SheetCount = ActiveWorkbook.Worksheets.count   '共有几个Sheet页
    For SheetNum = 1 To SheetCount
        If UCase(Worksheets(SheetNum).name) = UCase(sheetName) Then
            IsSheetExist = True
            Exit Function
        End If
    Next SheetNum
    IsSheetExist = False
End Function

Public Sub CreateNewSiteSheet_i(NodeBName As String)
    'Application.StatusBar = "create sheet:[" + NodeBName + "] begin"
    'If IsSheetExist(NodeBName) Then
    '    Sheets(NodeBName).Delete
    'End If
    
    Sheets.Add after:=ThisWorkbook.ActiveSheet
    ThisWorkbook.ActiveSheet.name = NodeBName
    
    ThisWorkbook.ActiveSheet.Tab.colorIndex = BluePrintSheetColor
    
    With ThisWorkbook.ActiveSheet.Cells.Interior
        .colorIndex = 15
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
End Sub

Public Function CreateNewSiteSheet()
    Dim siteName As String
    
    Dim siteIndex As Long
    siteIndex = 0
    
    Dim Iter As Long
    Sheets(GetMainSheetName()).Activate

    siteName = ChosenSiteArray(0)
    If siteName = "" Then
        Exit Function
    End If
    
    Call CreateNewSiteSheet_i(siteName)
End Function

'按第基站名列算最大有值行
Public Function GetSheetUsedRows(sheet As Worksheet) As Long
    GetSheetUsedRows = 0
    
    Dim nameCol As Long
    nameCol = siteNameColNum(sheet)
    If nameCol = -1 Then
        nameCol = controllerNameColNum(sheet)
    End If
    
    If nameCol = -1 Then nameCol = 1
    
    GetSheetUsedRows = sheet.Range(C(nameCol) & "1048576").End(xlUp).row
End Function

'按第一行算最大有值列
Public Function GetSheetUsedColumns(sheet As Worksheet) As Long
  Dim MaxColumn As Long
  MaxColumn = 1
  
  Do While Trim(sheet.Cells(2, MaxColumn + 1)) <> ""
    MaxColumn = MaxColumn + 1
  Loop
  
  GetSheetUsedColumns = MaxColumn
End Function

'Public Function GetLastRow(sheetName) As Long
'        GetLastRow = Worksheets(sheetName).UsedRange.rows.row
'End Function

''copy Pattern页后，定位到NodeB页，返回title行数
'Public Function LocateLastRow(NodeBName) As Long
'    Dim sheet As Worksheet
'    Set sheet = Sheets(NodeBName)
'
'    Dim NodeBBeginRow As Long
'    Dim NodeBEndRow As Long
'    NodeBBeginRow = 1
'    NodeBEndRow = 1000
'
'    Dim NodeBRow As Long
'    NodeBRow = NodeBBeginRow
'
'    Do While NodeBRow <= NodeBEndRow
'        If Application.WorksheetFunction.CountA(sheet.rows(NodeBRow)) = 0 Then 'when the row is empty
'            sheet.Activate
'            sheet.Cells(NodeBRow, 2).Select
'            LocateLastRow = NodeBRow
'            Exit Function
'        End If
'        NodeBRow = NodeBRow + 1
'    Loop
'End Function

'Public Function GetMocNameFromPatternName(PatternName As String)
'    GetMocNameFromPatternName = Replace(PatternName, "Pattern", "")
'End Function

'Public Sub ConvertPattern_i(sheetName As String, NodeBName As String)
'    Application.StatusBar = "ConvertPattern_i:" + sheetName + "NodeBName:" + NodeBName
'
'    Dim PatternSheet As Worksheet
'    Set PatternSheet = Sheets(sheetName)
'    Dim PatternSheetBeginRow As Long
'    Dim PatternSheetEndRow As Long
'    PatternSheetBeginRow = 2
'    PatternSheetEndRow = GetSheetUsedRows(PatternSheet)
'
'    Dim PatternSheetEndColumn As Long
'    PatternSheetEndColumn = GetSheetUsedColumns(PatternSheet)
'
'    If PatternSheetEndRow > 2 Then
'        PatternSheet.Activate
'        PatternSheet.Range("A2:" + C(PatternSheetEndColumn) + CStr(PatternSheetEndRow)).Select
'        Selection.Copy
'
'        Dim TitleRow As Long
'        TitleRow = LocateLastRow(NodeBName)
'        ThisWorkbook.ActiveSheet.Paste
'        ThisWorkbook.ActiveSheet.Cells(TitleRow, 1) = GetMocNameFromPatternName(sheetName)
'
'        With ThisWorkbook.ActiveSheet.Cells(TitleRow, 1)
'            .Merge
'            .ColumnWidth = 20
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'            .Font.name = "Microsoft Sans Serif"
'            .Font.Bold = True
'            .Font.Size = 10
'            .Font.Underline = xlUnderlineStyleNone
'            .Font.colorIndex = 2
'            .WrapText = True
'            .Interior.colorIndex = 9
'            .Interior.Pattern = xlGray8
'            .Interior.PatternColorIndex = xlAutomatic
'        End With
'    End If
'End Sub

'Public Sub ConvertPattern(PatternSheetName As String)
'    Application.StatusBar = "ConvertPattern:" + PatternSheetName
'
'    Dim siteName As String
'    Dim Iter As Long
'    For Iter = 0 To MaxChosenSiteNum - 1
'        siteName = ChosenSiteArray(Iter)
'        If siteName = "" Then
'            Exit For
'        End If
'        Call ConvertPattern_i(PatternSheetName, siteName)
'    Next Iter
'End Sub

'Public Function IsPatternLink(str As String) As Boolean
'    IsPatternLink = False
'
'    Dim ArrData() As String
'    ArrData = Split(str, "\")
'    If UBound(ArrData) >= 2 Then
'        IsPatternLink = True
'    End If
'End Function

'Public Function GetValueFromBaseSheet(NodeBName As String, RealColumn As Long, value As String) As String
'    On Error GoTo ErrorHandler:
'    Dim sheet As Worksheet
'    Set sheet = Sheets(getResByKey("BaseTransPort"))
'
'    Dim NodeBBeginRow As Long
'    Dim NodeBEndRow As Long
'    NodeBBeginRow = g_NodeBBeginRow
'    NodeBEndRow = sheet.UsedRange.rows.count
'
'    Dim NodeBNameColumn As Long
'    NodeBNameColumn = 1
'
'    Dim NodeBRow As Long
'    NodeBRow = NodeBBeginRow
'
'    Do While NodeBBeginRow <= NodeBEndRow And NodeBBeginRow < MaxChosenSiteNum
'        If sheet.Cells(NodeBBeginRow, NodeBNameColumn) = NodeBName Then
'            GetValueFromBaseSheet = sheet.Cells(NodeBBeginRow, RealColumn)
'            Exit Function
'        End If
'        NodeBBeginRow = NodeBBeginRow + 1
'    Loop
'    GetValueFromBaseSheet = value
'    Exit Function
'ErrorHandler:
'    GetValueFromBaseSheet = value
'End Function

'Public Function GetPatternRealValue(NodeBName As String, value As String) As String
'    Dim Vec() As String
'    Dim sheetName As String
'    Dim GroupName As String
'    Dim colName As String
'
'    Vec = Split(value, "\")
'    sheetName = Vec(0)
'    GroupName = Vec(1)
'    colName = Vec(2)
'
'    Dim RealColumn As Long
'    RealColumn = Get_RefCol(sheetName, 2, GroupName, colName)
'
'    GetPatternRealValue = GetValueFromBaseSheet(NodeBName, RealColumn, value)
'End Function

'Public Sub RefreshPatternValue_i(NodeBName As String)
'    Application.StatusBar = "refresh nodeb:" + NodeBName
'
'    Dim NodeBSheet As Worksheet
'    Set NodeBSheet = Sheets(NodeBName)
'
'    Dim endRow As Long
'    Dim EndColumn As Long
'    endRow = NodeBSheet.UsedRange.rows.count
'    EndColumn = NodeBSheet.UsedRange.columns.count
'
'    Dim IterRow As Long
'    Dim IterColumn As Long
'
'    Dim value As String
'    For IterRow = 1 To endRow
'        For IterColumn = 1 To EndColumn
'            value = NodeBSheet.Cells(IterRow, IterColumn)
'            NodeBSheet.Cells(IterRow, IterColumn).Hyperlinks.Delete
'            If value <> "" And IsPatternLink(value) Then
'                With NodeBSheet.Cells(IterRow, IterColumn).Validation
'                    .Delete
'                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
'                    .inputTitle = getResByKey("Reference Address")
'                    .inputMessage = NodeBSheet.Cells(IterRow, IterColumn)
'                    .ShowInput = True
'                    .ShowError = False
'                End With
'                NodeBSheet.Cells(IterRow, IterColumn) = GetPatternRealValue(NodeBName, value)
'                With NodeBSheet.Cells(IterRow, IterColumn).Interior
'                    .colorIndex = HyperLinkColorIndex
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
'                End With
'            End If
'        Next IterColumn
'    Next IterRow
'End Sub

''针对每个NodeB Sheet页，进行PatternLink值的刷新
'Public Sub RefreshPatternValue()
'    Dim siteName As String
'    Dim Iter As Long
'    For Iter = 0 To MaxChosenSiteNum - 1
'        siteName = ChosenSiteArray(Iter)
'        If siteName = "" Then
'            Exit For
'        End If
'
'        Call RefreshPatternValue_i(siteName)
'    Next Iter
'End Sub

Public Function JudgingEmptyDataSheet(ListSheetName As String) As Boolean
    Dim sheet As Worksheet
    Set sheet = Worksheets(ListSheetName)

    JudgingEmptyDataSheet = True
    
    If sheet.UsedRange.rows.count > 2 Then
        JudgingEmptyDataSheet = False
    End If
End Function

Public Sub writeData()
    Dim startRow As Long
    Dim endRow As Long
    startRow = 1
    Dim IsEmptyDataSheet As Boolean
    Dim Iter As Long
    Dim ListSheetName As String
    Dim sheetDef As Worksheet
    Set sheetDef = Worksheets("SHEET DEF")
    
    Dim shtNameCol, startRowCol, endRowCol As Long
    shtNameCol = shtNameColNumInShtDef
    startRowCol = startRowColNumInShtDef
    endRowCol = endRowColNumInShtDef
    If shtNameCol = -1 Or startRowCol = -1 Or endRowCol = -1 Then Exit Sub
    
    For Iter = 2 To sheetDef.Range("a1048576").End(xlUp).row
        If sheetDef.Cells(Iter, startRowCol).value <> "" And sheetDef.Cells(Iter, endRowCol).value <> "" Then
            startRow = CStr(sheetDef.Cells(Iter, startRowCol).value)
            endRow = CStr(sheetDef.Cells(Iter, endRowCol).value)
            ListSheetName = sheetDef.Cells(Iter, shtNameCol).value
            If IsSheetExist(ListSheetName) Then
                If FILE_TYPE = "1" Or Not isTrasnPortSheet(ListSheetName) Then
                    IsEmptyDataSheet = JudgingEmptyDataSheet(ListSheetName)
                    Call writeListData(ListSheetName, startRow, endRow, IsEmptyDataSheet)
                    Application.DisplayAlerts = False
                    Sheets(ListSheetName).Delete
                    Application.DisplayAlerts = True
                End If
            End If
        End If
    Next
End Sub

Public Sub writeListData(ByVal strSheetName As String, ByVal startRow As Long, ByVal endRow As Long, ByVal IsEmptyDataSheet As Boolean)
    Dim vListSht As Worksheet
    Dim vName As String
    Dim vMocEndRowNum As Long
    Dim nameCol As Long
    
    Set vListSht = Sheets(strSheetName)
    
    'the first column may not be the site name or controller name
    nameCol = siteNameColNum(vListSht)
    If nameCol = -1 Then
        nameCol = controllerNameColNum(vListSht)
    End If
    
    If nameCol = -1 Then Exit Sub
    
    vName = GetCell(vListSht, g_TitleRowNo, nameCol)
    '只处理NodeB&BTS对象
    If (Not is_SiteorController(vName)) Or (isTrasnPortSheet(strSheetName) And FILE_TYPE = "0") Then
        Exit Sub
    End If
    
    vMocEndRowNum = GetSheetUsedRows(vListSht)
    If (vMocEndRowNum <= g_TitleRowNo) Then
        vMocEndRowNum = g_TitleRowNo + 1
    End If

    If is_Controller(vName) Then
        Call CopyControllerData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet)
    Else
       Call CopySiteData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet)
    End If
End Sub

Public Sub InvokeConversion()
    Dim startRow As Long
    Dim endRow As Long
    startRow = 1
    Dim IsEmptyDataSheet As Boolean
    Dim Iter As Long
    Dim ListSheetName As String
    Dim sheetDef As Worksheet
    Set sheetDef = Worksheets("SHEET DEF")
    
    'write starRow and endRow
    Call WriteSheeDefTitle
    
    Dim shtNameCol, shtTypeCol, startRowCol, endRowCol As Long
    shtNameCol = shtNameColNumInShtDef
    shtTypeCol = shtTypeColNumInShtDef
    startRowCol = startRowColNumInShtDef
    endRowCol = endRowColNumInShtDef
    
    Dim bscFlag As Boolean, rncFlag As Boolean, btsFlag As Boolean, alreadyAddEmptyRow As Boolean
    bscFlag = False
    rncFlag = False
    btsFlag = False
    alreadyAddEmptyRow = False
    
    For Iter = 2 To sheetDef.Range("a1048576").End(xlUp).row
        If UCase(sheetDef.Cells(Iter, shtTypeCol).value) = UCase("List") Then
            ListSheetName = sheetDef.Cells(Iter, shtNameCol).value
            If IsSheetExist(ListSheetName) Then
                If FILE_TYPE = "1" Or Not isTrasnPortSheet(ListSheetName) Then
                    Dim tmpSht As Worksheet
                    Set tmpSht = Worksheets(ListSheetName)
                    Dim nameCol As Long
                    nameCol = siteNameColNum(tmpSht)
                    If nameCol = -1 Then
                        nameCol = controllerNameColNum(tmpSht)
                    End If
                    If nameCol = -1 Then GoTo NextLoop
                    
                    If is_SiteorController(GetCell(tmpSht, g_TitleRowNo, nameCol)) Then
                        IsEmptyDataSheet = JudgingEmptyDataSheet(ListSheetName)
                        If (IsEmptyDataSheet) Then
                            endRow = startRow + 1
                        Else
                            endRow = startRow + GetMaxCountPerSite(ListSheetName)
                        End If
                        
                        Call setBscRncBtsFlag(bscFlag, rncFlag, btsFlag, ListSheetName)
                        
                        '如果先找到了控制器再找到基站，则增加一行空行
                        If alreadyAddEmptyRow = False And (bscFlag = True Or rncFlag = True) And btsFlag = True Then
                            startRow = startRow + 1
                            endRow = endRow + 1
                            alreadyAddEmptyRow = True
                        End If
                        Call ConvertList(ListSheetName, startRow, endRow, IsEmptyDataSheet)
                
                        sheetDef.Cells(Iter, startRowCol) = CStr(startRow)
                        sheetDef.Cells(Iter, endRowCol) = CStr(endRow)
                        startRow = endRow + 1
                    End If
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Private Sub setBscRncBtsFlag(ByRef bscFlag As Boolean, ByRef rncFlag As Boolean, ByRef btsFlag As Boolean, ByRef ListSheetName As String)
    If InStr(ListSheetName, "(BSC)") <> 0 Then
        bscFlag = True
    ElseIf InStr(ListSheetName, "(RNC)") <> 0 Then
        rncFlag = True
    Else
        btsFlag = True
    End If
End Sub

'Added by chenjun
Public Sub InitSiteNameArray()
    g_NodeIndex = 0
    Dim siteIndex As Long
    For siteIndex = 0 To UBound(ChosenSiteArray)
        NodeNameArray(siteIndex) = ""
        RatNameArray(siteIndex) = ""
    Next siteIndex
End Sub

'Added by chenjun
Public Sub FormatSiteSheet()
    Dim siteIndex As Long
    Dim SiteSht As Worksheet
    Dim siteName As String
    
    siteIndex = 0
    siteName = ChosenSiteArray(siteIndex)
    'While ("" <> SiteName And IsChosenSite(SiteName))
        Set SiteSht = Sheets(siteName)
        With SiteSht.Cells
            .ColumnWidth = 20
            .WrapText = False
        End With
        '冻结首列
        SiteSht.Select
        With ActiveWindow
            .SplitColumn = 1
            .SplitRow = 0
        End With
        ActiveWindow.FreezePanes = True
        
        'SiteIndex = SiteIndex + 1
        'SiteName = ChosenSiteArray(SiteIndex)
    'Wend
End Sub

Private Sub clearChosenSiteArray()
        Dim index As Long
        For index = LBound(ChosenSiteArray) To UBound(ChosenSiteArray)
                ChosenSiteArray(index) = ""
        Next
End Sub

Private Function siteNameExist(name As String) As Boolean
        Dim index As Long
        For index = LBound(ChosenSiteArray) To UBound(ChosenSiteArray)
                If LCase(ChosenSiteArray(index)) = LCase(name) Then
                    siteNameExist = True
                    Exit Function
                End If
        Next
        siteNameExist = False
End Function

Private Sub makeSheetNameCol(ByRef sheetNameCol As Collection)
    Dim ws As Worksheet
    Dim sheetName As String
    For Each ws In ThisWorkbook.Worksheets
        sheetName = UCase(ws.name)
        If Not Contains(sheetNameCol, sheetName) Then
            sheetNameCol.Add Item:=sheetName, key:=sheetName
        End If
    Next
End Sub

Private Function checkExistingSheetName() As Boolean
    checkExistingSheetName = True
    Call InitTemplateVersion
    
    '得到所有页签名称容器，key和item都是页签名称的大写
    Dim sheetNameCol As New Collection
    Call makeSheetNameCol(sheetNameCol)
    
    Dim NodeBBeginRow As Long, NodeBEndRow As Long, rowNumber As Long
    NodeBBeginRow = g_NodeBBeginRow
    
    Dim baseStationSheet As Worksheet
    Set baseStationSheet = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    NodeBEndRow = GetSheetUsedRows(baseStationSheet)
    
    Dim siteName As String, ucaseSiteName As String
    
    Dim existingNeNameCol As New Collection
    
    Dim nameCol As Long
    nameCol = siteNameColNum(baseStationSheet)
    
    Dim sitesNameCol As New Collection
    For rowNumber = NodeBBeginRow To NodeBEndRow
        siteName = Trim(baseStationSheet.Cells(rowNumber, nameCol).value)
        ucaseSiteName = UCase(siteName)
        If siteName = "" Then GoTo NextLoop
        If Contains(sheetNameCol, ucaseSiteName) Then
            existingNeNameCol.Add Item:=siteName, key:=siteName
        End If
NextLoop:
    Next rowNumber

    '如果count>0，说明有基站名称与已有页签重名，输出结果，不进行转换
    Dim errorMsg As String
    If existingNeNameCol.count <> 0 Then
        checkExistingSheetName = False
        errorMsg = getErrorMsg(existingNeNameCol)
    End If
End Function

Private Function getErrorMsg(ByRef existingNeNameCol As Collection) As String
    Dim maxNumber As Long, index As Long
    Dim errorMsg As String
    If existingNeNameCol.count > 10 Then
        maxNumber = 10
    Else
        maxNumber = existingNeNameCol.count
    End If
    
    For index = 1 To maxNumber
        errorMsg = connectIndividualString(errorMsg, existingNeNameCol(index), vbCrLf)
    Next
    
    If existingNeNameCol.count > 10 Then
        errorMsg = errorMsg & vbCrLf & getResByKey("MSG_TOO_LONG")
    End If
    
    errorMsg = getResByKey("SheetNameExist") & vbCrLf & errorMsg
    
    Call MsgBox(errorMsg, vbCritical, getResByKey("ErrorInfo"))
End Function

'拼接字符串代码
Private Function connectIndividualString(ByVal wholeString As String, ByVal eachString As String, Optional ByVal delimeter As String = ",") As String
    If wholeString = "" Then
        connectIndividualString = eachString
    Else
        connectIndividualString = wholeString & delimeter & eachString
    End If
End Function

Public Function InitChosenSiteArray() As Boolean
    InitChosenSiteArray = True
    
    'fill in ChosenSiteArray
    Dim NodeBBaseSheet As Worksheet
    Set NodeBBaseSheet = Sheets(getResByKey("BaseTransPort"))
    
    Dim NodeBBeginRow As Long
    Dim NodeBEndRow As Long
    NodeBBeginRow = g_NodeBBeginRow
    NodeBEndRow = GetSheetUsedRows(Worksheets(getResByKey("BaseTransPort")))
    If NodeBEndRow > MaxChosenSiteNum Then
        '如果超出最大基站个数，则提示无法转换，退出
        Call MsgBox(getResByKey("SitesExceedsMaxNumber"), vbOKOnly + vbExclamation, getResByKey("ErrorInfo"))
        InitChosenSiteArray = False
        Exit Function
        'NodeBEndRow = MaxChosenSiteNum
    ElseIf NodeBEndRow <= g_NodeBBeginRow - 1 Then 'if no sites has been found, abort with msg
        If bIsEng Then
            MsgBox "No site has been found.", vbOKOnly
        Else
            MsgBox getResByKey("A161"), vbOKOnly
        End If
        InitChosenSiteArray = False
        Exit Function
    End If

    Call clearChosenSiteArray
    
    Dim NodeBNameColumn As Long
    NodeBNameColumn = siteNameColNum(NodeBBaseSheet)
    
    Dim NodeBRow As Long
    NodeBRow = NodeBBeginRow
    
    Dim siteName As String
    Dim IgnoredSiteNum As Long
    IgnoredSiteNum = 0
    Do While NodeBBeginRow <= NodeBEndRow
        siteName = NodeBBaseSheet.Cells(NodeBBeginRow, NodeBNameColumn)
        If siteName = "" Then
            Exit Do
        End If
        
        If IsValidSiteName(siteName) = False Then
            InitChosenSiteArray = False
            Exit Function
        ElseIf siteNameExist(siteName) Then
            If bIsEng Then
                MsgBox "The name of [" + siteName + "] is the same as another NE name, please modify the name.", vbCritical
            Else
                MsgBox getResByKey("A162") + siteName + getResByKey("A163"), vbCritical
            End If
            InitChosenSiteArray = False
            Exit Function
        Else
            If Len(siteName) > 31 Then
              siteName = ForNewSiteName(siteName, NodeBBeginRow)
            End If
            ChosenSiteArray(NodeBBeginRow - g_NodeBBeginRow - IgnoredSiteNum) = Trim(siteName)
            If Not Contains(ChosenSiteNames, siteName) Then
                ChosenSiteNames.Add Item:=(Trim(siteName)), key:=(Trim(siteName))
            End If
            NodeBBeginRow = NodeBBeginRow + 1
        End If
    Loop
    
    '检查是否基站名称与已有List页签重名，有则不进行转换
    If checkExistingSheetName = False Then
        InitChosenSiteArray = False
        Exit Function
    End If
End Function

Function IsValidSiteName(siteName As String) As Boolean
  IsValidSiteName = True
  
  If Len(siteName) > 64 _
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
     IsValidSiteName = False
  End If

End Function

Sub GenIubFormatReport()
    Call InitTemplateVersion
    If InitChosenSiteArray Then
        Dim response
        response = MsgBox(getResByKey("IUBTips"), vbExclamation + vbOKCancel)
        If response = vbCancel Then
            Exit Sub
        End If
    
        Application.ScreenUpdating = False
        
        'Call SortListLine
        Call GenIubFormatReport_i
        
        Application.ScreenUpdating = True
    End If
End Sub

Private Sub createIubStyleWorkSheet(wb As Workbook, iubStyleSheet As Worksheet)
    On Error GoTo ErrorHandler
    Dim index As Long
    Dim count As Long
    Dim arrayIndex As Long
    
    Dim sheet As Worksheet
    
    Dim tempSheetName As String
    tempSheetName = "My Sheet"
    iubStyleSheet.name = tempSheetName
    
    Call setDataRowsWrap(tempSheetName)   '无数据行的自动换行原来没有，新增为其设置上
    
    For index = LBound(ChosenSiteArray) + 1 To UBound(ChosenSiteArray)
           If ChosenSiteArray(index) = "" Then
                Exit For
           End If
           iubStyleSheet.Copy after:=iubStyleSheet
    Next
    count = wb.Worksheets.count
    arrayIndex = 0
    For index = 1 To count
        If wb.Worksheets(index).Tab.colorIndex = BluePrintSheetColor And arrayIndex <= UBound(ChosenSiteArray) Then
            If ChosenSiteArray(arrayIndex) = "" Then
                Exit For
            End If
            wb.Worksheets(index).name = ChosenSiteArray(arrayIndex)
            arrayIndex = arrayIndex + 1
            If Not sheet Is Nothing Then
                wb.Worksheets(index).Move after:=sheet
            End If
            Set sheet = wb.Worksheets(index)
        End If
    Next
    Exit Sub
ErrorHandler:
End Sub

Private Sub setDataRowsWrap(ByRef sheetName As String)
    Dim sheetDefSheet As Worksheet, iubSheet As Worksheet
    Set sheetDefSheet = ThisWorkbook.Worksheets("SHEET DEF")
    Set iubSheet = ThisWorkbook.Worksheets(sheetName)
    
    Dim rowIndex As Long, maxColumnNumber As Long
    Dim startRow As String, endRow As String
    Dim dataRange As Range
    Dim startRowCol, endRowCol As Long
    startRowCol = startRowColNumInShtDef
    endRowCol = endRowColNumInShtDef
    
    '遍历Sheet Def页，把每个IUB页签中的空行都设置上自动换行
    For rowIndex = 2 To sheetDefSheet.Range("A1048576").End(xlUp).row
        startRow = sheetDefSheet.Cells(rowIndex, startRowCol).value
        endRow = sheetDefSheet.Cells(rowIndex, endRowCol).value
        If startRow <> "" And endRow <> "" Then
            maxColumnNumber = iubSheet.Range("XFD" & startRow).End(xlToLeft).column
            Set dataRange = iubSheet.Range("B" & (startRow + 1) & ":" & C(maxColumnNumber) & endRow)
            Call setRangeWrap(dataRange, True)
        End If
    Next rowIndex
End Sub


Sub GenIubFormatReport_i()
    '如果第一个基站名为空，则直接退出，不处理
    If ChosenSiteArray(0) = "" Then Exit Sub
    Call refreshStart
    Call InitSiteNameArray
    Call InitTemplateVersion
    
    '第一步，生成NodeB Sheet页
    Call CreateNewSiteSheet

    '第二步，递归所有数据页，调用ConvertList或ConvertPattern
    Call InvokeConversion

    Call FormatSiteSheet
    Call createIubStyleWorkSheet(ThisWorkbook, ThisWorkbook.Worksheets(ChosenSiteArray(0)))
    Call writeData
    Call ContructPositionInfo
    Call refreshEnd
    If bIsEng Then
        MsgBox "Finish generating view.", vbOKOnly
    Else
        MsgBox getResByKey("A164"), vbOKOnly
    End If
End Sub

Public Sub InitTemplateVersion()
    bIsEng = getResByKey("Cover") = "Cover"
End Sub

Public Function GetBluePrintSheetName() As String '当前只支持一个
    GetBluePrintSheetName = ""
    
    Dim SheetNum, SheetCount As Long
    SheetCount = ActiveWorkbook.Worksheets.count
    For SheetNum = 1 To SheetCount
        If Worksheets(SheetNum).Tab.colorIndex = BluePrintSheetColor Then
            GetBluePrintSheetName = Worksheets(SheetNum).name
            Exit Function
        End If
    Next SheetNum
End Function

Public Function IsBluePrintSheetName(sheetName As String) As Boolean
    IsBluePrintSheetName = (Sheets(sheetName).Tab.colorIndex = BluePrintSheetColor)
End Function

Public Function HasNoBluePrintSheet() As Boolean
    HasNoBluePrintSheet = (GetBluePrintSheetName = "")
End Function

Public Function IsChosenSite(ByVal chkSiteName As String) As Boolean
    IsChosenSite = False
    Dim siteName As String
    Dim Iter As Long
    For Iter = 0 To MaxChosenSiteNum - 1
        siteName = ChosenSiteArray(Iter)
        If siteName = "" Then
            Exit For
        End If
        
        If UCase(siteName) = UCase(chkSiteName) Then
            IsChosenSite = True
            Exit For
        End If
    
    Next Iter
End Function

'Public Sub test()
'    Dim i As Long
'    i = GetMaxCountPerSite(ThisWorkbook.ActiveSheet.name)
'    MsgBox CStr(i), vbOKOnly
'End Sub

Public Function IndexInArray(ByRef SiteNames() As String, siteName As String) As Long
    IndexInArray = -1
    Dim Iter As Long
    For Iter = 0 To UBound(SiteNames)
        If SiteNames(Iter) = siteName Then
            IndexInArray = Iter
            Exit Function
        End If
    Next Iter
End Function

Public Function GetMaxCountPerSite_i(ByRef listSheet As Worksheet) As Long
    'fill in SiteNames and MocCount
    Dim SiteNames(MaxChosenSiteNum) As String
    Dim MocCounts(MaxChosenSiteNum) As Long
    
    Dim siteNameColumn As Long
    siteNameColumn = siteNameColNum(listSheet)
    
    Dim siteBeginRow As Long
    Dim siteEndRow As Long
    siteBeginRow = 3
    siteEndRow = GetSheetUsedRows(listSheet)
    
    Dim index As Long
    Dim siteName As String
    Dim SiteCount As Long
    SiteCount = 0
    Do While siteBeginRow <= siteEndRow
        siteName = listSheet.Cells(siteBeginRow, siteNameColumn)
        
        'If Not Contains(ChosenSiteNames, siteName) Then
        If Not Contains(ChosenSiteNames, siteName) Then
            Dim nodeNames As Collection
            Set nodeNames = getSiteSheetNames(siteName, listSheet, siteNameColumn)
            
            Dim containsFlag As Boolean
            containsFlag = False
            Dim nodeName
            For Each nodeName In nodeNames
                If Contains(ChosenSiteNames, CStr(nodeName)) Then
                    containsFlag = True
                    Exit For
                End If
            Next
            
            If Not containsFlag Then GoTo NextLoop
        End If
        
        index = IndexInArray(SiteNames, siteName)
        If index = -1 Then
            SiteNames(SiteCount) = siteName
            MocCounts(SiteCount) = 1
            SiteCount = SiteCount + 1
        Else
            MocCounts(index) = MocCounts(index) + 1
        End If
NextLoop:
        siteBeginRow = siteBeginRow + 1
    Loop
    
    'get the max count
    Dim MaxCount As Long
    MaxCount = 0
    
    Dim Iter As Long
    For Iter = 0 To UBound(MocCounts)
        If MaxCount < MocCounts(Iter) Then
            MaxCount = MocCounts(Iter)
        End If
    Next Iter
    
    GetMaxCountPerSite_i = MaxCount + 1
End Function

Public Function GetMaxCountPerSite(ListName As String) As Long
    If (Not IsSheetExist(ListName)) Then
        GetMaxCountPerSite = 0
        Exit Function
    End If
    
    Dim listSheet As Worksheet
    Set listSheet = Sheets(ListName)
    
    If siteNameColNum(listSheet) <> -1 Then                 'Site related List
        GetMaxCountPerSite = GetMaxCountPerSite_i(listSheet)
    ElseIf controllerNameColNum(listSheet) <> -1 Then       'Controller related list
        Dim row As Long
        Dim count As Long
        count = listSheet.columns.count
        For row = 1 To listSheet.UsedRange.rows.count
            If Application.WorksheetFunction.CountBlank(listSheet.rows(row)) = count Then
                Exit For
            End If
        Next
        GetMaxCountPerSite = row
    End If
End Function

Public Function GetSheetUsedColumnsForRow(sheet As Worksheet, row As Long) As Long
  Dim MaxColumn As Long
  MaxColumn = 1
  
  Do While Trim(sheet.Cells(row, MaxColumn + 1)) <> ""
    MaxColumn = MaxColumn + 1
  Loop
  
  GetSheetUsedColumnsForRow = MaxColumn
End Function

Public Sub WriteSheeDefTitle()
    Dim sheetDef As Worksheet
    Set sheetDef = Sheets("SHEET DEF")
    
    Dim startColumn As Long
    startColumn = GetSheetUsedColumnsForRow(sheetDef, 1)
    sheetDef.Cells(1, startColumn + 1) = "StartRow"
    sheetDef.Cells(1, startColumn + 2) = "EndRow"
End Sub

Public Sub ContructPositionInfo()
    Dim sheet As Worksheet
    Set sheet = Sheets(getResByKey("BaseTransPort"))
    sheet.Activate
    
    Dim SiteMergeCount As Long
    With sheet
        rows("3:3").Select
        Selection.Insert Shift:=xlDown
        rows("3:3").Hidden = True
        If GetCell(sheet, 2, 2) <> "Sheet Name for Site" And GetCell(sheet, 2, 2) <> getResByKey("A165") Then
            Range("B2").Select
            Selection.EntireColumn.Insert
        
            If bIsEng Then
                Cells(2, 2) = "Referenced Site"
            Else
                Cells(2, 2) = getResByKey("A166")
            End If
        Else
            Range("C2").Select
            Selection.EntireColumn.Insert
        
            If bIsEng Then
                Cells(2, 3) = "Referenced Site"
            Else
                Cells(2, 3) = getResByKey("A167")
            End If
        End If
               
        Range("A1").Select
        
        If Selection.count = 1 Then
            Range("A1:B1").Select
            Selection.Merge
        End If
    End With
End Sub









