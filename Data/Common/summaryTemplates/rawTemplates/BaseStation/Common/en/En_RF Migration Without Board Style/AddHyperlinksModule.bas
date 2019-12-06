Attribute VB_Name = "AddHyperlinksModule"
Option Explicit

Private sheetsHyperlinksCol As New Collection
Private Const DecouplingSheetName As String = "DecouplingSheet"
Private transportSheetChangeFlag As Boolean
Private cellSheetHyperlinkFlag As New CCellSheetHyperlinkFlag '处理小区页监控传输页是否更改过，以刷新小区页超链接的类

'先将行高设置好，防止加超链接时，自动调整行高速度很慢
Public Sub setHyperlinkSheetsRowHeight()
    Dim ws As Worksheet
    
        'DTS2017011105086
    Dim clipBoardData As DataObject
    Set clipBoardData = New DataObject

    clipBoardData.Clear
    clipBoardData.GetFromClipboard
        
    For Each ws In ThisWorkbook.Worksheets
        If isBoardStyleSheet(ws) Or ws.name = GetMainSheetName() Or isCellSheet(ws.name) Then
            If ws.Rows(1).RowHeight < 14 Then
                ws.Cells.RowHeight = 14
            End If
            ws.Cells.HorizontalAlignment = xlCenter
            ws.Cells.VerticalAlignment = xlTop
        End If
    Next ws
        
        'DTS2017011105086
    With clipBoardData
        .SetText ""
        .PutInClipboard
    End With
End Sub

Public Sub addHyperlinksForSheets(ByRef sheet As Worksheet)
    If isBoardStyleSheet(sheet) Then
        Call addBoardStyleHyperlinks_SheetActive(sheet) '给单板样式页签的引用单元格增加超链接
    ElseIf sheet.name = GetMainSheetName() Then
        Call addTransportSheetHyperlinks_SheetActive(sheet) '给传输页增加超链接
    ElseIf isCellSheet(sheet.name) Or isEuCellSectorEqmSht(sheet.name) Or isEuPrbSectorEqmSht(sheet.name) Then
        Call addCellSheetHyperlinks_SheetActive(sheet) '给小区页签RXU Ant No.增加超链接
    End If
End Sub

'引用单元格超链接的添加
Public Sub addReferenceRangeHyperlinks_SheetChange(ByRef ws As Worksheet, ByRef target As range)
    If target.count <> 1 Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    If isReferenceValue(target.value) Then Exit Sub
    
    Dim groupName As String, columnName As String
    Dim rowNumber As Long, columnNumber As Long
    rowNumber = target.row
    columnNumber = target.column
    
    Dim targetInRecordsRangeFlag As Boolean, targetIsInListBoxFlag As Boolean
    '判断修改的单元格是不是有效单元格
    targetInRecordsRangeFlag = getRangeGroupAndColumnName(ws, rowNumber, columnNumber, groupName, columnName)
    
    Dim referencedString As String
    Dim currentBoardStyleMappingDefData As CBoardStyleMappingDefData
    '判断修改的单元格是不是需要下拉框的
    targetIsInListBoxFlag = getReferecedString(groupName, columnName, referencedString, currentBoardStyleMappingDefData) '判断选定的列是否需要增加自动下拉框

    If targetInRecordsRangeFlag = False Or targetIsInListBoxFlag = False Then
        Exit Sub
    End If
    
    Call initBoardNoManagerPublic
    Dim referencedBoardNoAddress As String
    referencedBoardNoAddress = boardNoManager.getBoardNoAddress(target.value)
    If referencedBoardNoAddress <> "" Then
        Call addHyperlink(ws, target.address, ws, referencedBoardNoAddress)
    Else '如果是空的话， 是把引用单板编号清空了，需要将字体格式设置为正常，以防手工填值时样式不正常
        Call deleteHyperlink(ws, target.address)
        Call setCertainRangeFont(target)
    End If
End Sub

Public Sub setRangeFont(ByRef certainRange As range, ByRef fontName As String)
    certainRange.Font.name = fontName
End Sub

Private Sub setCertainRangeFont(ByRef certainRange As range)
    With certainRange
        .WrapText = True '设置自动换行
        With .Font
            .name = "Arial"
            .Underline = xlUnderlineStyleNone '无下划线
            .colorIndex = -4105 '黑色
        End With
    End With
End Sub

Public Sub addHyperlink(ByRef srcWs As Worksheet, ByRef srcAddress As String, ByRef dstWs As Worksheet, ByRef dstAddress As String, Optional ByVal fontName As String = "Arial", Optional ByVal fontSize As Long = 10)
    On Error Resume Next
    Dim srcRange As range
    Set srcRange = srcWs.range(srcAddress)

    Dim dstSheetName As String, subAddressString As String
    dstSheetName = dstWs.name
    subAddressString = "'" & dstSheetName & "'!" & dstAddress
    
    srcWs.Hyperlinks.Add Anchor:=srcRange, address:="", _
        SubAddress:=subAddressString
    '修改超链接单元格的字体大小和字体，可能会导致速度慢引起性能问题
    'DTS2017011105086
'    Dim clipBoardData As DataObject
'    Set clipBoardData = New DataObject
'
'    clipBoardData.Clear
'    clipBoardData.GetFromClipboard
    
'    With srcWs.range(srcAddress).Font
'        .Size = fontSize
'        .name = fontName
'    End With
    If Not isCellSheet(srcWs.name) Then srcWs.range(srcAddress).WrapText = False '如果不是小区页，则将自动换行置否，否则点击选择不方便
    srcWs.range(srcAddress).WrapText = False
    'DTS2017011105086
'    With clipBoardData
'        .SetText ""
'        .PutInClipboard
'    End With
End Sub

Public Sub deleteHyperlink(ByRef srcWs As Worksheet, ByRef srcAddress As String, Optional ByVal fontName As String = "Arial", Optional ByVal fontSize As Long = 10)
    On Error Resume Next
    'DTS2017011105086
'    Dim clipBoardData As DataObject
'    Set clipBoardData = New DataObject
'
'    clipBoardData.Clear
'    clipBoardData.GetFromClipboard
    
    With srcWs.range(srcAddress)
        If .Hyperlinks.count > 0 Then '如果该单元格上有超链接，则删除
            .Hyperlinks.Delete
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'            .WrapText = True
'            .NumberFormatLocal = "@" '设置单元格格式为文本
            Call setBorders(srcWs.range(srcAddress))
        End If
'        With .Font
'            .Size = fontSize
'            .name = fontName
'        End With
    End With
    'DTS2017011105086
'    With clipBoardData
'        .SetText ""
'        .PutInClipboard
'    End With
End Sub

Private Function sheetHyperlinksShouldAdd(ByRef ws As Worksheet, ByRef sheetsHyperlinksCol As Collection, Optional ByRef cellSheetFlag As Boolean = False) As Boolean
    Dim returnFlag As Boolean
    Dim sheetName As String
    sheetName = ws.name
    If Not Contains(sheetsHyperlinksCol, sheetName) Then '如果之前没有处理过该页签，则将页签名称加入col中
        sheetsHyperlinksCol.Add Item:=sheetName, key:=sheetName
        returnFlag = True
        If cellSheetFlag Then Call cellSheetHyperlinkFlag.setSheetFlag(ws.name, True) '如果是小区页，则将小区页增加到处理类中
    Else '如果处理过了，就直接退出，无需重复添加超链接了，提高效率
        returnFlag = False
    End If
    sheetHyperlinksShouldAdd = returnFlag
End Function

'单板样式页签单板编号的超链接在页签激活时添加
Public Sub addBoardStyleHyperlinks_SheetActive(ByRef ws As Worksheet)
    '如果不是单板样式页签或包含解耦页签（说明是不完整表格），退出
    If (Not isBoardStyleSheet(ws)) Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    
    '看页签是否加过超链接，如果加过，直接退出
    'If sheetHyperlinksShouldAdd(ws, sheetsHyperlinksCol) = False And forceRefreshFlag = False Then Exit Sub
    
    Set currentSheet = ws
    
    '声明一个引用单元格管理类
    Dim referenceRangeManager As New CReferenceRangeManager
    Call referenceRangeManager.generateBoardNoReferenceAddressMap
    
    Call initBoardNoManagerPublic '初始化单板编号管理类
    
    Call referenceRangeManager.setReferenceAddressHyperlinks(currentSheet, boardNoManager)
End Sub

'传输页在页签激活时添加Board Style Name的超链接
Public Sub addTransportSheetHyperlinks_SheetActive(ByRef ws As Worksheet)
    '如果不是传输页或包含解耦页签（说明不是完整表格），退出
    If ws.name <> GetMainSheetName() Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    
    '看页签是否加过超链接，如果加过，直接退出
    If sheetHyperlinksShouldAdd(ws, sheetsHyperlinksCol) = False Then Exit Sub
    
    Dim boardStyleNameColumnNumber As Long
    boardStyleNameColumnNumber = getBoardStyleNameColumnNumber(ws) '得到BoardStyleName列号
    '如果没有找到单板样式名称字段，则退出
    If boardStyleNameColumnNumber = -1 Then Exit Sub
    
    Dim maxRowNumber As Long, rowIndex As Long
    maxRowNumber = ws.Cells(1048576, boardStyleNameColumnNumber).End(xlUp).row
    
    Dim dstBoardStyleSheetName As String
    For rowIndex = 3 To maxRowNumber
        With ws.Cells(rowIndex, boardStyleNameColumnNumber)
            dstBoardStyleSheetName = .value '得到单板样式名称值
            Call addHyperlinkOfBoardStyleName(ws, .address, dstBoardStyleSheetName, "A1")
        End With
    Next rowIndex
End Sub

Private Function getBoardStyleNameColumnNumber(ByRef transportSheet As Worksheet) As Long
    getBoardStyleNameColumnNumber = -1
    Dim flag As Boolean
    flag = getBoardStyleInfo(transportSheet.name) '判断在RelationDef是否有Board Style Name的字段
    If flag = False Then
        Exit Function
    End If
    getBoardStyleNameColumnNumber = findCertainValColumnNumber(transportSheet, 2, board_style)
End Function

Private Sub addHyperlinkOfBoardStyleName(ByRef srcWs As Worksheet, ByRef srcAddress As String, ByRef dstWsName As String, ByRef dstAddress As String)
    Dim dstBoardStyleSheet As Worksheet
    If containsASheet(ThisWorkbook, dstWsName) Then '如果有这个单板样式页签，则添加超链接
        Set dstBoardStyleSheet = ThisWorkbook.Worksheets(dstWsName)
        Call addHyperlink(srcWs, srcAddress, dstBoardStyleSheet, dstAddress) '将超链接添加到目标单板样式的A1格
    End If
End Sub

'传输页在Board Style Name值改变时添加超链接
Public Sub addBoardStyleNameHyperlinks_SheetChange(ByRef ws As Worksheet, ByRef target As range)
    If target.count <> 1 Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    Dim flag As Boolean
    flag = getBoardStyleInfo(ws.name) '判断在RelationDef是否有Board Style Name的字段
    If flag = False Then
        Exit Sub
    End If
    
    Dim rowNumber As Long, columnNumber As Long
    rowNumber = target.row
    columnNumber = target.column
    flag = isBoardStyleCol(ws, rowNumber, columnNumber) '判断当前改变的单元格是否是Board Style Name字段
    If flag = False Then Exit Sub
    
    Dim dstBoardStyleSheetName As String
    dstBoardStyleSheetName = target.value
    
    '如果名称为清，即被清空时，则退出
    If dstBoardStyleSheetName = "" Then Exit Sub
    
    Call addHyperlinkOfBoardStyleName(ws, target.address, dstBoardStyleSheetName, "A1")
End Sub

'小区页在RXU Ant No.值改变时添加超链接
Public Sub addRxuAntNoHyperlinks_SheetChange(ByRef ws As Worksheet, ByRef target As range)
    If target.count <> 1 Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    
    Dim targetRowNumber As Long, targetColumnNumber As Long
    targetRowNumber = target.row
    targetColumnNumber = target.column
    
    Dim rxuAntNoName As String
    rxuAntNoName = getRxuAntNoName(ws.name) '得到RXU Ant No.的列名
    If rxuAntNoName = "" Then Exit Sub '如果未找到列名，则退出
    
    Dim rxuAntNoNameColumnNumber As Long
    rxuAntNoNameColumnNumber = findCertainValColumnNumber(ws, 2, rxuAntNoName)
    
    If target.value = "" Or rxuAntNoNameColumnNumber <> targetColumnNumber Then 'targetColumnNumber肯定不为-1，所以如果rxuAntNoNameColumnNumber未找到为-1，则直接退出
        Exit Sub '如果目标单元格为空或改变的列不是RXU Ant No.列，则退出
    End If
    
    Dim bsBoardStyleMap As New CMap
    Call makeBsBoardStyleMap(bsBoardStyleMap, ws) '得到基站名称和BoardStyle的Map
    
    Dim dstBoardStyleSheetName As String
    dstBoardStyleSheetName = getBoardStyleName(ws, targetRowNumber, bsBoardStyleMap) '得到小区页修改行所对应的Board Style
    If dstBoardStyleSheetName = "" Then
        Call deleteHyperlink(ws, target.address) '没有找到对应的BoardStyleName，则将该单元格的超链接清空
        Exit Sub '如果boardStyleName为空，则退出
    End If
    Call addHyperlinkOfBoardStyleName(ws, target.address, dstBoardStyleSheetName, "A1")
End Sub

'小区页在激活时给RXU Ant No.添加超链接
Private Sub addCellSheetHyperlinks_SheetActive(ByRef ws As Worksheet)
    Dim applicationFlag As New CApplicationFlags
    Call applicationFlag.init
    
    If Not (isCellSheet(ws.name) Or isEuCellSectorEqmSht(ws.name) Or isEuPrbSectorEqmSht(ws.name)) Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub '如果不是小区页或包含解耦页签（说明表格不完整），则退出
    
    '看页签是否加过超链接，如果加过，直接退出
    If sheetHyperlinksShouldAdd(ws, sheetsHyperlinksCol, True) = False And cellSheetHyperlinkFlag.getSheetFlag(ws.name) = False Then Exit Sub
    
    Dim rxuAntNoName As String
    rxuAntNoName = getRxuAntNoName(ws.name) '得到RXU Ant No.的列名
    If rxuAntNoName = "" Then Exit Sub '如果未找到列名，则退出
    
    Dim rxuAntNoNameColumnNumber As Long
    rxuAntNoNameColumnNumber = findCertainValColumnNumber(ws, 2, rxuAntNoName)
    
    If rxuAntNoNameColumnNumber = -1 Then Exit Sub '如果没找到该列，则退出
    Dim maxRowNumber As Long, rowIndex As Long
    maxRowNumber = ws.Cells(1048576, rxuAntNoNameColumnNumber).End(xlUp).row
    Dim dstBoardStyleSheetName As String
    
    Dim bsBoardStyleMap As New CMap
    Call makeBsBoardStyleMap(bsBoardStyleMap, ws) '得到基站名称和BoardStyle的Map
    
    For rowIndex = 3 To maxRowNumber
        With ws.Cells(rowIndex, rxuAntNoNameColumnNumber)
            If .value = "" Then GoTo NextLoop  '如果RXU Ant No.为空，则进入下一个循环
            
            dstBoardStyleSheetName = getBoardStyleName(ws, rowIndex, bsBoardStyleMap) '得到小区页修改行所对应的Board Style
            If dstBoardStyleSheetName = "" Then
                Call deleteHyperlink(ws, .address) '如果没有找到BoardStyleName，则将该单元格的超链接清空
                GoTo NextLoop '如果boardStyleName为空，则进入下一个循环
            End If
            Call addHyperlinkOfBoardStyleName(ws, .address, dstBoardStyleSheetName, "A1")
        End With
NextLoop:
    Next rowIndex
    
    Call cellSheetHyperlinkFlag.setSheetFlag(ws.name, False)
End Sub

Public Sub recordTransportSheetChange(ByRef flag As Boolean)
    transportSheetChangeFlag = flag
End Sub

'记录传输页board style是否有改变，如果变了，则在小区页激活时去刷新小区页的天线端口号超链接
Public Sub recordTransportSheetChangeIncludingBoardStyle(ByRef ws As Worksheet, ByRef target As range)
    If cellSheetHyperlinkFlag.allSheetsFlag = True Then Exit Sub '如果传输页的board style列已经记录为修改过，则不需要再次查询记录了
    Dim columnRange As range
    Dim flag As Boolean
    For Each columnRange In target.Columns
        flag = isBoardStyleCol(ws, 3, columnRange.column) '判断该单元格是否是Board Style Name字段
        If flag = True Then
            Call cellSheetHyperlinkFlag.setAllSheetsFlag(True) '设置传输页board style列为修改过，并退出
            Exit Sub
        End If
    Next columnRange
End Sub

Private Sub getAntenneMocNameAndAttrByCellSheetName(ByRef mocName As String, ByRef attrName As String, ByRef CellSheetName As String)
    If CellSheetName = getResByKey("A170") Or CellSheetName = "GSM Cell" Then
        mocName = "GTRXGROUPSECTOREQM"
        attrName = "SECTORANTENNA"
    ElseIf CellSheetName = getResByKey("A171") Or CellSheetName = "UMTS Cell" Then
        mocName = "ULOCELLSECTOREQM"
        attrName = "SECTORANTENNA"
    ElseIf CellSheetName = getResByKey("A172") Or CellSheetName = "LTE Cell" Then
        mocName = "eUCellSectorEqm"
        attrName = "SECTORANTENNA"
    ElseIf CellSheetName = getResByKey("EUCELLSECTOREQM") Or CellSheetName = "EUCELLSECTOREQM" Then
        mocName = "eUCellSectorEqm"
        attrName = "SECTORANTENNA"
    ElseIf CellSheetName = getResByKey("EUPRBSECTOREQM") Or CellSheetName = "EUPRBSECTOREQM" Then
        mocName = "EuPrbSectorEqm"
        attrName = "SECTORANTENNA"
    End If
End Sub

Private Function getRxuAntNoName(ByRef sheetName As String) As String
    Dim mocName As String, attrName As String
    Call getAntenneMocNameAndAttrByCellSheetName(mocName, attrName, sheetName)
    getRxuAntNoName = findColumnFromRelationDef(sheetName, mocName, attrName)
End Function

Private Function getBoardStyleName(ByRef ws As Worksheet, ByRef targetRowNumber As Long, ByRef bsBoardStyleMap As CMap) As String
    getBoardStyleName = ""
    
    Dim bsName As String
    Dim btsNameColIndex As Long
    Dim mocName As String
    Dim attrName As String
    
    CELL_TYPE = cellSheetType(ws.name)
    
    Call getBaseStationMocNameAndAttrName(mocName, attrName)
    btsNameColIndex = getColNum(ws.name, 2, attrName, mocName)
    
'    bsName = ws.range("B" & targetRowNumber).value '得到基站名
    bsName = ws.Cells(targetRowNumber, btsNameColIndex).value '得到基站名
    If bsName = "" Then Exit Function '如果该基站名称为空，则退出
    If Not bsBoardStyleMap.hasKey(bsName) Then Exit Function ' 如果传输页没有该基站名称，则退出
    
    getBoardStyleName = bsBoardStyleMap.GetAt(bsName)
End Function

Private Sub makeBsBoardStyleMap(ByRef bsBoardStyleMap As CMap, ByRef cellsheet As Worksheet)
    Dim bsNameColumnName As String
    Dim btsNameColIndex As Long
    Dim mocName As String
    Dim attrName As String
    
    CELL_TYPE = cellSheetType(cellsheet.name)
    
    Call getBaseStationMocNameAndAttrName(mocName, attrName)
    btsNameColIndex = getColNum(cellsheet.name, 2, attrName, mocName)
    
'    bsNameColumnName = cellsheet.range("B2").value '得到基站名称或rat名称列，如“*BTS Name”
    bsNameColumnName = cellsheet.Cells(2, btsNameColIndex).value '得到基站名称或rat名称列，如“*BTS Name”
    
    Dim bsNameColumnNumberInTransportSheet As Long
    Dim transportSheetName As String
    Dim transportSheet As Worksheet
    transportSheetName = GetMainSheetName()
    Set transportSheet = ThisWorkbook.Worksheets(transportSheetName) '得到传输页
    '在传输页找到RXU Ant No.字段所在列数字
    bsNameColumnNumberInTransportSheet = findCertainValColumnNumber(transportSheet, 2, bsNameColumnName)
    If bsNameColumnNumberInTransportSheet = -1 Then Exit Sub '如果没找到则，退出
    
    Dim maxRowNumber As Long, boardStyleNameColumnNumber As Long, rowIndex As Long
    maxRowNumber = transportSheet.Cells(1048576, bsNameColumnNumberInTransportSheet).End(xlUp).row '得到传输页最大行
    
    boardStyleNameColumnNumber = getBoardStyleNameColumnNumber(transportSheet) '得到BoardStyleName列号
    If boardStyleNameColumnNumber = -1 Then Exit Sub '如果在传输页没找到BoardStyleName字段，则退出
    
    Dim bsNameInTransportSheet As String, boardstyleName As String
    For rowIndex = 3 To maxRowNumber
        bsNameInTransportSheet = transportSheet.Cells(rowIndex, bsNameColumnNumberInTransportSheet).value
        If bsNameInTransportSheet = "" Then GoTo NextLoop
        boardstyleName = transportSheet.Cells(rowIndex, boardStyleNameColumnNumber).value
        
        Call bsBoardStyleMap.SetAt(bsNameInTransportSheet, boardstyleName)
NextLoop:
    Next rowIndex
End Sub

