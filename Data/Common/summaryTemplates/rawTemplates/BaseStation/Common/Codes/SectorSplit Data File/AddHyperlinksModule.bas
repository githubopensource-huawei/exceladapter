Attribute VB_Name = "AddHyperlinksModule"
Option Explicit

Private sheetsHyperlinksCol As New Collection
Private Const DecouplingSheetName As String = "DecouplingSheet"
Public board_style As String
Public board_pattern As String
Public Const BasebandReferenceBoardNoDelimeter As String = ";" '处理单板编号的分隔符

Public Sub addHyperlinksForSheets(ByRef sheet As Worksheet)
    If isBoardStyleSheet(sheet) Then
        Call addBoardStyleHyperlinks_SheetActive(sheet) '给单板样式页签的引用单元格增加超链接
    ElseIf sheet.name = GetMainSheetName() Then
        Call addTransportSheetHyperlinks_SheetActive(sheet) '给传输页增加超链接
    End If
End Sub

Public Sub addHyperlink(ByRef srcWs As Worksheet, ByRef srcAddress As String, ByRef dstWs As Worksheet, ByRef dstAddress As String, Optional ByVal fontName As String = "Arial", Optional ByVal fontSize As Long = 10)
    Dim srcRange As range
    Set srcRange = srcWs.range(srcAddress)

    Dim dstSheetName As String, subAddressString As String
    dstSheetName = dstWs.name
    subAddressString = "'" & dstSheetName & "'!" & dstAddress
    
    srcWs.Hyperlinks.Add Anchor:=srcRange, address:="", _
        subAddress:=subAddressString
    '修改超链接单元格的字体大小和字体，可能会导致速度慢引起性能问题
    With srcWs.range(srcAddress).Font
        .Size = fontSize
        .name = fontName
    End With
    srcWs.range(srcAddress).WrapText = False '如果不是小区页，则将自动换行置否，否则点击选择不方便
End Sub

Public Sub deleteHyperlink(ByRef srcWs As Worksheet, ByRef srcAddress As String, Optional ByVal fontName As String = "Arial", Optional ByVal fontSize As Long = 10)
    With srcWs.range(srcAddress)
        If .Hyperlinks.count > 0 Then '如果该单元格上有超链接，则删除
            .Hyperlinks.Delete
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .NumberFormatLocal = "@" '设置单元格格式为文本
            Call setBorders(srcWs.range(srcAddress))
        End If
        With .Font
            .Size = fontSize
            .name = fontName
        End With
    End With
End Sub

Private Function sheetHyperlinksShouldAdd(ByRef ws As Worksheet, ByRef sheetsHyperlinksCol As Collection, Optional ByRef cellSheetFlag As Boolean = False) As Boolean
    Dim returnFlag As Boolean
    Dim sheetName As String
    sheetName = ws.name
    If Not Contains(sheetsHyperlinksCol, sheetName) Then '如果之前没有处理过该页签，则将页签名称加入col中
        sheetsHyperlinksCol.Add Item:=sheetName, key:=sheetName
        returnFlag = True
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
    
    Set g_CurrentSheet = ws
    
    '声明一个引用单元格管理类
    Dim referenceRangeManager As New CReferenceRangeManager
    Call referenceRangeManager.generateBoardNoReferenceAddressMap
    
    Call initBoardNoManagerPublic '初始化单板编号管理类
    
    Call referenceRangeManager.setReferenceAddressHyperlinks(g_CurrentSheet, boardNoManager)
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
    maxRowNumber = ws.Cells(65535, boardStyleNameColumnNumber).End(xlUp).row
    
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

Private Function getBoardStyleName(ByRef ws As Worksheet, ByRef targetRowNumber As Long, ByRef bsBoardStyleMap As CMap) As String
    getBoardStyleName = ""
    
    Dim bsName As String
    bsName = ws.range("A" & targetRowNumber).value '得到基站名
    If bsName = "" Then Exit Function '如果该基站名称为空，则退出
    If Not bsBoardStyleMap.hasKey(bsName) Then Exit Function ' 如果传输页没有该基站名称，则退出
    
    getBoardStyleName = bsBoardStyleMap.GetAt(bsName)
End Function
