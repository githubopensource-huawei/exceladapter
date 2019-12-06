Attribute VB_Name = "AddHyperlinksModule"
Option Explicit

Private sheetsHyperlinksCol As New Collection
Private Const DecouplingSheetName As String = "DecouplingSheet"
Public board_style As String
Public board_pattern As String
Public Const BasebandReferenceBoardNoDelimeter As String = ";" '�������ŵķָ���
Private cellSheetHyperlinkFlag As New CCellSheetHyperlinkFlag '����С��ҳ��ش���ҳ�Ƿ���Ĺ�����ˢ��С��ҳ�����ӵ���


Public Sub addHyperlinksForSheets(ByRef sheet As Worksheet)
    If isBoardStyleSheet(sheet) Then
        Call addBoardStyleHyperlinks_SheetActive(sheet) '��������ʽҳǩ�����õ�Ԫ�����ӳ�����
    ElseIf sheet.name = GetMainSheetName() Then
        Call addTransportSheetHyperlinks_SheetActive(sheet) '������ҳ���ӳ�����
    ElseIf isCellSheet(sheet.name) Then
        'Call mergeSectorAndRXUAnt(sheet)
        Call addCellSheetHyperlinks_SheetActive(sheet) '��С��ҳǩRXU Ant No.���ӳ�����
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
    '�޸ĳ����ӵ�Ԫ��������С�����壬���ܻᵼ���ٶ���������������
    With srcWs.range(srcAddress).Font
        .Size = fontSize
        .name = fontName
    End With
    srcWs.range(srcAddress).WrapText = False '�������С��ҳ�����Զ������÷񣬷�����ѡ�񲻷���
End Sub

Public Sub deleteHyperlink(ByRef srcWs As Worksheet, ByRef srcAddress As String, Optional ByVal fontName As String = "Arial", Optional ByVal fontSize As Long = 10)
    With srcWs.range(srcAddress)
        If .Hyperlinks.count > 0 Then '����õ�Ԫ�����г����ӣ���ɾ��
            .Hyperlinks.Delete
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .NumberFormatLocal = "@" '���õ�Ԫ���ʽΪ�ı�
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
    If Not Contains(sheetsHyperlinksCol, sheetName) Then '���֮ǰû�д������ҳǩ����ҳǩ���Ƽ���col��
        sheetsHyperlinksCol.Add Item:=sheetName, key:=sheetName
        returnFlag = True
    Else '���������ˣ���ֱ���˳��������ظ���ӳ������ˣ����Ч��
        returnFlag = False
    End If
    sheetHyperlinksShouldAdd = returnFlag
End Function

'������ʽҳǩ�����ŵĳ�������ҳǩ����ʱ���
Public Sub addBoardStyleHyperlinks_SheetActive(ByRef ws As Worksheet)
    '������ǵ�����ʽҳǩ���������ҳǩ��˵���ǲ�������񣩣��˳�
    If (Not isBoardStyleSheet(ws)) Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    
    '��ҳǩ�Ƿ�ӹ������ӣ�����ӹ���ֱ���˳�
    'If sheetHyperlinksShouldAdd(ws, sheetsHyperlinksCol) = False And forceRefreshFlag = False Then Exit Sub
    
    Set currentSheet = ws
    
    '����һ�����õ�Ԫ�������
    Dim referenceRangeManager As New CReferenceRangeManager
    Call referenceRangeManager.generateBoardNoReferenceAddressMap
    
    Call initBoardNoManagerPublic '��ʼ�������Ź�����
    
    Call referenceRangeManager.setReferenceAddressHyperlinks(currentSheet, boardNoManager)
End Sub

'����ҳ��ҳǩ����ʱ���Board Style Name�ĳ�����
Public Sub addTransportSheetHyperlinks_SheetActive(ByRef ws As Worksheet)
    '������Ǵ���ҳ���������ҳǩ��˵������������񣩣��˳�
    If ws.name <> GetMainSheetName() Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    
    '��ҳǩ�Ƿ�ӹ������ӣ�����ӹ���ֱ���˳�
    If sheetHyperlinksShouldAdd(ws, sheetsHyperlinksCol) = False Then Exit Sub
    
    Dim boardStyleNameColumnNumber As Long
    boardStyleNameColumnNumber = getBoardStyleNameColumnNumber(ws) '�õ�BoardStyleName�к�
    '���û���ҵ�������ʽ�����ֶΣ����˳�
    If boardStyleNameColumnNumber = -1 Then Exit Sub
    
    Dim maxRowNumber As Long, rowIndex As Long
    maxRowNumber = ws.Cells(65535, boardStyleNameColumnNumber).End(xlUp).row
    
    Dim dstBoardStyleSheetName As String
    For rowIndex = 3 To maxRowNumber
        With ws.Cells(rowIndex, boardStyleNameColumnNumber)
            dstBoardStyleSheetName = .value '�õ�������ʽ����ֵ
            Call addHyperlinkOfBoardStyleName(ws, .address, dstBoardStyleSheetName, "A1")
        End With
    Next rowIndex
End Sub

'С��ҳ�ڼ���ʱ��RXU Ant No.��ӳ�����
Private Sub addCellSheetHyperlinks_SheetActive(ByRef ws As Worksheet)
    Dim applicationFlag As New CApplicationFlags
    Call applicationFlag.init
    
    If (Not isCellSheet(ws.name)) Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub '�������С��ҳ���������ҳǩ��˵����������������˳�
    
    '��ҳǩ�Ƿ�ӹ������ӣ�����ӹ���ֱ���˳�
    If sheetHyperlinksShouldAdd(ws, sheetsHyperlinksCol, True) = False And cellSheetHyperlinkFlag.getSheetFlag(ws.name) = False Then Exit Sub
    
    Dim rxuAntNoName As String
    rxuAntNoName = getRxuAntNoName(ws.name) '�õ�RXU Ant No.������
    If rxuAntNoName = "" Then Exit Sub '���δ�ҵ����������˳�
    
    Dim rxuAntNoNameColumnNumber As Long
    rxuAntNoNameColumnNumber = findCertainValColumnNumber(ws, 2, rxuAntNoName)
    
    If rxuAntNoNameColumnNumber = -1 Then Exit Sub '���û�ҵ����У����˳�
    Dim maxRowNumber As Long, rowIndex As Long
    maxRowNumber = ws.Cells(65535, rxuAntNoNameColumnNumber).End(xlUp).row
    Dim dstBoardStyleSheetName As String
    
    Dim bsBoardStyleMap As New CMap
    Call makeBsBoardStyleMap(bsBoardStyleMap, ws) '�õ���վ���ƺ�BoardStyle��Map
    
    For rowIndex = 3 To maxRowNumber
        With ws.Cells(rowIndex, rxuAntNoNameColumnNumber)
            If .value = "" Then GoTo NextLoop  '���RXU Ant No.Ϊ�գ��������һ��ѭ��
            
            dstBoardStyleSheetName = getBoardStyleName(ws, rowIndex, bsBoardStyleMap) '�õ�С��ҳ�޸�������Ӧ��Board Style
            If dstBoardStyleSheetName = "" Then
                Call deleteHyperlink(ws, .address) '���û���ҵ�BoardStyleName���򽫸õ�Ԫ��ĳ��������
                GoTo NextLoop '���boardStyleNameΪ�գ��������һ��ѭ��
            End If
            Call addHyperlinkOfBoardStyleName(ws, .address, dstBoardStyleSheetName, "A1")
        End With
NextLoop:
    Next rowIndex
    
    Call cellSheetHyperlinkFlag.setSheetFlag(ws.name, False)
End Sub

Private Function getBoardStyleNameColumnNumber(ByRef transportSheet As Worksheet) As Long
    getBoardStyleNameColumnNumber = -1
    Dim flag As Boolean
    flag = getBoardStyleInfo(transportSheet.name) '�ж���RelationDef�Ƿ���Board Style Name���ֶ�
    If flag = False Then
        Exit Function
    End If
    getBoardStyleNameColumnNumber = findCertainValColumnNumber(transportSheet, 2, board_style)
End Function

Private Sub addHyperlinkOfBoardStyleName(ByRef srcWs As Worksheet, ByRef srcAddress As String, ByRef dstWsName As String, ByRef dstAddress As String)
    Dim dstBoardStyleSheet As Worksheet
    If containsASheet(ThisWorkbook, dstWsName) Then '��������������ʽҳǩ������ӳ�����
        Set dstBoardStyleSheet = ThisWorkbook.Worksheets(dstWsName)
        Call addHyperlink(srcWs, srcAddress, dstBoardStyleSheet, dstAddress) '����������ӵ�Ŀ�굥����ʽ��A1��
    End If
End Sub

Private Function getBoardStyleName(ByRef ws As Worksheet, ByRef targetRowNumber As Long, ByRef bsBoardStyleMap As CMap) As String
    getBoardStyleName = ""
    
    Dim bsName As String
    bsName = ws.range("A" & targetRowNumber).value '�õ���վ��
    If bsName = "" Then Exit Function '����û�վ����Ϊ�գ����˳�
    If Not bsBoardStyleMap.hasKey(bsName) Then Exit Function ' �������ҳû�иû�վ���ƣ����˳�
    
    getBoardStyleName = bsBoardStyleMap.GetAt(bsName)
End Function

Private Function getRxuAntNoName(ByRef sheetName As String) As String
    Dim mocName As String, attrName As String
    mocName = "ULOCELLSECTOREQM"
    attrName = "SECTORANTENNA"
    getRxuAntNoName = findColumnFromRelationDef(sheetName, mocName, attrName)
End Function

Private Sub makeBsBoardStyleMap(ByRef bsBoardStyleMap As CMap, ByRef cellsheet As Worksheet)
    Dim bsNameColumnName As String
    bsNameColumnName = cellsheet.range("A2").value '�õ���վ���ƻ�rat�����У��硰*BTS Name��
    
    Dim bsNameColumnNumberInTransportSheet As Long
    Dim transportSheetName As String
    Dim transportSheet As Worksheet
    transportSheetName = GetMainSheetName()
    Set transportSheet = ThisWorkbook.Worksheets(transportSheetName) '�õ�����ҳ
    '�ڴ���ҳ�ҵ�RXU Ant No.�ֶ�����������
    bsNameColumnNumberInTransportSheet = findCertainValColumnNumber(transportSheet, 2, bsNameColumnName)
    If bsNameColumnNumberInTransportSheet = -1 Then Exit Sub '���û�ҵ����˳�
    
    Dim maxRowNumber As Long, boardStyleNameColumnNumber As Long, rowIndex As Long
    maxRowNumber = transportSheet.Cells(65535, bsNameColumnNumberInTransportSheet).End(xlUp).row '�õ�����ҳ�����
    
    boardStyleNameColumnNumber = getBoardStyleNameColumnNumber(transportSheet) '�õ�BoardStyleName�к�
    If boardStyleNameColumnNumber = -1 Then Exit Sub '����ڴ���ҳû�ҵ�BoardStyleName�ֶΣ����˳�
    
    Dim bsNameInTransportSheet As String, boardStyleName As String
    For rowIndex = 3 To maxRowNumber
        bsNameInTransportSheet = transportSheet.Cells(rowIndex, bsNameColumnNumberInTransportSheet).value
        If bsNameInTransportSheet = "" Then GoTo NextLoop
        boardStyleName = transportSheet.Cells(rowIndex, boardStyleNameColumnNumber).value
        
        Call bsBoardStyleMap.SetAt(bsNameInTransportSheet, boardStyleName)
NextLoop:
    Next rowIndex
End Sub


