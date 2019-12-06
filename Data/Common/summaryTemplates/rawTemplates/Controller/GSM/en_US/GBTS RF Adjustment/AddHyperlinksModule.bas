Attribute VB_Name = "AddHyperlinksModule"
Option Explicit

Private sheetsHyperlinksCol As New Collection
Private Const DecouplingSheetName As String = "DecouplingSheet"

Public Sub addHyperlinksForSheets(ByRef sheet As Worksheet)
    If Not hyperLintFlag Then
        Exit Sub
    Else
        If isBoardStyleSheet(sheet) Then
            Call addBoardStyleHyperlinks_SheetActive(sheet) '��������ʽҳǩ�����õ�Ԫ�����ӳ�����
        ElseIf sheet.name = GetMainSheetName() Then
            Call addTransportSheetHyperlinks_SheetActive(sheet) '������ҳ���ӳ�����
        ElseIf isCellSheet(sheet.name) Then
            Call addCellSheetHyperlinks_SheetActive(sheet) '��С��ҳǩRXU Ant No.���ӳ�����
        End If
    End If
End Sub

'���õ�Ԫ�����ӵ����
Public Sub addReferenceRangeHyperlinks_SheetChange(ByRef ws As Worksheet, ByRef Target As range)
    If Target.count <> 1 Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    Dim groupName As String, columnName As String
    Dim rowNumber As Long, columnNumber As Long
    rowNumber = Target.row
    columnNumber = Target.column
    
    Dim targetInRecordsRangeFlag As Boolean, targetIsInListBoxFlag As Boolean
    '�ж��޸ĵĵ�Ԫ���ǲ�����Ч��Ԫ��
    targetInRecordsRangeFlag = getRangeGroupAndColumnName(ws, rowNumber, columnNumber, groupName, columnName)
    
    Dim referencedString As String
    Dim currentBoardStyleMappingDefData As CBoardStyleMappingDefData
    '�ж��޸ĵĵ�Ԫ���ǲ�����Ҫ�������
    targetIsInListBoxFlag = getReferecedString(groupName, columnName, referencedString, currentBoardStyleMappingDefData) '�ж�ѡ�������Ƿ���Ҫ�����Զ�������

    If targetInRecordsRangeFlag = False Or targetIsInListBoxFlag = False Then
        Exit Sub
    End If
    
    Call initBoardNoManagerPublic
    Dim referencedBoardNoAddress As String
    referencedBoardNoAddress = boardNoManager.getBoardNoAddress(Target.value)
    If referencedBoardNoAddress <> "" Then
        Call addHyperlink(ws, Target.address, ws, referencedBoardNoAddress)
    Else '����ǿյĻ��� �ǰ����õ���������ˣ���Ҫ�������ʽ����Ϊ�������Է��ֹ���ֵʱ��ʽ������
        Call deleteHyperlink(ws, Target.address)
        Call setCertainRangeFont(Target)
    End If
End Sub

Public Sub setRangeFont(ByRef certainRange As range, ByRef fontName As String)
    certainRange.Font.name = fontName
End Sub

Private Sub setCertainRangeFont(ByRef certainRange As range)
    With certainRange
        .WrapText = True '�����Զ�����
        With .Font
            .name = "Arial"
            .Underline = xlUnderlineStyleNone '���»���
            .colorIndex = -4105 '��ɫ
        End With
    End With
End Sub

Public Sub addHyperlink(ByRef srcWs As Worksheet, ByRef srcAddress As String, ByRef dstWs As Worksheet, ByRef dstAddress As String, Optional ByVal fontName As String = "Arial")
    On Error Resume Next
    srcWs.Hyperlinks.Add Anchor:=srcWs.range(srcAddress), address:="", _
        SubAddress:="'" & dstWs.name & "'!" & dstAddress
    '��������Ӱ�����ܣ�ɾ���������ӻ�վ��ͬ��������Ҫע�⣬��Ҫ��ͬ��������
    'DTS2017011105086
    Dim clipBoardData As DataObject
    Set clipBoardData = New DataObject
    clipBoardData.Clear
    clipBoardData.GetFromClipboard
    If Not isCellSheet(srcWs.name) Then srcWs.range(srcAddress).WrapText = False '�������С��ҳ�����Զ������÷񣬷�����ѡ�񲻷���
    'DTS2017011105086
    With clipBoardData
        .SetText ""
        .PutInClipboard
    End With
End Sub

Public Sub deleteHyperlink(ByRef srcWs As Worksheet, ByRef srcAddress As String)
    On Error Resume Next
    'DTS2017011105086
    Dim clipBoardData As DataObject
    Set clipBoardData = New DataObject
    
    clipBoardData.Clear
    clipBoardData.GetFromClipboard
    With srcWs.range(srcAddress)
        If .Hyperlinks.count > 0 Then '����õ�Ԫ�����г����ӣ���ɾ��
            .Hyperlinks.Delete
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .NumberFormatLocal = "@" '���õ�Ԫ���ʽΪ�ı�
            Call setBorders(srcWs.range(srcAddress))
        End If
    End With
    'DTS2017011105086
    With clipBoardData
        .SetText ""
        .PutInClipboard
    End With
End Sub

Private Function sheetHyperlinksShouldAdd(ByRef ws As Worksheet, ByRef sheetsHyperlinksCol As Collection) As Boolean
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
    maxRowNumber = ws.Cells(1048576, boardStyleNameColumnNumber).End(xlUp).row
    
    Dim dstBoardStyleSheetName As String
    For rowIndex = 3 To maxRowNumber
        With ws.Cells(rowIndex, boardStyleNameColumnNumber)
            dstBoardStyleSheetName = .value '�õ�������ʽ����ֵ
            Call addHyperlinkOfBoardStyleName(ws, .address, dstBoardStyleSheetName, "A1")
        End With
    Next rowIndex
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

'����ҳ��Board Style Nameֵ�ı�ʱ��ӳ�����
Public Sub addBoardStyleNameHyperlinks_SheetChange(ByRef ws As Worksheet, ByRef Target As range)
    If Target.count <> 1 Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    Dim flag As Boolean
    flag = getBoardStyleInfo(ws.name) '�ж���RelationDef�Ƿ���Board Style Name���ֶ�
    If flag = False Then
        Exit Sub
    End If
    
    Dim rowNumber As Long, columnNumber As Long
    rowNumber = Target.row
    columnNumber = Target.column
    flag = isBoardStyleCol(ws, rowNumber, columnNumber) '�жϵ�ǰ�ı�ĵ�Ԫ���Ƿ���Board Style Name�ֶ�
    If flag = False Then Exit Sub
    
    Dim dstBoardStyleSheetName As String
    dstBoardStyleSheetName = Target.value
    
    '�������Ϊ�壬�������ʱ�����˳�
    If dstBoardStyleSheetName = "" Then Exit Sub
    
    Call addHyperlinkOfBoardStyleName(ws, Target.address, dstBoardStyleSheetName, "A1")
End Sub

'С��ҳ��RXU Ant No.ֵ�ı�ʱ��ӳ�����
Public Sub addRxuAntNoHyperlinks_SheetChange(ByRef ws As Worksheet, ByRef Target As range)
    If Target.count <> 1 Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub
    
    Dim targetRowNumber As Long, targetColumnNumber As Long
    targetRowNumber = Target.row
    targetColumnNumber = Target.column
    
    Dim rxuAntNoName As String
    rxuAntNoName = getRxuAntNoName(ws.name) '�õ�RXU Ant No.������
    If rxuAntNoName = "" Then Exit Sub '���δ�ҵ����������˳�
    
    Dim rxuAntNoNameColumnNumber As Long
    rxuAntNoNameColumnNumber = findCertainValColumnNumber(ws, 2, rxuAntNoName)
    
    If Target.value = "" Or rxuAntNoNameColumnNumber <> targetColumnNumber Then 'targetColumnNumber�϶���Ϊ-1���������rxuAntNoNameColumnNumberδ�ҵ�Ϊ-1����ֱ���˳�
        Exit Sub '���Ŀ�굥Ԫ��Ϊ�ջ�ı���в���RXU Ant No.�У����˳�
    End If
    
    Dim dstBoardStyleSheetName As String
    dstBoardStyleSheetName = getBoardStyleName(ws, targetRowNumber) '�õ�С��ҳ�޸�������Ӧ��Board Style
    If dstBoardStyleSheetName = "" Then
        Call deleteHyperlink(ws, Target.address) 'û���ҵ���Ӧ��BoardStyleName���򽫸õ�Ԫ��ĳ��������
        Exit Sub '���boardStyleNameΪ�գ����˳�
    End If
    Call addHyperlinkOfBoardStyleName(ws, Target.address, dstBoardStyleSheetName, "A1")
End Sub

'С��ҳ�ڼ���ʱ��RXU Ant No.��ӳ�����
Private Sub addCellSheetHyperlinks_SheetActive(ByRef ws As Worksheet)
    If (Not isCellSheet(ws.name)) Or containsASheet(ThisWorkbook, DecouplingSheetName) Then Exit Sub '�������С��ҳ���������ҳǩ��˵����������������˳�
    
    '��ҳǩ�Ƿ�ӹ������ӣ�����ӹ���ֱ���˳�
    'If sheetHyperlinksShouldAdd(ws, sheetsHyperlinksCol) = False Then Exit Sub
    
    Dim rxuAntNoName As String
    rxuAntNoName = getRxuAntNoName(ws.name) '�õ�RXU Ant No.������
    If rxuAntNoName = "" Then Exit Sub '���δ�ҵ����������˳�
    
    Dim rxuAntNoNameColumnNumber As Long
    rxuAntNoNameColumnNumber = findCertainValColumnNumber(ws, 2, rxuAntNoName)
    
    If rxuAntNoNameColumnNumber = -1 Then Exit Sub '���û�ҵ����У����˳�
    Dim maxRowNumber As Long, rowIndex As Long
    maxRowNumber = ws.Cells(1048576, rxuAntNoNameColumnNumber).End(xlUp).row
    Dim dstBoardStyleSheetName As String
    For rowIndex = 3 To maxRowNumber
        With ws.Cells(rowIndex, rxuAntNoNameColumnNumber)
            If .value = "" Then GoTo NextLoop  '���RXU Ant No.Ϊ�գ��������һ��ѭ��
            
            dstBoardStyleSheetName = getBoardStyleName(ws, rowIndex) '�õ�С��ҳ�޸�������Ӧ��Board Style
            If dstBoardStyleSheetName = "" Then
                Call deleteHyperlink(ws, .address) '���û���ҵ�BoardStyleName���򽫸õ�Ԫ��ĳ��������
                GoTo NextLoop '���boardStyleNameΪ�գ��������һ��ѭ��
            End If
            Call addHyperlinkOfBoardStyleName(ws, .address, dstBoardStyleSheetName, "A1")
        End With
NextLoop:
    Next rowIndex
End Sub


Private Sub getAntenneMocNameAndAttrByCellSheetName(ByRef mocName As String, ByRef attrName As String, ByRef CellSheetName As String)
    If CellSheetName = getResByKey("GSM Logic Cell") Or CellSheetName = "GSM Cell" Then
        mocName = "GTRXGROUPSECTOREQM"
        attrName = "SECTORANTENNA"
    ElseIf CellSheetName = getResByKey("UMTS Logic Cell") Or CellSheetName = "UMTS Cell" Then
        mocName = "ULOCELLSECTOREQM"
        attrName = "SECTORANTENNA"
    ElseIf CellSheetName = getResByKey("LTE Cell") Or CellSheetName = "LTE Cell" Then
        mocName = "eUCellSectorEqm"
        attrName = "SECTORANTENNA"
    End If
End Sub

Private Function getRxuAntNoName(ByRef sheetName As String) As String
    Dim mocName As String, attrName As String
    Call getAntenneMocNameAndAttrByCellSheetName(mocName, attrName, sheetName)
    getRxuAntNoName = findColumnFromRelationDef(sheetName, mocName, attrName)
End Function

Private Function getBoardStyleName(ByRef ws As Worksheet, ByRef targetRowNumber As Long) As String
    getBoardStyleName = ""
    
    Dim bsNameColumnName As String, bsName As String
    bsNameColumnName = ws.range("A2").value '�õ���վ���ƻ�rat�����У��硰*BTS Name��
    bsName = ws.range("A" & targetRowNumber).value '�õ���վ��
    If bsName = "" Then Exit Function '����û�վ����Ϊ�գ����˳�
    
    Dim bsNameColumnNumberInTransportSheet As Long
    Dim transportSheetName As String
    Dim transportSheet As Worksheet
    transportSheetName = GetMainSheetName()
    Set transportSheet = ThisWorkbook.Worksheets(transportSheetName) '�õ�����ҳ
    '�ڴ���ҳ�ҵ�RXU Ant No.�ֶ�����������
    bsNameColumnNumberInTransportSheet = findCertainValColumnNumber(transportSheet, 2, bsNameColumnName)
    If bsNameColumnNumberInTransportSheet = -1 Then Exit Function '���û�ҵ����˳�
    
    
    Dim maxRowNumber As Long, boardStyleNameColumnNumber As Long, rowIndex As Long
    maxRowNumber = transportSheet.Cells(1048576, bsNameColumnNumberInTransportSheet).End(xlUp).row '�õ�����ҳ�����
    
    boardStyleNameColumnNumber = getBoardStyleNameColumnNumber(transportSheet) '�õ�BoardStyleName�к�
    If boardStyleNameColumnNumber = -1 Then Exit Function '����ڴ���ҳû�ҵ�BoardStyleName�ֶΣ����˳�
    
    Dim bsNameInTransportSheet As String
    For rowIndex = 3 To maxRowNumber
        bsNameInTransportSheet = transportSheet.Cells(rowIndex, bsNameColumnNumberInTransportSheet).value
        If bsNameInTransportSheet = bsName Then  '�ڴ���ҳ�ҵ��û�վ����
            getBoardStyleName = transportSheet.Cells(rowIndex, boardStyleNameColumnNumber).value
            Exit Function
        End If
    Next rowIndex
End Function
