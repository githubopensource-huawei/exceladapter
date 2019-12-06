VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomizeTemplateForm 
   Caption         =   "Customize Template"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   OleObjectBlob   =   "CustomizeTemplateForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "CustomizeTemplateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InitGUI()
    Dim optionBt As Control
    init ThisWorkbook
    CustomizeTemplateForm.Caption = gCaption_CustomizeTemplate
    FuctionFrame.Caption = gCaptionSceneFrame
    For Each optionBt In FuctionFrame.Controls
        Select Case optionBt.Name
            Case gCreateBTS
                optionBt.Caption = gCaptionCreateBTS
            Case gRpsTDMInBSC
                optionBt.Caption = gCaptionRpsTDMInBSC
            Case gRpsBetweenBSC
                optionBt.Caption = gCaptionRpsBetweenBSC
            Case "AllOptionBt"
                optionBt.Caption = gCaptionAll
        End Select
    Next optionBt
    OKBt.Caption = gCaption_OKButton
    CancelBt.Caption = gCaption_CancelButton

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub CancelBt_Click()
    CustomizeTemplateForm.Hide
End Sub
Private Function isSpecColSupported(ByVal colName As String) As Boolean
    Dim sceneRange, specFieldRange As Range
    Dim curSceneCol As Integer
    isSpecColSupported = True
    Set specFieldRange = ThisWorkbook.Sheets(gShtNameSpecialFields).Cells.Find(Trim(colName), LookIn:=xlValues, LookAt:=xlWhole)
    Set sceneRange = ThisWorkbook.Sheets(gShtNameSpecialFields).Cells.Find(gCurScene)
    If Not specFieldRange Is Nothing And Not sceneRange Is Nothing Then
        If ThisWorkbook.Sheets(gShtNameSpecialFields).Cells(specFieldRange.row, sceneRange.Column) <> "YES" Then
            isSpecColSupported = False
        End If
    End If
End Function

Private Function isMocSupported(ByVal mocName As String) As Boolean
    Dim mocRange As Range
    Dim sceneRange As Range
    Dim curSceneCol As Integer
    isMocSupported = True
    If Trim(mocName) = "" Then
        Exit Function
    End If
    'Summary��վ����Moc��֧��
    If gCurScene = gCreateBTS Then
        Exit Function
    End If
    '���⴦��914֮ǰ�汾��BSC��ǨHomeҳ��û�в𿪶���
    If gCurScene = gRpsBetweenBSC Then
        If mocName = "GCELLOSPMAP" Or mocName = "BTSSHARING" Then
            isMocSupported = True
            Exit Function
        End If
    End If
    'TDM�ڰ�Ǩֻ֧��BTSCONNECT
    If gCurScene = gRpsTDMInBSC Then
        If mocName = "BTSCONNECT" Then
            isMocSupported = True
        Else
            isMocSupported = False
        End If
        Exit Function
    End If
    Set sceneRange = ThisWorkbook.Sheets(gShtNameFuctionMocs).Cells.Find(gCurScene, LookIn:=xlValues, LookAt:=xlWhole)
    If Not sceneRange Is Nothing Then
        Set mocRange = ThisWorkbook.Sheets(gShtNameFuctionMocs).Columns(sceneRange.Column).Find(mocName, LookIn:=xlValues, LookAt:=xlWhole)
        If mocRange Is Nothing Then
            isMocSupported = False
        End If
    End If
End Function

'�ж��Ƿ񱻵�ǰ����֧��
Private Function isColSupported(ByVal colName As String, ByVal groupName As String, ByVal mocName As String, ByVal attrName As String) As Boolean
    Dim supported As Boolean
    If UCase(gCurScene) = "ALL" Then
        isColSupported = True
        Exit Function
    End If
    '�����ֶ��Ҳ�����ǰ����֧��ʱ������False
    supported = isSpecColSupported(colName)
    '�ж�Moc�Ƿ񱻵�ǰ����֧��
    If supported Then
        supported = isMocSupported(mocName)
    End If
    '��վ���ƶ�֧��
    If mocName = "BTS" And (attrName = "BTSNAME" Or attrName = "BSCName") Then
        supported = True
    End If
    '�޸Ļ�վ���ơ�С������ֻ�п�BSC��Ǩ֧��
    If (attrName = "MODBTSNAME" Or attrName = "MODCELLNAME") And gCurScene <> gRpsBetweenBSC Then
        supported = False
    End If
    isColSupported = supported
End Function

Private Sub restoreInvalidColor()
    Dim row As Integer
    Dim colNameCol, groupNameCol, shtNameCol As Integer
    Dim findRange As Range
    row = 2
    If isShtExists(gShtNameInvalidFields) Then
        colNameCol = getColNum(gShtNameInvalidFields, "", gColName_srcColName)
        groupNameCol = getColNum(gShtNameInvalidFields, "", gColName_groupName)
        shtNameCol = getColNum(gShtNameInvalidFields, "", gColName_srcShtName)
        With ThisWorkbook.Sheets(gShtNameInvalidFields)
            While Trim(.Cells(row, colNameCol)) <> ""
                Set findRange = getPosRange(.Cells(row, shtNameCol), .Cells(row, groupNameCol), .Cells(row, colNameCol))
                If Not findRange Is Nothing Then
                '�ָ�ΪĬ����ɫ
                    findRange.Interior.ColorIndex = 40
                    findRange.Interior.Pattern = xlSolid
                    findRange.Interior.PatternColorIndex = xlAutomatic
                End If
                row = row + 1
            Wend
        End With
    End If
End Sub

Private Sub refreshInvalidFields()
    Dim mapColNameCol, mapGroupNameCol, mapShtNameCol, mapMocNameCol, mapAttrNameCol As Integer
    Dim mapRow, invalidRow As Integer
    
    If isShtExists(gShtNameInvalidFields) Then
        ThisWorkbook.Sheets(gShtNameInvalidFields).Cells.ClearContents
    Else
        ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ActiveSheet.Name = gShtNameInvalidFields
        'Type:=gShtNameInvalidFields
    End If
    ThisWorkbook.Sheets(gShtNameInvalidFields).Visible = False
    '������ͷ
    ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(1, InvalidColNameCol) = gColName_srcColName
    ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(1, InvalidGroupNameCol) = gColName_groupName
    ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(1, InvalidShtNameCol) = gColName_srcShtName
    
    mapColNameCol = getColNum(gMappingDefShtName, "", gColName_srcColName)
    mapGroupNameCol = getColNum(gMappingDefShtName, "", gColName_groupName)
    mapShtNameCol = getColNum(gMappingDefShtName, "", gColName_srcShtName)
    mapMocNameCol = getColNum(gMappingDefShtName, "", gColName_dstShtName)
    mapAttrNameCol = getColNum(gMappingDefShtName, "", gColName_dstColName)
    
    mapRow = 2
    invalidRow = 2
    Dim validFields() As String
    ReDim validFields(0) As String
    With ThisWorkbook.Sheets(gMappingDefShtName)
        While Trim(ThisWorkbook.Sheets(gMappingDefShtName).Cells(mapRow, mapColNameCol)) <> ""
            '�����ǰ������֧�֣���д����Ч�ֶ�ҳ����
            If Not isColSupported(.Cells(mapRow, mapColNameCol), .Cells(mapRow, mapGroupNameCol), .Cells(mapRow, mapMocNameCol), .Cells(mapRow, mapAttrNameCol)) Then
                ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(invalidRow, InvalidColNameCol) = .Cells(mapRow, mapColNameCol)
                ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(invalidRow, InvalidGroupNameCol) = .Cells(mapRow, mapGroupNameCol)
                ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(invalidRow, InvalidShtNameCol) = .Cells(mapRow, mapShtNameCol)
                invalidRow = invalidRow + 1
            Else
                '��Ч�ֶα�������
                insertList validFields, .Cells(mapRow, mapColNameCol) & "_" & .Cells(mapRow, mapGroupNameCol) & "_" & .Cells(mapRow, mapShtNameCol)
            End If
            mapRow = mapRow + 1
        Wend
    End With
    '����һ��ӳ�䵽����ֶ�ʱ�����ܻ��в���ӳ����Ч������ӳ����Ч����������������Ϊ����Ч
    Dim keys() As String
    Dim index As Integer
    For index = 1 To UBound(validFields)
        If validFields(index) <> "" Then
            keys = Split(validFields(index), "_")
            invalidRow = 2
            Do While Trim(ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(invalidRow, InvalidColNameCol)) <> ""
                If ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(invalidRow, InvalidColNameCol) = keys(0) And _
                    ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(invalidRow, InvalidGroupNameCol) = keys(1) And _
                    ThisWorkbook.Sheets(gShtNameInvalidFields).Cells(invalidRow, InvalidShtNameCol) = keys(2) Then
                    '����Чӳ����ɾ��������Чӳ��
                    ThisWorkbook.Sheets(gShtNameInvalidFields).Rows(invalidRow).Delete shift:=xlUp
                    invalidRow = invalidRow - 1
                End If
                invalidRow = invalidRow + 1
            Loop
        End If
    Next index
End Sub

Private Sub setInvalidColor()
    Dim row As Integer
    Dim colNameCol, groupNameCol, shtNameCol As Integer
    Dim findRange As Range
    colNameCol = getColNum(gShtNameInvalidFields, "", gColName_srcColName)
    groupNameCol = getColNum(gShtNameInvalidFields, "", gColName_groupName)
    shtNameCol = getColNum(gShtNameInvalidFields, "", gColName_srcShtName)
    row = 2
    If isShtExists(gShtNameInvalidFields) Then
        With ThisWorkbook.Sheets(gShtNameInvalidFields)
            While Trim(.Cells(row, colNameCol)) <> ""
                Set findRange = getPosRange(.Cells(row, shtNameCol), .Cells(row, groupNameCol), .Cells(row, colNameCol))
                If Not findRange Is Nothing Then
                '����Ϊ��ɫ
                    findRange.Interior.ColorIndex = 48
                    findRange.Interior.Pattern = xlLightUp
                    findRange.Interior.PatternColorIndex = xlAutomatic
                End If
                row = row + 1
            Wend
        End With
    End If
    setRxuspecSht
End Sub
Private Sub setRxuspecSht()
    Dim col As Integer
    If Not isShtExists(gRxuSpecShtName) Then
        Exit Sub
    End If
    col = 2
    'ֻ�н�վ֧��RXUSPECҳ
    With ThisWorkbook.Sheets(gRxuSpecShtName)
        If gCurScene = gCreateBTS Or UCase(gCurScene) = "ALL" Then
            While Trim(.Cells(2, col)) <> ""
                '����ΪĬ����Ч��ɫ
                .Cells(2, col).Interior.ColorIndex = 40
                .Cells(2, col).Interior.Pattern = xlSolid
                .Cells(2, col).Interior.PatternColorIndex = xlAutomatic
                col = col + 1
            Wend
        Else
            While Trim(.Cells(2, col)) <> ""
                '����Ϊ��ɫ
                .Cells(2, col).Interior.ColorIndex = 48
                .Cells(2, col).Interior.Pattern = xlLightUp
                .Cells(2, col).Interior.PatternColorIndex = xlAutomatic
                col = col + 1
            Wend
        End If
    End With
End Sub

Private Sub OKBt_Click()
    Dim controlBt As Control
    Dim optionBt As OptionButton
    If GSM_SUMMARY_CREATEBTS.Value = True Then
        gCurScene = gCreateBTS
    ElseIf GSM_BTS_REPARENT.Value = True Then
        gCurScene = gRpsBetweenBSC
    ElseIf GSM_BTS_REPARENT_TDM_INBSC.Value = True Then
        gCurScene = gRpsTDMInBSC
    Else
        gCurScene = "ALL"
    End If

    '��ԭ��Ч�ֶε���ɫ
    restoreInvalidColor
    '���ݵ�ǰ��������������Ч�ֶ��б�
    refreshInvalidFields
    '������Ч�ֶ��б�����������Ч�ֶ���ɫ
    setInvalidColor
    CustomizeTemplateForm.Hide
End Sub

