Attribute VB_Name = "IubDeleteButtonModule"
Option Explicit
Private Const OutOfRangeGrayColor As Variant = 12632256 '��ɫ��ɫ
Private Const MocTitleRedColor As Variant = 128 'Moc A�к�ɫ
Private Const AttributeOrangeColor As Variant = 10079487 '���Ը��ɫ

Public Sub changeIubSheetDeleteButtonFunction(ByRef ws As Worksheet)
    If isIubStyleWorkSheet(ws.name) Then
        'IUBҳǩ�����·���Delete��������
        Application.OnKey "{DEL}", "OnIubSheetDel"
    Else
        '��IUBҳǩ���ָܻ�Ĭ��
        Application.OnKey "{DEL}"
    End If
End Sub

Private Sub OnIubSheetDel()
    Dim selectionRange As range, rowRange As range
    Set selectionRange = Selection
    For Each rowRange In selectionRange.rows
        Call clearEachRowContents(rowRange)
    Next rowRange
End Sub

Private Sub clearEachRowContents(ByRef rowRange As range)
    '����������أ���ֱ���˳�
    If rowRange.EntireRow.Hidden Then Exit Sub
    
    Dim cellRange As range
    Dim cellRangeColor As Variant
    For Each cellRange In rowRange.Cells
        cellRangeColor = cellRange.Interior.Color
        If cellRangeColor = MocTitleRedColor Or cellRangeColor = AttributeOrangeColor Or _
            cellRange.count <> cellRange.MergeArea.count Then
            'cellRange.MergeCells Then'���������������Ϊ������жϵ�ʱ�򣬻�Ѻϲ���Ԫ��ѡ��
            '�����Moc���ɫ�����Ե�Ԫ�񣬻����Ǻϲ���Ԫ��������
            GoTo NextLoop
        ElseIf cellRangeColor = OutOfRangeGrayColor Then
            '����ǻ�ɫ��ɫ��������Ѿ�����Χ���ˣ�����ֱ���˳�
            Exit Sub
        End If
        
        cellRange.ClearContents
NextLoop:
    Next cellRange
End Sub


