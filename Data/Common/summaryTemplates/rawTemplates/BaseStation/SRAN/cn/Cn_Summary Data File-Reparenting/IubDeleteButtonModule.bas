Attribute VB_Name = "IubDeleteButtonModule"
Option Explicit
Private Const OutOfRangeGrayColor As Variant = 12632256 '灰色底色
Private Const MocTitleRedColor As Variant = 128 'Moc A列红色
Private Const AttributeOrangeColor As Variant = 10079487 '属性格橙色

Public Sub changeIubSheetDeleteButtonFunction(ByRef ws As Worksheet)
    If isIubStyleWorkSheet(ws.name) Then
        'IUB页签则重新分配Delete按键功能
        Application.OnKey "{DEL}", "OnIubSheetDel"
    Else
        '非IUB页签功能恢复默认
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
    '如果该行隐藏，则直接退出
    If rowRange.EntireRow.Hidden Then Exit Sub
    
    Dim cellRange As range
    Dim cellRangeColor As Variant
    For Each cellRange In rowRange.Cells
        cellRangeColor = cellRange.Interior.Color
        If cellRangeColor = MocTitleRedColor Or cellRangeColor = AttributeOrangeColor Or _
            cellRange.count <> cellRange.MergeArea.count Then
            'cellRange.MergeCells Then'不用这个方法是因为用这个判断的时候，会把合并单元格选定
            '如果是Moc大红色或属性单元格，或者是合并单元格则跳过
            GoTo NextLoop
        ElseIf cellRangeColor = OutOfRangeGrayColor Then
            '如果是灰色底色，则该行已经到范围外了，可以直接退出
            Exit Sub
        End If
        
        cellRange.ClearContents
NextLoop:
    Next cellRange
End Sub


