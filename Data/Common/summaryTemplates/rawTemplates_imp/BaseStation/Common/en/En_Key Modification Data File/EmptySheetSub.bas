Attribute VB_Name = "EmptySheetSub"
Sub hiddenEmptySheet()
    Dim index As Long
    Dim column As Long
    Dim isEmpty As Boolean
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    For index = 2 To sheetDef.Range("a65536").End(xlUp).row
        isEmpty = True
        If UCase(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo)) <> "MAIN" And UCase(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo)) <> "COMMON" Then
            For column = 1 To ThisWorkbook.Worksheets(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtNameColNo).value).Range("IV2").End(xlToLeft).column
                If ThisWorkbook.Worksheets(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtNameColNo).value).Cells(3, column).value <> "" Then
                       isEmpty = False
                       Exit For
                End If
            Next
            If isEmpty Then
                ThisWorkbook.Worksheets(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtNameColNo).value).Visible = 0
            End If
        End If
    Next
End Sub

Sub showEmptySheet()
    Dim index As Long
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    For index = 2 To sheetDef.Range("a65536").End(xlUp).row
        If UCase(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo)) <> "MAIN" And UCase(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo)) <> "COMMON" Then
            ThisWorkbook.Worksheets(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtNameColNo).value).Visible = -1
        End If
    Next
End Sub


