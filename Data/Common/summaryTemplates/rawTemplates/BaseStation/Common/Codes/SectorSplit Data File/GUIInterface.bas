Attribute VB_Name = "GUIInterface"
Option Explicit

'************************************************************************
'this macro is called by JNI
'************************************************************************
Public Sub adjustColumnPosition()
    If isUMTS Then
        Dim tmpShtName As String
        tmpShtName = g_CurrentSheet.name
        
        Dim cellSht As Worksheet
        Set cellSht = Sheets(getResByKey("UMTSCellSheet"))
        cellSht.Activate
        
        Application.DisplayAlerts = False
        Call adjustLgcCellPosInCellSht(cellSht)
        Call adjustSplitPosInCellSht(cellSht)
        Application.DisplayAlerts = True
        
        Set g_CurrentSheet = Sheets(tmpShtName)
        g_CurrentSheet.Activate
    End If
End Sub

Private Sub adjustSplitPosInCellSht(cellSht As Worksheet)
    Dim copySrcStartCol As Integer
    Dim copySrcEndCol As Integer
    Dim copySrcStartRow As Integer
    Dim copySrcEndRow As Integer
    Dim copyDstPos As Integer
    
    copySrcStartCol = colNumByAttr(cellSht, getResByKey("RRUChainStrategy"))
    copySrcEndCol = colNumByAttr(cellSht, getResByKey("NewCellAntNo"))
    copySrcStartRow = listShtAttrRow
    copySrcEndRow = cellSht.UsedRange.rows.count
    copyDstPos = getListShtUsedCol(cellSht) + 1
    
    With cellSht
        .range(getColStr(copySrcStartCol) & copySrcStartRow & ":" & getColStr(copySrcEndCol) & copySrcEndRow).Cut
        .range(getColStr(copyDstPos) & listShtAttrRow).Insert shift:=xlToRight
    End With
End Sub

Private Sub adjustLgcCellPosInCellSht(cellSht As Worksheet)
    Dim copySrcStartCol As Integer
    Dim copySrcEndCol As Integer
    Dim copyDstPos As Integer
    
    copySrcStartCol = colNumByAttr(cellSht, getResByKey("CellSplitInfo"), listShtGrpRow)
    copySrcEndCol = copySrcStartCol + Cells(listShtGrpRow, copySrcStartCol).MergeArea.columns.count - 1
    copyDstPos = getListShtUsedCol(cellSht) + 1

    If copyDstPos <> copySrcEndCol + 1 Then 'if already the last column, no need to move
        With cellSht
            .columns(getColStr(copySrcStartCol) & ":" & getColStr(copySrcEndCol)).EntireColumn.Cut
            .columns(getColStr(copyDstPos)).Insert shift:=xlToRight
        End With
    End If
End Sub

