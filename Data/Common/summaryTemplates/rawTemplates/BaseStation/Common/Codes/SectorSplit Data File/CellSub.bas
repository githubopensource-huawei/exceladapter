Attribute VB_Name = "CellSub"
Option Explicit
Public g_CellSheet As Worksheet
Public g_ExpParaSheet As Worksheet
Public g_OldCellAntCol As Long
Public g_NewCellAntCol As Long

Public Sub hideParaInCellSheet()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim exposeParas As New Collection
    Set exposeParas = getExposeParas
    
    Dim maxCol As Integer
    maxCol = g_CellSheet.range("IV2").End(xlToLeft).column
    
    Dim cellShtName As String
    Dim colIdx As Integer
    Dim c As Variant
    
    With g_CellSheet
        cellShtName = g_CellSheet.name
        For Each c In g_CellSheet.comments
            c.Shape.Placement = 1
            c.Shape.TextFrame.AutoSize = True
        Next c
        
        For colIdx = 2 To maxCol
            Dim grpName As String
            Dim colName As String
            
            colName = .Cells(2, colIdx)
            grpName = getGroupNameFromMappingDef(cellShtName, colName)
            
            If grpName = getResByKey("CellSplitInfo") Then
                Exit Sub
            End If
            
            If Not Contains(exposeParas, grpName & "_" & colName) Then
                .columns(colIdx).EntireColumn.Hidden = True
            End If
        Next
    End With
    Set exposeParas = Nothing
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Debug.Print "some exception in hideParaInCellSheet, " & Err.Description
End Sub

Public Sub showParaInCellSheet()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim colIdx As Integer
    With g_CellSheet
        For colIdx = 2 To g_CellSheet.range("IV2").End(xlToLeft).column
            If .columns(colIdx).EntireColumn.Hidden = True Then
                .columns(colIdx).EntireColumn.Hidden = False
            End If
        Next
    End With
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Debug.Print "some exception in showParaInCellSheet, " & Err.Description
End Sub

Public Function getExposeParas() As Collection
    Dim grpNameCol As Integer
    Dim colNameCol As Integer
    If g_Language = "CN" Then
        grpNameCol = 1
        colNameCol = 2
    Else
        grpNameCol = 3
        colNameCol = 4
    End If
    
    Dim rowIdx As Integer
    Dim exposeParas As New Collection
    
    For rowIdx = 3 To g_ExpParaSheet.range("a65536").End(xlUp).row
        Dim grpName As String
        Dim colName As String
        With g_ExpParaSheet
            grpName = .Cells(rowIdx, grpNameCol)
            colName = .Cells(rowIdx, colNameCol)
        End With
        exposeParas.Add Item:=(Trim(grpName) & "_" & Trim(colName)), key:=(Trim(grpName) & "_" & Trim(colName))
    Next
    
    Set getExposeParas = exposeParas
End Function

Public Sub genBoardNoList(ByRef ws As Worksheet, ByVal target As range)
    Dim siteName As String
    Dim targetBoardStyleCol As Long
    Dim targetBoardStyleName As String
    
    siteName = ws.Cells(target.row, siteNameCol(ws))
    targetBoardStyleName = getTargetBoardStyleName(siteName)
    
    Dim validInfo As String
    validInfo = getTargetBoardNos(targetBoardStyleName)
    
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=validInfo
    End With
End Sub

Public Function getTargetBoardStyleName(ByRef siteName As String) As String
    Dim transportSht As Worksheet
    Dim targetBoardStyleCol As Long
    
    Set transportSht = ThisWorkbook.Worksheets(GetMainSheetName())
    targetBoardStyleCol = colNum(transportSht, getResByKey("targetBoardStyleName"))
    
    If targetBoardStyleCol = -1 Then
        getTargetBoardStyleName = ""
        Exit Function
    End If
    
    Dim ratNameCol As Long
    ratNameCol = siteNameCol(transportSht) 'function name, not node name when SRAN
    
    Dim row As Long
    For row = g_ListShtColNameRow + 1 To getUsedRow(transportSht)
        If transportSht.Cells(row, ratNameCol) = siteName Then
            getTargetBoardStyleName = transportSht.Cells(row, targetBoardStyleCol)
            Exit Function
        End If
    Next
End Function

Public Function getTargetBoardNos(ByRef boardStyleName As String) As String
On Error GoTo ErrorHandle:
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(boardStyleName)
    
    Dim bbpBeginRow As Long, bbpEndRow As Long
    bbpEndRow = -1
    
    Dim row As Long
    For row = 1 To getUsedRow(ws, "Pattern")
        If ws.Cells(row, 1) = getResByKey("BBP") Then
            bbpBeginRow = row
            bbpEndRow = row
        End If
        If bbpEndRow <> -1 Then
            If ws.Cells(row, 1) <> "" Then
                bbpEndRow = row
            Else
                Exit For
            End If
        End If
    Next
    
    If bbpEndRow - bbpBeginRow < 2 Then Exit Function
    
    Dim result As String
    For row = bbpBeginRow + 2 To bbpEndRow
        result = result & ws.Cells(row, 1) & ","
    Next
    
    getTargetBoardNos = Left(result, Len(result) - 1)
    
ErrorHandle:

End Function

Public Sub getOldCellAntNoCol()
    g_OldCellAntCol = colNum(g_CurrentSheet, getResByKey("OldCellAntNo"))
End Sub

Public Sub getNewCellAntNoCol()
    g_NewCellAntCol = colNum(g_CurrentSheet, getResByKey("NewCellAntNo"))
End Sub

Public Function cellAntSelected(ByRef ws As Worksheet, ByRef target As range) As Boolean
    cellAntSelected = False
    Call getOldCellAntNoCol
    Call getNewCellAntNoCol
    
    If target.column = g_OldCellAntCol Or target.column = g_NewCellAntCol Then
        cellAntSelected = True
    End If
End Function



