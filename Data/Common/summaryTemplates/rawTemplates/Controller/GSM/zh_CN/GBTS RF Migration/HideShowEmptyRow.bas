Attribute VB_Name = "HideShowEmptyRow"
Const MocName_Column = 1
Public Function IsSystemSheet(curSheet As Worksheet) As Boolean
    IsSystemSheet = False
    If curSheet.name = GetMainSheetName _
       Or curSheet.name = getResByKey("Cover") _
       Or curSheet.name = getResByKey("Comm Data") _
         Then
        IsSystemSheet = True
    End If
End Function
Private Sub HideRow(startRow As Long, endRow As Long, curSheet As Worksheet)
  Dim index As Long
  Dim TitleRow As Long
  Dim ExistsData As Boolean
  ExistsData = False
  
  TitleRow = 1
  Dim count As Long
  count = curSheet.columns.count
  For index = startRow To endRow
     If Application.WorksheetFunction.CountBlank(curSheet.Rows(index)) <> count And curSheet.Cells(index, MocName_Column) = "" Then 'Data row is not empty
       ExistsData = True
     End If
     
     If Application.WorksheetFunction.CountBlank(curSheet.Rows(index)) = count Then 'Empty Row
       curSheet.Rows(index).Hidden = True
     End If
     
     If curSheet.Cells(index, MocName_Column) <> "" Then 'Next Title Row
       'If last Title Row have not data, hide it
       If ExistsData = False Then
         curSheet.Rows(TitleRow).Hidden = True
       End If
       ExistsData = False
       TitleRow = index
     End If
  Next index
  
  'Last MOC
  If ExistsData = False Then
    curSheet.Rows(TitleRow).Hidden = True
  End If
  
  
End Sub

Function get_EndRow() As Long
  Dim iRow As Long
  Dim ShetDef_Shet As Worksheet
  Set ShetDef_Shet = Sheets("SHEET DEF")
  
  get_EndRow = 1
  iRow = 2
  Do While ShetDef_Shet.Cells(iRow, 1) <> ""
    If ShetDef_Shet.Cells(iRow, 5).value <> "" Then
      If Int(ShetDef_Shet.Cells(iRow, 5)) > get_EndRow Then
              get_EndRow = Int(ShetDef_Shet.Cells(iRow, 5).value)
      End If
    End If
    iRow = iRow + 1
  Loop
End Function

Sub HideEmptyRow()
  Dim curSheet As Worksheet
  Dim endRow As Long
  Dim startRow As Long
  
  Dim index As Long

  If IsSystemSheet(ActiveSheet) Then Exit Sub
  Set curSheet = ActiveSheet

  startRow = 2
  endRow = get_EndRow()
  
  Application.ScreenUpdating = False
  
  Call HideRow(startRow, endRow, curSheet)
  Application.ScreenUpdating = True

End Sub

Sub ShowEmptyRow()
  Dim curSheet As Worksheet
  If IsSystemSheet(ActiveSheet) Then Exit Sub
  Set curSheet = ActiveSheet
  curSheet.Cells.Select
  curSheet.Cells.EntireRow.Hidden = False
  'CurSheet.rows("1:1").Select
  'Selection.EntireRow.Hidden = True
End Sub










