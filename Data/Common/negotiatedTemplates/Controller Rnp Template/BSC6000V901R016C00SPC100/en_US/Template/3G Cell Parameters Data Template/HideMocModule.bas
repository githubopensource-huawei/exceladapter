Attribute VB_Name = "HideMocModule"
'°üº¬Ä³¸öÒ³Ç©´úÂë
Private Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Dim sheet As Worksheet
    Set sheet = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsASheet = False
End Function

Private Function getSheetRowNum(ByRef sheet As Worksheet, ByVal col As Integer) As Integer
    For i = 1 To 1000
      Value = Trim(sheet.Cells(i, 1).Value)
      If Value = "" Then
        Exit For
      End If
      
    Next i
    getSheetRowNum = i - 1
End Function

Private Sub removeHomeItem(ByRef mocs As Collection)
    Dim sheetHOME As Worksheet
    
    Set sheetHOME = ThisWorkbook.Sheets("Home")
    rowNum = getSheetRowNum(sheetHOME, 1)
    For i = rowNum To 2 Step -1
      cell = Trim(sheetHOME.Cells(i, 1).Value)
      found = False
      
      For j = 1 To mocs.Count
        If mocs.Item(j) = cell Then
          found = True
          Exit For
        End If
      Next j
      
      If Not found Then
        sheetHOME.Rows(i).Delete
      End If
    Next i
End Sub

Private Function isReverseItem(ByRef name As String)
    Dim reverse(7) As String
    reverse(0) = "CMETemplateInfo"
    reverse(1) = "Refresh"
    reverse(2) = "ValidInfo"
    reverse(3) = "TableInfo"
    reverse(4) = "UserSelectMoc"
    reverse(5) = "Home"
    reverse(6) = "Cover"
    
    found = False
    For i = LBound(reverse) To UBound(reverse)
      If reverse(i) = name Then
        found = True
        Exit For
      End If
    Next i
    
    isReverseItem = found
End Function

Private Function inSelectMoc(ByRef mocs As Collection, ByRef name As String)
    found = False
    For i = 1 To mocs.Count
      If mocs.Item(i) = name Then
        found = True
        Exit For
      End If
    Next i
    
    inSelectMoc = found
End Function

Private Sub hideMocsSheet(ByRef mocs As Collection)

    Dim sheet As Worksheet
    For i = 1 To ThisWorkbook.Worksheets.Count
      Set sheet = ThisWorkbook.Worksheets.Item(i)
      If (Not isReverseItem(sheet.name)) And (Not inSelectMoc(mocs, sheet.name)) Then
        sheet.Visible = xlSheetHidden
      End If
    Next i
    
End Sub

Public Sub HideNoSelectMoc()
  Dim mocs As New Collection
  
  If containsASheet(ThisWorkbook, "UserSelectMoc") Then
    Dim sheetSelectMoc As Worksheet, sheetHOME As Worksheet
    
    Set sheetSelectMoc = ThisWorkbook.Sheets("UserSelectMoc")
    rowNum = getSheetRowNum(sheetSelectMoc, 1)
    For i = 2 To rowNum
      moc = Trim(sheetSelectMoc.Cells(i, 1).Value)
      If moc = "" Then
        Exit For
      End If
      
      mocs.Add moc
    Next i

    removeHomeItem mocs
    
    hideMocsSheet mocs

  End If
  
End Sub

