Attribute VB_Name = "Home_Cover"
Private Sub CommandButton1_Click()
Dim MOInfosheet As Worksheet
Dim HomeSheet As Worksheet
Dim iRow, iMO As Integer
Dim TableName, MOName, MOFullENG, MOFullCHS As String


Set MOInfosheet = ThisWorkbook.Sheets("MOInfo")
Set HomeSheet = ThisWorkbook.Sheets("HOME")
For iRow = 2 To 30000
    If HomeSheet.Cells(iRow, 1).value = "" Then
        Exit For
      End If
    TableName = HomeSheet.Cells(iRow, 1).value
     For iMO = 4 To 30000
          If MOInfosheet.Cells(iMO, 1).value = "" Then
            Exit For
          End If
         MOName = MOInfosheet.Cells(iMO, 1).value
          If TableName = MOName Then
             MOFullENG = MOInfosheet.Cells(iMO, 2).value
             MOFullCHS = MOInfosheet.Cells(iMO, 8).value
             HomeSheet.Cells(iRow, 2).value = MOFullENG
             HomeSheet.Cells(iRow, 3).value = MOFullCHS
           End If
     Next iMO
 Next iRow
End Sub
