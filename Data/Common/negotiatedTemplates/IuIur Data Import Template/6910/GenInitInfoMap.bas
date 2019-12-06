Attribute VB_Name = "GenInitInfoMap"
Option Explicit
Const StartTblDataRow = 15  'should not be modified

Public Sub GenInitTableMap()
  Const TableNameCol = 18
  Const StartRowCol = 12
  Const EndRowCol = 13
  
  Dim iRow As Integer
  Dim iGenRow As Integer
  
  Dim PreTableName As String
  PreTableName = "jack"

  Dim sht As Worksheet
  Set sht = Sheets("TableDef")
  Dim tosht As Worksheet
  Set tosht = Sheets("InitTableMap")
  tosht.Cells.ClearContents

  iGenRow = 1
  tosht.Cells(iGenRow, 1) = "ImpOption"
  tosht.Cells(iGenRow, 2) = "TableName"
  tosht.Cells(iGenRow, 3) = "SheetName"
  tosht.Cells(iGenRow, 4) = "StartRow"
  tosht.Cells(iGenRow, 5) = "EndRow"
  
  iGenRow = 2
  iRow = StartTblDataRow
  Do Until sht.Cells(iRow, TableNameCol) = "" And sht.Cells(iRow, StartRowCol) = "" And sht.Cells(iRow, EndRowCol) = ""
    If PreTableName <> sht.Cells(iRow, TableNameCol) Then
      tosht.Cells(iGenRow, 1) = "YES"
      tosht.Cells(iGenRow, 2) = sht.Cells(iRow, TableNameCol)
      tosht.Cells(iGenRow, 3) = sht.Cells(iRow, iMocName + 1)
      tosht.Cells(iGenRow, 4) = CStr(CInt(sht.Cells(iRow, StartRowCol)) + 1)
      tosht.Cells(iGenRow, 5) = sht.Cells(iRow, EndRowCol)

      PreTableName = sht.Cells(iRow, TableNameCol)
      iGenRow = iGenRow + 1
    End If
    iRow = iRow + 1
  Loop
End Sub

Public Sub GenInitFieldMap()
  'only need these columns
  Const FieldNameCol = 9
  
  Const TableNameCol = 17
  Const ColNameCol = 18
  
  Dim iRow As Integer
  Dim iGenRow As Integer

  Dim sht As Worksheet
  Set sht = Sheets("TableDef")
  Dim tosht As Worksheet
  Set tosht = Sheets("InitFieldMap")
  tosht.Cells.ClearContents

  iGenRow = 1
  tosht.Cells(iGenRow, 1) = "SheetName"
  tosht.Cells(iGenRow, 2) = "TableName"
  tosht.Cells(iGenRow, 3) = "FieldName"
  tosht.Cells(iGenRow, 4) = "ColName"
  
  tosht.Cells(iGenRow, 5) = "ColumnType"
  tosht.Cells(iGenRow, 6) = "Min"
  tosht.Cells(iGenRow, 7) = "Max"
  tosht.Cells(iGenRow, 8) = "ListValue"
  tosht.Cells(iGenRow, 9) = "CheckNull"
  tosht.Cells(iGenRow, 10) = "ColumnType2"

  iGenRow = 2
  iRow = StartTblDataRow
  Do Until sht.Cells(iRow, ColNameCol) = "" And sht.Cells(iRow, TableNameCol) = "" And sht.Cells(iRow, FieldNameCol) = ""

    tosht.Cells(iGenRow, 1) = sht.Cells(iRow, iMocName + 1)
    tosht.Cells(iGenRow, 2) = sht.Cells(iRow, ColNameCol)
    tosht.Cells(iGenRow, 3) = sht.Cells(iRow, iMapFieldName + 1)
    tosht.Cells(iGenRow, 4) = sht.Cells(iRow, FieldNameCol)
    tosht.Cells(iGenRow, 5) = sht.Cells(iRow, iColumnType + 1)
    tosht.Cells(iGenRow, 6) = sht.Cells(iRow, iMin + 1)
    tosht.Cells(iGenRow, 7) = sht.Cells(iRow, iMax + 1)
    tosht.Cells(iGenRow, 8) = sht.Cells(iRow, iListValue + 1)
    tosht.Cells(iGenRow, 9) = sht.Cells(iRow, iCheckNull + 1)
    tosht.Cells(iGenRow, 10) = sht.Cells(iRow, iColumnType2 + 1)
    
    iGenRow = iGenRow + 1
    iRow = iRow + 1
  Loop
End Sub

