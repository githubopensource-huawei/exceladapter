Attribute VB_Name = "GenInitInfoMap"
Option Explicit
Const StartTblDataRow = 15  'should not be modified

Public Sub GenInitTableMap()
  Const TableNameCol = 17
  Const StartRowCol = 11
  Const EndRowCol = 12
  
  Dim iRow As Integer
  Dim iGenRow As Integer
  
  Dim PreTableName As String
  PreTableName = "jack"

  Dim Sht As Worksheet
  Set Sht = Sheets("TableDef")
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
  Do Until Sht.Cells(iRow, TableNameCol) = "" And Sht.Cells(iRow, StartRowCol) = "" And Sht.Cells(iRow, EndRowCol) = ""
    If PreTableName <> Sht.Cells(iRow, TableNameCol) Then
      tosht.Cells(iGenRow, 1) = "YES"
      tosht.Cells(iGenRow, 2) = Sht.Cells(iRow, TableNameCol)
      tosht.Cells(iGenRow, 3) = Sht.Cells(iRow, iMocName + 1)
      tosht.Cells(iGenRow, 4) = CStr(CInt(Sht.Cells(iRow, StartRowCol)) + 1)
      tosht.Cells(iGenRow, 5) = Sht.Cells(iRow, EndRowCol)

      PreTableName = Sht.Cells(iRow, TableNameCol)
      iGenRow = iGenRow + 1
    End If
    iRow = iRow + 1
  Loop
End Sub

Public Sub GenInitFieldMap()
  'only need these columns
  Const FieldNameCol = 8
  
  Const TableNameCol = 16
  Const ColNameCol = 17
  
  Dim iRow As Integer
  Dim iGenRow As Integer

  Dim Sht As Worksheet
  Set Sht = Sheets("TableDef")
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
  Do Until Sht.Cells(iRow, ColNameCol) = "" And Sht.Cells(iRow, TableNameCol) = "" And Sht.Cells(iRow, FieldNameCol) = ""

    tosht.Cells(iGenRow, 1) = Sht.Cells(iRow, iMocName + 1)
    tosht.Cells(iGenRow, 2) = Sht.Cells(iRow, ColNameCol)
    tosht.Cells(iGenRow, 3) = Sht.Cells(iRow, iMapFieldName + 1)
    tosht.Cells(iGenRow, 4) = Sht.Cells(iRow, FieldNameCol)
    tosht.Cells(iGenRow, 5) = Sht.Cells(iRow, iColumnType + 1)
    tosht.Cells(iGenRow, 6) = Sht.Cells(iRow, iMin + 1)
    tosht.Cells(iGenRow, 7) = Sht.Cells(iRow, iMax + 1)
    tosht.Cells(iGenRow, 8) = Sht.Cells(iRow, iListValue + 1)
    tosht.Cells(iGenRow, 9) = Sht.Cells(iRow, iCheckNull + 1)
    tosht.Cells(iGenRow, 10) = Sht.Cells(iRow, iColumnType2 + 1)
    
    iGenRow = iGenRow + 1
    iRow = iRow + 1
  Loop
End Sub

