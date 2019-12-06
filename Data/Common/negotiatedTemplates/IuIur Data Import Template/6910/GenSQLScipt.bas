Attribute VB_Name = "GenSQLScipt"
Option Explicit
Const StartTblDataRow = 15  'should not be modified

'generate SQL Scripts for create table£¬invoked by button click
Public Sub GenSQLScripts()
  Const FileName = "Init_IUIUR_Auto.sql"
  Dim Path As String
  Dim txtFile As Object

  Dim sSqlStrLine As String

  Set txtFile = CreatePathFileObject(ActiveWorkbook.Path, FileName)
  If txtFile Is Nothing Then
    Exit Sub
  End If
  
  Dim iCount As Integer
  iCount = 0

  txtFile.Write GenTableLevelSQL
  txtFile.Write GenInitSQL
  txtFile.Write GenOrderSQL
  txtFile.Close
End Sub

'part 1: generate table level SQLs
Public Function GenTableLevelSQL() As String
  Const TableNameCol = 18
  Const TableFieldCol = 19

  Dim tSQL As String
  tSQL = ""

  Dim iRow As Integer
  Dim sTableSql As String
  Dim sPreTableName As String
  sPreTableName = "jack"

  Dim sht As Worksheet
  Set sht = Sheets("TableDef")

  iRow = StartTblDataRow
  Do Until sht.Cells(iRow, TableNameCol) = ""
    If sPreTableName <> sht.Cells(iRow, TableNameCol) Then
      sPreTableName = sht.Cells(iRow, TableNameCol)
      
      sTableSql = "if exists (select * from sysobjects where name = '" _
           + sPreTableName + "') drop table " + sPreTableName + " " + vbCrLf _
           + "go" + vbCrLf + "create table " + sPreTableName + "(" + vbCrLf _
           + "    PlanID int null ," + vbCrLf _
           + "    CMENEID int null ," + vbCrLf _
           + "    RowIdx varchar(255) null ," + vbCrLf _
           + "    SheetName varchar(255) null ," + vbCrLf
           
      Do Until sPreTableName <> sht.Cells(iRow + 1, TableNameCol)
        sTableSql = sTableSql + "    " + sht.Cells(iRow, TableFieldCol) + " varchar(255)  null," + vbCrLf
                  
        iRow = iRow + 1
      Loop
      
      sTableSql = sTableSql + "    " + sht.Cells(iRow, TableFieldCol) + " varchar(255)  null" + vbCrLf
      iRow = iRow + 1
    End If
    
    tSQL = tSQL + sTableSql + ")" + vbCrLf + "go" + vbCrLf + vbCrLf
  Loop

  GenTableLevelSQL = tSQL
End Function

Public Function GenInitSQL() As String
  Const ColXLSTableName = 18
  Const ColXLSFieldName = 19
  Const ColXLSCol = 9
  Const ColXLSStartRow = 12
  Const ColXLSEndRow = 13
  'Const Version = "16"
  
  Dim strSQL As String
  Dim strLineSQL As String
  Dim strInsertHead As String
  Dim strInsertTail As String
  Dim iRow As Integer
  Dim ShtTableDef As Worksheet
  Dim Version As String
  
  Version = Sheets("Cover").Cells(2, "E").Value
  strSQL = ""
  strLineSQL = ""
  strSQL = "delete from t_IUIUR_xlsInfo where XLSVersion = '" + Version + "'" + vbCrLf + "go" + vbCrLf
  strInsertHead = "insert into t_IUIUR_xlsInfo(XLSTableName, XLSFieldName, XLSCol, XLSStartRow, XLSEndRow, XLSVersion) values ("
      
  strInsertTail = ") " + vbCrLf
  iRow = 15
  Set ShtTableDef = Sheets("TableDef")
  
  Do Until ShtTableDef.Cells(iRow, ColXLSTableName) = ""
    strLineSQL = strInsertHead + "'" + Trim(ShtTableDef.Cells(iRow, ColXLSTableName)) + "', '" + Trim(ShtTableDef.Cells(iRow, ColXLSFieldName)) _
                               + "', '" + Trim(ShtTableDef.Cells(iRow, ColXLSCol)) + "', " + Trim(ShtTableDef.Cells(iRow, ColXLSStartRow)) _
                               + ", " + Trim(ShtTableDef.Cells(iRow, ColXLSEndRow)) + ", '" + Version + "'" _
                               + strInsertTail
    strSQL = strSQL + strLineSQL
    iRow = iRow + 1
  Loop

  strSQL = strSQL + vbCrLf + vbCrLf
  GenInitSQL = strSQL
End Function

Public Function GenOrderSQL() As String
  Const TableNameCol = 18
  Const TableFieldCol = 19
  Const COL_CATEGORY = 2
  Const COL_MOC = 3

  Dim tSQL As String
  tSQL = ""

  Dim iRow As Integer
  Dim sTableSql As String
  Dim sPreTableName As String
  sPreTableName = "jack"

  Dim sht As Worksheet
  Set sht = Sheets("TableDef")

  Dim Version As String
  Version = Sheets("Cover").Cells(2, "E").Value
  iRow = StartTblDataRow
  
  Dim iOrder As Integer
  iOrder = 0
  tSQL = "delete from t_IuIurCompare_DispOrder where Version = '" + Version + "'" + vbCrLf + "go" + vbCrLf
  Do Until sht.Cells(iRow, TableNameCol) = ""
    If sPreTableName <> sht.Cells(iRow, TableNameCol) Then
      sPreTableName = sht.Cells(iRow, TableNameCol)
      
      sTableSql = sTableSql + "insert into t_IuIurCompare_DispOrder(Version, CategoryName, MocName, OrderID)" _
        + " values('" + Version + "', '" + Trim(sht.Cells(iRow, COL_CATEGORY)) + "', '" + Trim(sht.Cells(iRow, COL_MOC)) _
        + "', " + CStr(iOrder) + ")" + vbCrLf

      iOrder = iOrder + 1
    End If
    iRow = iRow + 1
  Loop
  
  tSQL = tSQL + sTableSql + "go" + vbCrLf + vbCrLf

  GenOrderSQL = tSQL
End Function

