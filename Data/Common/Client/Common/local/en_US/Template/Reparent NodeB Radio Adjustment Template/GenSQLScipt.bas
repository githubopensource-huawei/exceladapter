Attribute VB_Name = "GenSQLScipt"
Option Explicit
Const StartTblDataRow = 15  'should not be modified

'generate SQL Scripts for create table£¬invoked by button click
Public Sub GenSQLScripts()
  Const FileName = "FeedbackSQL.sql"
  Dim Path As String
  Dim txtFile As Object

  Dim sSqlStrAll As String
  Dim sSqlStrLine As String

  Set txtFile = CreatePathFileObject(ActiveWorkbook.Path, FileName)
  If txtFile Is Nothing Then
    Exit Sub
  End If
  
  Dim iCount As Integer
  iCount = 0

  txtFile.Write GenTableLevelSQL
  txtFile.Write GenInitSQL
  txtFile.Write sSqlStrAll
  txtFile.Close
End Sub

'part 1: generate table level SQLs
Public Function GenTableLevelSQL() As String
  Const TableNameCol = 17
  Const TableFieldCol = 18

  Dim tSQL As String
  tSQL = ""

  Dim iRow As Integer
  Dim sTableSql As String
  Dim sPreTableName As String
  sPreTableName = "jack"

  Dim Sht As Worksheet
  Set Sht = Sheets("TableDef")

  iRow = StartTblDataRow
  Do Until Sht.Cells(iRow, TableNameCol) = ""
    If sPreTableName <> Sht.Cells(iRow, TableNameCol) Then
      sPreTableName = Sht.Cells(iRow, TableNameCol)
      
      sTableSql = "if exists (select * from sysobjects where name = '" _
           + sPreTableName + "') drop table " + sPreTableName + " " + vbCrLf _
           + "go" + vbCrLf + "create table " + sPreTableName + "(" + vbCrLf _
           + "    PlanID int null ," + vbCrLf _
           + "    CMENEID int null ," + vbCrLf _
           + "    RowIdx varchar(255) null ," + vbCrLf _
           + "    SheetName varchar(255) null ," + vbCrLf
           
      Do Until sPreTableName <> Sht.Cells(iRow + 1, TableNameCol)
        sTableSql = sTableSql + "    " + Sht.Cells(iRow, TableFieldCol) + " varchar(255)  null," + vbCrLf
                  
        iRow = iRow + 1
      Loop
      
      sTableSql = sTableSql + "    " + Sht.Cells(iRow, TableFieldCol) + " varchar(255)  null" + vbCrLf
      iRow = iRow + 1
    End If
    
    tSQL = tSQL + sTableSql + ")" + vbCrLf + "go" + vbCrLf + vbCrLf
  Loop

  GenTableLevelSQL = tSQL
End Function

Public Function GenInitSQL() As String
  Const ColXLSTableName = 17
  Const ColXLSFieldName = 18
  Const ColXLSCol = 8
  Const ColXLSStartRow = 11
  Const ColXLSEndRow = 12
  Const Version = "11"
  
  Dim strSQL As String
  Dim strLineSQL As String
  Dim strInsertHead As String
  Dim strInsertTail As String
  Dim iRow As Integer
  Dim ShtTableDef As Worksheet
    
  
  strSQL = ""
  strLineSQL = ""
  strInsertHead = "insert into t_Rpt_xlsInfo(XLSTableName, XLSFieldName, XLSCol, XLSStartRow, XLSEndRow, XLSVersion) values ("
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

