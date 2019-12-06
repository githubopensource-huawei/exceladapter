Attribute VB_Name = "ExportXMLModule"
Option Explicit

Public Sub ExportXML()
  Dim FileName As String
  
  FileName = ActiveWorkbook.Name
  FileName = Application.WorksheetFunction.Replace(FileName, Len(FileName) - 2, 3, "xml")
  
  Dim Path As String
  Dim txtFile As Object
  
  Dim sSqlStrAll As String
  Dim sSqlStrLine As String
  
  Set txtFile = CreatePathFileObject(ActiveWorkbook.Path, FileName)
  If txtFile Is Nothing Then
    Exit Sub
  End If

  txtFile.Write GenExportXML
  txtFile.Close
End Sub

Public Function GenExportXML() As String
  GenExportXML = "<?xml version=""1.0"" encoding=""utf-8""?>" + vbCrLf + "<ExcelTemplate Name=""Reparent NodeB Feedback"" Version=""1.0"">" + vbCrLf
  
  Dim strTable As String
  Dim strField As String
  
  Dim iRow As Integer
  Dim iParaNameRow As Integer
  
  Dim sPreTableName As String
  sPreTableName = "jack"
  
  Dim Sht As Worksheet
  Set Sht = Sheets("TableDef")
  
  iRow = StartTblDataRow
  iParaNameRow = StartTblDataRow - 2
  Do Until Sht.Cells(iRow, iMapTableName + 1) = ""
    If sPreTableName <> Sht.Cells(iRow, iMapTableName + 1) Then
      sPreTableName = Sht.Cells(iRow, iMapTableName + 1)
  
      strTable = "    <Table Name = """ _
           + sPreTableName + """>" + vbCrLf _
           
      strField = ""
      Do Until sPreTableName <> Sht.Cells(iRow + 1, iMapTableName + 1)
        strField = strField + "        <Field " _
                 + "Name" + "=""" + CStr(Sht.Cells(iRow, iMapFieldName + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iMocName + 1)) + "=""" + CStr(Sht.Cells(iRow, iMocName + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iColumnType + 1)) + "=""" + CStr(Sht.Cells(iRow, iColumnType + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iMin + 1)) + "=""" + CStr(Sht.Cells(iRow, iMin + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iMax + 1)) + "=""" + CStr(Sht.Cells(iRow, iMax + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iListValue + 1)) + "=""" + CStr(Sht.Cells(iRow, iListValue + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iFieldBeginColumn + 1)) + "=""" + CStr(Sht.Cells(iRow, iFieldBeginColumn + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iTitleBeginRow + 1)) + "=""" + CStr(Sht.Cells(iRow, iTitleBeginRow + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iContentEndRow + 1)) + "=""" + CStr(Sht.Cells(iRow, iContentEndRow + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iMapTableName + 1)) + "=""" + CStr(Sht.Cells(iRow, iMapTableName + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iMapFieldName + 1)) + "=""" + CStr(Sht.Cells(iRow, iMapFieldName + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iCheckNull + 1)) + "=""" + CStr(Sht.Cells(iRow, iCheckNull + 1)) + """ " _
                 + CStr(Sht.Cells(iParaNameRow, iColumnType2 + 1)) + "=""" + CStr(Sht.Cells(iRow, iColumnType2 + 1)) + """" _
                 + ">" + vbCrLf + "        </Field>" + vbCrLf
  
        iRow = iRow + 1
      Loop
  
      strTable = strTable + strField + "    </Table>"
      GenExportXML = GenExportXML + strTable + vbCrLf
      iRow = iRow + 1
    End If
  Loop
  
  GenExportXML = GenExportXML + "</ExcelTemplate>"
End Function




