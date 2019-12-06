Attribute VB_Name = "TableDef_Sheet"
Private Sub ExportXMLCommandButton_Click()
 ''
  Call ExportXML
End Sub

Private Sub GenExcelInitCommandButton_Click()
  Call GenInitTableMap
  Call GenInitFieldMap
  
  MsgBox "Generate Initial Information in Excel succeeded.", vbOKOnly
End Sub

Private Sub GenNegotiatedFile_Click()
  Call DefineNegotiatedFile.GenNegotiatedFile
End Sub

Private Sub GenSQLButton_Click()
  Call GenSQLScripts
End Sub



