Attribute VB_Name = "TableDef_Code"


Private Sub cmdBtnLoadParams_Click()
  Dim xx As New CSupportDataReader
  xx.InitDataSource (SDR_Language_Chinese)
  xx.InitDataSource (SDR_Language_English)
  MsgBox "Íê³É"
  Application.StatusBar = ""
End Sub

Private Sub CommandButton1_Click()
  Dim CurrentSheet As Worksheet
  Dim TableName As String
  Dim i As Integer
  
  Set CurrentSheet = ThisWorkbook.Sheets("TableList")
  
  For i = 2 To 30000
    TableName = CurrentSheet.Cells(i, 1).Value
    If TableName = "" Then
        Exit For
      End If
      
  If SheetExist(ThisWorkbook, TableName) Then
      Call AddSheetWithVBA(ThisWorkbook, TableName, "T_" + TableName)
   End If
  Next i
  CommandButton1.Visible = False
  
End Sub

Private Sub GenNegotiatedFile_Click()
  Call DefineNegotiatedFile.GenNegotiatedFile
  CommandButton1.Visible = True
  
End Sub
Private Sub AddSheetWithVBA(wk As Workbook, ByVal shtName As String, ByVal CurrentSheet As String)
    Dim sht As Worksheet
    Dim s As String
    
    Set sht = ThisWorkbook.Sheets(shtName)
    
    With wk.VBProject.VBComponents(sht.CodeName)
        '.Name = CurrentSheet
        s = s + GenWorksheetChangeEvent()
        .CodeModule.AddFromString (s)
        
    End With
End Sub
Private Function GenWorksheetChangeEvent() As String
    Dim s As String
    s = vbLf
    s = s + "Private Sub Worksheet_Change(ByVal rTarget As Range)" + vbLf
    s = s + "    Call SetWorksheetChange(rTarget)" + vbLf
    s = s + "End Sub" + vbLf
    GenWorksheetChangeEvent = s
End Function
Public Function SheetExist(wk As Workbook, ByVal SheetName As String) As Boolean
    On Error GoTo E
    Dim sht As Worksheet
    Set sht = wk.Sheets(SheetName)
    SheetExist = True
    Exit Function
E:
    SheetExist = False
End Function


