Attribute VB_Name = "SystemConf"
Option Explicit



Public Function IsGBTSTemplate() As Boolean
    
    IsGBTSTemplate = False
    
    If IsExistsSheet("FUNCTION_SHEET") Then
        If ("DATAITF_BTS" = ThisWorkbook.Sheets("FUNCTION_SHEET").Cells(1, 1).value) Then
            IsGBTSTemplate = True
        End If
    End If
End Function

Function IsExistsSheet(sheetName As String) As Boolean
  Dim ShtIdx As Long
  Dim OpSht As Worksheet
  
  ShtIdx = 1
  Do While (ShtIdx <= ActiveWorkbook.Sheets.count)
      Set OpSht = ActiveWorkbook.Sheets(ShtIdx)
      If OpSht.name = sheetName Then
        IsExistsSheet = True
        Exit Function
      End If
      ShtIdx = ShtIdx + 1
  Loop
  IsExistsSheet = False
End Function


