Attribute VB_Name = "SetRange"
Const NullColorIdx = 0
Const TableBeginCol = 2
Const TableEndCol = 27 + 1
Const TableBeginRow = 5
Const TableEndRow = 3920 + 1



'color
Const SolidColorIdx = 16
Const SolidPattern = xlGray16
Const NullPattern = xlNone

Const AllRange = "B4:BZ300"
Public Sub SetInValidationMode(FieldRange As String)
    Range(FieldRange).Interior.ColorIndex = SolidColorIdx
    Range(FieldRange).Interior.Pattern = SolidPattern
    Call ClearCell(Range(FieldRange))
End Sub
Public Sub SetAvailabilityMode(FieldRange As String)
    Range(FieldRange).Interior.ColorIndex = NullPattern
    Range(FieldRange).Interior.Pattern = NullPattern
End Sub
Public Sub ClearCell(CurCell As Range)
    If (Trim(CurCell.Value) <> "") Then
        CurCell.Value = ""
    End If
End Sub
Public Sub SetFieldRange(TableName As String, FieldName As String, FieldValue As String, iRows As Integer)
    Dim sheetValidDef As Worksheet
    Dim iRow As Integer
    
    Set sheetValidDef = ThisWorkbook.Sheets("ValidDef")
    For iRow = 4 To 30000
          If (Trim(sheetValidDef.Cells(iRow, 1).Value) = "") Or (Trim(sheetValidDef.Cells(iRow, 2).Value) = "") Then
               Exit For
            End If
            
          If (Trim(sheetValidDef.Cells(iRow, 1).Value) = TableName) And (Trim(sheetValidDef.Cells(iRow, 2).Value) = FieldName) Then
          
              If (Trim(sheetValidDef.Cells(iRow, 8).Value) <> "") Then
                   
                   If IsSubStr(FieldValue, Trim(sheetValidDef.Cells(iRow, 6).Value)) Then
                        Call SetInValidationMode(Trim(sheetValidDef.Cells(iRow, 8).Value) + CStr(iRows))
                      Else
                        Call SetAvailabilityMode(Trim(sheetValidDef.Cells(iRow, 8).Value) + CStr(iRows))
                      End If
                      
                 End If
              
            End If
          
    Next iRow
End Sub
Public Function IsSubStr(substr As String, str As String) As Boolean
    Dim ArrData() As String
    Dim i As Integer
    
    IsSubStr = False
    
    ArrData = Split(str, ",")
    For i = 0 To UBound(ArrData)
        If Trim(ArrData(i)) = Trim(substr) Then
            IsSubStr = True
            Exit Function
        End If
    Next
End Function
Public Sub SetWorksheetChange(Target As Range)
  '  Dim CurSheet As Worksheet
    Dim sFieldName As String
    Dim sTableName As String
    Dim sValue As String
        
    If GeneratingFlag = 1 Then  '刷新时不进入
        Exit Sub
    End If
    
     
    If Target.Row > TableEndRow Or Target.Row < TableBeginRow Or Target.Column > TableEndCol Or Target.Column < TableBeginCol Then
        Exit Sub
    End If
        
    'Call Ensure_NoValue(Target)
  
    Set CurSheet = ActiveWindow.ActiveSheet
    
    '无效格直接退出
    If Target.Interior.ColorIndex = 16 Then
         Call ClearCell(Target)
         Exit Sub
     End If
      
    sTableName = CurSheet.Name
    sFieldName = CurSheet.Cells(5, Target.Column).Value
    sValue = CurSheet.Cells(Target.Row, Target.Column).Value
    
    Call SetFieldRange(sTableName, sFieldName, sValue, Target.Row)
    

    On Error Resume Next
End Sub
