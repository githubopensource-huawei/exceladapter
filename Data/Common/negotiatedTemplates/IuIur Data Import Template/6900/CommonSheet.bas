Attribute VB_Name = "CommonSheet"
Option Explicit

Const TableBeginCol = 2
Const TableEndCol = 52 + 1
Const TableBeginRow = 8

'color
Const SolidColorIdx = 16
Const SolidPattern = xlGray16
Const NullPattern = xlNone


Public Sub Common_Ensure_NoValue(ByVal Target As Range)
  Dim CurRange As Range
  
  For Each CurRange In Target
    If CurRange.Cells(1, 1) <> "" And CurRange.Cells(1, 1).Interior.ColorIndex = SolidColorIdx And CurRange.Cells(1, 1).Interior.Pattern = SolidPattern Then
      NoValueNeeded ("B" + CStr(CurRange.Row))
      CurRange.Cells(1, 1).ClearContents
    End If
  Next CurRange
End Sub

Public Sub Common_RefreshThisSheet(ByVal Target As Range, isActivateFlag As Boolean)
    
    If GeneratingFlag = 1 Then  '刷新时不进入
        Exit Sub
    End If
    
    Dim CurSheet As Worksheet
    Set CurSheet = ActiveWindow.ActiveSheet
    
    If CurSheet.Cells(Target.Row, 1).MergeArea.Cells(1, 1) = "" Or Target.Row < TableBeginRow Or Target.Column > TableEndCol Or Target.Column < TableBeginCol Then
        Exit Sub
    End If
        
    Call Common_Ensure_NoValue(Target)
  
    Call UnprotectSheet(CurSheet)
    On Error Resume Next
    
    Call GetValidDefineData
    Call SetInvalidateField(Target, isActivateFlag, CurSheet.Name)
    Call SetFieldValidation(Target, CurSheet.Name)

    Call ProtectSheet(CurSheet)
End Sub


