Attribute VB_Name = "ID_Modification_eNodeB"
'Option Explicit

Public Sub eNodeBShtChange(sht As Worksheet, target As Range)
   Dim c As Variant
   For Each c In target
    If c.Column = 3 And c.Row > 2 Then
     If c.text <> "" Then
         If isNum(c.text) = False Then
            errID_eNodeBID = 1
            c.Value = ""
         ElseIf Val(c.Value) > 1048575 Or Val(c.Value) < 0 Then
            errID_eNodeBID = 1
            c.Value = ""
         End If
      End If
    End If
   Next c

   If errID_eNodeBID = 1 Then
    rc = MsgBox("eNodeBID input error. The value range must be from 0 to 1048575.", vbCritical + vbOKOnly, "Error")
   End If
End Sub


Public Sub cellShtChange(sht As Worksheet, target As Range)
   Dim c As Variant
   For Each c In target
    If c.Column = 3 And c.Row > 2 Then
     If c.text <> "" Then
         If isNum(c.Value) = False Then
            errID_CellID = 2
            c.Value = ""
         ElseIf Val(c.Value) > 255 Or Val(c.Value) < 0 Then
            errID_LocalCellID = 1
            c.Value = ""
         End If
      End If
    End If
    
    If c.Column = 4 And c.Row > 2 Then
     If c.text <> "" Then
         If isNum(c.Value) = False Then
            errID_CellID = 2
            c.Value = ""
         ElseIf Val(c.Value) > 255 Or Val(c.Value) < 0 Then
            errID_CellID = 2
            c.Value = ""
         End If
      End If
    End If
   Next c
   
   If errID_LocalCellID = 1 Then
    rc = MsgBox("LocalCellId input error. The value range must be from 0 to 255.", vbCritical + vbOKOnly, "Error")
   End If
   
   If errID_CellID = 2 Then
    rc = MsgBox("CellId input error. The value range must be from 0 to 255.", vbCritical + vbOKOnly, "Error")
   End If
   
End Sub


Private Function isNum(text As String) As Boolean
         Dim sItem As String
         Dim nLoop As Long
    
         For nLoop = 1 To Len(Trim(text))
            sItem = Right(Left(Trim(text), nLoop), 1)
            If sItem < "0" Or sItem > "9" Then
                If nLoop = 1 And sItem = "-" Then
                    bFlag = True
                Else
                    bFlag = False
                    isNum = False
                    Exit Function
                End If
            End If
         Next
         isNum = True
End Function

