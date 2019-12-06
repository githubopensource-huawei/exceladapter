Attribute VB_Name = "GenerateSheet"
Public Sub RebuildSheet()
    Dim i As Integer, s As String, CheckFlags As String
    Dim CurSheet As Worksheet
    For i = ActiveWorkbook.Sheets.Count To 1 Step -1
       Set CurSheet = ActiveWorkbook.Sheets(i)
       If StrComp(CurSheet.Name, "Refresh", vbTextCompare) <> 0 And StrComp(CurSheet.Name, "TableInfo", vbTextCompare) <> 0 And StrComp(CurSheet.Name, "Cover", vbTextCompare) <> 0 _
       And StrComp(CurSheet.Name, "ValidInfo", vbTextCompare) <> 0 _
       And StrComp(CurSheet.Name, "Home", vbTextCompare) _
       And StrComp(CurSheet.Name, "CHS", vbTextCompare) _
       And StrComp(CurSheet.Name, "ENG", vbTextCompare) Then
         CurSheet.Unprotect Password:=GetSheetsPass
         With ActiveWorkbook.VBProject.VBComponents(CurSheet.CodeName)
             If CurSheet.Name = "CELL" Then
                 CheckFlags = CStr("True, True, True, False, True")
             Else
                 CheckFlags = CStr("True, True, False, True, False")
             End If
             s = GenWorksheetActivateEvent(CheckFlags)
         .CodeModule.AddFromString (s)
         
         '.CodeModule.DeleteLines (1)
   
         
         End With
        End If
   Next i
End Sub

Private Function GenWorksheetActivateEvent(CheckFlags As String) As String
    Dim s As String
    s = vbLf
    s = s + "Private Sub Worksheet_Activate()" + vbLf
    s = s + "    Call Do_Worksheet_Activate(Me)" + vbLf
    s = s + "End Sub" + vbLf
    s = s + "Private Sub Worksheet_Change(ByVal Target As Range)" + vbLf
    s = s + "    Call Do_Worksheet_Change(Me, Target," + CheckFlags + ")" + vbLf
    s = s + "End Sub" + vbLf
    GenWorksheetActivateEvent = s
End Function
Public Sub Do_Worksheet_Activate(sht As Worksheet)
'取得当前表的定义参数在TableDef表中的起始行索引
    Dim CurrSheet As Worksheet
    Set CurrSheet = sht
    Dim sSheetName As String
    sSheetName = CurrSheet.Name
    
    Dim nDefSheetIndex As Integer
    Dim sID As String, sDefSheetName As String
    sDefSheetName = ""

    For nDefSheetIndex = 0 To UBound(SheetDefine) - 1
        sID = Trim(SheetDefine(nDefSheetIndex, 0))
        If sID <> "" Then
            sDefSheetName = Trim(SheetDefine(nDefSheetIndex, 1))
            If sSheetName = sDefSheetName Then
                m_nDefSheetIndex = nDefSheetIndex
                TableEndCol = GetSheetColCount(Trim(SheetDefine(nDefSheetIndex + 1, 1)))
                Exit For
            End If
        End If
    Next
    
End Sub

Public Sub Do_Worksheet_Change(sht As Worksheet, ByVal Target As Range, CheckFlag1 As Boolean, CheckFlag2 As Boolean, CheckFlag3 As Boolean, CheckFlag4 As Boolean, CheckFlag5 As Boolean)
    Call Do_Worksheet_Activate(sht)

    If GeneratingFlag = 1 Then  '刷新时不进入
        Exit Sub
    End If
    
    Dim CurSheet As Worksheet
    If Target.Row > TableEndRow Or Target.Row < TableBeginRow Or Target.Column < TableBeginCol Then
        Exit Sub
    End If
    Call Ensure_NoValue(Target)
    
    Set CurSheet = sht
    On Error Resume Next
    If CheckFlag1 Then
        Call GetValidDefineData
    End If
    If CheckFlag2 Then
        Call SetInvalidateField(Target, CurSheet.Name)
    End If
    If CheckFlag3 Then
        Call SetFieldValidation(Target, CurSheet.Name)
    End If
    '检查输入的值是否符合数据有效性规则
    If CheckFlag4 Then
        Call CheckFieldData(m_nDefSheetIndex, Target)
    End If
    If CheckFlag5 Then
        Call CellCheckFieldData(m_nDefSheetIndex, Target)
    End If
End Sub

