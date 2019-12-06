Attribute VB_Name = "T_Worksheet"
Option Explicit

Public Const CODE_REMARK = "'The following codes are generated automatically after refresh template."

Private Const TableBeginRow = 3
Private Const TableEndRow = 65535 + 1
Private Const TableBeginCol = 1

Private TableEndCol As Long
Private m_nDefSheetIndex As Long   '当前表的定义参数在TableDef表中的起始行索引

Public Sub RebuildSheet(wk As Workbook)
    Dim Sheets As Variant, SheetsAfter As Variant, SheetsPrefix As Variant, IsGetValidDefineData As Variant
    Dim SheetsCheckFlag1 As Variant, SheetsCheckFlag2 As Variant, SheetsCheckFlag3 As Variant, SheetsCheckFlag4 As Variant, SheetsCheckFlag5 As Variant
    Sheets = Array(MOC_BSCINFO, MOC_NODEB, MOC_CELL, MOC_NRNCCELL, MOC_INTRAFREQNCELL, MOC_INTERFREQNCELL, _
        MOC_GSMCELL, MOC_GSMNCELL, MOC_LTECELL, MOC_LTENCELL, MOC_SMLCCELL, MOC_PHY_NB_RADIO, MOC_WHOLE_NETWORK_CELL, MOC_DEL_INTERFREQNCELL)
        
    SheetsAfter = Array(MOC_NODEB, MOC_CELL, MOC_NRNCCELL, MOC_INTRAFREQNCELL, MOC_INTERFREQNCELL, MOC_GSMCELL, _
        MOC_GSMNCELL, MOC_LTECELL, MOC_LTENCELL, MOC_SMLCCELL, MOC_PHY_NB_RADIO, MOC_DOUBLE_FREQ_CELL, MOC_DEL_INTERFREQNCELL, SHT_CONVERT_TEMPLATE)
    SheetsPrefix = Array("T_", "T_", "T_", "T_", "T_", "T_", "T_", "T_", "T_", "T_", "T_", "T_", "DF_", "DF_")
    
    SheetsCheckFlag1 = Array(False, True, True, True, True, True, True, True, True, True, True, False, True, True) 'GetValidDefineData
    SheetsCheckFlag2 = Array(False, True, True, True, True, True, True, True, True, True, True, False, True, True) 'SetInvalidateField
    SheetsCheckFlag3 = Array(False, False, True, True, False, False, False, False, False, False, False, True, True, False) 'SetFieldValidation
    SheetsCheckFlag4 = Array(True, True, False, False, True, True, True, True, True, True, True, True, False, True) 'CheckFieldData
    SheetsCheckFlag5 = Array(False, False, True, True, False, False, False, False, False, False, False, False, True, False) 'CellCheckFieldData

    Dim i As Integer, s As String
    For i = UBound(Sheets) To LBound(Sheets) Step -1
        If SheetExist(wk, SheetsAfter(i)) Then
            s = CStr(SheetsCheckFlag1(i)) + ", " + CStr(SheetsCheckFlag2(i)) + ", " + CStr(SheetsCheckFlag3(i)) + ", " + CStr(SheetsCheckFlag4(i)) + ", " + CStr(SheetsCheckFlag5(i))
            Call AddSheetWithVBA(wk, Sheets(i), SheetsAfter(i), SheetsPrefix(i), s)
        Else
            MsgBox "After sheet '" + SheetsAfter(i) + "' does not exist."
        End If
    Next i
End Sub

Private Sub AddSheetWithVBA(wk As Workbook, ByVal CurrentSheet As String, ByVal AfterSheet As String, ByVal ModuleNamePrefix As String, CheckFlags As String)
    Dim sht As Worksheet, sht2 As Worksheet
    If SheetExist(wk, CurrentSheet) Then
        Dim isDisplayAlerts As Boolean
        isDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        wk.Sheets(CurrentSheet).Delete
        Application.DisplayAlerts = isDisplayAlerts
    End If
    Set sht2 = wk.Sheets(AfterSheet)
    Set sht = wk.Sheets.Add(Before:=sht2)
    sht.Name = CurrentSheet
    Dim s As String
    With wk.VBProject.VBComponents(sht.CodeName)
        .Name = ModuleNamePrefix + sht.Name
        s = CODE_REMARK + vbLf
        s = s + GenWorksheetActivateEvent
        s = s + GenWorksheetChangeEvent(CheckFlags)
        .CodeModule.AddFromString (s)
    End With
End Sub

Private Function GenWorksheetActivateEvent() As String
    Dim s As String
    s = vbLf
    s = s + "Private Sub Worksheet_Activate()" + vbLf
    s = s + "    Call Do_Worksheet_Activate(Me)" + vbLf
    s = s + "End Sub" + vbLf
    GenWorksheetActivateEvent = s
End Function

Private Function GenWorksheetChangeEvent(CheckFlags As String) As String
    Dim s As String
    s = vbLf
    s = s + "Private Sub Worksheet_Change(ByVal Target As Range)" + vbLf
    s = s + "    Call Do_Worksheet_Change(Me, Target," + CheckFlags + ")" + vbLf
    s = s + "End Sub" + vbLf
    GenWorksheetChangeEvent = s
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
    If Target.Row > TableEndRow Or Target.Row < TableBeginRow Or Target.Column > TableEndCol Or Target.Column < TableBeginCol Then
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

