Attribute VB_Name = "DoubleFrequencyCell"
Option Explicit

Const TableBeginRow = 4
Const TableEndRow = 65535 + 1
Const TableBeginCol = 1
Private TableEndCol As Long

Private m_nDefSheetIndex As Long   '当前表的定义参数在TableDef表中的起始行索引

Public Sub cmdConfigGSMNCell_Click()
    Call ConfigNCellRelation(rtDiffSystem, RSC_STR_GSMNCELL, otADD)
End Sub

Public Sub cmdConfigInterNCellDiffSector_Click()
    Call ConfigNCellRelation(rtDiffFreqDiffSector, RSC_STR_INTRANCELL_DIFF_SECTOR, otADD)
End Sub

Public Sub cmdConfigInterNCellSameSector_Click()
    Call ConfigNCellRelation(rtDiffFreqSameSector, RSC_STR_INTERNCELL_SAME_SECTOR, otADD)
End Sub

Public Sub cmdConfigIntraNCell_Click()
    Call ConfigNCellRelation(rtSameFreq, RSC_STR_INTRANCELL, otADD)
End Sub

Public Sub cmdCopyDataFromCELL_Click()
    Application.ScreenUpdating = False

    Dim iEnd As Long, sht1 As Worksheet, sht2 As Worksheet, Rng1 As String, Rng2 As String
    Set sht1 = Sheets(MOC_CELL)
    Set sht2 = Sheets(MOC_DOUBLE_FREQ_CELL)
    iEnd = GetLastRowIndex(sht1)
    If iEnd > ROW_DATA_HW Then
        Rng1 = "A3:B" + CStr(iEnd)
        Rng2 = "A4:B" + CStr(iEnd + 1)
        sht1.Range(Rng1).Copy Destination:=sht2.Range(Rng2)
        Rng1 = "C3:V" + CStr(iEnd)
        Rng2 = "D4:W" + CStr(iEnd + 1)
        sht1.Range(Rng1).Copy Destination:=sht2.Range(Rng2)
    End If
    sht2.Activate

    Dim i As Integer, s As String
    Const COL_SECTORID = 3
    Const COL_CELLID = 4
    For i = ROW_DATA_HW + 1 To iEnd + 1
        s = sht2.Cells(i, COL_CELLID).Value
        If Len(s) > 0 Then
            sht2.Cells(i, COL_SECTORID).Value = GetSectorID(s)
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

Public Sub cmdCopyDataToCELL_Click()
    Application.ScreenUpdating = False

    Dim iEnd As Long, sht1 As Worksheet, sht2 As Worksheet, Rng1 As String, Rng2 As String
    Set sht1 = Sheets(MOC_DOUBLE_FREQ_CELL)
    Set sht2 = Sheets(MOC_CELL)
    iEnd = GetLastRowIndex(sht1)
    If iEnd > ROW_DATA_HW Then
        Rng1 = "A4:B" + CStr(iEnd) + ",D4:W" + CStr(iEnd)
        Rng2 = "A3:V" + CStr(iEnd - 1)
        sht1.Range(Rng1).Copy Destination:=sht2.Range(Rng2)
    End If
    sht2.Activate

    Application.ScreenUpdating = True
End Sub

Private Sub cmdDeleteInterNCellDiffSector_Click()
    Call ConfigNCellRelation(rtDiffFreqDiffSector, RSC_STR_INTRANCELL_DIFF_SECTOR, otDEL)
End Sub

Public Sub cmdSetFormula_Click()
    frmSetFormula.txtFormula.Text = GetFormula
    frmSetFormula.Show
End Sub

Public Sub DoubleFrequency_Worksheet_Activate(sht As Worksheet)
'取得当前表的定义参数在TableDef表中的起始行索引
    Dim sSheetName As String
    sSheetName = sht.Name
    
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

Public Sub DoubleFrequency_Worksheet_Change(sht As Worksheet, ByVal Target As Range)
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
    Call GetValidDefineData
    Call SetInvalidateField(Target, CurSheet.Name)
    Call SetFieldValidation(Target, CurSheet.Name)
    
    '检查输入的值是否符合数据有效性规则
    Call CellCheckFieldData(m_nDefSheetIndex, Target)
    
End Sub


