Attribute VB_Name = "DF_ConfigDiffSystem"
Option Explicit

Private Const SHT_TEMP_4 = "TempSheet4"
Private Const SHT_TEMP_5 = "TempSheet5"

Private m_clsGSMCELL_SMC As New TableClass '缓存数据到内存表格中
Private m_clsGSMCELL_SC As New TableClass '缓存数据到内存表格中
Private m_clsMocGSMNCELL As New MocClass 'Mapping defination of GSMNCELL Moc, only stores key attr, and used to filter data.
Private m_clsMocGSMCELL As New MocClass 'Mapping defination of GSMCELL Moc, only stores key attr, and used to filter data.

Public Sub ConfigDiffSystem(BSCName As String, RNCID As String, SelectedCellID As String, SectorMappingCellID As String)
    Call PrintToDebugger("Processing DiffSystem......")

    Call InitModuleVars

    If GetLastRowIndex(Sheets(MOC_GSMNCELL)) < ROW_DATA_HW Then
        Call PrintToReport(FormatStr(RSC_STR_SHEET_NO_RECORD, MOC_GSMNCELL))
        Exit Sub
    End If

    If GetLastRowIndex(Sheets(MOC_GSMCELL)) < ROW_DATA_HW Then
        Call PrintToReport(FormatStr(RSC_STR_SHEET_NO_RECORD, MOC_GSMCELL))
        Exit Sub
    End If

    '查找扇区映射小区的GSM邻区关系，并在此基础上查找其GSM小区，最后收集到内存表格中
    Call GetGSMCell(Sheets(SHT_TEMP_4), Sheets(SHT_TEMP_4), SectorMappingCellID, m_clsGSMCELL_SMC, "SectorMappingCell")
    If m_clsGSMCELL_SMC.GetRecordCount <= 0 Then
        Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_GSMNCELL))
        Exit Sub
    End If
    
    '查找当前选中小区的GSM邻区关系，并在此基础上查找其GSM小区，最后收集到内存表格中
    Call GetGSMCell(Sheets(SHT_TEMP_5), Sheets(SHT_TEMP_5), SelectedCellID, m_clsGSMCELL_SC, "SelectedCell")

    '删除已经存在GSM邻区关系的GSM小区
    Call RemoveGSMCellExisted(m_clsGSMCELL_SMC, m_clsGSMCELL_SC)
    If m_clsGSMCELL_SMC.GetRecordCount <= 0 Then
        Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_GSMNCELL))
        Exit Sub
    End If

    '删除NCC BCC BCCHARFCN BANDIND相同的GSM小区
    Call RemoveInvalidGSMCell(m_clsGSMCELL_SMC, m_clsGSMCELL_SMC, True)
    Call RemoveInvalidGSMCell(m_clsGSMCELL_SMC, m_clsGSMCELL_SC, False)
    If m_clsGSMCELL_SMC.GetRecordCount <= 0 Then
        Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_GSMNCELL))
        Exit Sub
    End If

    '写数据到相应Sheet页面
    Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, m_clsGSMCELL_SMC.GetRecordCount, MOC_GSMNCELL))
    Call WriteData(BSCName, RNCID, SelectedCellID)
End Sub

Private Sub GetGSMCell(Dst1Sheet As Worksheet, Dst2Sheet As Worksheet, CellID As String, Tbl As TableClass, DebugInfo As String)
    Call PrintToDebugger("Searching GSMNCell of " + DebugInfo + "......")
    Call DoAdvancedFilter(Sheets(MOC_GSMNCELL), Dst1Sheet, m_clsMocGSMNCELL, Array(ATTR_CELLID + "=" + CellID))
    
    Dim i As Long, iEnd As Long, Conditions() As String, iIndex As Long
    iEnd = GetLastRowIndex(Dst1Sheet)
    iIndex = 0
    Dim Attr1 As AttrClass, Attr2 As AttrClass, s As String, Value As String
    Set Attr1 = m_clsMocGSMNCELL.Attr(ATTR_BSCNAME)
    Set Attr2 = m_clsMocGSMNCELL.Attr(ATTR_GSMCELLINDEX)
    For i = ROW_DATA_HW To iEnd
        Value = Dst1Sheet.Cells(i, Attr1.ColIndex).Value
        s = Attr1.Name + "=" + Value
        Value = Dst1Sheet.Cells(i, Attr2.ColIndex).Value
        s = s + "," + Attr2.Name + "=" + Value
        ReDim Preserve Conditions(iIndex)
        Conditions(iIndex) = s
        iIndex = iIndex + 1
    Next i

    Call PrintToDebugger("Searching GSMCell of " + DebugInfo + "......")
    If iIndex > 0 Then
        Call DoAdvancedFilter2(Sheets(MOC_GSMCELL), Dst2Sheet, m_clsMocGSMCELL, Conditions)

        iEnd = GetLastRowIndex(Dst2Sheet)
        Dim j As Long, iRecordIndex As Long
        For i = ROW_DATA_HW To iEnd
            Tbl.AddRecord
            iRecordIndex = Tbl.GetRecordCount
            For j = 1 To Tbl.GetFieldCount
                Set Attr1 = Tbl.GetValueAt(iRecordIndex, j)
                Attr1.Value = Dst2Sheet.Cells(i, Attr1.ColIndex).Value
            Next j
        Next i
    Else
        Call PrintToDebugger("No Condition.")
    End If
End Sub

Private Sub RemoveGSMCellExisted(Tbl As TableClass, Tbl2 As TableClass)
    Dim i As Long, j As Long, isFound As Boolean, isRemoved As Boolean
    isRemoved = False
    For i = 1 To Tbl.GetRecordCount
        isFound = False
        For j = 1 To Tbl2.GetRecordCount
            If Tbl.GetValueAt(i, ATTR_BSCNAME).Value = Tbl2.GetValueAt(j, ATTR_BSCNAME).Value _
                And Tbl.GetValueAt(i, ATTR_GSMCELLINDEX).Value = Tbl2.GetValueAt(j, ATTR_GSMCELLINDEX).Value Then
                isFound = True
                Exit For
            End If
        Next j
        If isFound Then
            Tbl.Remove (i)
            isRemoved = True
            Exit For
        End If
    Next i

    If isRemoved Then
        Call RemoveGSMCellExisted(Tbl, Tbl2)
    End If
End Sub

Private Sub RemoveInvalidGSMCell(Tbl As TableClass, Tbl2 As TableClass, IsSameTable As Boolean)
    Dim i As Long, j As Long, isFound As Boolean, isRemoved As Boolean
    isRemoved = False
    For i = 1 To Tbl.GetRecordCount
        isFound = False
        For j = 1 To Tbl2.GetRecordCount
            If Not (IsSameTable And j = i) Then
                If Tbl.GetValueAt(i, ATTR_NCC).Value = Tbl2.GetValueAt(j, ATTR_NCC).Value _
                    And Tbl.GetValueAt(i, ATTR_BCC).Value = Tbl2.GetValueAt(j, ATTR_BCC).Value _
                    And Tbl.GetValueAt(i, ATTR_BCCHARFCN).Value = Tbl2.GetValueAt(j, ATTR_BCCHARFCN).Value _
                    And Tbl.GetValueAt(i, ATTR_BANDIND).Value = Tbl2.GetValueAt(j, ATTR_BANDIND).Value Then
                    isFound = True
                    Exit For
                End If
            End If
        Next j
        If isFound Then
            Tbl.Remove (i)
            isRemoved = True
            Exit For
        End If
    Next i

    If isRemoved Then
        Call RemoveInvalidGSMCell(Tbl, Tbl2, IsSameTable)
    End If
End Sub

Private Sub WriteData(BSCName As String, RNCID As String, CellID As String)
    Dim iRow1 As Long, iRow2 As Long, iRecordIndex As Long
    iRow1 = GetLastRowIndex(Sheets(MOC_GSMNCELL))
    iRow1 = iRow1 + 1
    iRow2 = iRow1 - 1 + m_clsGSMCELL_SMC.GetRecordCount
    iRecordIndex = 1
    Dim i As Long, j As Integer, Moc As MocClass, Attr As AttrClass, GSMCellIndex As String
    Set Moc = g_c_clsMocInfos.Moc(MOC_GSMNCELL)

    With Sheets(MOC_GSMNCELL)
        For i = iRow1 To iRow2
            .Cells(i, m_clsMocGSMNCELL.Attr(ATTR_BSCNAME).ColIndex).Value = BSCName
            .Cells(i, m_clsMocGSMNCELL.Attr(ATTR_RNCID).ColIndex).Value = RNCID
            .Cells(i, m_clsMocGSMNCELL.Attr(ATTR_CELLID).ColIndex).Value = CellID
            GSMCellIndex = m_clsGSMCELL_SMC.GetValueAt(iRecordIndex, ATTR_GSMCELLINDEX).Value
            .Cells(i, m_clsMocGSMNCELL.Attr(ATTR_GSMCELLINDEX).ColIndex).Value = GSMCellIndex

            For j = 1 To Moc.Count
                Set Attr = Moc.Attr(j)
                If (Not m_clsMocGSMNCELL.Exists(Attr.Name)) And Attr.DefaultValue <> "" Then
                    .Cells(i, Attr.ColIndex).Value = Attr.DefaultValue
                End If
            Next j

            Call PrintToReport(FormatStr(RSC_STR_NAME_VALUE_PAIR_4, ATTR_BSCNAME, BSCName, ATTR_RNCID, RNCID, ATTR_CELLID, CellID, ATTR_GSMCELLINDEX, GSMCellIndex))
            iRecordIndex = iRecordIndex + 1
        Next i
    End With
End Sub

Private Sub InitModuleVars()
    Dim Moc As MocClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_GSMNCELL)
    m_clsMocGSMNCELL.Name = Moc.Name
    If m_clsMocGSMNCELL.Count <= 0 Then
        With m_clsMocGSMNCELL
            .Add Attr:=Moc.Attr(ATTR_BSCNAME)
            .Add Attr:=Moc.Attr(ATTR_RNCID)
            .Add Attr:=Moc.Attr(ATTR_CELLID)
            .Add Attr:=Moc.Attr(ATTR_GSMCELLINDEX)
        End With
    End If

    Set Moc = g_c_clsMocInfos.Moc(MOC_GSMCELL)
    m_clsMocGSMCELL.Name = Moc.Name
    If m_clsMocGSMCELL.Count <= 0 Then
        With m_clsMocGSMCELL
            .Add Attr:=Moc.Attr(ATTR_BSCNAME)
            .Add Attr:=Moc.Attr(ATTR_GSMCELLINDEX)
            .Add Attr:=Moc.Attr(ATTR_NCC)
            .Add Attr:=Moc.Attr(ATTR_BCC)
            .Add Attr:=Moc.Attr(ATTR_BCCHARFCN)
            .Add Attr:=Moc.Attr(ATTR_BANDIND)
        End With
    End If
    
    Call InitTable(m_clsMocGSMCELL, m_clsGSMCELL_SMC)
    Call InitTable(m_clsMocGSMCELL, m_clsGSMCELL_SC)
End Sub

Private Sub InitTable(Moc As MocClass, Tbl As TableClass)
    Dim i As Long, Attr As AttrClass
    If Tbl.GetFieldCount <= 0 Then
        For i = 1 To m_clsMocGSMCELL.Count
            Set Attr = Moc.Attr(i)
            Set Attr = Attr.Clone
            Attr.Value = ""
            Tbl.AddField Attr
        Next i
    End If
    
    Tbl.Clear
End Sub
