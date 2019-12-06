Attribute VB_Name = "DF_ConfigSameFreq"
Option Explicit

Private Const SHT_TEMP_4 = "TempSheet4"
Private Const SHT_TEMP_5 = "TempSheet5"
Private Const SHT_TEMP_6 = "TempSheet6"

Public Sub ConfigSameFreq(BSCName As String, RNCID As String, SelectedCell As MocClass, SectorMappingCellID As String)
    Call PrintToDebugger("Processing SameFreq......")

    Dim CacheList As New Collection
    '查找扇区映射小区的同频邻区关系
    Call PrintToDebugger("Searching sector mapping cell's " + MOC_INTRAFREQNCELL + "......")
    Call DoAdvancedFilter(Sheets(MOC_INTRAFREQNCELL), Sheets(SHT_TEMP_4), g_df_clsIntraFreqNCell, _
        Array(ATTR_BSCNAME + "=" + BSCName + "," + ATTR_CELLID + "=" + SectorMappingCellID))

    If GetLastRowIndex(Sheets(SHT_TEMP_4)) < ROW_DATA_HW Then
        Call PrintToReport(FormatStr(RSC_STR_NO_INTRA_FREQ_NCELL, SectorMappingCellID, BSCName))
        Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_INTRAFREQNCELL))
    Else
        '在WholeNetworkCell中查找上述同频邻区关系对端小区
        Call GetPeerCells(Sheets(MOC_WHOLE_NETWORK_CELL), Sheets(SHT_TEMP_4), Sheets(SHT_TEMP_4), g_df_clsIntraFreqNCell, CacheList, g_df_clsWNCell)
        Call PrintCacheListToReport(Sheets(SHT_TEMP_4), CacheList, Sheets(SHT_TEMP_5))
        If GetLastRowIndex(Sheets(SHT_TEMP_4)) < ROW_DATA_HW Then
            Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_INTRAFREQNCELL))
            Exit Sub
        End If

        '在WholeNetworkCell中查找上述对端小区的扇区映射小区
        Call GetSectorMappingCells(Sheets(MOC_WHOLE_NETWORK_CELL), Sheets(SHT_TEMP_4), Sheets(SHT_TEMP_4), g_df_clsWNCell, g_df_clsWNCell)
        If GetLastRowIndex(Sheets(SHT_TEMP_4)) < ROW_DATA_HW Then
            Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_INTRAFREQNCELL))
            Exit Sub
        End If
        
        '在上述扇区映射小区的基础上查找与选中小区频点相同、扰码不同的小区
        Call PrintToDebugger("Searching valid cell that it and selected cell is same frequency and different scrambling code......")
        Dim s As String
        s = ATTR_UARFCNUPLINK + "=" + SelectedCell.Attr(ATTR_UARFCNUPLINK).Value + "," + _
            ATTR_UARFCNDOWNLINK + "=" + SelectedCell.Attr(ATTR_UARFCNDOWNLINK).Value + "," + _
            ATTR_PSCRAMBCODE + "=<>" + SelectedCell.Attr(ATTR_PSCRAMBCODE).Value
        Call DoAdvancedFilter(Sheets(SHT_TEMP_4), Sheets(SHT_TEMP_5), g_df_clsWNCell, Array(s))
        If GetLastRowIndex(Sheets(SHT_TEMP_5)) < ROW_DATA_HW Then
            Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_INTRAFREQNCELL))
            Exit Sub
        End If

        '查找选中小区的同频邻区关系
        Call PrintToDebugger("Searching selected cell's " + MOC_INTRAFREQNCELL + "......")
        Call DoAdvancedFilter(Sheets(MOC_INTRAFREQNCELL), Sheets(SHT_TEMP_4), g_df_clsIntraFreqNCell, _
            Array(ATTR_BSCNAME + "=" + BSCName + "," + ATTR_CELLID + "=" + SelectedCell.Attr(ATTR_CELLID).Value))
    
        '在WholeNetworkCell中查找选中小区的同频邻区关系对端小区
        If GetLastRowIndex(Sheets(SHT_TEMP_4)) < ROW_DATA_HW Then
            Call PrintToDebugger(FormatStr(RSC_STR_NO_INTRA_FREQ_NCELL, SelectedCell.Attr(ATTR_CELLID).Value, BSCName))
        Else
            Call GetPeerCells(Sheets(MOC_WHOLE_NETWORK_CELL), Sheets(SHT_TEMP_4), Sheets(SHT_TEMP_4), g_df_clsIntraFreqNCell, CacheList, g_df_clsWNCell)
            Call PrintCacheListToReport(Sheets(SHT_TEMP_4), CacheList, Sheets(SHT_TEMP_6))
        End If
    
        Call WriteData(BSCName, RNCID, SelectedCell.Attr(ATTR_CELLID).Value, Sheets(SHT_TEMP_5), GetLastRowIndex(Sheets(SHT_TEMP_4)) >= ROW_DATA_HW)
    End If
End Sub

Private Sub GetSectorMappingCells(SrcSheet As Worksheet, DstSheet As Worksheet, ConditionSheet As Worksheet, ConditionMoc As MocClass, Moc As MocClass)
    Call PrintToDebugger("Searching sector mapping cell in sheet '" + MOC_WHOLE_NETWORK_CELL + "'......")

    Dim i As Long, iRow1 As Long, iRow2 As Long
    iRow1 = ROW_DATA_HW
    iRow2 = GetLastRowIndex(ConditionSheet)
    
    If iRow2 >= iRow1 Then
        Dim Groups As New Collection, Conditions As Collection, s As String
        Set Conditions = New Collection
        Groups.Add Conditions

        With ConditionSheet
            For i = iRow1 To iRow2
                s = ATTR_BSCNAME + "=" + CStr(.Cells(i, ConditionMoc.Attr(ATTR_BSCNAME).ColIndex).Value)
                s = s + "," + ATTR_NODEBNAME + "=" + CStr(.Cells(i, ConditionMoc.Attr(ATTR_NODEBNAME).ColIndex).Value)
                s = s + "," + ATTR_SECTORID + "=" + CStr(.Cells(i, ConditionMoc.Attr(ATTR_SECTORID).ColIndex).Value)
                s = s + "," + ATTR_UARFCNUPLINK + "=<>" + CStr(.Cells(i, ConditionMoc.Attr(ATTR_UARFCNUPLINK).ColIndex).Value)
                Conditions.Add s
            Next i
    
            Set Conditions = New Collection
            Groups.Add Conditions
            For i = iRow1 To iRow2
                s = ATTR_BSCNAME + "=" + CStr(.Cells(i, ConditionMoc.Attr(ATTR_BSCNAME).ColIndex).Value)
                s = s + "," + ATTR_NODEBNAME + "=" + CStr(.Cells(i, ConditionMoc.Attr(ATTR_NODEBNAME).ColIndex).Value)
                s = s + "," + ATTR_SECTORID + "=" + CStr(.Cells(i, ConditionMoc.Attr(ATTR_SECTORID).ColIndex).Value)
                s = s + "," + ATTR_UARFCNDOWNLINK + "=<>" + CStr(.Cells(i, ConditionMoc.Attr(ATTR_UARFCNDOWNLINK).ColIndex).Value)
                Conditions.Add s
            Next i
        End With
    End If

    Call DoAdvancedFilter3(SrcSheet, DstSheet, Moc, Groups)
    
    Do While Groups.Count > 0
        Set Conditions = Groups.Item(1)
        Do While Conditions.Count > 0
            Conditions.Remove (1)
        Loop
        Groups.Remove (1)
        Set Conditions = Nothing
    Loop
    Set Groups = Nothing
End Sub

Private Sub WriteData(BSCName As String, RNCID As String, CellID As String, sht As Worksheet, HasIntraFreqNCell As Boolean)
    Dim iEnd1 As Long, iEnd2 As Long, iEnd3 As Long, iEnd4 As Long, isValid As Boolean
    iEnd1 = GetLastRowIndex(sht)
    iEnd2 = GetLastRowIndex(Sheets(MOC_INTRAFREQNCELL))
    iEnd4 = GetLastRowIndex(Sheets(MOC_NRNCCELL))

    Dim Moc As MocClass, Attr As AttrClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_INTRAFREQNCELL)
    Dim i As Long, j As Integer, Value As String, ReportInfos As New Collection, ReportInfos2 As New Collection, NCellRNCID As String, NCellID As String
    
    With Sheets(MOC_INTRAFREQNCELL)
        For i = ROW_DATA_HW To iEnd1
            isValid = True
            If HasIntraFreqNCell Then
                Value = sht.Cells(i, g_df_clsWNCell.Attr(ATTR_PSCRAMBCODE).ColIndex).Value
                Call DoAdvancedFilter(Sheets(SHT_TEMP_4), Sheets(SHT_TEMP_6), g_df_clsWNCell, Array(ATTR_PSCRAMBCODE + "=" + Value), False)
                iEnd3 = GetLastRowIndex(Sheets(SHT_TEMP_6))
                If iEnd3 >= ROW_DATA_HW Then '若当前小区与选中小区已存在邻区扰码相同，则不能建立邻区关系
                    isValid = False
                End If
            End If

            Value = sht.Cells(i, g_df_clsWNCell.Attr(ATTR_BSCNAME).ColIndex).Value
            NCellRNCID = C_MocModule.g_c_colBSCNAME_RNCID.Item(Value)
            NCellID = sht.Cells(i, g_df_clsWNCell.Attr(ATTR_CELLID).ColIndex).Value
            If isValid And (Not (NCellRNCID = RNCID And NCellID = CellID)) Then
                iEnd2 = iEnd2 + 1
                .Cells(iEnd2, g_df_clsIntraFreqNCell.Attr(ATTR_BSCNAME).ColIndex).Value = BSCName
                .Cells(iEnd2, g_df_clsIntraFreqNCell.Attr(ATTR_RNCID).ColIndex).Value = RNCID
                .Cells(iEnd2, g_df_clsIntraFreqNCell.Attr(ATTR_CELLID).ColIndex).Value = CellID
                .Cells(iEnd2, g_df_clsIntraFreqNCell.Attr(ATTR_NCELLRNCID).ColIndex).Value = NCellRNCID
                .Cells(iEnd2, g_df_clsIntraFreqNCell.Attr(ATTR_NCELLID).ColIndex).Value = NCellID

                For j = 1 To Moc.Count
                    Set Attr = Moc.Attr(j)
                    If (Not g_df_clsIntraFreqNCell.Exists(Attr.Name)) And Attr.DefaultValue <> "" Then
                        .Cells(iEnd2, Attr.ColIndex).Value = Attr.DefaultValue
                    End If
                Next j

                ReportInfos.Add FormatStr(RSC_STR_NAME_VALUE_PAIR_5, ATTR_BSCNAME, BSCName, ATTR_RNCID, RNCID, ATTR_CELLID, CellID, _
                    ATTR_NCELLRNCID, NCellRNCID, ATTR_NCELLID, NCellID)

                Call AddNRNCCell(Sheets(MOC_NRNCCELL), iEnd4, sht, i, Sheets(SHT_TEMP_6), BSCName, RNCID, NCellRNCID, NCellID, ReportInfos2)
            End If
        Next i
    End With

    Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, ReportInfos.Count, MOC_INTRAFREQNCELL))
    For i = 1 To ReportInfos.Count
        Call PrintToReport(ReportInfos.Item(i))
    Next i

    If ReportInfos2.Count > 0 Then
        Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, ReportInfos2.Count, MOC_NRNCCELL))
        For i = 1 To ReportInfos2.Count
            Call PrintToReport(ReportInfos2.Item(i))
        Next i
    End If

    Call ClearCollection(ReportInfos)
    Call ClearCollection(ReportInfos2)
End Sub
