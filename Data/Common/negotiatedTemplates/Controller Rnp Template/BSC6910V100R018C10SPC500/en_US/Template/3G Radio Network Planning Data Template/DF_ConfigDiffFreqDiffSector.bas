Attribute VB_Name = "DF_ConfigDiffFreqDiffSector"
Option Explicit

Private Const SHT_TEMP_4 = "TempSheet4"
Private Const SHT_TEMP_5 = "TempSheet5"
Private Const SHT_TEMP_6 = "TempSheet6"
Private Const SHT_TEMP_7 = "TempSheet7"
Private m_colReportInfos As New Collection

Public Sub ConfigDiffFreqDiffSector(ot As dfOperationType, BSCName As String, RNCID As String, SelectedCell As MocClass, SectorMappingCellID As String)
    Call PrintToDebugger("Processing DiffFreqDiffSector......")

    '查找扇区映射小区的同频邻区关系
    Call PrintToDebugger("Searching sector mapping cell's " + MOC_INTRAFREQNCELL + "......")
    Call DoAdvancedFilter(Sheets(MOC_INTRAFREQNCELL), Sheets(SHT_TEMP_4), g_df_clsIntraFreqNCell, _
        Array(ATTR_BSCNAME + "=" + BSCName + "," + ATTR_CELLID + "=" + SectorMappingCellID))

    Dim CacheList As New Collection
    If GetLastRowIndex(Sheets(SHT_TEMP_4)) < ROW_DATA_HW Then
        Call PrintToReport(FormatStr(RSC_STR_NO_INTRA_FREQ_NCELL, SectorMappingCellID, BSCName))
        Select Case ot
            Case otADD
                Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_INTERFREQNCELL))
            Case otDEL
                Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_DEL_INTERFREQNCELL))
        End Select
    Else
        '在WholeNetworkCell中查找上述同频邻区关系对端小区
        Call GetPeerCells(Sheets(MOC_WHOLE_NETWORK_CELL), Sheets(SHT_TEMP_4), Sheets(SHT_TEMP_4), g_df_clsIntraFreqNCell, CacheList, g_df_clsWNCell)
        Call PrintCacheListToReport(Sheets(SHT_TEMP_4), CacheList, Sheets(SHT_TEMP_5))
        If GetLastRowIndex(Sheets(SHT_TEMP_4)) < ROW_DATA_HW Then
            Select Case ot
                Case otADD
                    Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_INTERFREQNCELL))
                Case otDEL
                    Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_DEL_INTERFREQNCELL))
            End Select
            Exit Sub
        End If

        '在WholeNetworkCell中查找上述对端小区的扇区映射小区，若不存在且操作类型为增加则认为是ValidInterFreqNCell
        '若存在且操作类型为删除，则在InterFreqNCell中查找选中小区与上述对端小区构成的异频邻区关系，若存在则拷贝至DeleteInterNCellRelation中
        Call GetValidInterFreqNCell(ot, Sheets(SHT_TEMP_4), Sheets(SHT_TEMP_5), RNCID, SelectedCell)
        If ot = otDEL Then '如果是删除操作，则此处退出即可
            Exit Sub
        End If

        If GetLastRowIndex(Sheets(SHT_TEMP_5)) < ROW_DATA_HW Then
            Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_INTERFREQNCELL))
            Exit Sub
        End If

        Call DeleteSameScramblingCode(Sheets(SHT_TEMP_5), Sheets(SHT_TEMP_4))

        Dim sht As Worksheet
        Set sht = Sheets(SHT_TEMP_4)
        '查找选中小区的异频邻区关系
        Call PrintToDebugger("Searching selected cell's " + MOC_INTERFREQNCELL + "......")
        Call DoAdvancedFilter(Sheets(MOC_INTERFREQNCELL), Sheets(SHT_TEMP_5), g_df_clsInterFreqNCell, _
            Array(ATTR_BSCNAME + "=" + BSCName + "," + ATTR_CELLID + "=" + SelectedCell.Attr(ATTR_CELLID).Value))
        If GetLastRowIndex(Sheets(SHT_TEMP_5)) >= ROW_DATA_HW Then
            '在WholeNetworkCell中查找上述异频邻区关系对端小区
            Call GetPeerCells(Sheets(MOC_WHOLE_NETWORK_CELL), Sheets(SHT_TEMP_5), Sheets(SHT_TEMP_5), g_df_clsInterFreqNCell, CacheList, g_df_clsWNCell)
            Call PrintCacheListToReport(Sheets(SHT_TEMP_5), CacheList, Sheets(SHT_TEMP_6))
            If GetLastRowIndex(Sheets(SHT_TEMP_5)) >= ROW_DATA_HW Then
                Call DeleteInvalidInterFreqNCell(Sheets(SHT_TEMP_4), Sheets(SHT_TEMP_5), Sheets(SHT_TEMP_6))
                Set sht = Sheets(SHT_TEMP_6)
            End If
        End If
        
        If GetLastRowIndex(sht) < ROW_DATA_HW Then
            Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_INTERFREQNCELL))
        Else
            Call AddData(sht, BSCName, RNCID, SelectedCell.Attr(ATTR_CELLID).Value)
        End If
    End If
End Sub

Private Sub DeleteSameScramblingCode(SrcSheet As Worksheet, DstSheet As Worksheet)
    Dim i As Long, iEnd As Long, iEnd2 As Long, iEnd3 As Long, s As String
    DstSheet.UsedRange.Clear
    iEnd2 = 0
    For i = 1 To ROW_DATA_HW
        iEnd2 = iEnd2 + 1
        SrcSheet.Rows(i).Copy Destination:=DstSheet.Cells(iEnd2, 1)
    Next i
    
    iEnd = GetLastRowIndex(SrcSheet)
    For i = ROW_DATA_HW + 1 To iEnd
        s = ATTR_PSCRAMBCODE + "=" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_PSCRAMBCODE).ColIndex).Value)
        Call DoAdvancedFilter(DstSheet, Sheets(SHT_TEMP_6), g_df_clsWNCell, Array(s), False)
        iEnd3 = GetLastRowIndex(Sheets(SHT_TEMP_6))
        If iEnd3 < ROW_DATA_HW Then
            iEnd2 = iEnd2 + 1
            SrcSheet.Rows(i).Copy Destination:=DstSheet.Cells(iEnd2, 1)
        End If
    Next i
End Sub

Private Sub DeleteInvalidInterFreqNCell(SrcSheet As Worksheet, ConditionSheet As Worksheet, DstSheet As Worksheet)
    Dim i As Long, iEnd As Long, iEnd2 As Long, iEnd3 As Long, s As String
    iEnd = GetLastRowIndex(SrcSheet)
    DstSheet.UsedRange.Clear
    iEnd3 = 0
    For i = 1 To ROW_DATA_HW - 1
        iEnd3 = iEnd3 + 1
        SrcSheet.Rows(i).Copy Destination:=DstSheet.Cells(iEnd3, 1)
    Next i

    For i = ROW_DATA_HW To iEnd
        s = ATTR_PSCRAMBCODE + "=" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_PSCRAMBCODE).ColIndex).Value) + "," + _
            ATTR_UARFCNUPLINK + "=" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_UARFCNUPLINK).ColIndex).Value) + "," + _
            ATTR_UARFCNDOWNLINK + "=" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_UARFCNDOWNLINK).ColIndex).Value)
        Call DoAdvancedFilter(ConditionSheet, Sheets(SHT_TEMP_7), g_df_clsWNCell, Array(s), False)
        iEnd2 = GetLastRowIndex(Sheets(SHT_TEMP_7))
        If iEnd2 < ROW_DATA_HW Then '若在选中小区已存在异频邻区中没有找到同频同扰，则认为是有效的
            iEnd3 = iEnd3 + 1
            SrcSheet.Rows(i).Copy Destination:=DstSheet.Cells(iEnd3, 1)
        End If
    Next i
End Sub

Private Sub GetValidInterFreqNCell(ot As dfOperationType, SrcSheet As Worksheet, DstSheet As Worksheet, RNCID As String, SelectedCell As MocClass)
    Dim SelectedNodeBName As String
    SelectedNodeBName = SelectedCell.Attr(ATTR_NODEBNAME).Value
    
    Dim i As Long, iEnd As Long, iEnd2 As Long, iEnd3 As Long, s As String, s1 As String, s2 As String, RNCs As New Collection
    iEnd = GetLastRowIndex(SrcSheet)
    iEnd3 = 0
    DstSheet.UsedRange.Clear
    For i = 1 To ROW_DATA_HW - 1
        iEnd3 = iEnd3 + 1
        SrcSheet.Rows(i).Copy Destination:=DstSheet.Cells(iEnd3, 1)
    Next i

    Call PrintToDebugger("Begin to validate peer cell(s)......")
    For i = ROW_DATA_HW To iEnd
        s = ATTR_BSCNAME + "=" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_BSCNAME).ColIndex).Value) + _
            "," + ATTR_NODEBNAME + "=" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_NODEBNAME).ColIndex).Value) + _
            "," + ATTR_CELLID + "=<>" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_CELLID).ColIndex).Value) + _
            "," + ATTR_SECTORID + "=" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_SECTORID).ColIndex).Value)
        s1 = s + "," + ATTR_UARFCNUPLINK + "=<>" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_UARFCNUPLINK).ColIndex).Value)
        s2 = s + "," + ATTR_UARFCNDOWNLINK + "=<>" + CStr(SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_UARFCNDOWNLINK).ColIndex).Value)

        Call DoAdvancedFilter(Sheets(MOC_WHOLE_NETWORK_CELL), Sheets(SHT_TEMP_6), g_df_clsWNCell, Array(s1, s2), True)
        iEnd2 = GetLastRowIndex(Sheets(SHT_TEMP_6))
        If iEnd2 >= ROW_DATA_HW Then
            If ot = otDEL Then
                s1 = SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_BSCNAME).ColIndex).Value
                If IsInCollection(s1, g_c_colBSCNAME_RNCID) Then
                    s1 = g_c_colBSCNAME_RNCID.Item(s1)
                    s2 = SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_CELLID).ColIndex).Value
                    Call DeleteData(SelectedCell.Attr(ATTR_BSCNAME).Value, RNCID, SelectedCell.Attr(ATTR_CELLID).Value, s1, s2)
                Else
                    If Not IsInCollection(s1, RNCs) Then '避免重复提示
                        Call PrintToReport(FormatStr(RSC_STR_RNC_NOT_FOUND, s1, MOC_BSCINFO))
                        RNCs.Add Item:=s1, Key:=s1
                    End If
                End If
            End If
        Else
            s = SrcSheet.Cells(g_df_clsWNCell.Attr(ATTR_NODEBNAME).ColIndex).Value
            If ot = otADD And s <> SelectedNodeBName Then '同一个NodeB中小区不能建立异频非同覆盖小区
                s = SrcSheet.Cells(i, g_df_clsWNCell.Attr(ATTR_BSCNAME).ColIndex).Value
                If IsInCollection(s, g_c_colBSCNAME_RNCID) Then '提前校验BSCNAME与RNCID映射是否存在
                    iEnd3 = iEnd3 + 1
                    SrcSheet.Rows(i).Copy Destination:=DstSheet.Cells(iEnd3, 1)
                Else
                    If Not IsInCollection(s, RNCs) Then '避免重复提示
                        Call PrintToReport(FormatStr(RSC_STR_RNC_NOT_FOUND, s, MOC_BSCINFO))
                        RNCs.Add Item:=s, Key:=s
                    End If
                End If
            End If
        End If
    Next i
    Call PrintToDebugger("Finished to validate peer cell(s).")
    
    If ot = otDEL Then
        Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, m_colReportInfos.Count, MOC_DEL_INTERFREQNCELL))
        For i = 1 To m_colReportInfos.Count
            Call PrintToReport(m_colReportInfos.Item(i))
        Next i
        Call ClearCollection(m_colReportInfos)
    End If
End Sub

Private Sub DeleteData(BSCName As String, RNCID As String, CellID As String, NCellRNCID As String, NCellID As String)
    Dim sht As Worksheet, s As String, iEnd As Long, iEnd2 As Long
    Set sht = Sheets(MOC_DEL_INTERFREQNCELL)
    iEnd2 = GetLastRowIndex(sht)
    s = ATTR_BSCNAME + "=" + BSCName + "," + ATTR_RNCID + "=" + RNCID + "," + ATTR_CELLID + "=" + CellID + "," + _
        ATTR_NCELLRNCID + "=" + NCellRNCID + "," + ATTR_NCELLID + "=" + NCellID
    Call DoAdvancedFilter(Sheets(MOC_INTERFREQNCELL), Sheets(SHT_TEMP_6), g_df_clsInterFreqNCell, Array(s))
    iEnd = GetLastRowIndex(Sheets(SHT_TEMP_6))
    If iEnd >= ROW_DATA_HW Then '最多应该只能找到一行数据
        iEnd2 = iEnd2 + 1
        Call CopySheetRow(Sheets(SHT_TEMP_6), iEnd, sht, iEnd2, g_c_clsMocInfos.Moc(MOC_DEL_INTERFREQNCELL))
        m_colReportInfos.Add s
    End If
End Sub

Private Sub AddData(sht As Worksheet, BSCName As String, RNCID As String, CellID As String)
    Dim i As Long, j As Integer, iEnd As Long, iEnd2 As Long, iEnd3 As Long, s As String, NCellRNCID As String, NCellID As String, colReportInfos2 As New Collection
    iEnd2 = GetLastRowIndex(Sheets(MOC_INTERFREQNCELL))
    iEnd = GetLastRowIndex(sht)
    iEnd3 = GetLastRowIndex(Sheets(MOC_NRNCCELL))
    Dim Moc As MocClass, Attr As AttrClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_INTERFREQNCELL)

    For i = ROW_DATA_HW To iEnd
        s = sht.Cells(i, g_df_clsWNCell.Attr(ATTR_BSCNAME).ColIndex).Value
        NCellRNCID = g_c_colBSCNAME_RNCID.Item(s)
        NCellID = sht.Cells(i, g_df_clsWNCell.Attr(ATTR_CELLID).ColIndex).Value
        
        iEnd2 = iEnd2 + 1
        With Sheets(MOC_INTERFREQNCELL)
            .Cells(iEnd2, g_df_clsInterFreqNCell.Attr(ATTR_BSCNAME).ColIndex).Value = BSCName
            .Cells(iEnd2, g_df_clsInterFreqNCell.Attr(ATTR_RNCID).ColIndex).Value = RNCID
            .Cells(iEnd2, g_df_clsInterFreqNCell.Attr(ATTR_CELLID).ColIndex).Value = CellID
            .Cells(iEnd2, g_df_clsInterFreqNCell.Attr(ATTR_NCELLRNCID).ColIndex).Value = NCellRNCID
            .Cells(iEnd2, g_df_clsInterFreqNCell.Attr(ATTR_NCELLID).ColIndex).Value = NCellID

            For j = 1 To Moc.Count
                Set Attr = Moc.Attr(j)
                If (Not g_df_clsInterFreqNCell.Exists(Attr.Name)) And Attr.DefaultValue <> "" Then
                    .Cells(iEnd2, Attr.ColIndex).Value = Attr.DefaultValue
                End If
            Next j
        End With

        s = FormatStr(RSC_STR_NAME_VALUE_PAIR_5, ATTR_BSCNAME, BSCName, ATTR_RNCID, RNCID, ATTR_CELLID, CellID, ATTR_NCELLRNCID, NCellRNCID, ATTR_NCELLID, NCellID)
        m_colReportInfos.Add s
        
        Call AddNRNCCell(Sheets(MOC_NRNCCELL), iEnd3, sht, i, Sheets(SHT_TEMP_5), BSCName, RNCID, NCellRNCID, NCellID, colReportInfos2)
    Next i

    Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, m_colReportInfos.Count, MOC_INTERFREQNCELL))
    For i = 1 To m_colReportInfos.Count
        Call PrintToReport(m_colReportInfos.Item(i))
    Next i
    If colReportInfos2.Count > 0 Then
        Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, colReportInfos2.Count, MOC_NRNCCELL))
        For i = 1 To colReportInfos2.Count
            Call PrintToReport(colReportInfos2.Item(i))
        Next i
    End If

    Call ClearCollection(m_colReportInfos)
    Call ClearCollection(colReportInfos2)
End Sub

