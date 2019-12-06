Attribute VB_Name = "DF_ConfigDiffFreqSameSector"
Option Explicit

Private Const SHT_TEMP_4 = "TempSheet4"
Private Const SHT_TEMP_5 = "TempSheet5"

Public Sub ConfigDiffFreqSameSector(BSCName As String, RNCID As String, SelectedCellID As String, SectorMappingCell As MocClass)
    Call PrintToDebugger("Processing DiffFreqSameSector......")

    '查找选中小区的异频邻区关系
    Call PrintToDebugger("Searching selected cell's " + MOC_INTERFREQNCELL + "......")
    Call DoAdvancedFilter(Sheets(MOC_INTERFREQNCELL), Sheets(SHT_TEMP_4), g_df_clsInterFreqNCell, _
        Array(ATTR_BSCNAME + "=" + BSCName + "," + ATTR_CELLID + "=" + SelectedCellID))

    '若在选中小区已存在异频邻区关系中找到频点、扰码与扇区映射小区相同的对端小区，则这样的扇区映射小区不能成为选中小区的异频邻区关系
    If IsInvalidInterFreqNCell(SectorMappingCell) Then
        Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 0, MOC_INTERFREQNCELL))
        Exit Sub
    Else
        Call PrintToReport(FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, 1, MOC_INTERFREQNCELL))
        Call WriteData(BSCName, RNCID, SelectedCellID, SectorMappingCell.Attr(ATTR_CELLID).Value)
    End If
End Sub

Private Function IsInvalidInterFreqNCell(SectorMappingCell As MocClass) As Boolean
    IsInvalidInterFreqNCell = True

    Dim i As Long, iEnd As Long, iEnd2 As Long, Value As String
    iEnd = GetLastRowIndex(Sheets(SHT_TEMP_4))
    If iEnd >= ROW_DATA_HW Then
        Dim Cells As New Collection, Attr As AttrClass
        Set Attr = g_df_clsInterFreqNCell.Attr(ATTR_NCELLID)
        For i = ROW_DATA_HW To iEnd
            Value = Sheets(SHT_TEMP_4).Cells(i, Attr.ColIndex).Value
            Cells.Add Item:=Value, Key:=Value '收集当前选中小区所有异频邻区关系对端小区
        Next i

        '在WholeNetworkCell中查找上述异频邻区关系对端小区
        Call PrintToDebugger("Searching peer cell in " + MOC_WHOLE_NETWORK_CELL + "......")
        Call GetCell(Sheets(MOC_WHOLE_NETWORK_CELL), Sheets(SHT_TEMP_5), g_df_clsWNCell)
        iEnd2 = GetLastRowIndex(Sheets(SHT_TEMP_5))
        If IsSameFreqSameScramb(Sheets(SHT_TEMP_5), iEnd2, g_df_clsWNCell, SectorMappingCell, Cells) Then
            Exit Function
        End If

        If iEnd2 < iEnd Then '在NRNCCELL中继续查找上述异频邻区关系对端小区
            iEnd = GetLastRowIndex(Sheets(MOC_NRNCCELL))
            If iEnd < ROW_DATA_HW Then
                Call PrintToReport(FormatStr(RSC_STR_SHEET_HAS_NO_RECORD, MOC_NRNCCELL))
            Else
                Call PrintToDebugger("Searching peer cell in " + MOC_NRNCCELL + "......")
                Call GetCell(Sheets(MOC_NRNCCELL), Sheets(SHT_TEMP_5), g_df_clsNCell)
                iEnd2 = GetLastRowIndex(Sheets(SHT_TEMP_5))
                If IsSameFreqSameScramb(Sheets(SHT_TEMP_5), iEnd2, g_df_clsNCell, SectorMappingCell, Cells) Then
                    Exit Function
                End If
            End If
        End If
    End If

    If Cells.Count > 0 Then
        For i = 1 To Cells.Count
            Call PrintToReport(FormatStr(RSC_STR_PEER_CELL_NOT_FOUND, Cells.Item(i), MOC_WHOLE_NETWORK_CELL, MOC_NRNCCELL))
        Next i
    Else
        IsInvalidInterFreqNCell = False
    End If
End Function

Private Function IsSameFreqSameScramb(sht As Worksheet, RowIndex As Long, Moc As MocClass, SectorMappingCell As MocClass, Cells As Collection)
    IsSameFreqSameScramb = False
    Dim i As Long, s As String
    For i = ROW_DATA_HW To RowIndex
        s = sht.Cells(i, Moc.Attr(ATTR_CELLID).ColIndex).Value
        If IsInCollection(s, Cells) Then
            Cells.Remove (s)
        End If

        If Not IsSameFreqSameScramb Then
            If CStr(sht.Cells(i, Moc.Attr(ATTR_UARFCNUPLINK).ColIndex).Value) = SectorMappingCell.Attr(ATTR_UARFCNUPLINK).Value _
                And CStr(sht.Cells(i, Moc.Attr(ATTR_UARFCNDOWNLINK).ColIndex).Value) = SectorMappingCell.Attr(ATTR_UARFCNDOWNLINK).Value _
                And CStr(sht.Cells(i, Moc.Attr(ATTR_PSCRAMBCODE).ColIndex).Value) = SectorMappingCell.Attr(ATTR_PSCRAMBCODE).Value Then
                IsSameFreqSameScramb = True
            End If
        End If
    Next i
End Function

Private Sub GetCell(SrcSheet As Worksheet, DstSheet As Worksheet, Moc As MocClass)
    Dim i As Long, iRow2 As Long, Conditions() As String, iIndex As Long
    iRow2 = GetLastRowIndex(Sheets(SHT_TEMP_4))
    iIndex = 0
    
    Dim Attr1 As AttrClass, Attr2 As AttrClass, s As String, Value As String
    Set Attr1 = g_df_clsInterFreqNCell.Attr(ATTR_BSCNAME)
    Set Attr2 = g_df_clsInterFreqNCell.Attr(ATTR_NCELLID)

    With Sheets(SHT_TEMP_4)
        For i = ROW_DATA_HW To iRow2
            Value = .Cells(i, Attr1.ColIndex).Value
            s = Attr1.Name + "=" + Value
            Value = .Cells(i, Attr2.ColIndex).Value
            s = s + "," + ATTR_CELLID + "=" + Value
            ReDim Preserve Conditions(iIndex)
            Conditions(iIndex) = s
            iIndex = iIndex + 1
        Next i
    End With
    
    Call DoAdvancedFilter2(SrcSheet, DstSheet, Moc, Conditions)
End Sub

Private Sub WriteData(BSCName As String, RNCID As String, CellID As String, NCellID As String)
    Dim iEnd As Long, i As Integer, Moc As MocClass, Attr As AttrClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_INTERFREQNCELL)
    iEnd = GetLastRowIndex(Sheets(MOC_INTERFREQNCELL))
    iEnd = iEnd + 1

    With Sheets(MOC_INTERFREQNCELL)
        .Cells(iEnd, g_df_clsInterFreqNCell.Attr(ATTR_BSCNAME).ColIndex).Value = BSCName
        .Cells(iEnd, g_df_clsInterFreqNCell.Attr(ATTR_RNCID).ColIndex).Value = RNCID
        .Cells(iEnd, g_df_clsInterFreqNCell.Attr(ATTR_CELLID).ColIndex).Value = CellID
        .Cells(iEnd, g_df_clsInterFreqNCell.Attr(ATTR_NCELLRNCID).ColIndex).Value = RNCID
        .Cells(iEnd, g_df_clsInterFreqNCell.Attr(ATTR_NCELLID).ColIndex).Value = NCellID
        
        For i = 1 To Moc.Count
            Set Attr = Moc.Attr(i)
            If (Not g_df_clsInterFreqNCell.Exists(Attr.Name)) And Attr.DefaultValue <> "" Then
                .Cells(iEnd, Attr.ColIndex).Value = Attr.DefaultValue
            End If
        Next i
    End With

    Call PrintToReport(FormatStr(RSC_STR_NAME_VALUE_PAIR_5, ATTR_BSCNAME, BSCName, ATTR_RNCID, RNCID, ATTR_CELLID, CellID, _
        ATTR_NCELLRNCID, RNCID, ATTR_NCELLID, NCellID))
End Sub
