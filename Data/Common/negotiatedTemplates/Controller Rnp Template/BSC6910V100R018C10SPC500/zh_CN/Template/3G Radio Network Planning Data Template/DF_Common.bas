Attribute VB_Name = "DF_Common"
Option Explicit

'邻区关系类型
Enum dfNCellRelationType
     rtSameFreq = 0
     rtDiffSystem = 1
     rtDiffFreqSameSector = 2
     rtDiffFreqDiffSector = 3
End Enum

'配置模式
Enum dfConfigureMode
     cmSingleSelected = 0
     cmMultiSelected = 1
End Enum

'操作类型
Enum dfOperationType
    otADD = 0
    otDEL = 1
End Enum

'定义全局常量
Public Const SHT_DOUBLE_FREQ_CELL_SETTING = "DoubleFrequencyCellSetting"
Public Const DF_ROW_MOC = 3  'MOC开始行
Public Const DF_COL_MOC = 2
Public Const DF_COL_ATTR = 3
Public Const DF_COL_ATTR_DFT_VALUE = 4
Public Const DF_ROW_PROPERTY = 3 'Property开始行
Public Const DF_COL_PROPERTY_NAME = 6
Public Const DF_COL_PROPERTY_VALUE = 7
Public Const MAX_SHT_ROW = 65536
Public Const FORMAT_TEXT = "@"
Public Const FORMAT_GENERAL = "G/通用格式"

'全局变量定义
Public g_df_emConfigureMode As dfConfigureMode '存储配置模式
Public g_df_clsWNCell As New MocClass 'Mapping defination of WholeNetworkCell Moc, only stores key attr, and used to filter data.
Public g_df_clsIntraFreqNCell As New MocClass
Public g_df_clsInterFreqNCell As New MocClass
Public g_df_clsNCell As New MocClass
Public g_df_colMocNames As New Collection

'模块常量定义
Private Const DEFAULT_FORMULA_SECTOR = "Mod(Mod(x,10),3)" '默认的扇区ID计算公式
Private Const SHT_BSC = "TempSheet2" '存放在WholeNetworkCell页面找到属于一个BSC的所有数据
Private Const SHT_NODEB = "TempSheet3" '存放在SHT_BSC结果集上找到属于一个NodeB的所有数据
Private Const SHT_CELL = "TempSheet4" '存放在SHT_NODEB结果集上找到与选中小区频点不同的所有数据

'模块变量定义
Private m_colReportInfos As New Collection '收集处理报告信息
Private m_clsDataCell As New TableClass '存放在m_shtDataCell结果集

Private m_isSelectedFreq As Boolean '是否选择了频点
Private m_strSelectedFreq As String 'UL Freq和DL Freq

Private m_colSelectedRows As New Collection '收集选中的行索引
Private m_colBSCs As New NameCollectionClass '在m_colSelectedRows基础上，收集有多少个BSC，每个BSC下有多少NodeB，每个NodeB下有多少个选中行索引

Private m_strBSCName As String '当前正在处理的BSC
Private m_strRNCID As String '当前正在处理的BSC对应的RNCID
Private m_strNodeBName As String '当前正在处理的NodeB
Private m_strCellID As String '当前正在处理的Cell
Private m_isReportEnable As Boolean

Private m_clsSelectedCell As New MocClass 'Mapping defination of DoubleFrequencyCell Moc, only stores key attrs, and used to store the record selected in DoubleFrequencyCell sheet.
Private m_clsSectorMappingCell As New MocClass '当前正在处理小区的扇区映射小区

Private m_colCellFormatAttrs As New Collection

Public Sub SetGUI_DF()
    Dim sht As Worksheet
    Set sht = Sheets(MOC_DOUBLE_FREQ_CELL)
    sht.Rows(1).Insert Shift:=xlDown
    sht.Rows(1).RowHeight = 63
    
    Call DF_Common.PrepareMocInfos_DF
    
    Dim Moc As MocClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_DOUBLE_FREQ_CELL)
    Dim iCol As Integer
    iCol = Moc.Attr(Moc.Count).ColIndex
    sht.Range(sht.Cells(1, 1), sht.Cells(1, iCol)).Merge

    Dim fTop As Single, fLeft As Single, fWidth As Single, fHeight As Single, fHGap As Single, fVGap As Single
    fWidth = sht.Shapes("cmdConfigIntraNCell").Width
    fHeight = sht.Shapes("cmdConfigIntraNCell").Height
    fTop = 10
    fLeft = 10
    fHGap = 1
    fVGap = 1

    'row 1
    With sht.Shapes("cmdConfigIntraNCell")
        .Top = fTop + fHeight * 0 + fVGap * 0
        .Left = fLeft + fWidth * 0 + fHGap * 0
    End With

    With sht.Shapes("cmdConfigGSMNCell")
        .Top = fTop + fHeight * 0 + fVGap * 0
        .Left = fLeft + fWidth * 1 + fHGap * 1
    End With

    With sht.Shapes("cmdConfigInterNCellSameSector")
        .Top = fTop + fHeight * 0 + fVGap * 0
        .Left = fLeft + fWidth * 2 + fHGap * 2
    End With

    With sht.Shapes("cmdConfigInterNCellDiffSector")
        .Top = fTop + fHeight * 0 + fVGap * 0
        .Left = fLeft + fWidth * 3 + fHGap * 3
    End With

    With sht.Shapes("cmdDeleteInterNCellDiffSector")
        .Top = fTop + fHeight * 0 + fVGap * 0
        .Left = fLeft + fWidth * 4 + fHGap * 4
        .Visible = msoFalse
    End With

    'row 2
    With sht.Shapes("cmdCopyDataToCELL")
        .Top = fTop + fHeight * 1 + fVGap * 1
        .Left = fLeft + fWidth * 0 + fVGap * 0
    End With

    With sht.Shapes("cmdCopyDataFromCELL")
        .Top = fTop + fHeight * 1 + fVGap * 1
        .Left = fLeft + fWidth * 1 + fHGap * 1
    End With

    With sht.Shapes("cmdSetFormula")
        .Top = fTop + fHeight * 1 + fVGap * 1
        .Left = fLeft + fWidth * 2 + fHGap * 2
    End With
End Sub

Public Function GetFormula() As String
    GetFormula = DEFAULT_FORMULA_SECTOR

    Dim r As Range, s As String
    Set r = Sheets(MOC_DOUBLE_FREQ_CELL).Range("C3")
    If Not r.Comment Is Nothing Then
        s = r.Comment.Text
    Else
        s = ""
    End If
    Dim i As Integer
    i = InStr(s, RSC_STR_FORMULA)
    If i > 0 Then
        s = Mid(s, i + Len(RSC_STR_FORMULA + Chr(10)))
    Else
        s = ""
    End If
    
    If "" <> s Then
        GetFormula = s
    End If
End Function

'获取扇区ID
Public Function GetSectorID(CellID As String, Optional Formula As String = "") As Integer
    If Not IsNumeric(CellID) Then
        MsgBox RSC_STR_CELLID_NOT_NUMERIC
    End If
    
    Dim strFormula As String
    If Formula = "" Then
        strFormula = GetFormula
    Else
        strFormula = Formula
    End If

    Dim r As Range, vFormula As Variant
    Set r = Sheets(SHT_CONDITION).Range("X1")
    vFormula = r.Formula
    On Error GoTo FormulaError
    r.Formula = "=" + Replace(strFormula, "x", CellID)
    GetSectorID = r.Value
    r.Formula = vFormula

    If GetSectorID < 0 Or GetSectorID > 5 Then
        MsgBox RSC_STR_SECTORID_RANGE_ERR
    End If
    
    r.Clear
    Exit Function

FormulaError:
    MsgBox RSC_STR_FORMULA_ERR
    End
End Function

'响应配置双载波邻区关系按钮点击事件
Public Sub ConfigNCellRelation(rt As dfNCellRelationType, Title As String, ot As dfOperationType)
    Application.ScreenUpdating = False
    Set g_c_wkCurrent = Application.ActiveWorkbook

    '设置配置过程中出现窗体的标题
    Select Case ot
        Case otADD
            frmSelectFreq.caption = "Add " + Title
            frmReport.caption = "Add " + Title
        Case otDEL
            frmSelectFreq.caption = "Delete " + Title
            frmReport.caption = "Delete " + Title
    End Select

    Call PrepareMocInfos_DF

    Call PrepareTemplateMapInfos

    Call PrepareBSCInfos(Application.ActiveWorkbook)

    Call InitModule
    
    If rt = rtSameFreq Or rt = rtDiffFreqDiffSector Then
        Call RefreshSectorID
    End If

    If rt = rtDiffFreqDiffSector And ot = otDEL Then
        If Not CheckDeleteInterNCellRelation Then
            Exit Sub
        End If
    End If

    '检查选中小区信息是否合法
    If Not CheckSelectedCellInfos Then '必须在m_colDFCellInfo赋值后调用
        Exit Sub
    End If

    If Not CheckWholeNetworkCell Then
        Exit Sub
    End If

    '准备BSC信息，因为我们要支持多网元
    Call PrepareBSCs

    Call ProcessBSCs(rt, ot)

    Sheets(MOC_DOUBLE_FREQ_CELL).Activate
    Application.StatusBar = RSC_STR_FINISHED
    Application.ScreenUpdating = True
    
    Call ShowReport
End Sub

Public Function GetSectorMappingCell(SelectedFreq As String) As Boolean
    GetSectorMappingCell = False

    m_isSelectedFreq = True
    m_strSelectedFreq = SelectedFreq

    Dim Freqs() As String
    Freqs = Split(SelectedFreq, " ")
    Dim Sector1 As String, Sector2 As String
    Sector1 = m_clsSelectedCell.Attr(ATTR_SECTORID).Value
    Dim Attr1 As AttrClass, Attr2 As AttrClass
    Dim i As Long, j As Long
    m_clsSectorMappingCell.Attr(ATTR_CELLID).Value = ""
    For i = 1 To m_clsDataCell.GetRecordCount
        Set Attr1 = m_clsDataCell.GetValueAt(i, ATTR_CELLID)
        Sector2 = GetSectorID(Attr1.Value)
        Set Attr1 = m_clsDataCell.GetValueAt(i, ATTR_UARFCNUPLINK)
        Set Attr2 = m_clsDataCell.GetValueAt(i, ATTR_UARFCNDOWNLINK)
        If Attr1.Value = Freqs(0) And Attr2.Value = Freqs(1) And Sector2 = Sector1 Then
            For j = 1 To m_clsDataCell.GetFieldCount
                Set Attr1 = m_clsDataCell.GetValueAt(i, j)
                m_clsSectorMappingCell.Attr(j).Value = Attr1.Value
            Next j
            Exit For
        End If
    Next i
    
    If m_clsSectorMappingCell.Attr(ATTR_CELLID).Value = "" Then
        Select Case g_df_emConfigureMode
            Case cmSingleSelected
                MsgBox RSC_STR_NO_SECTOR_MAPPING_CELL
            Case cmMultiSelected
                Call PrintToReport(FormatStr(RSC_STR_NO_SECTOR_MAPPING_CELL))
        End Select
        Exit Function
    End If
    
    GetSectorMappingCell = True
End Function

Public Sub PrintToReport(s As String)
    m_colReportInfos.Add s
    Call PrintToDebugger("REPORT: " + s)
End Sub

'Used for multi filter condition group, the relation of different group is OR, but every group has only one condition, for example: the condition of first group is "UARFCNUPLINK=<>9787", the one of second group is "UARFCNDOWNLINK=<>10737", ...
Public Sub DoAdvancedFilter(SrcSheet As Worksheet, DstSheet As Worksheet, Moc As MocClass, Conditions As Variant, Optional DebuggerEnable As Boolean = True)
    Dim col1 As Collection, col2 As New Collection
    Dim i As Integer
    For i = 0 To UBound(Conditions)
        Set col1 = New Collection
        col2.Add Item:=col1 ' multi filter condition group
        col1.Add Item:=Conditions(i) 'One group has only one condition
    Next i
    
    Call DoAdvancedFilter3(SrcSheet, DstSheet, Moc, col2, DebuggerEnable)
    
    Do While col2.Count > 0
        Set col1 = col2.Item(1)
        Do While col1.Count > 0
            col1.Remove (1)
        Loop
        col2.Remove (1)
        Set col1 = Nothing
    Loop
    Set col2 = Nothing
End Sub

'Used for only one filter condition group, but this group has multi condition, for example: first condition is "RNCID=429,CELLID=1", second is "RNCID=429,CELLID=2", ...
Public Sub DoAdvancedFilter2(SrcSheet As Worksheet, DstSheet As Worksheet, Moc As MocClass, Conditions() As String, Optional DebuggerEnable As Boolean = True)
    Dim col1 As New Collection, col2 As New Collection
    col2.Add Item:=col1 'only one filter condition group
    Dim i As Integer
    On Error GoTo E: 'Conditions可能越界，即没有数据元素
    For i = 0 To UBound(Conditions)
        col1.Add Item:=Conditions(i) 'One group has multi condition
    Next i

    Call DoAdvancedFilter3(SrcSheet, DstSheet, Moc, col2, DebuggerEnable)
    
    Do While col1.Count > 0
        col1.Remove (1)
    Loop
    col2.Remove (1)
    Set col1 = Nothing
    Set col2 = Nothing
    Exit Sub
E:
    DstSheet.UsedRange.Clear
    PrintToDebugger ("No conditions.")
End Sub

Public Sub DoAdvancedFilter3(SrcSheet As Worksheet, DstSheet As Worksheet, Moc As MocClass, Groups As Collection, Optional DebuggerEnable As Boolean = True)
    DstSheet.UsedRange.Clear
    If Groups.Count < 0 Then
        PrintToDebugger ("No conditions.")
        Exit Sub
    End If
    Sheets(SHT_CONDITION).UsedRange.Clear

    Dim i As Integer, j As Integer, k As Integer
    Dim iCol1 As Integer, iCol2 As Integer, iRow1 As Integer, iRow2 As Integer
    Dim SubConditions() As String, Condition As String, Conditions As Collection
    Dim Values() As String ' Pair of Name and Value
    iRow1 = 2
    iCol1 = 1
    For i = 1 To Groups.Count 'The relation of different group is OR
        Set Conditions = Groups.Item(i)
        iCol2 = iCol1
        For j = 1 To Conditions.Count
            iCol1 = iCol2
            Condition = Conditions.Item(j)
            SubConditions = Split(Condition, ",")
            For k = 0 To UBound(SubConditions) ' The relation of different SubCondition is AND
                Values = Split(SubConditions(k), "=")
                Sheets(SHT_CONDITION).Cells(1, iCol1).Value = Moc.Attr(Values(0)).caption 'Caption of Field Name
                Sheets(SHT_CONDITION).Cells(iRow1, iCol1).Value = Values(1) 'Field Value, and it maybe includes operator

                iCol1 = iCol1 + 1
            Next k
            iRow1 = iRow1 + 1
        Next j
    Next i

    iRow1 = ROW_DATA_HW - 1
    iRow2 = GetLastRowIndex(SrcSheet)
    If iRow2 > iRow1 Then
        Call GetColIndex(iCol1, iCol2, Moc)
        Call SetCellFormat(SrcSheet, Moc, FORMAT_GENERAL)
        On Error GoTo E
        SrcSheet.Range(SrcSheet.Cells(iRow1, iCol1), SrcSheet.Cells(iRow2, iCol2)).AdvancedFilter action:=xlFilterCopy, _
            CriteriaRange:=Sheets(SHT_CONDITION).UsedRange, CopyToRange:=DstSheet.Range(DstSheet.Cells(iRow1, iCol1), DstSheet.Cells(iRow2, iCol2)), Unique:=True
        Call SetCellFormat(SrcSheet, Moc, FORMAT_TEXT)
    Else
        Call PrintToDebugger("Source sheet has no data.")
    End If

    If DebuggerEnable Then
        Call PrintToDebugger("Filter condition:", Sheets(SHT_CONDITION))
        Call PrintToDebugger("Filtered data:", DstSheet)
    End If

    Exit Sub
E:
    Call SetCellFormat(SrcSheet, Moc, FORMAT_TEXT)
End Sub

Public Sub GetPeerCells(SrcSheet As Worksheet, DstSheet As Worksheet, ConditionSheet As Worksheet, ConditionMoc As MocClass, CacheList As Collection, Moc As MocClass)
    Do While CacheList.Count > 0
        CacheList.Remove (1)
    Loop
    
    Dim i As Long, iEnd As Long
    iEnd = GetLastRowIndex(ConditionSheet)
    With ConditionSheet
        Dim RNCs As New Collection
        Dim Conditions() As String, iIndex As Long
        iIndex = 0
        Dim Attr1 As AttrClass, Attr2 As AttrClass, s As String, Value As String
        Set Attr1 = ConditionMoc.Attr(ATTR_NCELLRNCID)
        Set Attr2 = ConditionMoc.Attr(ATTR_NCELLID)
        For i = ROW_DATA_HW To iEnd
            Value = .Cells(i, Attr1.ColIndex).Value
            If IsInCollection(Value, g_c_colRNCID_BSCNAME) Then
                Value = g_c_colRNCID_BSCNAME.Item(Value)
                s = ATTR_BSCNAME + "=" + Value
                Value = .Cells(i, Attr2.ColIndex).Value
                s = ATTR_CELLID + "=" + Value + "," + s
                ReDim Preserve Conditions(iIndex)
                Conditions(iIndex) = s
                iIndex = iIndex + 1
                CacheList.Add Item:=s, Key:=s
            Else
                If Not IsInCollection(Value, RNCs) Then
                    Call PrintToReport(FormatStr(RSC_STR_BSC_NOT_FOUND, Value, MOC_BSCINFO))
                    RNCs.Add Item:=Value, Key:=Value
                End If
            End If
        Next i
    End With

    Call PrintToDebugger("Searching peer cell in sheet '" + MOC_WHOLE_NETWORK_CELL + "'......")
    Call DoAdvancedFilter2(SrcSheet, DstSheet, Moc, Conditions)
End Sub

Public Sub CopySheetRow(SrcSheet As Worksheet, SrcRowIndex As Long, DstSheet As Worksheet, DstRowIndex As Long, Moc As MocClass)
    Dim i As Integer, Attr As AttrClass
    For i = 1 To Moc.Count
        Set Attr = Moc.Attr(i)
        DstSheet.Cells(DstRowIndex, Attr.ColIndex).Value = SrcSheet.Cells(SrcRowIndex, Attr.ColIndex).Value
    Next i
End Sub

Public Sub AddNRNCCell(DstSheet As Worksheet, DstRowIndex As Long, SrcSheet As Worksheet, SrcRowIndex As Long, TempSheet As Worksheet, _
    BSCName As String, RNCID As String, NCellRNCID As String, NCellID As String, ReportInfos As Collection)
    Dim Moc As MocClass, Attr As AttrClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_NRNCCELL)

    Dim iEnd As Long, i As Integer
    If NCellRNCID <> RNCID Then 'if local rnc and peer rnc is the same, can not add NRNCCell
        Call DoAdvancedFilter(Sheets(MOC_NRNCCELL), TempSheet, g_df_clsNCell, _
            Array(ATTR_BSCNAME + "=" + BSCName + "," + ATTR_NRNCID + "=" + NCellRNCID + "," + ATTR_CELLID + "=" + NCellID), False)
        iEnd = GetLastRowIndex(TempSheet)
        If iEnd < ROW_DATA_HW Then 'if no NRNCCell, should add one.
            DstRowIndex = DstRowIndex + 1
            With DstSheet
                '为主键字段赋值
                .Cells(DstRowIndex, g_df_clsNCell.Attr(ATTR_BSCNAME).ColIndex).Value = BSCName
                .Cells(DstRowIndex, g_df_clsNCell.Attr(ATTR_NRNCID).ColIndex).Value = NCellRNCID
                .Cells(DstRowIndex, g_df_clsNCell.Attr(ATTR_CELLID).ColIndex).Value = NCellID
                
                '为重要字段赋值，主键除外
                For i = 1 To g_df_clsNCell.Count
                    Set Attr = g_df_clsNCell.Attr(i)
                    If Attr.Name <> ATTR_BSCNAME And Attr.Name <> ATTR_NRNCID And Attr.Name <> ATTR_CELLID Then
                        .Cells(DstRowIndex, Attr.ColIndex).Value = SrcSheet.Cells(SrcRowIndex, g_df_clsWNCell.Attr(Attr.Name).ColIndex).Value
                    End If
                Next i

                '为次要字段赋默认值
                For i = 1 To Moc.Count
                    Set Attr = Moc.Attr(i)
                    If (Not g_df_clsNCell.Exists(Attr.Name)) And Attr.DefaultValue <> "" Then
                        .Cells(DstRowIndex, Attr.ColIndex).Value = Attr.DefaultValue
                    End If
                Next i
            End With
            ReportInfos.Add ATTR_BSCNAME + "=" + BSCName + ", " + ATTR_NRNCID + "=" + NCellRNCID + ", " + ATTR_CELLID + "=" + NCellID
        End If
    End If
End Sub

Public Sub PrepareMocInfos_DF()
    If g_df_colMocNames.Count > 0 Then
        Exit Sub
    End If

    Call PrepareMocInfos(Application.ActiveWorkbook)

    With g_df_colMocNames
        .Add Item:=MOC_NRNCCELL, Key:=MOC_NRNCCELL
        .Add Item:=MOC_INTRAFREQNCELL, Key:=MOC_INTRAFREQNCELL
        .Add Item:=MOC_INTERFREQNCELL, Key:=MOC_INTERFREQNCELL
        .Add Item:=MOC_GSMNCELL, Key:=MOC_GSMNCELL
    End With
End Sub

Public Sub PrintCacheListToReport(sht As Worksheet, CacheList As Collection, TempSheet As Worksheet)
    Dim i As Long, iEnd As Long, s As String
    For i = 1 To CacheList.Count
        Call DoAdvancedFilter(sht, TempSheet, g_df_clsWNCell, Array(CacheList.Item(i)), False)
        iEnd = GetLastRowIndex(TempSheet)
        If iEnd < ROW_DATA_HW Then
            s = Replace(CacheList.Item(i), ",", ", ")
            Call PrintToReport(FormatStr(RSC_STR_CELL_WAS_NOT_FOUND_2, s))
        End If
    Next i
End Sub

Private Sub SetCellFormat(sht As Worksheet, Moc As MocClass, FormatStr As String)
    Dim Attr As AttrClass
    Dim i As Integer, j As Long, iEnd As Long
    iEnd = GetLastRowIndex(sht)
    For i = 1 To Moc.Count
        Set Attr = Moc.Attr(i)
        If IsInCollection(Attr.Name, m_colCellFormatAttrs) And Attr.DataType = atInteger Then
            sht.Range(sht.Cells(ROW_DATA_HW, Attr.ColIndex), sht.Cells(MAX_SHT_ROW, Attr.ColIndex)).NumberFormatLocal = FormatStr
            For j = ROW_DATA_HW To iEnd
                sht.Cells(j, Attr.ColIndex).Value = sht.Cells(j, Attr.ColIndex).Value
            Next j
        End If
    Next i
End Sub

Private Sub PrepareTemplateMapInfos()
    Dim iEnd As Integer, sht As Worksheet
    Set sht = Sheets(SHT_DOUBLE_FREQ_CELL_SETTING)
    iEnd = GetLastRowIndex(sht)
    Dim Moc As MocClass, Attr As AttrClass
    Dim i As Integer, j As Long, iMocBegin As Integer, iMocEnd As Integer
    Dim MocName As String, strValue As String
    
    For i = 1 To g_df_colMocNames.Count
        MocName = g_df_colMocNames.Item(i)
        Set Moc = g_c_clsMocInfos.Moc(MocName)
        iMocBegin = 0
        iMocEnd = 0
        For j = DF_ROW_MOC To iEnd
            strValue = sht.Cells(j, DF_COL_MOC).Value
            If strValue = MocName Then '找到MOC开始行
                iMocBegin = j
            End If
            If iMocBegin <> 0 And IsEmptyRow(sht, j) Then '找到MOC结束行
                iMocEnd = j - 1
                Exit For
            End If
        Next j

        If iMocEnd = 0 Then
            iMocEnd = iEnd
        End If

        For j = iMocBegin To iMocEnd '读取属性映射
            strValue = sht.Cells(j, DF_COL_ATTR).Value
            If strValue = "" Then
                MsgBox "Failed to call PrepareTemplateMapInfos; MocName=" + MocName + "; ROW=" + CStr(j) + "; COL=" + CStr(DF_COL_ATTR)
            End If
            Set Attr = Moc.Attr(strValue)
            Attr.DefaultValue = sht.Cells(j, DF_COL_ATTR_DFT_VALUE).Value
        Next j
    Next i
End Sub

Private Sub ProcessBSCs(rt As dfNCellRelationType, ot As dfOperationType)
    '配置邻区关系
    Dim i As Integer
    Dim NodeBs As NameCollectionClass
    For i = 1 To m_colBSCs.Count
        Set NodeBs = m_colBSCs.Item(1)
        m_strBSCName = NodeBs.Name
        If i > 1 Then
            Call PrintToReport("")
        End If
        'Call PrintToReport(FormatStr(RSC_STR_RPT_SELECTED_BSC, m_strBSCName))

        If IsInCollection(m_strBSCName, g_c_colBSCNAME_RNCID) Then
            m_strRNCID = g_c_colBSCNAME_RNCID.Item(m_strBSCName)
            Application.StatusBar = FormatStr(RSC_STR_PROCESSING_BSC, m_strBSCName)
            Call PrintToDebugger(Application.StatusBar)
            If PrepareDataForBSC(m_strBSCName) Then
                Call ProcessBSC(NodeBs, rt, ot)
            Else
                Select Case g_df_emConfigureMode
                    Case cmSingleSelected
                        MsgBox FormatStr(RSC_STR_BSC_WAS_NOT_FOUND, m_strBSCName)
                    Case cmMultiSelected
                        Call PrintToReport(FormatStr(RSC_STR_BSC_WAS_NOT_FOUND, m_strBSCName))
                End Select
            End If
        Else
            Select Case g_df_emConfigureMode
                Case cmSingleSelected
                    MsgBox FormatStr(RSC_STR_RNC_NOT_FOUND, m_strBSCName, MOC_BSCINFO)
                Case cmMultiSelected
                    Call PrintToReport(FormatStr(RSC_STR_RNC_NOT_FOUND, m_strBSCName, MOC_BSCINFO))
            End Select
        End If
    Next i
End Sub

Private Sub ShowReport()
    Dim i As Long
    If m_colReportInfos.Count > 0 Then
        If Not (g_df_emConfigureMode = cmSingleSelected And m_isReportEnable = False) Then
            frmReport.lstReportInfos.Clear
            Dim Freqs() As String

            If m_strSelectedFreq <> "" Then
                Freqs = Split(m_strSelectedFreq, " ")
                frmReport.lstReportInfos.AddItem FormatStr(RSC_STR_SELECTED_FREQ, Freqs(0), Freqs(1))
                frmReport.lstReportInfos.AddItem ""
            End If

            For i = 1 To m_colReportInfos.Count
                frmReport.lstReportInfos.AddItem m_colReportInfos.Item(i)
            Next i
            frmReport.Show
        End If
    End If
End Sub

Private Function PrepareDataForBSC(BSCName As String) As Boolean
    PrepareDataForBSC = False

    Call PrintToDebugger("Searching selected BSC's " + MOC_WHOLE_NETWORK_CELL + "......")
    Call DoAdvancedFilter(Sheets(MOC_WHOLE_NETWORK_CELL), Sheets(SHT_BSC), g_df_clsWNCell, Array(ATTR_BSCNAME + "=" + BSCName))

    Dim iEnd As Long
    iEnd = GetLastRowIndex(Sheets(SHT_BSC))
    PrepareDataForBSC = IIf(iEnd > ROW_DATA_HW, True, False)
End Function

Private Sub GetColIndex(ByRef MinColIndex As Integer, ByRef MaxColIndex As Integer, Moc As MocClass)
    MinColIndex = 255
    MaxColIndex = 1
    Dim Attr As AttrClass
    Dim i As Integer
    For i = 1 To Moc.Count
        Set Attr = Moc.Attr(i)
        If Attr.ColIndex < MinColIndex Then
            MinColIndex = Attr.ColIndex
        End If
        If Attr.ColIndex > MaxColIndex Then
            MaxColIndex = Attr.ColIndex
        End If
    Next i
End Sub

Private Sub ProcessBSC(NodeBs As NameCollectionClass, rt As dfNCellRelationType, ot As dfOperationType)
    Dim i As Long, Cells As NameCollectionClass
    For i = 1 To NodeBs.Count
        Set Cells = NodeBs.Item(i)
        m_strNodeBName = Cells.Name
        If i > 1 Then
            Call PrintToReport("")
        End If
        'Call PrintToReport(FormatStr(RSC_STR_RPT_SELECTED_NODEB, m_strNodeBName))

        Application.StatusBar = FormatStr(RSC_STR_PROCESSING_NODEB, m_strBSCName, m_strNodeBName)
        Call PrintToDebugger(Application.StatusBar)
        If PrepareDataForNodeB(m_strNodeBName) Then
            Call ProcessNodeB(Cells, rt, ot)
        Else
            Select Case g_df_emConfigureMode
                Case cmSingleSelected
                    MsgBox FormatStr(RSC_STR_NODEB_WAS_NOT_FOUND, m_strNodeBName, m_strBSCName)
                Case cmMultiSelected
                    Call PrintToReport(FormatStr(RSC_STR_NODEB_WAS_NOT_FOUND, m_strNodeBName, m_strBSCName))
            End Select
        End If
    Next i
End Sub

Private Function PrepareDataForNodeB(NodeBName As String) As Boolean
    PrepareDataForNodeB = False
    
    Call PrintToDebugger("Searching selected NodeB's " + MOC_WHOLE_NETWORK_CELL + "......")
    Call DoAdvancedFilter(Sheets(SHT_BSC), Sheets(SHT_NODEB), g_df_clsWNCell, Array(ATTR_NODEBNAME + "=" + NodeBName))

    Dim iEnd As Long
    iEnd = GetLastRowIndex(Sheets(SHT_NODEB))
    PrepareDataForNodeB = IIf(iEnd >= ROW_DATA_HW, True, False)
End Function

Private Sub ProcessNodeB(Cells As NameCollectionClass, rt As dfNCellRelationType, ot As dfOperationType)
    Dim i As Long, j As Integer, iSelectedRow As Long, Attr As AttrClass
    For i = 1 To Cells.Count
        iSelectedRow = Cells.Item(i)
        For j = 1 To m_clsSelectedCell.Count
            Set Attr = m_clsSelectedCell.Attr(j)
            Attr.Value = Sheets(MOC_DOUBLE_FREQ_CELL).Cells(iSelectedRow, Attr.ColIndex)
        Next j
        m_strCellID = m_clsSelectedCell.Attr(ATTR_CELLID).Value
        Call PrintMocToDebugger("Selected Cell:", m_clsSelectedCell)
        Application.StatusBar = FormatStr(RSC_STR_PROCESSING_CELL, m_strBSCName, m_strNodeBName, m_strCellID)
        Call PrintToDebugger(Application.StatusBar)

        If i > 1 Then
            Call PrintToReport("")
        End If
        Call PrintToReport(FormatStr(RSC_STR_RPT_SELECTED_CELL, m_strCellID, m_strNodeBName, m_strBSCName, iSelectedRow))

        If PrepareDataForCell(m_clsSelectedCell) Then
            m_clsDataCell.Clear
            Call SaveDataForCell

            If PrepareSectorMappingCell Then
                If m_clsSectorMappingCell.Attr(ATTR_CELLID).Value <> "" Then
                    Call ProcessCell(rt, ot)
                End If
            End If
        Else
            Select Case g_df_emConfigureMode
                Case cmSingleSelected
                    MsgBox RSC_STR_NO_DIFF_FREQ
                Case cmMultiSelected
                    Call PrintToReport(FormatStr(RSC_STR_NO_DIFF_FREQ_CELL, m_strCellID))
            End Select
        End If
    Next i
End Sub

Private Function PrepareDataForCell(SelectedCell As MocClass) As Boolean
    PrepareDataForCell = False
    
    Dim Condition(1) As String, Attr As AttrClass
    Set Attr = SelectedCell.Attr(ATTR_UARFCNUPLINK)
    Condition(0) = Attr.Name + "=<>" + Attr.Value
    Set Attr = SelectedCell.Attr(ATTR_UARFCNDOWNLINK)
    Condition(1) = Attr.Name + "=<>" + Attr.Value

    Call PrintToDebugger("Searching selected cell's " + MOC_WHOLE_NETWORK_CELL + "......")
    Call DoAdvancedFilter(Sheets(SHT_NODEB), Sheets(SHT_CELL), g_df_clsWNCell, Condition())

    Dim iEnd As Long
    iEnd = GetLastRowIndex(Sheets(SHT_CELL))
    PrepareDataForCell = IIf(iEnd > ROW_DATA_HW, True, False)
End Function

Private Sub SaveDataForCell()
    Dim iRow2 As Integer, i As Long, j As Long, iRecordIndex As Long, Attr As AttrClass
    iRow2 = GetLastRowIndex(Sheets(SHT_CELL))
    For i = ROW_DATA_HW To iRow2
        m_clsDataCell.AddRecord
        iRecordIndex = m_clsDataCell.GetRecordCount
        For j = 1 To m_clsDataCell.GetFieldCount
            Set Attr = m_clsDataCell.GetValueAt(iRecordIndex, j)
            Attr.Value = Sheets(SHT_CELL).Cells(i, Attr.ColIndex).Value
        Next j
    Next i
End Sub

Private Function PrepareSectorMappingCell() As Boolean
    PrepareSectorMappingCell = False

    Dim Attr As AttrClass, s As String, i As Long, j As Long, isFound As Boolean
    frmSelectFreq.cmbFreq.Clear
    For i = 1 To m_clsDataCell.GetRecordCount
        Set Attr = m_clsDataCell.GetValueAt(i, ATTR_UARFCNUPLINK)
        s = Attr.Value
        Set Attr = m_clsDataCell.GetValueAt(i, ATTR_UARFCNDOWNLINK)
        s = s + " " + Attr.Value
        
        isFound = False
        For j = 0 To frmSelectFreq.cmbFreq.ListCount - 1
            If frmSelectFreq.cmbFreq.List(j, 0) = s Then
                isFound = True
                Exit For
            End If
        Next j
        If Not isFound Then
            frmSelectFreq.cmbFreq.AddItem (s)
        End If
    Next i

    If (m_isSelectedFreq = False) And (frmSelectFreq.cmbFreq.ListCount = 1) And UCase(GetPropertyValue("IsShowSelectOneFreqDialog")) = "FALSE" Then
        m_strSelectedFreq = frmSelectFreq.cmbFreq.List(0, 0)
        m_isSelectedFreq = True
    End If

    '当操作模式为多选时，若已经选过基准频点，则不再重复选择，使用相同的基准频点，但仍要做上述检查
    If m_isSelectedFreq Then
        PrepareSectorMappingCell = True
        Call GetSectorMappingCell(m_strSelectedFreq)
        Exit Function
    End If

    Application.ScreenUpdating = True
    frmSelectFreq.lblSelectedCell = FormatStr(RSC_STR_SELECTED_CELL, m_strBSCName, m_strNodeBName, m_strCellID)
    frmSelectFreq.Show
    Application.ScreenUpdating = False

    PrepareSectorMappingCell = True
End Function

Private Function GetPropertyValue(Name As String) As String
    GetPropertyValue = ""
    Dim i As Long, iEnd As Long, sht As Worksheet
    Set sht = Sheets(SHT_DOUBLE_FREQ_CELL_SETTING)
    iEnd = GetLastRowIndex(sht)
    For i = DF_ROW_PROPERTY To iEnd
        If CStr(sht.Cells(i, DF_COL_PROPERTY_NAME).Value) = Name Then
            GetPropertyValue = sht.Cells(i, DF_COL_PROPERTY_VALUE).Value
            Exit Function
        End If
    Next i
End Function

Private Sub ProcessCell(rt As dfNCellRelationType, ot As dfOperationType)
    m_isReportEnable = True

    Dim i As Long, Attr As AttrClass, strSectorMappingCellID As String
    strSectorMappingCellID = m_clsSectorMappingCell.Attr(ATTR_CELLID).Value
    Call PrintToReport(FormatStr(RSC_STR_SECTOR_MAPPING_CELL, strSectorMappingCellID, m_strBSCName))
    Call PrintMocToDebugger("Sector mapping cell:", m_clsSectorMappingCell)

    Select Case rt
        Case rtSameFreq
            Call ConfigSameFreq(m_strBSCName, m_strRNCID, m_clsSelectedCell, strSectorMappingCellID)
        Case rtDiffSystem
            Call ConfigDiffSystem(m_strBSCName, m_strRNCID, m_strCellID, strSectorMappingCellID)
        Case rtDiffFreqSameSector
            Call ConfigDiffFreqSameSector(m_strBSCName, m_strRNCID, m_strCellID, m_clsSectorMappingCell)
        Case rtDiffFreqDiffSector
            Call ConfigDiffFreqDiffSector(ot, m_strBSCName, m_strRNCID, m_clsSelectedCell, strSectorMappingCellID)
    End Select
End Sub

Private Function CheckSelectedCellInfos() As Boolean
    CheckSelectedCellInfos = False

    '清空选中行索引收集器
    Do While m_colSelectedRows.Count > 0
        m_colSelectedRows.Remove (1)
    Loop

    '计算选中多少行，并收集行索引
    Dim r As Range, i As Long
    Dim iSelectedCount As Long
    iSelectedCount = 0
    On Error GoTo E '若没有任何选中，则读取Selection.Areas会出现异常
    For Each r In Selection.Areas
        For i = 1 To r.Rows.Count
            If r.Rows(i).Columns.Count = 256 Then '整行选中才认为是选中
                iSelectedCount = iSelectedCount + 1
                m_colSelectedRows.Add Item:=r.Rows(i).Row
            End If
        Next i
    Next r

    '检查是否至少选中了一行数据，并设置配置模式
    If iSelectedCount > 0 Then
        If iSelectedCount > 1 Then
            g_df_emConfigureMode = cmMultiSelected
        Else
            g_df_emConfigureMode = cmSingleSelected
        End If
    Else
        GoTo E
    End If

    '检查是否选择了空白行或标题行，若是则重新操作
    Dim iEnd As Long, iRow As Long
    iEnd = GetLastRowIndex(Sheets(MOC_DOUBLE_FREQ_CELL))
    For i = 1 To m_colSelectedRows.Count
        iRow = m_colSelectedRows(i)
        If iRow > iEnd Or iRow < (ROW_DATA_HW + 1) Then
            MsgBox RSC_STR_CANNOT_SELECT_EMPTY_ROW
            Exit Function
        End If
    Next i

    '检查选中行数据主要字段是否字段值为空，若是则重新操作
    Dim Attr As AttrClass, j As Integer
    For i = 1 To m_colSelectedRows.Count
        iRow = m_colSelectedRows(i)
        For j = 1 To m_clsSelectedCell.Count
            Set Attr = m_clsSelectedCell.Attr(j)
            Attr.Value = Sheets(MOC_DOUBLE_FREQ_CELL).Cells(iRow, Attr.ColIndex).Value
            If Attr.Value = "" Then
                Sheets(MOC_DOUBLE_FREQ_CELL).Cells(iRow, Attr.ColIndex).Select
                MsgBox FormatStr(RSC_STR_ATTR_CANNOT_EMPTY, Attr.caption, Sheets(MOC_DOUBLE_FREQ_CELL).Name)
                Exit Function
            End If
        Next j
    Next i

    CheckSelectedCellInfos = True
    Exit Function
E:
    MsgBox RSC_STR_SELECT_AT_LEAST_1_ROW
End Function

Private Function CheckWholeNetworkCell() As Boolean
    CheckWholeNetworkCell = False
    
    Dim iEnd As Long
    iEnd = GetLastRowIndex(Sheets(MOC_WHOLE_NETWORK_CELL))
    If iEnd < ROW_DATA_HW Then
        MsgBox FormatStr(RSC_STR_SHEET_HAS_NO_RECORD, MOC_WHOLE_NETWORK_CELL)
        Exit Function
    End If

    CheckWholeNetworkCell = True
End Function

Private Function CheckDeleteInterNCellRelation()
    Dim iEnd As Long
    iEnd = GetLastRowIndex(Sheets(MOC_DEL_INTERFREQNCELL))
    If iEnd >= ROW_DATA_HW Then
        Application.ScreenUpdating = True
        Sheets(MOC_DEL_INTERFREQNCELL).Activate
        CheckDeleteInterNCellRelation = MsgBox(FormatStr(RSC_STR_SHEET_EXISTS_SOME_DATA, MOC_DEL_INTERFREQNCELL), vbYesNo, "Confirm") = VbMsgBoxResult.vbYes
        Sheets(MOC_DOUBLE_FREQ_CELL).Activate
        Application.ScreenUpdating = False
        Exit Function
    End If
    CheckDeleteInterNCellRelation = True
End Function

Private Sub PrepareBSCs()
    Dim Moc As MocClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_DOUBLE_FREQ_CELL)
    Dim AttrBSCName As AttrClass, AttrNodeBName As AttrClass
    Set AttrBSCName = Moc.Attr(ATTR_BSCNAME)
    Set AttrNodeBName = Moc.Attr(ATTR_NODEBNAME)
    
    Dim NodeBs As NameCollectionClass
    Do While m_colBSCs.Count > 0
        Set NodeBs = m_colBSCs.Item(1)
        Do While NodeBs.Count > 0
            NodeBs.Remove (1)
        Loop
        m_colBSCs.Remove (1)
    Loop
    
    Dim i As Long, iRow As Long, Rows As NameCollectionClass
    For i = 1 To m_colSelectedRows.Count
        iRow = m_colSelectedRows.Item(i)
        AttrBSCName.Value = Sheets(MOC_DOUBLE_FREQ_CELL).Cells(iRow, AttrBSCName.ColIndex).Value
        AttrNodeBName.Value = Sheets(MOC_DOUBLE_FREQ_CELL).Cells(iRow, AttrNodeBName.ColIndex).Value
        If IsInNameCollection(AttrBSCName.Value, m_colBSCs) Then
            Set NodeBs = m_colBSCs.Item(AttrBSCName.Value)
        Else
            Set NodeBs = New NameCollectionClass
            NodeBs.Name = AttrBSCName.Value
            m_colBSCs.Add Item:=NodeBs, Key:=NodeBs.Name
        End If
        
        If IsInNameCollection(AttrNodeBName.Value, NodeBs) Then
            Set Rows = NodeBs.Item(AttrNodeBName.Value)
        Else
            Set Rows = New NameCollectionClass
            Rows.Name = AttrNodeBName.Value
            NodeBs.Add Item:=Rows, Key:=Rows.Name
        End If
        Rows.Add Item:=iRow
    Next i
End Sub

Private Sub RefreshSectorID()
    Dim Moc As MocClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_WHOLE_NETWORK_CELL)
    Dim i As Long, iEnd As Long, CellID As String
    iEnd = GetLastRowIndex(Sheets(MOC_WHOLE_NETWORK_CELL))
    For i = ROW_DATA_HW To iEnd
        With Sheets(MOC_WHOLE_NETWORK_CELL)
            CellID = .Cells(i, Moc.Attr(ATTR_CELLID).ColIndex).Value
            .Cells(i, Moc.Attr(ATTR_SECTORID).ColIndex).Value = GetSectorID(CellID)
        End With
    Next i
End Sub

Private Sub InitModule()
    Dim Moc As MocClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_DOUBLE_FREQ_CELL)
    m_clsSelectedCell.Name = Moc.Name
    If m_clsSelectedCell.Count <= 0 Then
        With m_clsSelectedCell
            .Add Attr:=Moc.Attr(ATTR_BSCNAME)
            .Add Attr:=Moc.Attr(ATTR_NODEBNAME)
            .Add Attr:=Moc.Attr(ATTR_SECTORID)
            .Add Attr:=Moc.Attr(ATTR_CELLID)
            .Add Attr:=Moc.Attr(ATTR_CELLNAME)
            .Add Attr:=Moc.Attr(ATTR_UARFCNDOWNLINK)
            .Add Attr:=Moc.Attr(ATTR_UARFCNUPLINK)
            .Add Attr:=Moc.Attr(ATTR_PSCRAMBCODE)
        End With
    End If

    Set Moc = g_c_clsMocInfos.Moc(MOC_WHOLE_NETWORK_CELL)
    g_df_clsWNCell.Name = Moc.Name
    g_df_clsWNCell.Clear 'debug
    If g_df_clsWNCell.Count <= 0 Then
        With g_df_clsWNCell
            .Add Attr:=Moc.Attr(ATTR_BSCNAME)
            .Add Attr:=Moc.Attr(ATTR_NODEBNAME)
            .Add Attr:=Moc.Attr(ATTR_SECTORID)
            .Add Attr:=Moc.Attr(ATTR_CELLID)
            .Add Attr:=Moc.Attr(ATTR_CELLNAME)
            .Add Attr:=Moc.Attr(ATTR_BANDIND)
            .Add Attr:=Moc.Attr(ATTR_UARFCNUPLINK)
            .Add Attr:=Moc.Attr(ATTR_UARFCNDOWNLINK)
            .Add Attr:=Moc.Attr(ATTR_PSCRAMBCODE)
            .Add Attr:=Moc.Attr(ATTR_LAC)
            .Add Attr:=Moc.Attr(ATTR_RAC)
        End With
    End If
    
    Set Moc = g_c_clsMocInfos.Moc(MOC_NRNCCELL)
    g_df_clsNCell.Name = Moc.Name
    If g_df_clsNCell.Count <= 0 Then
        With g_df_clsNCell
            .Add Attr:=Moc.Attr(ATTR_BSCNAME)
            .Add Attr:=Moc.Attr(ATTR_NRNCID)
            .Add Attr:=Moc.Attr(ATTR_CELLID)
            .Add Attr:=Moc.Attr(ATTR_CELLNAME)
            .Add Attr:=Moc.Attr(ATTR_BANDIND)
            .Add Attr:=Moc.Attr(ATTR_UARFCNUPLINK)
            .Add Attr:=Moc.Attr(ATTR_UARFCNDOWNLINK)
            .Add Attr:=Moc.Attr(ATTR_PSCRAMBCODE)
            .Add Attr:=Moc.Attr(ATTR_LAC)
            .Add Attr:=Moc.Attr(ATTR_RAC)
        End With
    End If

    Set Moc = g_c_clsMocInfos.Moc(MOC_INTRAFREQNCELL)
    g_df_clsIntraFreqNCell.Name = Moc.Name
    If g_df_clsIntraFreqNCell.Count <= 0 Then
        With g_df_clsIntraFreqNCell
            .Add Attr:=Moc.Attr(ATTR_BSCNAME)
            .Add Attr:=Moc.Attr(ATTR_RNCID)
            .Add Attr:=Moc.Attr(ATTR_CELLID)
            .Add Attr:=Moc.Attr(ATTR_NCELLRNCID)
            .Add Attr:=Moc.Attr(ATTR_NCELLID)
        End With
    End If
    
    Set Moc = g_c_clsMocInfos.Moc(MOC_INTERFREQNCELL)
    g_df_clsInterFreqNCell.Name = Moc.Name
    If g_df_clsInterFreqNCell.Count <= 0 Then
        With g_df_clsInterFreqNCell
            .Add Attr:=Moc.Attr(ATTR_BSCNAME)
            .Add Attr:=Moc.Attr(ATTR_RNCID)
            .Add Attr:=Moc.Attr(ATTR_CELLID)
            .Add Attr:=Moc.Attr(ATTR_NCELLRNCID)
            .Add Attr:=Moc.Attr(ATTR_NCELLID)
        End With
    End If
    
    Dim i As Long, Attr As AttrClass
    If m_clsDataCell.GetFieldCount <= 0 Then
        For i = 1 To g_df_clsWNCell.Count
            Set Attr = g_df_clsWNCell.Attr(i)
            Set Attr = Attr.Clone
            Attr.Value = ""
            m_clsDataCell.AddField Attr
        Next i
    End If
    m_clsDataCell.Clear

    If m_clsSectorMappingCell.Count <= 0 Then
        For i = 1 To g_df_clsWNCell.Count
            Set Attr = g_df_clsWNCell.Attr(i)
            m_clsSectorMappingCell.Add Attr
        Next i
    End If
    m_clsSectorMappingCell.Attr(ATTR_CELLID).Value = ""

    m_isSelectedFreq = False
    m_strSelectedFreq = ""
    Do While m_colReportInfos.Count > 0
        m_colReportInfos.Remove (1)
    Loop
    m_strBSCName = ""
    m_strNodeBName = ""
    m_strCellID = ""
    m_isReportEnable = False

    Sheets(SHT_DEBUG).UsedRange.Clear
    Sheets(SHT_BSC).UsedRange.Clear
    Sheets(SHT_NODEB).UsedRange.Clear
    Sheets(SHT_CELL).UsedRange.Clear
    
    If m_colCellFormatAttrs.Count <= 0 Then
        With m_colCellFormatAttrs
            .Add Item:=ATTR_CELLID, Key:=ATTR_CELLID
            .Add Item:=ATTR_SECTORID, Key:=ATTR_SECTORID
            .Add Item:=ATTR_UARFCNDOWNLINK, Key:=ATTR_UARFCNDOWNLINK
            .Add Item:=ATTR_UARFCNUPLINK, Key:=ATTR_UARFCNUPLINK
            .Add Item:=ATTR_PSCRAMBCODE, Key:=ATTR_PSCRAMBCODE
            .Add Item:=ATTR_RNCID, Key:=ATTR_RNCID
            .Add Item:=ATTR_NRNCID, Key:=ATTR_NRNCID
            .Add Item:=ATTR_NCELLRNCID, Key:=ATTR_NCELLRNCID
            .Add Item:=ATTR_NCELLID, Key:=ATTR_NCELLID
            .Add Item:=ATTR_LAC, Key:=ATTR_LAC
            .Add Item:=ATTR_RAC, Key:=ATTR_RAC
            .Add Item:=ATTR_GSMCELLINDEX, Key:=ATTR_GSMCELLINDEX
            .Add Item:=ATTR_NCC, Key:=ATTR_NCC
            .Add Item:=ATTR_BCC, Key:=ATTR_BCC
            .Add Item:=ATTR_BCCHARFCN, Key:=ATTR_BCCHARFCN
        End With
    End If
End Sub
