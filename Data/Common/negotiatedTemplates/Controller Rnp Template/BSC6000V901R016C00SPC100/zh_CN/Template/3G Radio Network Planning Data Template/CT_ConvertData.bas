Attribute VB_Name = "CT_ConvertData"
'*****************************************************************************************************
'模块名称：CvtTemplateModule
'模块作用：转换VDF RNP模板数据至CME RNP模板，我们暴露了全MOC全参数，因此也支持其它运营商的RNP模板转换
'*****************************************************************************************************
Option Explicit

'为方便读取ConvertTemplate页面数据定义如下常量
Public Const SHT_CONVERT_TEMPLATE = "ConvertTemplate"
Public Const CT_ROW_MOC = 10  'MOC开始行
Public Const CT_COL_MOC_HW = 2  'MOC所在列
Public Const CT_COL_ATTR_HW = 3 '属性所在列
Public Const CT_COL_MOC_VDF = 4  'MOC所在列
Public Const CT_COL_ATTR_VDF = 5 '属性所在列
Public Const CT_COL_ATTR_DFT_VALUE = 6 '默认值所在列

Public g_ct_colMocNames As New Collection '收集模板转换涉及的MOC名称

'以下常量是存放用户输入的单元格信息
Public Const RNG_FILE_VDF = "D2" '存放VDF模板文件名称
Public Const RNG_FILE_HW = "D3" '存放HW模板文件名称
Private Const RNG_RNCID = "C6" '存放本端RNCID

'以下是模块变量
Private m_wkCvtTemplate As Workbook
Private m_shtCT As Worksheet 'ConvertTempalte Sheet
Private m_wkDataSourceVDF As Workbook
Private m_wkDataSourceHW As Workbook
Private m_colReportInfos As New Collection '收集处理报告信息

'以下常量用拷贝VDF模板数据至HW模板
Private Const ROW_NAME_VDF = 1 'VDF模板属性名所在行
Private Const ROW_DATA_VDF = 2 'VDF模板数据开始行

Private Const MIN_VALUE_GSMCELLINDEX = 0
Private Const MAX_VALUE_GSMCELLINDEX = 65535

Type FreqInfo
    RNCID As String
    CellID As String
    UARFCNUPLINK As String
    UARFCNDOWNLINK As String
End Type

Private m_FreqInfos() As FreqInfo '存储频点信息

'Select file name
Public Function SelectFileName(TypeIndex As Integer) As Boolean
    SelectFileName = False
    
    Dim strRange As String
    Select Case TypeIndex
        Case 0
            strRange = RNG_FILE_VDF
        Case 1
            strRange = RNG_FILE_HW
    End Select
    
    Dim fd As FileDialog
    Dim sht As Worksheet
  
    Set sht = Sheets(SHT_CONVERT_TEMPLATE)
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Filters.Clear
        .Filters.Add "Template Files", "*.xls"
        .AllowMultiSelect = False
        .Show
        If (.SelectedItems.Count = 0) Then
            Exit Function
        Else
            sht.Range(strRange).Value = .SelectedItems.Item(1)
            SelectFileName = True
        End If
    End With
End Function

'Prepare Moc Infos for ConvertTemplate
Public Sub PrepareMocInfos_CT(wk As Workbook)
    If g_ct_colMocNames.Count > 0 Then
        Exit Sub
    End If

    Call PrepareMocInfos(wk)

    '为GSMNCELL增加4个虚拟属性
    g_c_clsMocInfos.AddVirtualAttr MocName:="GSMNCELL", RefMocName:="GSMCELL", RefAttrName:="MCC"
    g_c_clsMocInfos.AddVirtualAttr MocName:="GSMNCELL", RefMocName:="GSMCELL", RefAttrName:="MNC"
    g_c_clsMocInfos.AddVirtualAttr MocName:="GSMNCELL", RefMocName:="GSMCELL", RefAttrName:="LAC"
    g_c_clsMocInfos.AddVirtualAttr MocName:="GSMNCELL", RefMocName:="GSMCELL", RefAttrName:="CID"
    g_c_clsMocInfos.AddVirtualAttr MocName:="GSMNCELL", RefMocName:="CELL", RefAttrName:="CELLNAME"

    '收集模板转换涉及的MOC
    With g_ct_colMocNames
        .Add Item:=MOC_CELL, Key:=MOC_CELL
        .Add Item:=MOC_NRNCCELL, Key:=MOC_NRNCCELL
        .Add Item:=MOC_INTRAFREQNCELL, Key:=MOC_INTRAFREQNCELL
        .Add Item:=MOC_INTERFREQNCELL, Key:=MOC_INTERFREQNCELL
        .Add Item:=MOC_GSMCELL, Key:=MOC_GSMCELL
        .Add Item:=MOC_GSMNCELL, Key:=MOC_GSMNCELL
    End With
End Sub

'Convert Data
Public Sub ConvertData()
    Call InitModule

    If Not IsValidInputInfo Then
        Exit Sub
    End If

    Dim strFile As String, OpenedWorkbooks As New Collection, s1 As String, s2 As String
   
    strFile = m_shtCT.Range(RNG_FILE_HW).Value
    Set m_wkDataSourceHW = GetWorkbook(strFile)
    If m_wkDataSourceHW Is Nothing Then
        Workbooks.Open FileName:=strFile, ReadOnly:=True
        Set m_wkDataSourceHW = ActiveWorkbook
        OpenedWorkbooks.Add Item:=m_wkDataSourceHW.Name, Key:=m_wkDataSourceHW.Name
    End If

    s1 = GetVersion(m_wkCvtTemplate)
    s2 = GetVersion(m_wkDataSourceHW)
    If s1 = "" Or s2 = "" Then
        Exit Sub
    End If
    If s1 <> s2 Then
        MsgBox FormatStr(RSC_STR_MAKE_SURE_TEMPLATE_SAME_VERSION, m_wkCvtTemplate.FullName, m_wkDataSourceHW.FullName)
        Exit Sub
    End If

    strFile = m_shtCT.Range(RNG_FILE_VDF).Value
    Set m_wkDataSourceVDF = GetWorkbook(strFile)
    If m_wkDataSourceVDF Is Nothing Then
        Workbooks.Open FileName:=strFile, ReadOnly:=True
        Set m_wkDataSourceVDF = ActiveWorkbook
        OpenedWorkbooks.Add Item:=m_wkDataSourceVDF.Name, Key:=m_wkDataSourceVDF.Name
    End If

    Call PrepareMocInfos_CT(m_wkCvtTemplate)

    Call PrepareTemplateMapInfos
    
    If Not ClearSheet(m_wkCvtTemplate) Then
        Exit Sub
    End If
    m_shtCT.Activate

    Application.StatusBar = RSC_STR_CONVERTING
    Application.ScreenUpdating = False
    
    Application.DisplayAlerts = False

    Call PrepareBSCInfos(m_wkDataSourceHW) '必须在m_wkDataSourceHW赋值后调用
    Application.DisplayAlerts = True
    
    Dim i As Integer
    For i = 1 To g_ct_colMocNames.Count
        Call ConvertMoc(g_ct_colMocNames.Item(i))
    Next i
    
    'VDF同频、异频邻区关系放在一个Sheet页中，因此，拷贝后需要过滤掉冗余的同频、异频邻区关系
    Dim iIndex As Long
    iIndex = 0
    Call PrepareFreqInfos(m_wkCvtTemplate, MOC_CELL, ATTR_BSCNAME, ATTR_CELLID, ATTR_UARFCNUPLINK, ATTR_UARFCNDOWNLINK, iIndex)
    Call PrepareFreqInfos(m_wkDataSourceHW, MOC_CELL, ATTR_BSCNAME, ATTR_CELLID, ATTR_UARFCNUPLINK, ATTR_UARFCNDOWNLINK, iIndex)
    Call PrepareFreqInfos(m_wkCvtTemplate, MOC_NRNCCELL, ATTR_NRNCID, ATTR_CELLID, ATTR_UARFCNUPLINK, ATTR_UARFCNDOWNLINK, iIndex)
    Call PrepareFreqInfos(m_wkDataSourceHW, MOC_NRNCCELL, ATTR_NRNCID, ATTR_CELLID, ATTR_UARFCNUPLINK, ATTR_UARFCNDOWNLINK, iIndex)

    Call RmvRedundantDataForNCellRelation(MOC_INTRAFREQNCELL)
    Call RmvRedundantDataForNCellRelation(MOC_INTERFREQNCELL)

    Call RmvRedundantNRNCCell

    Application.StatusBar = RSC_STR_CLOSING_WK
    If IsInCollection(m_wkDataSourceVDF.Name, OpenedWorkbooks) Then m_wkDataSourceVDF.Close SaveChanges:=False
    If IsInCollection(m_wkDataSourceHW.Name, OpenedWorkbooks) Then m_wkDataSourceHW.Close SaveChanges:=False
    m_wkCvtTemplate.Activate
    m_shtCT.Activate

    Application.ScreenUpdating = True
    Application.StatusBar = RSC_STR_FINISHED
    
    Call ShowReport(m_wkCvtTemplate)
End Sub

Private Sub ShowReport(wk As Workbook)
    frmReport.caption = RSC_STR_CONVERT_DATA
    frmReport.lstReportInfos.Clear
    Dim i As Long, sht As Worksheet, iEnd As Long, iCount As Long
    For i = 1 To g_ct_colMocNames.Count
        Set sht = wk.Sheets(g_ct_colMocNames.Item(i))
        iEnd = GetLastRowIndex(sht)
        iCount = iEnd - ROW_DATA_HW + 1
        frmReport.lstReportInfos.AddItem FormatStr(RSC_STR_INSERTED_DATA_INTO_SHEET, iCount, g_ct_colMocNames(i))
    Next i
    
    frmReport.lstReportInfos.AddItem ""
    For i = 1 To m_colReportInfos.Count
        frmReport.lstReportInfos.AddItem m_colReportInfos.Item(i)
    Next i

    frmReport.Show
End Sub

Private Function IsValidInputInfo()
    IsValidInputInfo = False

    Dim strFileName As String
    strFileName = m_shtCT.Range(RNG_FILE_VDF).Value
    If Not FileExists(strFileName) Then
        m_wkCvtTemplate.Activate
        m_shtCT.Activate
        m_shtCT.Range(RNG_FILE_VDF).Select
        MsgBox RSC_STR_SRC_TEMPLATE_FILE_NOT_FOUND
        Exit Function
    End If
    
    strFileName = m_shtCT.Range(RNG_FILE_HW).Value
    If Not FileExists(strFileName) Then
        m_wkCvtTemplate.Activate
        m_shtCT.Activate
        m_shtCT.Range(RNG_FILE_HW).Select
        MsgBox RSC_STR_SRC_HW_CME_RNP_DATA_FILE_NOT_FOUND
        Exit Function
    End If
    
    IsValidInputInfo = True
End Function

Private Sub PrepareTemplateMapInfos()
    Dim iEnd As Integer
    iEnd = GetLastRowIndex(m_shtCT)
    Dim Moc As MocClass, Attr As AttrClass
    Dim i As Integer, j As Long, iMocBegin As Integer, iMocEnd As Integer
    Dim MocName As String, strValue As String
    
    For i = 1 To g_ct_colMocNames.Count
        MocName = g_ct_colMocNames.Item(i)
        iMocBegin = 0
        iMocEnd = 0
        For j = CT_ROW_MOC To iEnd
            strValue = m_shtCT.Cells(j, CT_COL_MOC_HW).Value
            If strValue = MocName Then '找到MOC开始行
                Set Moc = g_c_clsMocInfos.Moc(MocName)
                Moc.SheetNameVDF = m_shtCT.Cells(j, CT_COL_MOC_VDF).Value '读取MOC映射
                iMocBegin = j
            End If
            If iMocBegin <> 0 And IsEmptyRow(m_shtCT, j) Then '找到MOC结束行
                iMocEnd = j - 1
                Exit For
            End If
        Next j
        
        If iMocEnd = 0 Then
            iMocEnd = iEnd
        End If
        
        For j = iMocBegin To iMocEnd '读取属性映射
            strValue = m_shtCT.Cells(j, CT_COL_ATTR_HW).Value
            If strValue = "" Then
                MsgBox "Failed to call PrepareTemplateMapInfos; MocName=" + MocName + "; ROW=" + CStr(j) + "; COL=" + CStr(CT_COL_ATTR_HW)
            End If
            Set Attr = Moc.Attr(strValue)
            Attr.ColNameVDF = m_shtCT.Cells(j, CT_COL_ATTR_VDF).Value
            Attr.DefaultValue = m_shtCT.Cells(j, CT_COL_ATTR_DFT_VALUE).Value
        Next j
    Next i
End Sub

Private Sub ConvertMoc(MocName As String)
    Application.StatusBar = FormatStr(RSC_STR_CONVERTING_MOC, MocName)

    Dim Moc As MocClass, MocGSMCell As MocClass, shtHW As Worksheet, shtVDF As Worksheet, EmptyRows As New Collection
    Set Moc = g_c_clsMocInfos.Moc(MocName)
    Set MocGSMCell = g_c_clsMocInfos.Moc(MOC_GSMCELL)
    Set shtHW = m_wkCvtTemplate.Sheets(Moc.Name)
    Set shtVDF = m_wkDataSourceVDF.Sheets(Moc.SheetNameVDF)

    If Not PrepareCopyDataFromVDF(shtVDF, Moc, EmptyRows) Then
        Exit Sub
    End If
    
    Dim Attr As AttrClass, i As Integer, j As Long, iEndVDF As Long, iRowHW As Long, r As Range, s As String

    '依次对每个属性从VDF模板拷贝至HW模板
    iEndVDF = GetLastRowIndex(shtVDF)
    For i = 1 To Moc.Count
        Set Attr = Moc.Attr(i)
        iRowHW = ROW_DATA_HW
        For j = ROW_DATA_VDF To iEndVDF
            Application.StatusBar = FormatStr(RSC_STR_CONVERTING_ATTR, MocName, Attr.Name, CStr(j))

            If Not IsInCollection(str(j), EmptyRows) Then
                If Attr.ColIndexVDF <> 0 Then
                    If MocName = MOC_NRNCCELL Then
                        If Not IsInvalidNCell(Moc, shtVDF, j) Then '如果是无效NRNCCELL，则所有记录都不需要拷贝
                            If Attr.Name = ATTR_BSCNAME Then
                                s = GetBSCName(shtVDF, j, Attr.ColIndexVDF) '我们认为用户提供的都是RNCID
                                shtHW.Cells(iRowHW, Attr.ColIndex).Value = s
                                iRowHW = iRowHW + 1
                            Else
                                s = shtVDF.Cells(j, Attr.ColIndexVDF).Value
                                shtHW.Cells(iRowHW, Attr.ColIndex).Value = s
                                iRowHW = iRowHW + 1
                            End If
                        End If
                    Else
                        If Attr.Name = ATTR_BSCNAME Then
                            s = GetBSCName(shtVDF, j, Attr.ColIndexVDF) '我们认为用户提供的都是RNCID
                            shtHW.Cells(iRowHW, Attr.ColIndex).Value = s
                            iRowHW = iRowHW + 1
                        Else
                            s = shtVDF.Cells(j, Attr.ColIndexVDF).Value '正常分支
                            shtHW.Cells(iRowHW, Attr.ColIndex).Value = s
                            iRowHW = iRowHW + 1
                        End If
                    End If
                Else '处理GSMNCELL对象两个特殊字段GSMCELLINDEX和CELLID
                    If MocName = MOC_GSMNCELL Then
                        If Attr.Name = ATTR_GSMCELLINDEX Then
                            'VDF只有四元组，没有GSMCellIndex
                            s = getGSMCellIndex(Moc, MocGSMCell, shtVDF, j, m_wkCvtTemplate)
                            If s = "" Then '数据源1中找不到，在数据源2中继续找
                                s = getGSMCellIndex(Moc, MocGSMCell, shtVDF, j, m_wkDataSourceHW)
                            End If
                            shtHW.Cells(iRowHW, Attr.ColIndex).Value = s
                            iRowHW = iRowHW + 1
                        ElseIf Attr.Name = ATTR_CELLID Then
                            'VDF只有CELLNAME，没有CELLID
                            s = getCellName(Moc, shtVDF, j, m_wkCvtTemplate)
                            If s = "" Then '数据源1中找不到，在数据源2中继续找
                                s = getCellName(Moc, shtVDF, j, m_wkDataSourceHW)
                            End If
                            shtHW.Cells(iRowHW, Attr.ColIndex).Value = s
                            iRowHW = iRowHW + 1
                        End If
                    End If
                End If
            End If
        Next j
    Next i
    
    '依次对每个属性设置默认值，这一步放在拷贝VDF数据之后处理，原因是需要知道设置多少行默认值
    Dim iEndHW As Long
    iEndHW = GetLastRowIndex(shtHW)
    For i = 1 To Moc.Count
        Set Attr = Moc.Attr(i)
        If Attr.ColNameVDF = "" And Attr.DefaultValue <> "" And iEndHW > ROW_DATA_HW Then
            Set r = shtHW.Range(shtHW.Cells(ROW_DATA_HW, Attr.ColIndex), shtHW.Cells(iEndHW, Attr.ColIndex))
            r.Value = Attr.DefaultValue
        End If
    Next i

    If Moc.Name = MOC_GSMCELL And MocGSMCell.Attr(ATTR_GSMCELLINDEX).ColNameVDF = "" Then
        Call GenerateGSMCellIndex '用户只提供了四元组，因此我们自己需要计算GSMCellIndex
    End If
End Sub

Private Function getCellName(MocGSMNCell As MocClass, SheetVDF As Worksheet, RowVDF As Long, BookHW As Workbook) As String
    getCellName = ""
    
    Dim MocCELL As MocClass
    Set MocCELL = g_c_clsMocInfos.Moc(MOC_CELL)

    Dim i As Long, iEnd As Long, shtCELL As Worksheet
    Set shtCELL = BookHW.Sheets(MOC_CELL)
    iEnd = GetLastRowIndex(shtCELL)

    Dim CELLNAME1 As String, CELLNAME2 As String
    For i = ROW_DATA_HW To iEnd
        CELLNAME1 = IIf(MocGSMNCell.Attr(ATTR_CELLNAME).ColIndexVDF <> 0, SheetVDF.Cells(RowVDF, MocGSMNCell.Attr(ATTR_CELLNAME).ColIndexVDF).Value, "")
        CELLNAME2 = shtCELL.Cells(i, MocCELL.Attr(ATTR_CELLNAME).ColIndex).Value
        If CELLNAME1 = CELLNAME2 Then
            getCellName = shtCELL.Cells(i, MocCELL.Attr(ATTR_CELLID).ColIndex).Value
            Exit For
        End If
    Next i
End Function

Private Function getGSMCellIndex(MocGSMNCell As MocClass, MocGSMCell As MocClass, SheetVDF As Worksheet, RowVDF As Long, BookHW As Workbook) As String
    getGSMCellIndex = ""

    Dim iEnd As Long, shtGSMCell As Worksheet
    Set shtGSMCell = BookHW.Sheets(MOC_GSMCELL)
    iEnd = GetLastRowIndex(shtGSMCell)

    Dim i As Long
    Dim MCC1 As String, MNC1 As String, LAC1 As String, CID1 As String
    Dim MCC2 As String, MNC2 As String, LAC2 As String, CID2 As String
    For i = ROW_DATA_HW To iEnd
        If MocGSMNCell.Attr(ATTR_MCC).ColIndexVDF <> 0 Then
            MCC1 = SheetVDF.Cells(RowVDF, MocGSMNCell.Attr(ATTR_MCC).ColIndexVDF).Value
        Else
            MCC1 = MocGSMNCell.Attr(ATTR_MCC).DefaultValue
        End If
        If MocGSMNCell.Attr(ATTR_MNC).ColIndexVDF <> 0 Then
            MNC1 = SheetVDF.Cells(RowVDF, MocGSMNCell.Attr(ATTR_MNC).ColIndexVDF).Value
        Else
            MNC1 = MocGSMNCell.Attr(ATTR_MNC).DefaultValue
        End If
        If MocGSMNCell.Attr(ATTR_LAC).ColIndexVDF <> 0 Then
            LAC1 = SheetVDF.Cells(RowVDF, MocGSMNCell.Attr(ATTR_LAC).ColIndexVDF).Value
        Else
            LAC1 = MocGSMNCell.Attr(ATTR_LAC).DefaultValue
        End If
        If MocGSMNCell.Attr(ATTR_CID).ColIndexVDF <> 0 Then
            CID1 = SheetVDF.Cells(RowVDF, MocGSMNCell.Attr(ATTR_CID).ColIndexVDF).Value
        Else
            CID1 = MocGSMNCell.Attr(ATTR_CID).DefaultValue
        End If
        MCC2 = shtGSMCell.Cells(i, MocGSMCell.Attr(ATTR_MCC).ColIndex).Value
        MNC2 = shtGSMCell.Cells(i, MocGSMCell.Attr(ATTR_MNC).ColIndex).Value
        LAC2 = shtGSMCell.Cells(i, MocGSMCell.Attr(ATTR_LAC).ColIndex).Value
        CID2 = shtGSMCell.Cells(i, MocGSMCell.Attr(ATTR_CID).ColIndex).Value
        If MCC1 = MCC2 And MNC1 = MNC2 And LAC1 = LAC2 And CID1 = CID2 Then
            getGSMCellIndex = shtGSMCell.Cells(i, MocGSMCell.Attr(ATTR_GSMCELLINDEX).ColIndex).Value
            Exit For
        End If
    Next i
End Function

Private Sub PrepareFreqInfos(wk As Workbook, MocName As String, AttrName1 As String, AttrName2 As String, AttrName3 As String, AttrName4 As String, FreqIndex As Long)
    Application.StatusBar = RSC_STR_PREPARING_FREQ
    Dim sht As Worksheet
    Set sht = wk.Sheets(MocName)
    
    Dim Moc As MocClass, Attr1 As AttrClass, Attr2 As AttrClass, Attr3 As AttrClass, Attr4 As AttrClass
    Set Moc = g_c_clsMocInfos.Moc(MocName)
    Set Attr1 = Moc.Attr(AttrName1)
    Set Attr2 = Moc.Attr(AttrName2)
    Set Attr3 = Moc.Attr(AttrName3)
    Set Attr4 = Moc.Attr(AttrName4)

    Dim i As Long, iEnd As Long
    iEnd = GetLastRowIndex(sht)
    For i = ROW_DATA_HW To iEnd
        Attr1.Value = sht.Cells(i, Attr1.ColIndex).Value
        If MocName = MOC_CELL Then
            Attr1.Value = GetRNCID(sht, i, Attr1.ColIndex)
        End If
        Attr2.Value = sht.Cells(i, Attr2.ColIndex).Value
        Attr3.Value = sht.Cells(i, Attr3.ColIndex).Value
        Attr4.Value = sht.Cells(i, Attr4.ColIndex).Value

        ReDim Preserve m_FreqInfos(FreqIndex)
        m_FreqInfos(FreqIndex).RNCID = Attr1.Value
        m_FreqInfos(FreqIndex).CellID = Attr2.Value
        m_FreqInfos(FreqIndex).UARFCNUPLINK = Attr3.Value
        m_FreqInfos(FreqIndex).UARFCNDOWNLINK = Attr4.Value
        Call PrintToDebugger(FormatStr(RSC_STR_FREQ_INFO, Attr1.Value, Attr2.Value, Attr3.Value, Attr4.Value))
        FreqIndex = FreqIndex + 1
    Next i
End Sub

Private Sub RmvRedundantDataForNCellRelation(MocName As String)
    Application.StatusBar = FormatStr(RSC_STR_DELETING_MOC, MocName)
    
    Dim sht As Worksheet
    Set sht = m_wkCvtTemplate.Sheets(MocName)
    
    Dim Moc As MocClass, Attr1 As AttrClass, Attr2 As AttrClass, Attr3 As AttrClass, Attr4 As AttrClass
    Set Moc = g_c_clsMocInfos.Moc(MocName)
    Set Attr1 = Moc.Attr(ATTR_RNCID)
    Set Attr2 = Moc.Attr(ATTR_CELLID)
    Set Attr3 = Moc.Attr(ATTR_NCELLRNCID)
    Set Attr4 = Moc.Attr(ATTR_NCELLID)

    Dim LocalCellFreq(1) As String, PeerCellFreq(1) As String
    Dim iEnd1 As Long, iEnd2 As Long, iRow As Long
    iEnd1 = GetLastRowIndex(sht)
    iEnd2 = UBound(m_FreqInfos)
    Dim i As Long, j As Long
    iRow = ROW_DATA_HW
    For i = ROW_DATA_HW To iEnd1
        Attr1.Value = sht.Cells(iRow, Attr1.ColIndex).Value
        Attr2.Value = sht.Cells(iRow, Attr2.ColIndex).Value
        Attr3.Value = sht.Cells(iRow, Attr3.ColIndex).Value
        Attr4.Value = sht.Cells(iRow, Attr4.ColIndex).Value
        LocalCellFreq(0) = ""
        LocalCellFreq(1) = ""
        PeerCellFreq(0) = ""
        PeerCellFreq(1) = ""
        For j = 0 To iEnd2
            If Attr1.Value = m_FreqInfos(j).RNCID And Attr2.Value = m_FreqInfos(j).CellID Then
                LocalCellFreq(0) = m_FreqInfos(j).UARFCNUPLINK
                LocalCellFreq(1) = m_FreqInfos(j).UARFCNDOWNLINK
            End If
            If Attr3.Value = m_FreqInfos(j).RNCID And Attr4.Value = m_FreqInfos(j).CellID Then
                PeerCellFreq(0) = m_FreqInfos(j).UARFCNUPLINK
                PeerCellFreq(1) = m_FreqInfos(j).UARFCNDOWNLINK
            End If

            If LocalCellFreq(0) <> "" And PeerCellFreq(0) <> "" Then Exit For
        Next j
        If LocalCellFreq(0) = "" And (Not IsInCollection(Attr2.Value + " " + Attr1.Value, m_colReportInfos)) Then
            m_colReportInfos.Add Item:=FormatStr(RSC_STR_FREQ_NOT_FOUND, Attr2.Value, Attr1.Value), Key:=Attr2.Value + " " + Attr1.Value
        End If
        If PeerCellFreq(0) = "" And (Not IsInCollection(Attr4.Value + " " + Attr3.Value, m_colReportInfos)) Then
            m_colReportInfos.Add Item:=FormatStr(RSC_STR_FREQ_NOT_FOUND, Attr4.Value, Attr3.Value), Key:=Attr4.Value + " " + Attr3.Value
        End If

        If (Not (LocalCellFreq(0) = PeerCellFreq(0) And LocalCellFreq(1) = PeerCellFreq(1))) And MocName = MOC_INTRAFREQNCELL Then
            sht.Rows(iRow).Delete
            iRow = iRow - 1
        End If
        
        If (LocalCellFreq(0) = PeerCellFreq(0) And LocalCellFreq(1) = PeerCellFreq(1)) And MocName = MOC_INTERFREQNCELL Then
            sht.Rows(iRow).Delete
            iRow = iRow - 1
        End If
        
        iRow = iRow + 1
    Next i
End Sub

Private Sub RmvRedundantNRNCCell()
    Dim sht As Worksheet
    Set sht = m_wkCvtTemplate.Sheets(MOC_NRNCCELL)
    Dim Moc As MocClass, Attr1 As AttrClass, Attr2 As AttrClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_NRNCCELL)
    Set Attr1 = Moc.Attr(ATTR_NRNCID)
    Set Attr2 = Moc.Attr(ATTR_CELLID)
    Dim iEnd As Long, i As Long, iRow As Long, NRNCCells As New Collection, s As String
    iEnd = GetLastRowIndex(sht)
    iRow = ROW_DATA_HW
    For i = ROW_DATA_HW To iEnd
        Attr1.Value = sht.Cells(iRow, Attr1.ColIndex).Value
        Attr2.Value = sht.Cells(iRow, Attr2.ColIndex).Value
        s = Attr1.Value + " " + Attr2.Value
        If IsInCollection(s, NRNCCells) Then
            sht.Rows(iRow).Delete
            iRow = iRow - 1
        Else
            NRNCCells.Add Item:=s, Key:=s
        End If

        iRow = iRow + 1
    Next i
    
    Call ClearCollection(NRNCCells)
End Sub

Private Function PrepareCopyDataFromVDF(SheetVDF As Worksheet, Moc As MocClass, EmptyRows As Collection) As Boolean
    PrepareCopyDataFromVDF = True

    Dim Attr As AttrClass, i As Long, j As Long, iEnd As Long, iCount As Integer, s As String
    If Moc.Name = MOC_GSMNCELL Then
        iCount = Moc.CountWithVirtualAttr
    Else
        iCount = Moc.Count
    End If
    
    '寻找VDF属性所在列索引
    For i = 1 To iCount
        Set Attr = Moc.Attr(i)
        Attr.ColIndexVDF = 0
        If Attr.ColNameVDF <> "" Then
            iEnd = SheetVDF.Columns.Count
            iEnd = SheetVDF.Columns(iEnd).Column
            For j = 1 To iEnd
                If SheetVDF.Cells(ROW_NAME_VDF, j) = Attr.ColNameVDF Then
                    Attr.ColIndexVDF = j
                End If
            Next j
            If Attr.ColIndexVDF = 0 Then '在VDF模板中没有找到相应属性列
                m_wkDataSourceVDF.Activate
                SheetVDF.Select
                SheetVDF.Cells(1, 1).Select
                s = FormatStr(RSC_STR_MSG_VDF_COL_NOT_FOUND, Attr.ColNameVDF, Moc.SheetNameVDF, m_wkDataSourceVDF.Name)
                If MsgBox(s, vbYesNo, RSC_STR_TITLE_CONFIRM) = VbMsgBoxResult.vbNo Then
                    m_wkCvtTemplate.Activate
                    m_shtCT.Activate
                    End
                End If
                PrepareCopyDataFromVDF = False
            End If
        End If
    Next i

    '寻找VDF数据中空白行
    Dim IsEmptyRow As Boolean
    iEnd = GetLastRowIndex(SheetVDF)
    For i = ROW_DATA_VDF To iEnd
        IsEmptyRow = True
        For j = 1 To iCount
            Set Attr = Moc.Attr(j)
            If Attr.ColNameVDF <> "" And Attr.ColIndexVDF <> 0 Then
                If SheetVDF.Cells(i, Attr.ColIndexVDF).Value <> "" Then
                    IsEmptyRow = False
                    Exit For
                End If
            End If
        Next j
        If IsEmptyRow Then
            EmptyRows.Add Item:=str(i), Key:=str(i)
        End If
    Next
End Function

Private Function IsInvalidNCell(Moc As MocClass, SheetVDF As Worksheet, RowVDF As Long)
    IsInvalidNCell = False

    Dim Attr1 As AttrClass, Attr2 As AttrClass
    Set Attr1 = Moc.Attr(ATTR_BSCNAME)
    Set Attr2 = Moc.Attr(ATTR_NRNCID)
    
    If Attr1.ColIndexVDF <> 0 And Attr2.ColIndexVDF <> 0 Then
        Attr1.Value = SheetVDF.Cells(RowVDF, Attr1.ColIndexVDF).Value
        Attr2.Value = SheetVDF.Cells(RowVDF, Attr2.ColIndexVDF).Value
        If Attr1.Value = Attr2.Value Then 'NRNCCELL中NRNCID不能与本RNCID相同
            IsInvalidNCell = True
        End If
    End If
End Function

Private Function ClearSheet(wk As Workbook)
    ClearSheet = False
    Dim isConfirmed As Boolean
    Dim s As String
    Dim sht As Worksheet
    Dim i As Integer, j As Long, iEnd As Long
    isConfirmed = False
    For i = 1 To g_ct_colMocNames.Count
        Set sht = wk.Sheets(g_ct_colMocNames.Item(i))
        iEnd = GetLastRowIndex(sht)
        If iEnd >= ROW_DATA_HW Then
            If isConfirmed Then
                For j = ROW_DATA_HW To iEnd
                    sht.Rows(ROW_DATA_HW).Delete
                Next j
            Else
                m_wkCvtTemplate.Activate
                sht.Activate
                s = FormatStr(RSC_STR_MSG_CLEAR_SHEET, m_wkCvtTemplate.Name)
                If MsgBox(s, vbYesNo, RSC_STR_TITLE_CONFIRM) = vbYes Then
                    isConfirmed = True
                    For j = ROW_DATA_HW To iEnd
                        sht.Rows(ROW_DATA_HW).Delete
                    Next j
                Else
                    Exit Function
                End If
            End If
        End If
    Next i
    ClearSheet = True
End Function

Private Sub GenerateGSMCellIndex()
    Application.StatusBar = RSC_STR_GENERATING_GSMCELLINDEX
    Dim Moc As MocClass, isValidIndex As Boolean
    Set Moc = g_c_clsMocInfos.Moc(MOC_GSMCELL)
    Dim sht1 As Worksheet, sht2 As Worksheet
    Set sht1 = m_wkCvtTemplate.Sheets(MOC_GSMCELL)
    Set sht2 = m_wkDataSourceHW.Sheets(MOC_GSMCELL)
    Dim i As Long, j As Long, k As Long, iEnd1 As Long, iEnd2 As Long
    iEnd1 = GetLastRowIndex(sht1)
    iEnd2 = GetLastRowIndex(sht2)

    Dim attrMCC As AttrClass, attrMNC As AttrClass, attrLAC As AttrClass, attrCID As AttrClass, attrGSMCELLINDEX As AttrClass
    Dim MCC As String, MNC As String, LAC As String, CID As String
    Set attrMCC = Moc.Attr(ATTR_MCC)
    Set attrMNC = Moc.Attr(ATTR_MNC)
    Set attrLAC = Moc.Attr(ATTR_LAC)
    Set attrCID = Moc.Attr(ATTR_CID)
    Set attrGSMCELLINDEX = Moc.Attr(ATTR_GSMCELLINDEX)
    For i = ROW_DATA_HW To iEnd1
        attrMCC.Value = sht1.Cells(i, attrMCC.ColIndex)
        attrMNC.Value = sht1.Cells(i, attrMNC.ColIndex)
        attrLAC.Value = sht1.Cells(i, attrLAC.ColIndex)
        attrCID.Value = sht1.Cells(i, attrCID.ColIndex)
        attrGSMCELLINDEX.Value = sht1.Cells(i, attrGSMCELLINDEX.ColIndex)
        If attrGSMCELLINDEX.Value = "" Then
            For j = ROW_DATA_HW To iEnd1
                MCC = sht2.Cells(j, attrMCC.ColIndex).Value
                MNC = sht2.Cells(j, attrMNC.ColIndex).Value
                LAC = sht2.Cells(j, attrLAC.ColIndex).Value
                CID = sht2.Cells(j, attrCID.ColIndex).Value
                If MCC = attrMCC.Value And MNC = attrMNC.Value And LAC = attrLAC.Value And CID = attrCID.Value Then
                    attrGSMCELLINDEX.Value = sht2.Cells(j, attrGSMCELLINDEX.ColIndex)
                    Exit For
                End If
            Next j
            If attrGSMCELLINDEX.Value = "" Then
                For j = MIN_VALUE_GSMCELLINDEX To MAX_VALUE_GSMCELLINDEX
                    isValidIndex = True
                    For k = ROW_DATA_HW To iEnd1
                        If sht1.Cells(k, attrGSMCELLINDEX.ColIndex).Value = j Then
                            isValidIndex = False
                            Exit For
                        End If
                    Next k
                    
                    If isValidIndex Then
                        For k = ROW_DATA_HW To iEnd2
                           If sht2.Cells(k, attrGSMCELLINDEX.ColIndex).Value = j Then
                                isValidIndex = False
                                Exit For
                            End If
                        Next k
                    End If
                    
                    If isValidIndex Then
                        attrGSMCELLINDEX.Value = CStr(j)
                        sht1.Cells(i, attrGSMCELLINDEX.ColIndex).Value = attrGSMCELLINDEX.Value
                        Exit For
                    End If
                Next j
            Else
                sht1.Cells(i, attrGSMCELLINDEX.ColIndex).Value = attrGSMCELLINDEX.Value
            End If
        End If
    Next i
End Sub

Private Sub InitModule()
    Set m_wkCvtTemplate = ActiveWorkbook
    Set m_shtCT = m_wkCvtTemplate.Sheets(SHT_CONVERT_TEMPLATE)
    Set g_c_wkCurrent = m_wkCvtTemplate
    g_c_wkCurrent.Sheets(SHT_DEBUG).UsedRange.Clear
    
    Call ClearCollection(m_colReportInfos)
End Sub
