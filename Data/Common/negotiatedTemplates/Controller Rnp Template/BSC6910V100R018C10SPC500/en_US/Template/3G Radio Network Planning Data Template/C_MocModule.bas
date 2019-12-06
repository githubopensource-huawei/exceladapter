Attribute VB_Name = "C_MocModule"
'***************************************************************************************
'模块名称：MocModule
'模块作用：收集RNP模板相关MOC及其属性信息，为双载波、VDF模板转换等功能模块开发提供支撑
'***************************************************************************************
Option Explicit
Option Base 1

Enum mocAttrDataType
    atInteger = 0 'INT
    atEnum = 1 'LIST
    atString = 2 'STRING
    atUnknown = 99 '未知
End Enum

Public Const SHT_DEBUG = "DebugSheet"
Public Const SHT_COVER = "Cover"
Public Const SHT_TABLE_DEF = "TableDef"
Public Const SHT_VALID_DEF = "ValidDef"
Public Const SHT_CONDITION = "TempSheet1"

'RNP模板涉及MOC名称及属性名称常量定义，必须保证这些名称与TableDef中一致
Public Const MOC_NODEB = "NODEB"
Public Const MOC_CELL = "CELL"
Public Const MOC_NRNCCELL = "NRNCCELL"
Public Const MOC_INTRAFREQNCELL = "INTRAFREQNCELL"
Public Const MOC_INTERFREQNCELL = "INTERFREQNCELL"
Public Const MOC_GSMCELL = "GSMCELL"
Public Const MOC_GSMNCELL = "GSMNCELL"
Public Const MOC_LTECELL = "LTECELL"
Public Const MOC_LTENCELL = "LTENCELL"
Public Const MOC_SMLCCELL = "SMLCCELL"
Public Const MOC_PHY_NB_RADIO = "PhyNBRadio"
Public Const MOC_BSCINFO = "BSCInfo"
Public Const MOC_DOUBLE_FREQ_CELL = "DoubleFrequencyCell"
Public Const MOC_WHOLE_NETWORK_CELL = "WholeNetworkCell"
Public Const MOC_DEL_INTERFREQNCELL = "DeleteInterNCellRelation"
Public Const ATTR_BSCNAME = "BSCName"
Public Const ATTR_GSMCELLINDEX = "GSMCELLINDEX"
Public Const ATTR_MCC = "MCC"
Public Const ATTR_MNC = "MNC"
Public Const ATTR_LAC = "LAC"
Public Const ATTR_CID = "CID"
Public Const ATTR_NCC = "NCC"
Public Const ATTR_BCC = "BCC"
Public Const ATTR_BCCHARFCN = "BCCHARFCN"
Public Const ATTR_BANDIND = "BANDIND"
Public Const ATTR_NODEBNAME = "NODEBNAME"
Public Const ATTR_SECTORID = "SECTORID"
Public Const ATTR_CELLNAME = "CELLNAME"
Public Const ATTR_CELLID = "CELLID"
Public Const ATTR_NRNCID = "NRNCID"
Public Const ATTR_RNCID = "RNCID"
Public Const ATTR_NCELLID = "NCELLID"
Public Const ATTR_NCELLRNCID = "NCELLRNCID"
Public Const ATTR_UARFCNDOWNLINK = "UARFCNDOWNLINK"
Public Const ATTR_UARFCNUPLINK = "UARFCNUPLINK"
Public Const ATTR_PSCRAMBCODE = "PSCRAMBCODE"
Public Const ATTR_RAC = "RAC"
Public Const ATTR_CELLHOSTTYPE = "CELLHOSTTYPE"
Public Const ATTR_UARFCNUPLINKIND = "UARFCNUPLINKIND"
Public Const ATTR_BIDIRECTION = "BIDIRECTION"

Public g_c_clsMocInfos As New MocCollectionClass '存储MOC及其属性信息
Public g_c_colBSCNAME_RNCID As New Collection '存储BSCNAME与RNCID映射, BSCNAME是Key
Public g_c_colRNCID_BSCNAME As New Collection '存储RNCID与BSCNAME映射, RNCID是Key
Public g_c_wkCurrent As Workbook

Public Const ROW_DATA_HW = 3 'HW模板（即CME RNP模板）数据开始行

'为读取TableDef页数据方便定义的一些常量
Private Const SHT_NAME_TABLE_DEF = "TableDef"
Private Const ROW_MOC = 15
Private Const COL_MOC_NAME = 2 'B
Private Const COL_ATTR_NAME = 3 'C
Private Const COL_ATTR_TYPE = 4 'D
Private Const COL_ATTR_COL = 5 'E
Private Const COL_ATTR_MIN = 6 'F
Private Const COL_ATTR_MAX = 7 'G
Private Const COL_ATTR_LIST = 8 'H
Private Const COL_ATTR_CAPTION = 13 'M

'为读取MOC页数据方便定义的一些常量
Private Const ROW_CAPTION = 2

Private Const RNG_VER_CME = "D15"
Private Const RNG_VER_NE = "D16"
Private Const RNG_VER_RNP = "D17"

Private m_strWorkbookName_BSCInfo As String
Private m_colNotFoundBSCs As New Collection
Private m_colNotFoundRNCs As New Collection

Public Sub PrepareMocInfos(wk As Workbook)
    If g_c_clsMocInfos.Count > 0 Then
        'Exit Sub '调试时候去掉
    End If

    g_c_clsMocInfos.Clear '为了调试方便

    Dim MocNames As New Collection
    With MocNames
        .Add Item:=MOC_NODEB, Key:=MOC_NODEB
        .Add Item:=MOC_CELL, Key:=MOC_CELL
        .Add Item:=MOC_NRNCCELL, Key:=MOC_NRNCCELL
        .Add Item:=MOC_INTRAFREQNCELL, Key:=MOC_INTRAFREQNCELL
        .Add Item:=MOC_INTERFREQNCELL, Key:=MOC_INTERFREQNCELL
        .Add Item:=MOC_GSMCELL, Key:=MOC_GSMCELL
        .Add Item:=MOC_GSMNCELL, Key:=MOC_GSMNCELL
        .Add Item:=MOC_SMLCCELL, Key:=MOC_SMLCCELL
        .Add Item:=MOC_LTECELL, Key:=MOC_LTECELL
        .Add Item:=MOC_LTENCELL, Key:=MOC_LTENCELL
        .Add Item:=MOC_PHY_NB_RADIO, Key:=MOC_PHY_NB_RADIO
        .Add Item:=MOC_BSCINFO, Key:=MOC_BSCINFO
        .Add Item:=MOC_DOUBLE_FREQ_CELL, Key:=MOC_DOUBLE_FREQ_CELL
        .Add Item:=MOC_WHOLE_NETWORK_CELL, Key:=MOC_WHOLE_NETWORK_CELL
        .Add Item:=MOC_DEL_INTERFREQNCELL, Key:=MOC_DEL_INTERFREQNCELL
    End With

    Dim sht As Worksheet
    Set sht = wk.Sheets(SHT_NAME_TABLE_DEF)
    Dim MocName As String
    Dim i As Integer
    For i = 1 To MocNames.Count
        MocName = MocNames.Item(i)
        Call PrepareMocInfo(wk, sht, MocName, MocNames)
    Next i
End Sub

Public Sub PrepareBSCInfos(wk As Workbook)
    Dim sht As Worksheet
    Set sht = wk.Sheets(MOC_BSCINFO)
    Do While g_c_colBSCNAME_RNCID.Count > 0
        g_c_colBSCNAME_RNCID.Remove (1)
    Loop
    Do While g_c_colRNCID_BSCNAME.Count > 0
        g_c_colRNCID_BSCNAME.Remove (1)
    Loop

    Dim i As Long, iRow1 As Long, iRow2 As Long
    iRow1 = ROW_DATA_HW
    iRow2 = GetLastRowIndex(sht)
    Dim Moc As MocClass, Attr1 As AttrClass, Attr2 As AttrClass
    Set Moc = g_c_clsMocInfos.Moc(MOC_BSCINFO)
    Set Attr1 = Moc.Attr(ATTR_BSCNAME)
    Set Attr2 = Moc.Attr(ATTR_RNCID)

    For i = iRow1 To iRow2
        Attr1.Value = sht.Cells(i, Attr1.ColIndex).Value
        Attr2.Value = sht.Cells(i, Attr2.ColIndex).Value
        If Attr1.Value = "" Or Attr2.Value = "" Then
            Exit For
        End If
        g_c_colBSCNAME_RNCID.Add Item:=Attr2.Value, Key:=Attr1.Value
        g_c_colRNCID_BSCNAME.Add Item:=Attr1.Value, Key:=Attr2.Value
    Next i

    m_strWorkbookName_BSCInfo = wk.Name
    Call ClearCollection(m_colNotFoundBSCs)
    Call ClearCollection(m_colNotFoundRNCs)
End Sub

Public Function GetBSCName(sht As Worksheet, RowIndex As Long, ColIndex As Long) As String
    GetBSCName = ""

    Dim s As String
    s = sht.Cells(RowIndex, ColIndex).Value
    If IsInCollection(s, g_c_colRNCID_BSCNAME) Then
        GetBSCName = g_c_colRNCID_BSCNAME.Item(s)
    Else
        If Not IsInCollection(s, m_colNotFoundBSCs) Then
            m_colNotFoundBSCs.Add Item:=s, Key:=s

            Dim isScreenUpdating As Boolean
            isScreenUpdating = Application.ScreenUpdating
            Application.ScreenUpdating = True

            Dim sht2 As Worksheet
            Set sht2 = Application.ActiveSheet
            sht.Activate
            sht.Cells(RowIndex, ColIndex).Select
            sht.Cells(RowIndex, ColIndex).Activate
            s = FormatStr(RSC_STR_BSC_NOT_FOUND_2, s, MOC_BSCINFO, m_strWorkbookName_BSCInfo)
            If MsgBox(s, VbMsgBoxStyle.vbYesNo) = VbMsgBoxResult.vbNo Then
                End
            Else
                sht2.Activate
            End If

            Application.ScreenUpdating = isScreenUpdating
        End If
    End If
End Function

Public Function GetRNCID(sht As Worksheet, RowIndex As Long, ColIndex As Long) As String
    GetRNCID = ""

    Dim s As String
    s = sht.Cells(RowIndex, ColIndex).Value
    If IsInCollection(s, g_c_colBSCNAME_RNCID) Then
        GetRNCID = g_c_colBSCNAME_RNCID.Item(s)
    Else
        If Not IsInCollection(s, m_colNotFoundRNCs) Then
            m_colNotFoundRNCs.Add Item:=s, Key:=s

            Dim isScreenUpdating As Boolean
            isScreenUpdating = Application.ScreenUpdating
            Application.ScreenUpdating = True

            Dim sht2 As Worksheet
            Set sht2 = Application.ActiveSheet
            sht.Activate
            sht.Cells(RowIndex, ColIndex).Select
            sht.Cells(RowIndex, ColIndex).Activate
            s = FormatStr(RSC_STR_RNC_NOT_FOUND_2, s, MOC_BSCINFO, m_strWorkbookName_BSCInfo)
            If MsgBox(s, VbMsgBoxStyle.vbYesNo) = VbMsgBoxResult.vbNo Then
                End
            Else
                sht2.Activate
            End If

            Application.ScreenUpdating = isScreenUpdating
        End If
    End If
End Function

Public Function GetVersion(wk As Workbook) As String
    GetVersion = ""
    Dim sht As Worksheet
    On Error GoTo E
    Set sht = wk.Sheets(SHT_COVER)
    Dim s As String
    s = sht.Range(RNG_VER_CME).Value
    s = s + " " + sht.Range(RNG_VER_NE).Value
    s = s + " " + sht.Range(RNG_VER_RNP).Value
    GetVersion = s
    Exit Function
E:
    MsgBox FormatStr(RSC_STR_MAKE_SURE_TEMPLATE_CORRECT, wk.FullName)
End Function

Public Sub PrintToDebugger(Optional Title As String = "", Optional sht As Worksheet = Nothing)
    Dim iEnd As Long, iEnd2 As Long
    iEnd = GetLastRowIndex(g_c_wkCurrent.Sheets(SHT_DEBUG))

    If Title <> "" Then
        iEnd = iEnd + 1
        If iEnd >= 65536 Then
            Exit Sub
        End If
        g_c_wkCurrent.Sheets(SHT_DEBUG).Range("A" + CStr(iEnd)).Value = Title
    End If
    If Not (sht Is Nothing) Then
        With sht.UsedRange.Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With
        iEnd = iEnd + 1
        
        iEnd2 = GetLastRowIndex(sht)
        If (iEnd + iEnd2 - 1) >= 65536 Then
            Exit Sub
        End If
        sht.UsedRange.Copy Destination:=g_c_wkCurrent.Sheets(SHT_DEBUG).Range("A" + CStr(iEnd))
    End If
End Sub

Public Sub PrintMocToDebugger(Title As String, Moc As MocClass)
    Dim i As Long, Attr As AttrClass
    With g_c_wkCurrent.Sheets(SHT_CONDITION)
        .UsedRange.Clear
        For i = 1 To Moc.Count
            Set Attr = Moc.Attr(i)
            .Cells(1, Attr.ColIndex).Value = Attr.caption
            .Cells(2, Attr.ColIndex).Value = Attr.Value
        Next i
    End With
    Call PrintToDebugger(Title, g_c_wkCurrent.Sheets(SHT_CONDITION))
End Sub

Private Sub PrepareMocInfo(wk As Workbook, sht As Worksheet, MocName As String, MocNames As Collection)
    Dim Moc As MocClass

    g_c_clsMocInfos.Add MocName
    Set Moc = g_c_clsMocInfos.Moc(MocName)
    
    Dim iRow1 As Integer, iRow2 As Integer
    iRow1 = ROW_MOC
    iRow2 = GetLastRowIndex(sht)
    
    Dim i As Integer
    i = iRow1
    Do While i < iRow2
        If sht.Cells(i, COL_MOC_NAME).Value = MocName Then
            iRow1 = i '找这个MOC的开始行
            Exit Do
        End If
        i = i + 1
    Loop
    
    Dim s As String
    i = iRow1 + 1
    Do While i < iRow2
        s = sht.Cells(i, COL_MOC_NAME).Value
        If (s <> "") And IsInCollection(s, MocNames) Then
            iRow2 = i - 1 '找这个MOC的结束行
            Exit Do
        End If
        i = i + 1
    Loop

    Dim strAttrName As String
    Dim shtMoc As Worksheet
    Set shtMoc = wk.Sheets(MocName)
    Dim Attr As AttrClass
    
    For i = iRow1 To iRow2
        strAttrName = sht.Cells(i, COL_ATTR_NAME)
        If strAttrName <> "" Then
            Set Attr = New AttrClass
            With Attr
                .Name = strAttrName
                .ColName = sht.Cells(i, COL_ATTR_COL)
                .DataTypeName = sht.Cells(i, COL_ATTR_TYPE)
                .MinValue = sht.Cells(i, COL_ATTR_MIN)
                .MaxValue = sht.Cells(i, COL_ATTR_MAX)
                .ValueList = sht.Cells(i, COL_ATTR_LIST)
                .caption = shtMoc.Cells(ROW_CAPTION, Moc.Count + 1)
                .MocName = MocName
                .DataType = atUnknown
            End With
            Moc.Add Attr
        End If
    Next i
End Sub
