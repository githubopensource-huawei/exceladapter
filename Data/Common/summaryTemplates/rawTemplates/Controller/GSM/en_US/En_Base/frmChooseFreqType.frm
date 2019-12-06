VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChooseFreqType 
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   OleObjectBlob   =   "frmChooseFreqType.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmChooseFreqType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








'*******************************************************************'
'************ 检查频点设置 Version: 1.4 2011-04-18 17:59 ***********'
'*******************************************************************'



' 取消频点检查
Private Sub btnCancel_Click()
    frmChooseFreqType.Hide
End Sub

' 频点检查
Private Sub btnOK_Click()
    Dim i As Integer
    Dim srcData() As String
    
    ' 清除上次的日志文件
    init ThisWorkbook
    Log bClear:=True
    
    ' 提取源数据 并 检查空数据
    buildSrcData srcData, ThisWorkbook
    setBand
    CheckError
    
    ' 检查数据
    CheckData srcData
    CheckError
  
    MsgBox gMsg_BandCheckPass
    Finish False
End Sub



' 读取源文件，提取需要的数据
Private Function buildSrcData(ByRef dstData() As String, ByRef srcWB As Workbook)
    Dim i, j As Integer, bFind As Boolean, bFlag As Boolean
    Dim btsData() As String, cellData() As String
    ReDim err(0) As String
    
    ' 分别提取 bts 和 cell 数据
    bFind = False
    bFlag = False
    For i = 1 To srcWB.Sheets.count
        If gShtName_bts = srcWB.Sheets(i).name Then
            getBtsData btsData, srcWB.Sheets(i)
            bFind = True
        ElseIf gShtName_cell = srcWB.Sheets(i).name Then
            getCellData cellData, srcWB.Sheets(i)
            bFlag = True
        End If
    Next i
    
    ' 检查 Sheet 页
    If False = bFind Then
        Log "Sheet[ " & gShtName_bts & " ] not exist."
        Exit Function
    End If
    If False = bFlag Then
        Log "Sheet[ " & gShtName_cell & " ] not exist."
        Exit Function
    End If

    ' 整合 bts 和 cell 数据
    For i = 1 To UBound(cellData)
        bFind = False
        For j = 1 To UBound(btsData)
            If cellData(i, 1) = btsData(j, 1) Then
                cellData(i, 2) = btsData(j, 2)
                bFind = True                            ' 可以匹配的 btsname
                btsData(j, 0) = "match"                 ' 可以匹配的 btsname
                Exit For
            End If
        Next j
        
        If False = bFind Then
            insertList err, cellData(i, 1)
        End If
    Next i
    
    ' 检查 bts name 不匹配的数据
    For i = 1 To UBound(btsData)
        If "match" <> btsData(i, 0) Then                ' bts 数据中存在但是 cell 数据中不存在的 btsname
            Log "BTS[ " & btsData(i, 1) & " ] don't in Sheet[ " & gShtName_cell & " ]."
        End If
    Next i
    For i = 1 To UBound(err)                            ' cell 数据中存在但是 bts 数据中不存在的 btsname
        Log "BTS[ " & err(i) & " ] don't in Sheet[ " & gShtName_bts & " ]."
    Next i
    
    dstData = cellData
End Function

' 提取 bts 数据
Private Function getBtsData(ByRef dstData() As String, ByRef sht As Worksheet)
    Dim sp As shtPage
    Dim i, cnt, col_btsname As Integer, col_btstype As Integer, row As Long
    
    ' 读取数据
    readSheetData sp, sht
    
    ' 定位标题行
    FindLocateFromSP sp, gColName_btsName, row, col_btsname
    FindLocateFromSP sp, gColName_btsType, row, col_btstype

    ' 处理列名前的 *
    delTitleAsterisk sp, row
    
    
    cnt = UBound(sp.shtData) - row
    If (0 = row) Or (0 = cnt) Then
        Log "Sheet[ " & sp.shtName & " ] is empty."
    End If
    If (0 = col_btsname) Then
        Log "Not find Column[ " & gColName_btsName & " ] of Sheet[ " & sp.shtName & " ]."
    End If
    If (0 = col_btstype) Then
        Log "Not find Column[ " & gColName_btsType & " ] of Sheet[ " & sp.shtName & " ]."
    End If
    
    ReDim dstData(cnt, 2)
    For i = 1 To cnt
        dstData(i, 1) = sp.shtData(row + i, col_btsname)
        dstData(i, 2) = sp.shtData(row + i, col_btstype)
        
        If "" = dstData(i, 1) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_btsName & " ], Line[ " & row + i & " ] is empty."
        End If
        If "" = dstData(i, 2) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_btsType & " ] is empty of BTS[ " & dstData(i, 1) & " ]."
        End If
    Next i
End Function

' 提取 cell 数据
Private Function getCellData(ByRef dstData() As String, ByRef sht As Worksheet)
    Dim sp As shtPage
    Dim i, cnt, row As Long, col_btsname As Integer, col_cellname As Integer, col_celltype As Integer, col_bcch As Integer, col_nobcch As Integer, col_fc As Integer
    
    ' 读取数据
    readSheetData sp, sht
    
    ' 定位标题行
    FindLocateFromSP sp, gColName_btsName, row, col_btsname

    ' 处理列名前的 *
    delTitleAsterisk sp, row
    
    
    
    FindLocateFromSP sp, gColName_fc, row, col_fc
    FindLocateFromSP sp, gColName_noBcch, row, col_nobcch
    FindLocateFromSP sp, gColName_bcch, row, col_bcch
    FindLocateFromSP sp, gColName_cellType, row, col_celltype
    
    FindLocateFromSP sp, gColName_cellName, row, col_cellname
    If 0 = col_cellname Then
        FindLocateFromSP sp, gColName_cellName2, row, col_cellname
    End If
    cnt = UBound(sp.shtData) - row
    
    ' 检查错误
    If (0 = row) Or (0 = cnt) Then
        Log "Sheet[ " & sp.shtName & " ] is empty."
    End If
    If (0 = col_btsname) Then
        Log "Not find Column[ " & gColName_btsName & " ] of Sheet[ " & sp.shtName & " ]."
    End If
    If (0 = col_cellname) Then
        Log "Not find Column[ " & gColName_cellName & " ] of Sheet[ " & sp.shtName & " ]."
    End If
    If (0 = col_celltype) Then
        Log "Not find Column[ " & gColName_cellType & " ] of Sheet[ " & sp.shtName & " ]."
    End If
    If (0 = col_bcch) Then
        Log "Not find Column[ " & gColName_bcch & " ] of Sheet[ " & sp.shtName & " ]."
    End If
    If (0 = col_nobcch) Then
        Log "Not find Column[ " & gColName_noBcch & " ] of Sheet[ " & sp.shtName & " ]."
    End If
    If (0 = col_fc) Then
        Log "Not find Column[ " & gColName_fc & " ] of Sheet[ " & sp.shtName & " ]."
    End If
    
    ' 提取数据，空一列保存 bts type
    ReDim dstData(cnt, 7)
    For i = 1 To cnt
        dstData(i, 1) = sp.shtData(row + i, col_btsname)        ' bts name 不能为空
        If "" = dstData(i, 1) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_btsName & " ] is empty of Cell[ " & dstData(i, 3) & " ]."
        End If
        
        dstData(i, 3) = sp.shtData(row + i, col_cellname)       ' cell name 不能为空
        If "" = dstData(i, 3) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_cellName & " ] is empty of Cell[ " & dstData(i, 1) & " ]."
        End If
        
        dstData(i, 4) = sp.shtData(row + i, col_celltype)       ' cell type 不能为空
        If "" = dstData(i, 4) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_cellType & " ] is empty of Cell[ " & dstData(i, 3) & " ]."
        End If
        
        dstData(i, 5) = sp.shtData(row + i, col_bcch)           ' Main BCCH 不能为空
        If "" = dstData(i, 4) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_bcch & " ] is empty of Cell[ " & dstData(i, 3) & " ]."
        End If
        
        dstData(i, 6) = sp.shtData(row + i, col_nobcch)         ' Non-main BCCH List 可以为空
        dstData(i, 7) = sp.shtData(row + i, col_fc)             ' Frequency Class 可以为空

        ' non-main bcch 和 frequency class 不能同时非空
        If ("" <> dstData(i, 6)) And ("" <> dstData(i, 7)) Then
            Log "Sheet[ " & sp.shtName & " ], Cell[ " & dstData(i, 3) & " ], Column[ " & gColName_noBcch & " ] and Column[ " & gColName_fc & " ] cannot have value at the same time."
        End If
    Next i
End Function

' 设置频段带宽
Private Function setBand()

    ' bts 3900 900M
    Select Case cbo_bts_900.text
        Case "15M"
            FreqType_BTS3900_900M = "74"
            
        Case "20M"
            FreqType_BTS3900_900M = "99"
            
        Case "20.2M"
            FreqType_BTS3900_900M = "100"
        
        Case Else
            Log "Band[ " & cbo_bts_900.text & " ] not support, please check."
    End Select
    
    ' bts 3900 1800M
    Select Case cbo_bts_1800.text
        Case "15M"
            FreqType_BTS3900_1800M = "74"
            
        Case "20M"
            FreqType_BTS3900_1800M = "99"
            
        Case "20.2M"
            FreqType_BTS3900_1800M = "100"
        
        Case Else
            Log "Band[ " & cbo_bts_1800.text & " ] not support, please check."
    End Select
    
    ' dbs 3900 900M
    Select Case cbo_dbs_900.text
        Case "12.5M"
            FreqType_DBS3900_900M = "62"
            
        Case "15M"
            FreqType_DBS3900_900M = "74"
    
        Case Else
            Log "Band[ " & cbo_dbs_900.text & " ] not support, please check."
    End Select
    
    ' dbs 3900 1800M
    Select Case cbo_dbs_900.text
        Case "12.5M"
            FreqType_DBS3900_1800M = "62"
            
        Case "15M"
            FreqType_DBS3900_1800M = "74"
    
        Case Else
            Log "Band[ " & cbo_dbs_1800.text & " ] not support, please check."
    End Select

End Function

' 检查频点数据
Private Function CheckData(ByRef srcData() As String)
    parseFrequencyClass srcData
    CheckFrequency srcData
    CheckFreq srcData
End Function

' 检查 Frequency of BCCH & Non-Main BCCH List 的值
Public Function CheckFrequency(ByRef srcData() As String)
    Dim i, j, row, col_celltype, col_cellname, col_bcch, col_nobcch  As Integer
    Dim lBcch As Long, lFreq As Long
    Dim sCellName, sBcch, sFreq, serr, str As String

    col_cellname = 3
    col_celltype = 4
    col_bcch = 5
    col_nobcch = 6
    
    For row = 1 To UBound(srcData)
        sCellName = srcData(row, col_cellname)
        
        ' 检查主B
        sBcch = srcData(row, col_bcch)
        If ("" <> sCellName) And ("" = sBcch) Then
            Log "The BCCH of Cell[ " & sCellName & " ] must have value."
        
        ElseIf ("" <> sCellName) And ("" <> sBcch) And (True = IsNumeric(sBcch)) And (0 = InStr(1, sBcch, ".", vbTextCompare)) Then
            lBcch = CLng(sBcch)
            If (0 > lBcch) Or (1024 <= lBcch) Then
                Log "Invaild frequency of Cell[ " & sCellName & " ], Frequency of BCCH[ " & sBcch & " ], please check."
            End If
            
        ElseIf ("" <> sCellName) And ("" <> sBcch) And (False = IsNumeric(sBcch) Or (0 < InStr(1, sBcch, ".", vbTextCompare))) Then
            Log "Invaild frequency of Cell[ " & sCellName & " ], Frequency of BCCH[ " & sBcch & " ], please check."
        End If
    
        ' 处理非主 B 的多余字符
        str = replacExt(srcData(row, col_nobcch), "", ",")      ' 常用字符替换
        srcData(row, col_nobcch) = str
        
        If "" <> str Then
            
            ' 检查非主B
            Dim sSplit() As String, x As Integer, bBcch As Boolean
            sSplit = Split(str, ",", , vbTextCompare)
            serr = ""
            For j = 0 To UBound(sSplit)
                sFreq = sSplit(j)
                If (True = IsNumeric(sFreq)) And (0 = InStr(1, sFreq, ".", vbTextCompare)) Then
                    lFreq = CLng(sFreq)
                    If (0 > lFreq) Or (1024 <= lFreq) Then
                        serr = serr & sFreq & ", "
                    End If
                Else
                    serr = serr & sFreq & ", "
                End If
                
                ' 非主 B 列表当中不能出现主 B 频点
                If sBcch = sFreq Then
                    Log "Main BCCH[ " & sBcch & " ] is not in Column[ " & gColName_noBcch & " ] of Cell[ " & sCellName & " ], please check."
                End If
            Next j
            
            ' 报告非主B频点错误
            If "" <> serr Then
                Log "Invaild frequency of Cell[ " & sCellName & " ], Non-Main BCCH Frequency[ " & Left(serr, Len(serr) - 2) & " ], please check."
            End If
        End If
    Next row
    
End Function

' 解析 Frequency Class：当 Frequency Class 非空且 Non-Main BCCH Frequency List 为空时，用 Frequency Class 填充 Non-Main BCCH Frequency List
Public Function parseFrequencyClass(ByRef srcData() As String) As Boolean
    Dim i, row, col_cellname, ol_celltype, col_bcch, col_nobcch, col_fc As Integer
    Dim sBcch, sCellName, str As String
    
    col_cellname = 3
    col_celltype = 4
    col_bcch = 5
    col_nobcch = 6
    col_fc = 7
    
    For row = 1 To UBound(srcData)
        str = srcData(row, col_fc)
        If "" <> str Then
            
            ' 检查 方括号
            str = Replace(str, "【", "[")
            str = Replace(str, "】", "]")
            If (0 = InStr(1, str, "[")) Or (0 = InStr(1, str, "]")) Then    ' 没有方括号
                Log "Not find ""[]"" in Column[ " & gColName_fc & " ] of Cell[ " & srcData(row, col_cellname) & " ]."
            End If
            
            ' 左右方括号个数不匹配
            If (Len(str) - Len(Replace(str, "[", ""))) <> (Len(str) - Len(Replace(str, "]", ""))) Then
                Log """["" number not match ""]"" number in Column[ " & gColName_fc & " ] of Cell[ " & srcData(row, col_cellname) & " ]."
            End If
            
            ' 检查 小括号
            str = Replace(str, "（", "(")
            str = Replace(str, "）", ")")
            
            ' 左右方括号个数不匹配
            If (Len(str) - Len(Replace(str, "(", ""))) <> (Len(str) - Len(Replace(str, ")", ""))) Then
                Log """("" number not match "")"" number in Column( " & gColName_fc & " ) of Cell[ " & srcData(row, col_cellname) & " ]."
            End If
            
            ' 处理多余字符
            str = replacExt(str, ",，;；:：", ",")
            str = Replace(str, "],", "]")
            srcData(row, col_fc) = str          ' 整理后回拷
            str = replacExt(str, "", ",")       ' 常用字符替换
            
            ' 如果有则删除 main bcch
            Dim sSplit() As String, x As Integer, bBcch As Boolean
            str = Replace(str, "(", "")
            str = Replace(str, ")", "")
            sSplit = Split(str, ",", , vbTextCompare)
            sBcch = srcData(row, col_bcch)
            str = ""
            bBcch = False
            For x = 0 To UBound(sSplit)
                If "" <> sSplit(x) Then
                    If sBcch = sSplit(x) Then
                        bBcch = True
                    Else
                        str = str & sSplit(x) & ","
                    End If
                End If
            Next x

            ' 用 FrequencyClass 填充 non-main bcch
            If "" = srcData(row, col_nobcch) Then
                str = trimComma(Replace(str, ",,", ",", , , vbTextCompare))
                srcData(row, col_nobcch) = str
            End If
            
            ' 当 Frequency Class 中没有 MainBCCH 时，保存并报错
            If ("" <> str) And (False = bBcch) Then
                sCellName = sCellName & srcData(row, col_cellname) & ","
            End If
        End If
    Next row
    
    ' 没有 main bcch 时报错
    If "" <> sCellName Then
        Log "Main BCCH not exist in Column[ " & gColName_fc & " ] of Cell[ " & trimComma(sCellName) & " ], please check."
        parseFrequencyClass = False
        Exit Function
    End If

    parseFrequencyClass = True
End Function

' 检查频点频段带宽
Private Function CheckFreq(ByRef srcData() As String)
    Dim i, j, k, bandwidth, boardNum As Integer
    Dim col_btsname, col_btstype, col_cellname, col_celltype, col_bcch, col_nobcch, col_fc As Integer
    Dim freqList() As String, boardList() As String
    Dim str, band, errFreq As String
    
    col_btsname = 1
    col_btstype = 2
    col_cellname = 3
    col_celltype = 4
    col_bcch = 5
    col_nobcch = 6
    col_fc = 7
    
    For i = 1 To UBound(srcData)
        
        ' 检查所有的频点的频段
        freqList = Split(srcData(i, col_bcch) & "," & srcData(i, col_nobcch), ",")
        If 0 <= UBound(freqList) Then
            CheckFreqListBand freqList, srcData(i, col_celltype), srcData(i, col_cellname)
        End If
        '当有AB通道通过小括号分离时,需要拆分开来作为单独的载频板处理
        srcData(i, col_fc) = Replace(srcData(i, col_fc), "[(", "[")
        srcData(i, col_fc) = Replace(srcData(i, col_fc), ")]", "]")
        srcData(i, col_fc) = Replace(srcData(i, col_fc), "),(", "][")
        ' 当 DBS3900 只有一个单板时，按照 A/B 信道拆成两个
        srcData(i, col_fc) = splitABBoard(srcData(i, col_fc), srcData(i, col_btstype))
        boardList = Split(srcData(i, col_fc), "][")
        
        ' 检查单板
        For j = 0 To UBound(boardList)
            freqList = Split(replacExt(boardList(j), "", ","), ",")
            If 0 <= UBound(freqList) Then
                ' 同一单板内，只能有一种频段
                checkBoardBand freqList, srcData(i, col_celltype), srcData(i, col_cellname)
            
                ' 同一单板内，检查带宽
                band = getBand(freqList(0), srcData(i, col_celltype))
                bandwidth = getBandwidth(srcData(i, col_btstype), band)     ' 获取带宽值
                CheckBandwidth freqList, bandwidth, srcData(i, col_cellname)
            End If
        Next j
        
    Next i
End Function

' 检查带宽
Private Function CheckBandwidth(ByRef freq() As String, ByVal bandwidth As Integer, ByVal cellname As String)
    Dim i, max, min As Integer
    If 1 > bandwidth Then
        Exit Function
    End If
    
    For i = 0 To UBound(freq)
        If freq(i) <= 124 Then      '判断频点值小于124时加1024
            freq(i) = freq(i) + 1024
        End If
    Next i
    
    sortArrayAsInteger freq         '升序排序
    min = freq(0)                   '最小频点
    max = freq(UBound(freq))        '最大频点
    
    ' 当频点差值超出带宽时报错
    If (max - min) > bandwidth Then
        If (max >= 1024) Then
            max = max - 1024
        End If
        If min >= 1024 Then
            min = min - 1024
        End If
        Log "Column[ " & gColName_fc & " ] overstep bandwidth of Cell[ " & cellname & " ]."
    End If
End Function

' 检查频点列表的频段
Public Function CheckFreqListBand(ByRef freqList() As String, ByVal celltype As String, ByVal cellname As String)
    Dim err As String
    err = ""
    
    For j = 0 To UBound(freqList)
        If False = checkBand(freqList(j), celltype) Then
            err = err & freqList(j) & ","
        End If
    Next j
    
    If 1 < Len(err) Then
        Log "Frequency[ " & Left(err, Len(err) - 1) & " ] band not match celltype[ " & celltype & " ] of Cell[ " & cellname & " ]."
    End If
End Function

' 同一单板内，只能有一种频段
Public Function checkBoardBand(ByRef freqList() As String, ByVal celltype As String, ByVal cellname As String)
    Dim i As Integer
    Dim band  As String
    
    ' 同一单板内，只能有一种频段
    band = getBand(freqList(0), celltype)
    
    For i = 1 To UBound(freqList)
        If getBand(freqList(i), celltype) <> band Then
            Log "There can not be two frequency segment in one board of Cell[ " & cellname & " ]."
            Exit For
        End If
    Next i
End Function
            
' 当 DBS3900 只有一个单板时，按照 A/B 信道拆成两个
Public Function splitABBoard(ByVal fc As String, ByVal BTSType As String) As String
    Dim i, boardNum As Integer
    Dim str, freqList() As String
    
    boardNum = Len(fc) - Len(Replace(fc, "]", ""))
    If (1 = boardNum) And ("DBS3900_GSM" = BTSType) Then
        freqList = Split(replacExt(fc, "", ","), ",")
        
        str = ""
        For i = 0 To UBound(freqList) Step 2        ' 偶数频点
            str = str & freqList(i) & ","
        Next i
        fc = "[" & str & "]"
        
        str = ""
        For i = 1 To UBound(freqList) Step 2        ' 奇数频点
            str = str & freqList(i) & ","
        Next i
        fc = fc & "[" & str & "]"
    End If
    
    splitABBoard = fc
End Function


' 检查频点的频段
Public Function checkBand(ByVal freq As String, ByVal celltype As String) As Boolean
    Dim band As String
    checkBand = False
    
    band = getBand(freq, celltype)
    band = Replace(band, "BAND_", "")       ' 提取频段数字，如 "BAND_900" 提出 "900"
    
    If ("" <> band) And (0 < InStr(1, celltype, band)) Then
        checkBand = True
    End If
End Function

' 输入 frequency，cell type，输出频段值
Public Function getBand(ByVal freq As String, ByVal celltype As String) As String
    getBand = ""
    If Not IsNumeric(freq) Then
        Exit Function
    End If
    
    If ((0 <= freq) And (124 >= freq)) Or ((955 <= freq) And (1023 >= freq)) And (("GSM900" = celltype) Or ("GSM900_DCS1800" = celltype)) Then
        getBand = "BAND_900"
        
    ElseIf ((128 <= freq) And (251 >= freq)) And (("GSM850" = celltype) Or ("GSM850_1800" = celltype) Or ("GSM850_1900" = celltype)) Then
        getBand = "BAND_850"
           
    ElseIf ((350 <= freq) And (425 >= freq)) Then
        getBand = "BAND_810"
        
    ElseIf ((512 <= freq) And (885 >= freq)) And (("DCS1800" = celltype) Or ("GSM900_DCS1800" = celltype) Or ("GSM850_1800" = celltype)) Then
        getBand = "BAND_1800"

    ElseIf ((512 <= freq) And (810 >= freq)) And (("PCS1900" = celltype) Or ("GSM850_1900" = celltype)) Then
        getBand = "BAND_1900"
    End If
End Function

' 获取最大带宽
Private Function getBandwidth(ByVal BTSType As String, ByVal freqBand As String) As Integer
    getBandwidth = 0
    
    If ("BTS3900_GSM" = BTSType) And ("BAND_900" = freqBand) Then
        getBandwidth = FreqType_BTS3900_900M
    
    ElseIf ("BTS3900_GSM" = BTSType) And ("BAND_1800" = freqBand) Then
        getBandwidth = FreqType_BTS3900_1800M
        
    ElseIf ("DBS3900_GSM" = BTSType) And ("BAND_900" = freqBand) Then
        getBandwidth = FreqType_DBS3900_900M
    
    ElseIf ("DBS3900_GSM" = BTSType) And ("BAND_1800" = freqBand) Then
        getBandwidth = FreqType_DBS3900_1800M
    End If
End Function

'*******************************************************************'
'************ 枚举模板文件 Version: 1.3 2010-12-18 12:59 ***********'
'*******************************************************************'

' 选择模板，创建枚举
Private Sub btnChooseBTSName_Click()
    Dim str, fileList() As String
    Dim row As Long, i, dstRow, col As Integer
    Dim btsSht As Worksheet
    Dim rowCnt As Long, colCnt As Integer, bFlag As Boolean
    
    '查找 BTS 页
    bFlag = False
    For i = 1 To ThisWorkbook.Sheets.count
        If gShtName_bts = ThisWorkbook.Sheets(i).name Then
            Set btsSht = ThisWorkbook.Sheets(i)
            bFlag = True
            Exit For
        End If
    Next
    If False = bFlag Then
        MsgBox "Sheet[ " & gShtName_bts & " ] not exist."
        Exit Sub
    End If
    
    '查找 BTS Template Name 单元格
    GetShtRange btsSht, rowCnt, colCnt
    FindLocateFromSht btsSht, gColName_btsTpltName, row, col
    
    '计算目的行数
    If 100 < rowCnt Then
        dstRow = rowCnt
    Else
        dstRow = 100
    End If
    
    '选择需要的文件
    ChooseFile fileList
    
    '生成枚举信息
    str = ""
    For i = 1 To UBound(fileList)
        If "" <> fileList(i) Then
            str = str & Dir(fileList(i)) & ","
        End If
    Next i
    If 2 > Len(str) Then
        Exit Sub
    End If
    str = Left(str, Len(str) - 1)

    setRangeEnum btsSht, row + 1, col, dstRow, col, str
End Sub

' 取消枚举
Private Sub btnCancelChoose_Click()
    Dim str, fileList() As String
    Dim row As Long, i, dstRow, col As Integer
    Dim btsSht As Worksheet
    Dim rowCnt As Long, colCnt As Integer, bFlag As Boolean
    
    '查找 BTS 页
    bFlag = False
    For i = 1 To ThisWorkbook.Sheets.count
        If gShtName_bts = ThisWorkbook.Sheets(i).name Then
            Set btsSht = ThisWorkbook.Sheets(i)
            bFlag = True
            Exit For
        End If
    Next
    If False = bFlag Then
        MsgBox "Sheet[ " & gShtName_bts & " ] not exist."
        Exit Sub
    End If
    
    '查找 BTS Template Name 单元格
    GetShtRange btsSht, rowCnt, colCnt
    FindLocateFromSht btsSht, gColName_btsTpltName, row, col
    
    '计算目的行数
    If 100 < rowCnt Then
        dstRow = rowCnt
    Else
        dstRow = 100
    End If
    
    ' 清除枚举设置
    ClearRangeEnum btsSht, row + 1, col, dstRow, col
End Sub

' 给选定的单元格设置枚举
Private Function setRangeEnum(ByRef sht As Worksheet, ByVal orgRow As Integer, ByVal orgCol As Integer, ByVal dstRow As Integer, ByVal dstCol As Integer, ByVal enumInfo As String) As Boolean
    If (1 > orgRow) Or (1 > orgCol) Or (1 > dstRow) Or (1 > dstCol) Or (gRowMax < orgRow) Or (gColMax < orgCol) Or (gRowMax < dstRow) Or (gColMax < dstCol) Then
        setRangeEnum = False
        Exit Function
    End If

    sht.Activate
    Range(sht.Cells(orgRow, orgCol), sht.Cells(dstRow, dstCol)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=enumInfo
        .IgnoreBlank = True
        .InCellDropdown = True
        .inputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "Invaild value, please check."
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    Range(sht.Cells(orgRow, orgCol), sht.Cells(orgRow, orgCol)).Select
    setRangeEnum = True
End Function

' 给选定的单元格删除枚举
Private Function ClearRangeEnum(ByRef sht As Worksheet, ByVal orgRow As Integer, ByVal orgCol As Integer, ByVal dstRow As Integer, ByVal dstCol As Integer) As Boolean
    If (1 > orgRow) Or (1 > orgCol) Or (1 > dstRow) Or (1 > dstCol) Or (gRowMax < orgRow) Or (gColMax < orgCol) Or (gRowMax < dstRow) Or (gColMax < dstCol) Then
        ClearRangeEnum = False
        Exit Function
    End If

    sht.Activate
    Range(sht.Cells(orgRow, orgCol), sht.Cells(dstRow, dstCol)).Select
    With Selection.Validation
        .Delete
    End With
    
    Range(sht.Cells(orgRow, orgCol), sht.Cells(orgRow, orgCol)).Select
    ClearRangeEnum = True
End Function

' 复制单元格的设置
Private Function copyCellSet(ByRef sht As Worksheet, ByVal orgRow, ByVal orgCol, ByVal dstRow, ByVal dstCol)
    If (1 > orgRow) Or (1 > orgCol) Or (1 > dstRow) Or (1 > dstCol) Or (gRowMax < orgRow) Or (gColMax < orgCol) Or (gRowMax < dstRow) Or (gColMax < dstCol) Then
        Exit Function
    End If
    
    sht.Activate
    Selection.AutoFill Destination:=Range(sht.Cells(orgRow, orgCol), sht.Cells(dstRow, dstCol)), Type:=xlFillDefault
    Range(sht.Cells(orgRow, orgCol), sht.Cells(orgRow, orgCol)).Select
End Function


Public Sub InitGUI()
  init ThisWorkbook
   Frame1.Caption = gCaption_BandTitle
   btnOK.Caption = gCaption_BandCheck
   btnCancel.Caption = gCaption_BandCancel

End Sub



