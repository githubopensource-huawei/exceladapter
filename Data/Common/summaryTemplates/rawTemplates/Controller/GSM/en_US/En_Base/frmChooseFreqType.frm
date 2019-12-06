VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChooseFreqType 
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   OleObjectBlob   =   "frmChooseFreqType.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "frmChooseFreqType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








'*******************************************************************'
'************ ���Ƶ������ Version: 1.4 2011-04-18 17:59 ***********'
'*******************************************************************'



' ȡ��Ƶ����
Private Sub btnCancel_Click()
    frmChooseFreqType.Hide
End Sub

' Ƶ����
Private Sub btnOK_Click()
    Dim i As Integer
    Dim srcData() As String
    
    ' ����ϴε���־�ļ�
    init ThisWorkbook
    Log bClear:=True
    
    ' ��ȡԴ���� �� ��������
    buildSrcData srcData, ThisWorkbook
    setBand
    CheckError
    
    ' �������
    CheckData srcData
    CheckError
  
    MsgBox gMsg_BandCheckPass
    Finish False
End Sub



' ��ȡԴ�ļ�����ȡ��Ҫ������
Private Function buildSrcData(ByRef dstData() As String, ByRef srcWB As Workbook)
    Dim i, j As Integer, bFind As Boolean, bFlag As Boolean
    Dim btsData() As String, cellData() As String
    ReDim err(0) As String
    
    ' �ֱ���ȡ bts �� cell ����
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
    
    ' ��� Sheet ҳ
    If False = bFind Then
        Log "Sheet[ " & gShtName_bts & " ] not exist."
        Exit Function
    End If
    If False = bFlag Then
        Log "Sheet[ " & gShtName_cell & " ] not exist."
        Exit Function
    End If

    ' ���� bts �� cell ����
    For i = 1 To UBound(cellData)
        bFind = False
        For j = 1 To UBound(btsData)
            If cellData(i, 1) = btsData(j, 1) Then
                cellData(i, 2) = btsData(j, 2)
                bFind = True                            ' ����ƥ��� btsname
                btsData(j, 0) = "match"                 ' ����ƥ��� btsname
                Exit For
            End If
        Next j
        
        If False = bFind Then
            insertList err, cellData(i, 1)
        End If
    Next i
    
    ' ��� bts name ��ƥ�������
    For i = 1 To UBound(btsData)
        If "match" <> btsData(i, 0) Then                ' bts �����д��ڵ��� cell �����в����ڵ� btsname
            Log "BTS[ " & btsData(i, 1) & " ] don't in Sheet[ " & gShtName_cell & " ]."
        End If
    Next i
    For i = 1 To UBound(err)                            ' cell �����д��ڵ��� bts �����в����ڵ� btsname
        Log "BTS[ " & err(i) & " ] don't in Sheet[ " & gShtName_bts & " ]."
    Next i
    
    dstData = cellData
End Function

' ��ȡ bts ����
Private Function getBtsData(ByRef dstData() As String, ByRef sht As Worksheet)
    Dim sp As shtPage
    Dim i, cnt, col_btsname As Integer, col_btstype As Integer, row As Long
    
    ' ��ȡ����
    readSheetData sp, sht
    
    ' ��λ������
    FindLocateFromSP sp, gColName_btsName, row, col_btsname
    FindLocateFromSP sp, gColName_btsType, row, col_btstype

    ' ��������ǰ�� *
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

' ��ȡ cell ����
Private Function getCellData(ByRef dstData() As String, ByRef sht As Worksheet)
    Dim sp As shtPage
    Dim i, cnt, row As Long, col_btsname As Integer, col_cellname As Integer, col_celltype As Integer, col_bcch As Integer, col_nobcch As Integer, col_fc As Integer
    
    ' ��ȡ����
    readSheetData sp, sht
    
    ' ��λ������
    FindLocateFromSP sp, gColName_btsName, row, col_btsname

    ' ��������ǰ�� *
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
    
    ' ������
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
    
    ' ��ȡ���ݣ���һ�б��� bts type
    ReDim dstData(cnt, 7)
    For i = 1 To cnt
        dstData(i, 1) = sp.shtData(row + i, col_btsname)        ' bts name ����Ϊ��
        If "" = dstData(i, 1) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_btsName & " ] is empty of Cell[ " & dstData(i, 3) & " ]."
        End If
        
        dstData(i, 3) = sp.shtData(row + i, col_cellname)       ' cell name ����Ϊ��
        If "" = dstData(i, 3) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_cellName & " ] is empty of Cell[ " & dstData(i, 1) & " ]."
        End If
        
        dstData(i, 4) = sp.shtData(row + i, col_celltype)       ' cell type ����Ϊ��
        If "" = dstData(i, 4) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_cellType & " ] is empty of Cell[ " & dstData(i, 3) & " ]."
        End If
        
        dstData(i, 5) = sp.shtData(row + i, col_bcch)           ' Main BCCH ����Ϊ��
        If "" = dstData(i, 4) Then
            Log "Sheet[ " & sp.shtName & " ], Column[ " & gColName_bcch & " ] is empty of Cell[ " & dstData(i, 3) & " ]."
        End If
        
        dstData(i, 6) = sp.shtData(row + i, col_nobcch)         ' Non-main BCCH List ����Ϊ��
        dstData(i, 7) = sp.shtData(row + i, col_fc)             ' Frequency Class ����Ϊ��

        ' non-main bcch �� frequency class ����ͬʱ�ǿ�
        If ("" <> dstData(i, 6)) And ("" <> dstData(i, 7)) Then
            Log "Sheet[ " & sp.shtName & " ], Cell[ " & dstData(i, 3) & " ], Column[ " & gColName_noBcch & " ] and Column[ " & gColName_fc & " ] cannot have value at the same time."
        End If
    Next i
End Function

' ����Ƶ�δ���
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

' ���Ƶ������
Private Function CheckData(ByRef srcData() As String)
    parseFrequencyClass srcData
    CheckFrequency srcData
    CheckFreq srcData
End Function

' ��� Frequency of BCCH & Non-Main BCCH List ��ֵ
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
        
        ' �����B
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
    
        ' ������� B �Ķ����ַ�
        str = replacExt(srcData(row, col_nobcch), "", ",")      ' �����ַ��滻
        srcData(row, col_nobcch) = str
        
        If "" <> str Then
            
            ' ������B
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
                
                ' ���� B �б��в��ܳ����� B Ƶ��
                If sBcch = sFreq Then
                    Log "Main BCCH[ " & sBcch & " ] is not in Column[ " & gColName_noBcch & " ] of Cell[ " & sCellName & " ], please check."
                End If
            Next j
            
            ' �������BƵ�����
            If "" <> serr Then
                Log "Invaild frequency of Cell[ " & sCellName & " ], Non-Main BCCH Frequency[ " & Left(serr, Len(serr) - 2) & " ], please check."
            End If
        End If
    Next row
    
End Function

' ���� Frequency Class���� Frequency Class �ǿ��� Non-Main BCCH Frequency List Ϊ��ʱ���� Frequency Class ��� Non-Main BCCH Frequency List
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
            
            ' ��� ������
            str = Replace(str, "��", "[")
            str = Replace(str, "��", "]")
            If (0 = InStr(1, str, "[")) Or (0 = InStr(1, str, "]")) Then    ' û�з�����
                Log "Not find ""[]"" in Column[ " & gColName_fc & " ] of Cell[ " & srcData(row, col_cellname) & " ]."
            End If
            
            ' ���ҷ����Ÿ�����ƥ��
            If (Len(str) - Len(Replace(str, "[", ""))) <> (Len(str) - Len(Replace(str, "]", ""))) Then
                Log """["" number not match ""]"" number in Column[ " & gColName_fc & " ] of Cell[ " & srcData(row, col_cellname) & " ]."
            End If
            
            ' ��� С����
            str = Replace(str, "��", "(")
            str = Replace(str, "��", ")")
            
            ' ���ҷ����Ÿ�����ƥ��
            If (Len(str) - Len(Replace(str, "(", ""))) <> (Len(str) - Len(Replace(str, ")", ""))) Then
                Log """("" number not match "")"" number in Column( " & gColName_fc & " ) of Cell[ " & srcData(row, col_cellname) & " ]."
            End If
            
            ' ��������ַ�
            str = replacExt(str, ",��;��:��", ",")
            str = Replace(str, "],", "]")
            srcData(row, col_fc) = str          ' �����ؿ�
            str = replacExt(str, "", ",")       ' �����ַ��滻
            
            ' �������ɾ�� main bcch
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

            ' �� FrequencyClass ��� non-main bcch
            If "" = srcData(row, col_nobcch) Then
                str = trimComma(Replace(str, ",,", ",", , , vbTextCompare))
                srcData(row, col_nobcch) = str
            End If
            
            ' �� Frequency Class ��û�� MainBCCH ʱ�����沢����
            If ("" <> str) And (False = bBcch) Then
                sCellName = sCellName & srcData(row, col_cellname) & ","
            End If
        End If
    Next row
    
    ' û�� main bcch ʱ����
    If "" <> sCellName Then
        Log "Main BCCH not exist in Column[ " & gColName_fc & " ] of Cell[ " & trimComma(sCellName) & " ], please check."
        parseFrequencyClass = False
        Exit Function
    End If

    parseFrequencyClass = True
End Function

' ���Ƶ��Ƶ�δ���
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
        
        ' ������е�Ƶ���Ƶ��
        freqList = Split(srcData(i, col_bcch) & "," & srcData(i, col_nobcch), ",")
        If 0 <= UBound(freqList) Then
            CheckFreqListBand freqList, srcData(i, col_celltype), srcData(i, col_cellname)
        End If
        '����ABͨ��ͨ��С���ŷ���ʱ,��Ҫ��ֿ�����Ϊ��������Ƶ�崦��
        srcData(i, col_fc) = Replace(srcData(i, col_fc), "[(", "[")
        srcData(i, col_fc) = Replace(srcData(i, col_fc), ")]", "]")
        srcData(i, col_fc) = Replace(srcData(i, col_fc), "),(", "][")
        ' �� DBS3900 ֻ��һ������ʱ������ A/B �ŵ��������
        srcData(i, col_fc) = splitABBoard(srcData(i, col_fc), srcData(i, col_btstype))
        boardList = Split(srcData(i, col_fc), "][")
        
        ' ��鵥��
        For j = 0 To UBound(boardList)
            freqList = Split(replacExt(boardList(j), "", ","), ",")
            If 0 <= UBound(freqList) Then
                ' ͬһ�����ڣ�ֻ����һ��Ƶ��
                checkBoardBand freqList, srcData(i, col_celltype), srcData(i, col_cellname)
            
                ' ͬһ�����ڣ�������
                band = getBand(freqList(0), srcData(i, col_celltype))
                bandwidth = getBandwidth(srcData(i, col_btstype), band)     ' ��ȡ����ֵ
                CheckBandwidth freqList, bandwidth, srcData(i, col_cellname)
            End If
        Next j
        
    Next i
End Function

' ������
Private Function CheckBandwidth(ByRef freq() As String, ByVal bandwidth As Integer, ByVal cellname As String)
    Dim i, max, min As Integer
    If 1 > bandwidth Then
        Exit Function
    End If
    
    For i = 0 To UBound(freq)
        If freq(i) <= 124 Then      '�ж�Ƶ��ֵС��124ʱ��1024
            freq(i) = freq(i) + 1024
        End If
    Next i
    
    sortArrayAsInteger freq         '��������
    min = freq(0)                   '��СƵ��
    max = freq(UBound(freq))        '���Ƶ��
    
    ' ��Ƶ���ֵ��������ʱ����
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

' ���Ƶ���б��Ƶ��
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

' ͬһ�����ڣ�ֻ����һ��Ƶ��
Public Function checkBoardBand(ByRef freqList() As String, ByVal celltype As String, ByVal cellname As String)
    Dim i As Integer
    Dim band  As String
    
    ' ͬһ�����ڣ�ֻ����һ��Ƶ��
    band = getBand(freqList(0), celltype)
    
    For i = 1 To UBound(freqList)
        If getBand(freqList(i), celltype) <> band Then
            Log "There can not be two frequency segment in one board of Cell[ " & cellname & " ]."
            Exit For
        End If
    Next i
End Function
            
' �� DBS3900 ֻ��һ������ʱ������ A/B �ŵ��������
Public Function splitABBoard(ByVal fc As String, ByVal BTSType As String) As String
    Dim i, boardNum As Integer
    Dim str, freqList() As String
    
    boardNum = Len(fc) - Len(Replace(fc, "]", ""))
    If (1 = boardNum) And ("DBS3900_GSM" = BTSType) Then
        freqList = Split(replacExt(fc, "", ","), ",")
        
        str = ""
        For i = 0 To UBound(freqList) Step 2        ' ż��Ƶ��
            str = str & freqList(i) & ","
        Next i
        fc = "[" & str & "]"
        
        str = ""
        For i = 1 To UBound(freqList) Step 2        ' ����Ƶ��
            str = str & freqList(i) & ","
        Next i
        fc = fc & "[" & str & "]"
    End If
    
    splitABBoard = fc
End Function


' ���Ƶ���Ƶ��
Public Function checkBand(ByVal freq As String, ByVal celltype As String) As Boolean
    Dim band As String
    checkBand = False
    
    band = getBand(freq, celltype)
    band = Replace(band, "BAND_", "")       ' ��ȡƵ�����֣��� "BAND_900" ��� "900"
    
    If ("" <> band) And (0 < InStr(1, celltype, band)) Then
        checkBand = True
    End If
End Function

' ���� frequency��cell type�����Ƶ��ֵ
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

' ��ȡ������
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
'************ ö��ģ���ļ� Version: 1.3 2010-12-18 12:59 ***********'
'*******************************************************************'

' ѡ��ģ�壬����ö��
Private Sub btnChooseBTSName_Click()
    Dim str, fileList() As String
    Dim row As Long, i, dstRow, col As Integer
    Dim btsSht As Worksheet
    Dim rowCnt As Long, colCnt As Integer, bFlag As Boolean
    
    '���� BTS ҳ
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
    
    '���� BTS Template Name ��Ԫ��
    GetShtRange btsSht, rowCnt, colCnt
    FindLocateFromSht btsSht, gColName_btsTpltName, row, col
    
    '����Ŀ������
    If 100 < rowCnt Then
        dstRow = rowCnt
    Else
        dstRow = 100
    End If
    
    'ѡ����Ҫ���ļ�
    ChooseFile fileList
    
    '����ö����Ϣ
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

' ȡ��ö��
Private Sub btnCancelChoose_Click()
    Dim str, fileList() As String
    Dim row As Long, i, dstRow, col As Integer
    Dim btsSht As Worksheet
    Dim rowCnt As Long, colCnt As Integer, bFlag As Boolean
    
    '���� BTS ҳ
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
    
    '���� BTS Template Name ��Ԫ��
    GetShtRange btsSht, rowCnt, colCnt
    FindLocateFromSht btsSht, gColName_btsTpltName, row, col
    
    '����Ŀ������
    If 100 < rowCnt Then
        dstRow = rowCnt
    Else
        dstRow = 100
    End If
    
    ' ���ö������
    ClearRangeEnum btsSht, row + 1, col, dstRow, col
End Sub

' ��ѡ���ĵ�Ԫ������ö��
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

' ��ѡ���ĵ�Ԫ��ɾ��ö��
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

' ���Ƶ�Ԫ�������
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



