Attribute VB_Name = "TrxInfoMgr"
Option Explicit

Public Const CELLBAND_850 = "GSM850"
Public Const CELLBAND_900 = "GSM900"
Public Const CELLBAND_1800 = "GSM1800"
Public Const CELLBAND_1900 = "GSM1900"
Private Const MOC_TRXINFO = "TRXINFO"
Private Const ATTR_MAGRP = "MAGRPFREQLIST"
Private Const ATTR_BRDNO = "BRDNO"
Private Const ATTR_TRXPN = "TRXPN"
Private Const ATTR_ANTPASSNO = "ANTPASSNO"
Private Const ATTR_ANTENNAGROUPID = "ANTENNAGROUPID"
Private Const ATTR_HSN = "HSN"
Private Const ATTR_HOPMODE = "HOPMODE"
Private Const ATTR_TRXNUM = "TRXNUM"
Private Const ATTR_TYPE = "TYPE"
Private Const ATTR_BCH_FREQ = "BCCHFREQ"
Private Const ATTR_TCH_FREQ = "NONBCCHFREQLIST"
Private Const ATTR_BTSNAME = "BTSNAME"
Private Const ATTR_CELLNAME = "CELLNAME"
Private Const ATTR_GLOCELLID = "GLocellId"
Private Const MOC_GLOCELL = "GLOCELL"
Private Const MOC_GCELL = "GCELL"
Private Const MOC_TRXBIND2PHYBRD = "TRXBIND2PHYBRD"
Private Const MOC_GCELLMAGRP = "GCELLMAGRP"
Private REMOVESUCCESS As Boolean
Private CellShtName As String
Private rowNumber As Long
Private NeedDeleteIndex As Long
Private freqCollection As Collection
Private hasTrx As Boolean
Public Sub test()
    Dim trxStr As String
    trxStr = ""
    trxStr = getFreqLstStrByGCell("GSM Cell", "MRAT3", "0")
 End Sub
 
    
'功能说明：根据基站名称BtsName和小区名称GCellName在页签GCellShtName获取频点信息
'注意点：
'1、载频个数TRXNUM必须已经填写；
'2、主B和非主B必须已经填写；
'3、不满足1和2情况，返回空；
'4、当实际的频点个数大于TRXNUM时，后面的频点被截掉，也就是不被绑定的频点
Public Function getFreqLstStrByGCell(GCellShtName As String, btsName As String, GLocellId As String) As String
    
    Dim cellsheet As Worksheet
    Set cellsheet = ThisWorkbook.Worksheets(GCellShtName)
    
    Dim trxNumIndex As Long
    trxNumIndex = getColNum(GCellShtName, 2, ATTR_TRXNUM, MOC_TRXINFO)
    
    Dim bchIndex As Long
    bchIndex = getColNum(GCellShtName, 2, ATTR_BCH_FREQ, MOC_TRXINFO)
    
    Dim tchIndex As Long
    tchIndex = getColNum(GCellShtName, 2, ATTR_TCH_FREQ, MOC_TRXINFO)
 
    Dim btsNameColIndex As Long
    btsNameColIndex = getColNum(GCellShtName, 2, ATTR_BTSNAME, MOC_GCELL)
        
    Dim gLocellIdColIndex As Long
    gLocellIdColIndex = getColNum(GCellShtName, 2, ATTR_GLOCELLID, MOC_GLOCELL)
 
    Dim index As Long
    
    For index = 3 To cellsheet.range("a1048576").End(xlUp).row
        If cellsheet.Cells(index, btsNameColIndex).value = btsName _
            And cellsheet.Cells(index, gLocellIdColIndex).value = GLocellId Then
            
            Dim trxNumStr As String
            trxNumStr = cellsheet.Cells(index, trxNumIndex).value
                        
            Dim trxNumArr() As String
            Dim elNo As Long
            Dim curCellTrxNum As Long
            curCellTrxNum = 0
            trxNumArr = Split(trxNumStr, ",")
            For elNo = LBound(trxNumArr) To UBound(trxNumArr)
                curCellTrxNum = curCellTrxNum + trxNumArr(elNo)
            Next
                      
            
            Dim tchFreqs As String
            tchFreqs = cellsheet.Cells(index, tchIndex).value
            If Not (Trim(tchFreqs) = "") Then
                tchFreqs = "," + tchFreqs
            End If
            getFreqLstStrByGCell = cutTail(cellsheet.Cells(index, bchIndex).value + tchFreqs, curCellTrxNum)
            Exit Function
        End If
    Next
    
    getFreqLstStrByGCell = ""
End Function
Public Function cutTail(str As String, elementNum As Long) As String
    Dim strArr() As String
    Dim finaStr As String
    
    
    strArr = Split(str, ",")
    finaStr = ""
    Dim iNo As Long
    iNo = 0
    While (iNo < elementNum And iNo <= UBound(strArr))
        If (0 = iNo) Then
            finaStr = strArr(iNo)
        Else
            finaStr = finaStr + "," + strArr(iNo)
        End If
        
        iNo = iNo + 1
    Wend
                   
    cutTail = finaStr
End Function

Public Sub deleteFreqAndAssocMo(CellSheetName As String, rowNum As Long, SelectedFreqIndex As Long, _
    freqCollect As Collection)
    On Error GoTo ErrorHandler
    CellShtName = CellSheetName
    rowNumber = rowNum
    NeedDeleteIndex = SelectedFreqIndex
    Set freqCollection = freqCollect
    
    Dim freq As String
    freq = freqCollection.Item(SelectedFreqIndex + 1)
    
    Call deleteTrx
    
    Call changeTrxNum(freq)
    
    hasTrx = True
    hasTrx = trxExist()
    
    Call deleteTrxChildMo
    
    Call changeHopType
    
    Call changeMaGrpList(freq)
    
    Call changeTrxBind
    
ErrorHandler:
    REMOVESUCCESS = False
    
End Sub
Private Function trxExist() As Boolean
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(CellShtName)
    
    Dim trxNumIndex As Long
    Dim trxNumStr As String
    trxExist = True
    
    trxNumIndex = getColNum(CellShtName, 2, ATTR_TRXNUM, MOC_TRXINFO)
    trxNumStr = ws.Cells(rowNumber, trxNumIndex)
    
    If (Trim(trxNumStr) = "") Or (Trim(trxNumStr) = "0") Or (Trim(trxNumStr) = "0,0") Then
        trxExist = False
    End If
End Function
Private Function deleteTrx()
    Dim bchIndex As Long
    Dim tchIndex As Long
    Dim tchStr As String
    Dim hasBch As Boolean
    Dim freq As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CellShtName)
    
    bchIndex = getColNum(CellShtName, 2, ATTR_BCH_FREQ, MOC_TRXINFO)
    tchIndex = getColNum(CellShtName, 2, ATTR_TCH_FREQ, MOC_TRXINFO)
    hasBch = existBch(ws, bchIndex)
    freq = freqCollection.Item(NeedDeleteIndex + 1)

    tchStr = ""
    
    If freq = ws.Cells(rowNumber, bchIndex) Then

        freqCollection.Remove (NeedDeleteIndex + 1)
        ws.Cells(rowNumber, bchIndex) = ""

    Else
        freqCollection.Remove (NeedDeleteIndex + 1)
        Dim tCount As Long
        If freqCollection.count >= 1 Then
            If hasBch Then
                For tCount = 2 To freqCollection.count
                    tchStr = tchStr + freqCollection.Item(tCount) + ","
                Next
            Else
                For tCount = 1 To freqCollection.count
                    tchStr = tchStr + freqCollection.Item(tCount) + ","
                Next
            End If
        End If
        If Trim(tchStr) <> "" Then
            ws.Cells(rowNumber, tchIndex) = Left(tchStr, Len(tchStr) - 1)
        Else
            ws.Cells(rowNumber, tchIndex) = ""
        End If
    End If
End Function
Private Function existBch(ws As Worksheet, bchIndex As Long) As Boolean
    Dim bch As String
    bch = ws.Cells(rowNumber, bchIndex).value
    If bch = "" Then
        existBch = False
    Else
        existBch = True
    End If
End Function

Private Function deleteTrxChildMo()
    Dim TrxChildMocCollection As Collection
    Set TrxChildMocCollection = New Collection
    
    TrxChildMocCollection.Add ("GTRXDEV")
    TrxChildMocCollection.Add ("GTRXRSVPARA")
    TrxChildMocCollection.Add ("GTRXIUO")
    TrxChildMocCollection.Add ("GTRXBASE")
    TrxChildMocCollection.Add ("GTRXFC")
    TrxChildMocCollection.Add ("GTRXRLALM")
    
    Dim mocName As Variant
    
    For Each mocName In TrxChildMocCollection
        changeAttrByMoc (mocName)
    Next
End Function
Private Function changeTrxNum(freq As String)
    Dim trxNumber As String
    Dim cellBand As String
    Dim trxNumIndex As Long
    Dim cellBandIndex As Long
    Dim trxNumArray() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CellShtName)
        
    trxNumIndex = getColNum(CellShtName, 2, ATTR_TRXNUM, MOC_TRXINFO)
    cellBandIndex = getColNum(CellShtName, 2, ATTR_TYPE, MOC_GCELL)
    
    trxNumber = ws.Cells(rowNumber, trxNumIndex)
    cellBand = ws.Cells(rowNumber, cellBandIndex)
    
    trxNumArray = Split(trxNumber, ",")
    Dim eNo As Long
    eNo = UBound(trxNumArray)
    
    If eNo = 0 Then
        ws.Cells(rowNumber, trxNumIndex) = trxNumber - 1
    Else
        Dim lBandNum As Long
        Dim uBandNum As Long
        lBandNum = trxNumArray(0)
        uBandNum = trxNumArray(1)
        
        Dim band As String
        band = getFreqBand(CStr(freq), cellBand)
            
        If (band = CELLBAND_850 Or band = CELLBAND_900) Then
            lBandNum = lBandNum - 1
        Else
            uBandNum = uBandNum - 1
        End If
        
        ws.Cells(rowNumber, trxNumIndex) = CStr(lBandNum) + "," + CStr(uBandNum)
    End If
End Function
Private Function changeHopType()
    Dim tchIndex As Long
    Dim trxNumIndex As Long
    Dim hopTypeIndex As Long
    
    Dim tchStr As String
    Dim trxNumStr As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CellShtName)
    tchIndex = getColNum(CellShtName, 2, ATTR_TCH_FREQ, MOC_TRXINFO)
    trxNumIndex = getColNum(CellShtName, 2, ATTR_TRXNUM, MOC_TRXINFO)
    hopTypeIndex = getColNum(CellShtName, 2, ATTR_HOPMODE, MOC_GCELLMAGRP)
    tchStr = ws.Cells(rowNumber, tchIndex)
    trxNumStr = ws.Cells(rowNumber, trxNumIndex)
    
    If Trim(tchStr) = "" Or CLng(trxNumStr) = 0 Or (Not hasTrx) Then
        ws.Cells(rowNumber, hopTypeIndex) = "NO_FH"
    End If
    
End Function
Private Function changeMaGrpList(needDeletefreq As String)
    Dim maGrpIndex As Long
    Dim maGrpStr As String
    Dim hopModeIndex As Long
    Dim hopModeStr As String
    Dim isRFHop As Boolean
    Dim newMaGrpStr As String
    Dim hsnIndex As Long
    Dim hsnStr As String
    Dim hsnArray() As String
    Dim newHsn As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CellShtName)
    
    newMaGrpStr = ""
    newHsn = ""
    
    maGrpIndex = getColNum(CellShtName, 2, ATTR_MAGRP, MOC_TRXINFO)
    maGrpStr = ws.Cells(rowNumber, maGrpIndex)
    
    hsnIndex = getColNum(CellShtName, 2, ATTR_HSN, MOC_GCELLMAGRP)
    hsnStr = ws.Cells(rowNumber, hsnIndex)
    
    hopModeIndex = getColNum(CellShtName, 2, ATTR_HOPMODE, MOC_GCELLMAGRP)
    hopModeStr = ws.Cells(rowNumber, hopModeIndex)
    
    If Trim(hopModeStr) = "NO_FH" Then
        ws.Cells(rowNumber, hsnIndex) = ""
        ws.Cells(rowNumber, maGrpIndex) = ""
    Else

        hsnArray = Split(hsnStr, ",")
    
        Dim hNo As Long
        hNo = UBound(hsnArray)
    
        isRFHop = False
        If hopModeStr = "RF_FH" Then
            isRFHop = True
        End If
    
        Dim maGrpArray() As String
        maGrpArray = Split(maGrpStr, "]")
    
    Dim index As Long
    For index = LBound(maGrpArray) To UBound(maGrpArray)
        Dim maFreqs As String
        Dim processFreqs As String
        maFreqs = maGrpArray(index)
        If Trim(maFreqs) <> "" Then
            processFreqs = deleteFreq(needDeletefreq, Right(maFreqs, Len(maFreqs) - 1), isRFHop)
                    
                If Trim(processFreqs) <> "" Then
                    newMaGrpStr = newMaGrpStr + "[" + Trim(processFreqs) + "]"
                    If hNo <> 0 Then
                        newHsn = newHsn + hsnArray(index) + ","
                    Else
                        newHsn = hsnStr + ","
                    End If
                End If
            End If
        Next
    
        ws.Cells(rowNumber, maGrpIndex) = newMaGrpStr
    
        If Trim(newMaGrpStr) <> "" Then
            ws.Cells(rowNumber, hsnIndex) = Left(newHsn, Len(newHsn) - 1)
        Else
            ws.Cells(rowNumber, hsnIndex) = ""
        End If
    End If
    
End Function
Private Function deleteFreq(needDeletefreq As String, inputMaFreqs As String, isRFHop As Boolean) As String
    Dim freqs() As String
    freqs = Split(inputMaFreqs, ",")
    Dim index As Long
    Dim onefreq As String
    
    deleteFreq = ""
    
    For index = LBound(freqs) To UBound(freqs)
        onefreq = freqs(index)
        Dim trimdFreq As String
        trimdFreq = onefreq
        If InStr(onefreq, "(") <> 0 Then
            Dim deliStr As String
            If InStr(onefreq, ":") <> 0 Then
                deliStr = ":"
            Else
                deliStr = ")"
            End If
            trimdFreq = Right(trimdFreq, Len(trimdFreq) - 1)
            trimdFreq = Left(trimdFreq, InStr(trimdFreq, deliStr) - 1)
            If trimdFreq <> needDeletefreq Then
                deleteFreq = deleteFreq + onefreq + ","
            Else
                If isRFHop Then
                    deleteFreq = deleteFreq + trimdFreq + ","
                End If
            End If
        Else
            If onefreq <> needDeletefreq Or isRFHop Then
                deleteFreq = deleteFreq + onefreq + ","
            End If
            
        End If
    Next
    If InStr(deleteFreq, ",") <> 0 Then
        deleteFreq = Left(deleteFreq, Len(deleteFreq) - 1)
    End If
End Function

Private Function changeAttrByMoc(mocName As String)
    Dim m_colNum As Long
    Dim m_rowNum As Long
    Dim attrName As String
    Dim mappingDef As Worksheet
    Dim ws As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For m_rowNum = 2 To mappingDef.range("a1048576").End(xlUp).row
        If UCase(CellShtName) = UCase(mappingDef.Cells(m_rowNum, 1).value) _
            And UCase(mocName) = UCase(mappingDef.Cells(m_rowNum, 4).value) Then
            attrName = mappingDef.Cells(m_rowNum, 5).value
            If GetDesStr(attrName) <> "CELLNAME" And GetDesStr(attrName) <> "BTSNAME" Then
                m_colNum = getColNum(CellShtName, 2, attrName, mocName)
                changeAttr (m_colNum)
            End If
            
        End If
    Next
End Function

Private Function changeAttr(colomn As Long)
    Dim ws As Worksheet
    Dim attrValue As String
    
    Set ws = ThisWorkbook.Worksheets(CellShtName)
    attrValue = ws.Cells(rowNumber, colomn).value
    
    ws.Cells(rowNumber, colomn).value = deleteIndex(attrValue)

End Function
Private Function deleteIndex(attrValue As String) As String
    Dim strArray() As String
    Dim iNo As Long
    Dim tCount As Long
    deleteIndex = ""
    If Trim(attrValue) = "" Then
        deleteIndex = attrValue
        Exit Function
    End If
    strArray = Split(attrValue, ",")
    tCount = UBound(strArray)
    If tCount = 0 Then
        If hasTrx Then
            deleteIndex = attrValue
            Exit Function
        Else
            deleteIndex = ""
            Exit Function
        End If
    Else
        For iNo = LBound(strArray) To tCount
            If iNo = NeedDeleteIndex Then
                strArray(iNo) = ""
            Else
                deleteIndex = deleteIndex + strArray(iNo) + ","
            End If
        Next
    End If
    
    deleteIndex = shrinkStr(Left(deleteIndex, Len(deleteIndex) - 1), ",")
    
End Function

Private Function shrinkStr(inputStr As String, deliStr As String) As String
        Dim fmtStr As String
        Dim appendStr As String
        
        fmtStr = inputStr
        If "]" = deliStr Then
            appendStr = "]"
            If Right(inputStr, 1) = "]" Then
                fmtStr = Left(inputStr, Len(inputStr) - 1)
            End If
        End If
        
        
        Dim strArray() As String
        strArray = Split(fmtStr, deliStr)

        Dim iNo As Long
        Dim tmpStr As String
        
        For iNo = LBound(strArray) To UBound(strArray)
            If (0 = iNo) Then
                tmpStr = strArray(iNo)
            Else
               If Not (tmpStr = strArray(iNo)) Then
                    shrinkStr = inputStr
                    Exit Function
               End If
            End If
        Next
        
        shrinkStr = tmpStr + appendStr
End Function


Private Function changeTrxBind()
    Dim brdNoIndex As Long
    Dim brdNoStr As String
    Dim trxPnIndex As Long
    Dim trxPnStr As String
    Dim antPassNoIndex As Long
    Dim antPassNoStr As String
    Dim antTennaGroupIdIndex As Long
    Dim antTennaGroupIdStr As String
    Dim rruFlag As Boolean
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CellShtName)
    brdNoIndex = getColNum(CellShtName, 2, ATTR_BRDNO, MOC_TRXBIND2PHYBRD)
    trxPnIndex = getColNum(CellShtName, 2, ATTR_TRXPN, MOC_TRXBIND2PHYBRD)
    antPassNoIndex = getColNum(CellShtName, 2, ATTR_ANTPASSNO, MOC_TRXBIND2PHYBRD)
    antTennaGroupIdIndex = getColNum(CellShtName, 2, ATTR_ANTENNAGROUPID, MOC_TRXBIND2PHYBRD)
    
    brdNoStr = ws.Cells(rowNumber, brdNoIndex)
    trxPnStr = ws.Cells(rowNumber, trxPnIndex)
    antPassNoStr = ws.Cells(rowNumber, antPassNoIndex)
    antTennaGroupIdStr = ws.Cells(rowNumber, antTennaGroupIdIndex)
    
    rruFlag = False
    If InStr(brdNoStr, "[") Then
        rruFlag = True
    End If
    
    If rruFlag Then
        brdNoStr = rruDelInd(brdNoStr)
        trxPnStr = rruDelInd(trxPnStr)
        antPassNoStr = rruDelInd(antPassNoStr)
        antTennaGroupIdStr = rruDelInd(antTennaGroupIdStr)
    Else
        brdNoStr = delInd(brdNoStr)
        trxPnStr = delInd(trxPnStr)
        antPassNoStr = delInd(antPassNoStr)
        antTennaGroupIdStr = delInd(antTennaGroupIdStr)
    End If
    ws.Cells(rowNumber, brdNoIndex) = brdNoStr
    ws.Cells(rowNumber, trxPnIndex) = trxPnStr
    ws.Cells(rowNumber, antPassNoIndex) = antPassNoStr
    ws.Cells(rowNumber, antTennaGroupIdIndex) = antTennaGroupIdStr
    
End Function
Private Function rruDelInd(inputStr As String) As String
    rruDelInd = ""
    Dim strArray() As String
    strArray = Split(inputStr, "]")
    If UBound(strArray) = 0 Then
        If hasTrx Then
            rruDelInd = inputStr
        Else
            rruDelInd = ""
        End If
    Else
        Dim index As Long
        For index = LBound(strArray) To UBound(strArray)
            If index <> NeedDeleteIndex And Trim(strArray(index)) <> "" Then
                rruDelInd = rruDelInd + strArray(index) + "]"
            End If
        Next
    End If
    rruDelInd = shrinkStr(rruDelInd, "]")
End Function
Private Function delInd(inputStr As String) As String
    Dim strArray() As String
    strArray = Split(inputStr, ",")
    delInd = ""
    If InStr(inputStr, ",") = 0 Then
        If hasTrx Then
            delInd = inputStr
        Else
            delInd = ""
        End If
    Else
        Dim index As Long
        For index = LBound(strArray) To UBound(strArray)
            If index <> NeedDeleteIndex Then
                delInd = delInd + strArray(index) + ","
            End If
        Next
    End If
    If InStr(delInd, ",") <> 0 Then
        delInd = Left(delInd, Len(delInd) - 1)
    End If
End Function
Public Function changeFreqs(ByRef freqs As String, trxNumStr As String, cellBand As String) As String
    Dim trxNumArr() As String
    Dim lBandNum As Long
    Dim uBandNum As Long
    Dim lbandIndex As Long
    Dim ubandIndex As Long
    Dim freq As String
    Dim band As String
    Dim newFreqs As String
    
    trxNumArr = Split(trxNumStr, ",")
    lBandNum = trxNumArr(0)
    uBandNum = trxNumArr(1)
    
    Dim freqsArr() As String
    freqsArr = Split(freqs, ",")
    
    Dim freqIndex As Long
    For freqIndex = LBound(freqsArr) To UBound(freqsArr)
        freq = freqsArr(freqIndex)
        band = getFreqBand(freq, cellBand)
        If (band = CELLBAND_850 Or band = CELLBAND_900) And (lbandIndex < lBandNum) Then
            newFreqs = newFreqs + freq + ","
            lbandIndex = lbandIndex + 1
        ElseIf (band = CELLBAND_1800 Or band = CELLBAND_1900) And (ubandIndex < uBandNum) Then
            newFreqs = newFreqs + freq + ","
            ubandIndex = ubandIndex + 1
        End If
    Next
    freqs = newFreqs
End Function
Public Function getFreqBand(freq As String, cellBand As String) As String
    If (freq >= 0 And freq <= 124) Or (freq >= 955 And freq <= 1023) Then
        getFreqBand = CELLBAND_900
    ElseIf (freq >= 128 And freq <= 251) Then
        getFreqBand = CELLBAND_850
    ElseIf (freq >= 512 And freq <= 885) And (InStr(cellBand, "1800") <> 0) Then
        getFreqBand = CELLBAND_1800
    ElseIf (freq >= 512 And freq <= 810) And (InStr(cellBand, "1900") <> 0) Then
        getFreqBand = CELLBAND_1900
    End If
End Function
Public Function getGcellBTSNameCol(shtname As String) As Long
    getGcellBTSNameCol = getColNum(shtname, 2, "BTSNAME", "GCELL")
End Function

Public Function getTransBTSNameCol(shtname As String) As Long
    getTransBTSNameCol = getColNum(shtname, 2, "BTSNAME", "BTS")
End Function
