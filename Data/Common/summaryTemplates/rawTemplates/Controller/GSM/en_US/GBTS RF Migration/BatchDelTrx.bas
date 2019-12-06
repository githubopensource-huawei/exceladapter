Attribute VB_Name = "BatchDelTrx"

Public FileNamesList As String
Private srcWorkbook As Workbook
Private LogFile As String
Private IsHint As Boolean
Private REMOVESUCCESS As Boolean
'当前页签小区名称集合
Private GCellName_Map As CMap


' 导入文件
Public Function ImportDataFile() As Boolean
    Dim i As Integer
    Dim FD As FileDialog
    Dim FDFilter As FileDialogFilter
    Dim FDFilters As FileDialogFilters
    Dim FileName As String
    
    ImportDataFile = True
    Set FD = Application.FileDialog(msoFileDialogOpen)
    Set FDFilters = FD.Filters
    FDFilters.Clear
    Set FDFilter = FDFilters.Add("Template Files", "*.csv")
    Set FDFilter = FDFilters.Add("Template Files", "*.xls")                                 '只可以选择xls类型文件
    
    With FD
        .AllowMultiSelect = False
        .Show
        If (.SelectedItems.count = 0) Then
            Exit Function
        Else
            FileNamesList = .SelectedItems(1)
            FileName = GetInputFileName(.SelectedItems(1))
            If VBA.Right(FileName, 4) <> ".xls" And VBA.Right(FileName, 4) <> ".csv" Then
                ImportDataFile = False
                Exit Function
            End If
        End If
    End With
End Function

Private Function GetInputFileName(ByVal FilePath As String) As String
    Dim Path As Integer
    
    Path = InStrRev(FilePath, "\", , vbTextCompare)
    GetInputFileName = VBA.Right(FilePath, Len(FilePath) - Path)
End Function
Private Sub initLogFile()
    Dim NowDate As String
    Dim FileName As String
    NowDate = Format(Now, "yyyy-mm-dd hh#mm#ss")
    FileName = "LogFile_" + NowDate + ".txt"
    If Dir(ThisWorkbook.Path + "\log", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path + "\log"
    End If
    
    LogFile = ThisWorkbook.Path & "\log" & "\" & FileName
    Call EndLogFile
    Open LogFile For Output As #1
End Sub
Private Sub WriteLogFile(Buffer As String)
    currentTime = "[" & Now & "]"
    Print #1, currentTime & Buffer
End Sub
Private Sub WriteLineLogFile(Buffer As String)
    Print #1, Buffer
End Sub
Private Sub EndLogFile()
    Close #1
End Sub

Public Sub BatchDelTrxMain(FilePath As String)
    Dim customFreqCollection As Collection
    Dim controllerAndBtsRelCollection As Collection
    Dim controllerNameString As String
    Dim notExistControllerName As String
    Dim controlNameArray() As String
    Dim index As Integer
    Dim notExistCellName As String
    Dim SourceFileName As String
    
    '获取当前页签小区名称
    Set GCellName_Map = New CMap
    Call getCurGCellNameMap
    
    '根据基站传输页，获取控制器->基站对应关系集合
    Call prepareControllerAndBtsRelation(controllerAndBtsRelCollection, controllerNameString)
    
   '文件判断
    If Trim(FilePath) = "" Then
        MsgBox getResByKey("PlsSelectFile")
        Exit Sub
    End If
    
    SourceFileName = GetInputFileName(FilePath)
    
    If WorkBookExists(SourceFileName) = True Then
        MsgBox SourceFileName & " " & getResByKey("PlsCloseFile")
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set srcWorkbook = Application.Workbooks.Open(FilePath)
    
    FileName = GetInputFileName(FilePath)
    If VBA.Right(FileName, 4) = ".xls" Then
    If WorksheetExists(srcWorkbook, getResByKey("SheetName")) = False Then
        MsgBox getResByKey("SheetNotExist")
        srcWorkbook.Close False
        Exit Sub
        End If
    End If
    Application.ScreenUpdating = True
    
     '记录日志
    Call initLogFile
    WriteLogFile ("Start to delete GTRX ...")
    
     '准备用户数据
    Call prepareUserDataCollectionForFreq(FilePath, customFreqCollection, controllerAndBtsRelCollection, notExistControllerName, notExistCellName)
    
    If customFreqCollection.count = 0 Then
        Exit Sub
    End If
    
    
    If notExistControllerName <> "" Then
        Call WriteLogFile("Error:" + getResByKey("NotExistControllerName"))
        'Call WriteLineLogFile("BSCName")
        Call WriteLineLogFile("RowNo,BSCName")
        Call WriteLineLogFile(notExistControllerName)
        WriteLogFile ("Delete GTRX end...")
        Call EndLogFile
        MsgBox getResByKey("deleteTrxFailed") & getResByKey("NotExistControllerNameFile") & vbCrLf & LogFile
        Exit Sub
    End If
    
    If notExistCellName <> "" Then
        Call WriteLogFile("Error:" + getResByKey("NotExistGCELL"))
        'Call WriteLineLogFile("BSCName,CellName")
        Call WriteLineLogFile("RowNo,BSCName,CellName")
        Call WriteLineLogFile(notExistCellName)
        WriteLogFile ("Delete GTRX end...")
        Call EndLogFile
        MsgBox getResByKey("deleteTrxFailed") & getResByKey("NotExistGCELLFile") & vbCrLf & LogFile
        Exit Sub
    End If
    
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    
    '关闭Excel功能提升效率
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '处理多控制器场景
    controlNameArray = Split(controllerNameString, ",")
    
    IsHint = False
    
    For index = LBound(controlNameArray) To UBound(controlNameArray)
        If Contains(customFreqCollection, controlNameArray(index)) Then
            '删除频点
            Call batchDeleteFreq(customFreqCollection, controlNameArray(index))
        End If
    Next
    
    WriteLogFile ("Delete GTRX end...")
    Call EndLogFile
    
    If IsHint Then
        MsgBox getResByKey("deleteTrxSuccess") & getResByKey("NotExistFreqFile") & vbCrLf & LogFile
    Else
        Call MsgBox(getResByKey("deleteTrxSuccess"), vbInformation, getResByKey("success"))
    End If
    
    '恢复Excel功能
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState

End Sub


Private Sub getCurGCellNameMap()
    Dim rowNum As Long
    Dim mocName As String
    Dim attrName As String
    Dim constCellTempCol As Long
    
    Call getGCellMocNameAndAttrName(mocName, attrName)
    constCellTempCol = getColNum(ThisWorkbook.ActiveSheet.name, 2, attrName, mocName)
    
    rowNum = ThisWorkbook.ActiveSheet.range(getColumnNameFromColumnNum(constCellTempCol) + CStr(1048576)).End(xlUp).row
    
    Dim index As Long
    Dim gcellName As String
    For index = 3 To rowNum
        gcellName = ThisWorkbook.ActiveSheet.Cells(index, constCellTempCol)
        If gcellName <> "" Then
            Call GCellName_Map.SetAt(gcellName, gcellName)
        End If
    Next
End Sub

Private Function getGCellMocNameAndAttrName(mocName As String, attrName As String)
    attrName = "CELLNAME"
    mocName = "GCELL"
End Function

Private Function changeTrxNum(freq As String, rowNum As Integer)
    Dim trxNumber As String
    Dim cellBand As String
    Dim trxNumIndex As Long
    Dim cellBandIndex As Long
    Dim trxNumArray() As String

    trxNumber = ThisWorkbook.ActiveSheet.Cells(rowNum, getTrxNumCol(ThisWorkbook.ActiveSheet.name))
    cellBand = ThisWorkbook.ActiveSheet.Cells(rowNum, getCellBandCol(ThisWorkbook.ActiveSheet.name))
    
    trxNumArray = Split(trxNumber, ",")
    Dim eNo As Long
    eNo = UBound(trxNumArray)
    
    If eNo = 0 Then
        ThisWorkbook.ActiveSheet.Cells(rowNum, getTrxNumCol(ThisWorkbook.ActiveSheet.name)).value = trxNumber - 1
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
        
        ThisWorkbook.ActiveSheet.Cells(rowNum, getTrxNumCol(ThisWorkbook.ActiveSheet.name)).value = CStr(lBandNum) + "," + CStr(uBandNum)
    End If
End Function

Private Function deleteTrxChildMo(rowNum As Integer, needDelIndexCollection As Collection)
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
        Call changeAttrByMoc(mocName, rowNum, needDelIndexCollection)
    Next
End Function

Private Sub deleteIstmptrx(rowNum As Integer, needDelIndexCollection As Collection)
    Dim m_colNum As Long
    Dim CellShtName As String
    CellShtName = ThisWorkbook.ActiveSheet.name
    '先获取ISTMPTRX所在列号
    m_colNum = getColNum(CellShtName, 2, ATTR_ISTMPTRX, MOC_TRXINFO)
    If m_colNum > 0 Then
        Call changeAttr(CellShtName, m_colNum, rowNum, needDelIndexCollection)
    End If
    
End Sub

Private Function changeAttrByMoc(ByVal mocName As String, rowNum As Integer, needDelIndexCollection As Collection)
    Dim m_colNum As Long
    Dim m_rowNum As Long
    Dim attrName As String
    Dim mappingDef As Worksheet
    Dim ws As Worksheet
    Dim CellShtName As String
    CellShtName = ThisWorkbook.ActiveSheet.name
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For m_rowNum = 2 To mappingDef.range("a1048576").End(xlUp).row
        If UCase(CellShtName) = UCase(mappingDef.Cells(m_rowNum, 1).value) _
            And UCase(mocName) = UCase(mappingDef.Cells(m_rowNum, 4).value) Then
            attrName = mappingDef.Cells(m_rowNum, 5).value
            If GetDesStr(attrName) <> "CELLNAME" And GetDesStr(attrName) <> "BTSNAME" Then
                m_colNum = getColNum(CellShtName, 2, attrName, mocName)
                Call changeAttr(CellShtName, m_colNum, rowNum, needDelIndexCollection)
            End If
            
        End If
    Next
End Function

Private Function changeAttr(CellShtName As String, colomn As Long, rowNum As Integer, needDelIndexCollection As Collection)
    Dim ws As Worksheet
    Dim attrValue As String
    Dim NeedDeleteIndex As Variant
    
    Set ws = ThisWorkbook.Worksheets(CellShtName)
    
    attrValue = ws.Cells(rowNum, colomn).value
    If needDelIndexCollection.count > 0 Then
        ws.Cells(rowNum, colomn).value = deleteIndex(attrValue, needDelIndexCollection, rowNum)
    End If

End Function

'判断小区页签是否存在频点
Private Function isTrxExist(rowNum As Integer) As Boolean
    
    Dim trxNumStr As String
    isTrxExist = True
    
    trxNumStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getTrxNumCol(ThisWorkbook.ActiveSheet.name))
    
    If (Trim(trxNumStr) = "") Or (Trim(trxNumStr) = "0") Or (Trim(trxNumStr) = "0,0") Then
        isTrxExist = False
    End If
End Function


Private Function deleteIndex(attrValue As String, needDelIndexCollection As Collection, rowNum As Integer) As String
    Dim strArray() As String
    Dim iNo As Long
    Dim tCount As Long
    Dim NeedDeleteIndex As Variant
    
    deleteIndex = ""
    If Trim(attrValue) = "" Then
        deleteIndex = attrValue
        Exit Function
    End If
    strArray = Split(attrValue, ",")
    tCount = UBound(strArray)
    If tCount = 0 Then
        If isTrxExist(rowNum) Then
            deleteIndex = attrValue
            Exit Function
        Else
            deleteIndex = ""
            Exit Function
        End If
    Else
        For Each NeedDeleteIndex In needDelIndexCollection
            If NeedDeleteIndex <> "" Then
                For iNo = LBound(strArray) To tCount
                    If iNo = NeedDeleteIndex Then
                        strArray(iNo) = ""
                    End If
                Next
            End If
        Next
        For index = LBound(strArray) To UBound(strArray)
            If strArray(index) <> "" Then
                If deleteIndex = "" Then
                    deleteIndex = strArray(index)
                Else
                    deleteIndex = deleteIndex + "," + strArray(index)
                End If
            End If
        Next
    End If
    
End Function

Private Function changeHopType(rowNum As Integer)
    Dim trxNumStr As String
    Dim curNotBcchFreq As String
    
    curNotBcchFreq = ThisWorkbook.ActiveSheet.Cells(rowNum, getNonBcchCol(ThisWorkbook.ActiveSheet.name)).value
    trxNumStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getTrxNumCol(ThisWorkbook.ActiveSheet.name))
    
    If Trim(curNotBcchFreq) = "" Or CLng(trxNumStr) = 0 Or (Not isTrxExist(rowNum)) Then
        ThisWorkbook.ActiveSheet.Cells(rowNum, getHopModeCol(ThisWorkbook.ActiveSheet.name)) = "NO_FH"
    End If
    
End Function

Private Sub changeMaGrpListMain(rowNum As Integer, curActualDelFreqCollection As Collection)
    Dim needDeletefreq As Variant
    For Each needDeletefreq In curActualDelFreqCollection
        Call changeMaGrpList(rowNum, needDeletefreq)
    Next
End Sub

Private Function changeMaGrpList(rowNum As Integer, ByVal needDeletefreq As String)
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
    
    
    newMaGrpStr = ""
    newHsn = ""
    
    maGrpStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getMaGrpListCol(ThisWorkbook.ActiveSheet.name))
    hsnStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getHsnCol(ThisWorkbook.ActiveSheet.name))
    hopModeStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getHopModeCol(ThisWorkbook.ActiveSheet.name))
    
    If Trim(hopModeStr) = "NO_FH" Then
        ThisWorkbook.ActiveSheet.Cells(rowNum, getHsnCol(ThisWorkbook.ActiveSheet.name)) = ""
        ThisWorkbook.ActiveSheet.Cells(rowNum, getMaGrpListCol(ThisWorkbook.ActiveSheet.name)) = ""
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
                
                '如果是基带跳频场景需要再调整偏移量
                If Not isRFHop Then
                    processFreqs = adjustBasebandFH(processFreqs)
                End If
                
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
    
        ThisWorkbook.ActiveSheet.Cells(rowNum, getMaGrpListCol(ThisWorkbook.ActiveSheet.name)) = newMaGrpStr
    
        If Trim(newMaGrpStr) <> "" Then
            ThisWorkbook.ActiveSheet.Cells(rowNum, getHsnCol(ThisWorkbook.ActiveSheet.name)) = Left(newHsn, Len(newHsn) - 1)
        Else
            ThisWorkbook.ActiveSheet.Cells(rowNum, getHsnCol(ThisWorkbook.ActiveSheet.name)) = ""
        End If
    End If
    
End Function

Private Function changeTrxBind(rowNum As Integer, needDelIndexCollection As Collection)
    On Error GoTo ErrorHandler
    Dim brdNoStr As String
    Dim trxPnStr As String
    Dim antPassNoStr As String
    Dim antTennaGroupIdStr As String
    Dim rruFlag As Boolean
        
    If needDelIndexCollection.count <= 0 Then Exit Function
    
    brdNoStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getBrdNoCol(ThisWorkbook.ActiveSheet.name))
    trxPnStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getTrxPnCol(ThisWorkbook.ActiveSheet.name))
    antPassNoStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getAntPassNoCol(ThisWorkbook.ActiveSheet.name))
    antTennaGroupIdStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getAntTennaGroupIdCol(ThisWorkbook.ActiveSheet.name))
    
    
'    rruFlag = False
'    If InStr(brdNoStr, "[") Then
'        rruFlag = True
'    End If
    
    If isMuliBrdBindScenario(brdNoStr, trxPnStr) Then
        brdNoStr = MuliBrdBindDelInd(rowNum, brdNoStr, needDelIndexCollection)
        trxPnStr = MuliBrdBindDelInd(rowNum, trxPnStr, needDelIndexCollection)
        antPassNoStr = MuliBrdBindDelInd(rowNum, antPassNoStr, needDelIndexCollection)
        antTennaGroupIdStr = MuliBrdBindDelInd(rowNum, antTennaGroupIdStr, needDelIndexCollection)
    ElseIf isRRUScenario(brdNoStr) Then
        brdNoStr = rruDelInd(rowNum, brdNoStr, needDelIndexCollection)
        trxPnStr = rruDelInd(rowNum, trxPnStr, needDelIndexCollection)
        antPassNoStr = rruDelInd(rowNum, antPassNoStr, needDelIndexCollection)
        antTennaGroupIdStr = rruDelInd(rowNum, antTennaGroupIdStr, needDelIndexCollection)
    Else
        brdNoStr = delInd(rowNum, brdNoStr, needDelIndexCollection)
        trxPnStr = delInd(rowNum, trxPnStr, needDelIndexCollection)
        antPassNoStr = delInd(rowNum, antPassNoStr, needDelIndexCollection)
        antTennaGroupIdStr = delInd(rowNum, antTennaGroupIdStr, needDelIndexCollection)
    End If
    
    ThisWorkbook.ActiveSheet.Cells(rowNum, getBrdNoCol(ThisWorkbook.ActiveSheet.name)) = brdNoStr
    ThisWorkbook.ActiveSheet.Cells(rowNum, getTrxPnCol(ThisWorkbook.ActiveSheet.name)) = trxPnStr
    ThisWorkbook.ActiveSheet.Cells(rowNum, getAntPassNoCol(ThisWorkbook.ActiveSheet.name)) = antPassNoStr
    ThisWorkbook.ActiveSheet.Cells(rowNum, getAntTennaGroupIdCol(ThisWorkbook.ActiveSheet.name)) = antTennaGroupIdStr
    
'    antTennaGroupIdStr = ThisWorkbook.ActiveSheet.Cells(rowNum, getAntTennaGroupIdCol(ThisWorkbook.ActiveSheet.name))
'    If rruFlag Then
'        antTennaGroupIdStr = rruDelInd(rowNum, antTennaGroupIdStr, needDelIndexCollection)
'    Else
'        antTennaGroupIdStr = delInd(rowNum, antTennaGroupIdStr, needDelIndexCollection)
'    End If
'    ThisWorkbook.ActiveSheet.Cells(rowNum, getAntTennaGroupIdCol(ThisWorkbook.ActiveSheet.name)) = antTennaGroupIdStr
    
ErrorHandler:
    REMOVESUCCESS = False
    
End Function
Private Function rruDelInd(rowNum As Integer, inputStr As String, needDelIndexCollection As Collection) As String
    rruDelInd = ""
    Dim strArray() As String
    strArray = Split(inputStr, "]")
    If UBound(strArray) = 0 Then
        If isTrxExist(rowNum) Then
            rruDelInd = inputStr
        Else
            rruDelInd = ""
        End If
    Else
        Dim index As Long
        Dim NeedDeleteIndex As Variant
        For Each NeedDeleteIndex In needDelIndexCollection
            If NeedDeleteIndex <> "" Then
                For index = LBound(strArray) To UBound(strArray)
                    If index = NeedDeleteIndex And Trim(strArray(index)) <> "" Then
                        strArray(index) = ""
                    End If
                Next
            End If
        Next
        Dim indexArr As Integer
        For indexArr = LBound(strArray) To UBound(strArray)
            If strArray(indexArr) <> "" Then
                If rruDelInd = "" Then
                    rruDelInd = strArray(indexArr)
                Else
                    rruDelInd = rruDelInd + "," + strArray(indexArr)
                End If
            End If
        Next
    End If
    
    rruDelInd = shrinkStr(rruDelInd, "]")

End Function
Private Function delInd(rowNum As Integer, inputStr As String, needDelIndexCollection As Collection) As String
    Dim strArray() As String
    strArray = Split(inputStr, ",")
    delInd = ""
    If InStr(inputStr, ",") = 0 Then
        If isTrxExist(rowNum) Then
            delInd = inputStr
        Else
            delInd = ""
        End If
    Else
        Dim index As Long
        Dim NeedDeleteIndex As Variant
        For Each NeedDeleteIndex In needDelIndexCollection
            If NeedDeleteIndex <> "" Then
                For index = LBound(strArray) To UBound(strArray)
                    If index = NeedDeleteIndex Then
                        strArray(index) = ""
                    End If
                Next
            End If
        Next
        Dim indexArr As Integer
        For indexArr = LBound(strArray) To UBound(strArray)
            If strArray(indexArr) <> "" Then
                If delInd = "" Then
                    delInd = strArray(indexArr)
                Else
                    delInd = delInd + "," + strArray(indexArr)
                End If
            End If
        Next
    End If

End Function

Private Function MuliBrdBindDelInd(rowNum As Integer, inputStr As String, needDelIndexCollection As Collection) As String
    Dim strArray() As String
    
    If inputStr <> "" Then
        '掐头去尾
        inputStr = Left(inputStr, Len(inputStr) - 1)
        inputStr = Right(inputStr, Len(inputStr) - 1)
    End If
    
    strArray = Split(inputStr, "][")
    MuliBrdBindDelInd = ""
    If InStr(inputStr, "][") = 0 Then
        If isTrxExist(rowNum) Then
            MuliBrdBindDelInd = inputStr
        Else
            MuliBrdBindDelInd = ""
        End If
    Else
        Dim index As Long
        Dim NeedDeleteIndex As Variant
        For Each NeedDeleteIndex In needDelIndexCollection
            If NeedDeleteIndex <> "" Then
                For index = LBound(strArray) To UBound(strArray)
                    If index = NeedDeleteIndex Then
                        strArray(index) = ""
                    End If
                Next
            End If
        Next
        Dim indexArr As Integer
        For indexArr = LBound(strArray) To UBound(strArray)
            If strArray(indexArr) <> "" Then
                If MuliBrdBindDelInd = "" Then
                    MuliBrdBindDelInd = strArray(indexArr)
                Else
                    MuliBrdBindDelInd = MuliBrdBindDelInd + "][" + strArray(indexArr)
                End If
            End If
        Next
    End If
    '加头和尾
    MuliBrdBindDelInd = "[" + MuliBrdBindDelInd + "]"
End Function

Private Sub batchDeleteFreq(ByRef customFreqCollection As Collection, ByRef controlName As String)
    Dim maxRow As Integer
    Dim curBtsName As String
    Dim curGCellName As String
    Dim curBcchFreq As String
    Dim curNotBcchFreq As String
    Dim notExistFreq As String
    Dim curNotBcchFreqArray() As String
    Dim curNotBcchFreqCollection As Collection
    Dim index As Integer
    Dim rowNum As Integer
    Dim temNotBcchFreq As Variant
    
    Dim customGCellCollection As Collection
    Dim curFreqCol As Collection
    
    If Contains(customFreqCollection, controlName) Then
        Set customGCellCollection = customFreqCollection(controlName)
    End If
    maxRow = ThisWorkbook.ActiveSheet.range("a1048576").End(xlUp).row
    notExistFreq = ""
    
    For rowNum = 3 To maxRow
        '获取基站和小区名称
        curBtsName = ThisWorkbook.ActiveSheet.Cells(rowNum, getGcellBTSNameCol(ThisWorkbook.ActiveSheet.name)).value
        curGCellName = ThisWorkbook.ActiveSheet.Cells(rowNum, getGcellCellNameCol(ThisWorkbook.ActiveSheet.name)).value
        '获取主B和非主B频点名称，非主B频点转换成集合操作
        curBcchFreq = ThisWorkbook.ActiveSheet.Cells(rowNum, getBcchCol(ThisWorkbook.ActiveSheet.name)).value
        curNotBcchFreq = ThisWorkbook.ActiveSheet.Cells(rowNum, getNonBcchCol(ThisWorkbook.ActiveSheet.name)).value
        curNotBcchFreqArray = Split(curNotBcchFreq, ",")
        
        Set curNotBcchFreqCollection = New Collection
        For index = 0 To UBound(curNotBcchFreqArray)
            If Not Contains(curNotBcchFreqCollection, curNotBcchFreqArray(index)) Then
                curNotBcchFreqCollection.Add Item:=curNotBcchFreqArray(index)
            End If
        Next
        
        '获取每一行实际删除的频点
        Dim curActualDelFreqCollection As Collection
        Set curActualDelFreqCollection = New Collection
        
        If Contains(customGCellCollection, curGCellName) Then
            Set curFreqCol = customGCellCollection(curGCellName)

            Dim freq As String
            'Dim afterDelFreq As String
            Dim indexInCol As Integer
            Dim needDelIndexCollection As Collection
            Set needDelIndexCollection = New Collection
            

            For indexInCol = 1 To curFreqCol.count
                freq = curFreqCol.Item(indexInCol)
                '主B频点只有一个，可以判断后直接删除
                If freq = curBcchFreq Then
                    ThisWorkbook.ActiveSheet.Cells(rowNum, getBcchCol(ThisWorkbook.ActiveSheet.name)).value = ""
                    '删除频点后调整频点个数，即Number of TRXs的值
                    Call changeTrxNum(freq, rowNum)
                    '如果删除了主B频点，则Index为0
                    needDelIndexCollection.Add Item:=0
                    curActualDelFreqCollection.Add Item:=freq
                Else
                    '非主B频点
                    If isInCollection(curNotBcchFreqCollection, freq) = True Then
                        Dim strIndex As Integer
                        
                        '维护数组，保存需要删除的非主B频点位置信息；并且删除主B频点
                        For arrIndex = LBound(curNotBcchFreqArray) To UBound(curNotBcchFreqArray)
                            If curNotBcchFreqArray(arrIndex) = freq Then
                                needDelIndexCollection.Add Item:=arrIndex + 1
                                '删除主B频点
                                curNotBcchFreqArray(arrIndex) = ""
                                curActualDelFreqCollection.Add Item:=freq
                                '修改频点个数
                                Call changeTrxNum(freq, rowNum)
                            End If
                        Next
                    Else
                        '处理用户配置的错误频点，即表格中不存在的频点
                        If notExistFreq = "" Then
                            notExistFreq = controlName + "," + curGCellName + "," + freq + vbCrLf
                        Else
                            notExistFreq = notExistFreq + controlName + "," + curGCellName + "," + freq + vbCrLf
                        End If
                    End If
                End If
            Next
            '处理非主B频点
            Dim temFreq As String
            temFreq = ""
            Dim temIndex As Integer
            For temIndex = LBound(curNotBcchFreqArray) To UBound(curNotBcchFreqArray)
                If curNotBcchFreqArray(temIndex) <> "" Then
                    If temFreq = "" Then
                        temFreq = curNotBcchFreqArray(temIndex)
                    Else
                        temFreq = temFreq + "," + curNotBcchFreqArray(temIndex)
                    End If
                End If
            Next
            '写入对应位置
            ThisWorkbook.ActiveSheet.Cells(rowNum, getNonBcchCol(ThisWorkbook.ActiveSheet.name)) = temFreq
            
            '当前for循环完成之后，频点删除完成，统一删除载频子对象
            Call deleteTrxChildMo(rowNum, needDelIndexCollection)
            '同步删除ISTMPTRX
            Call deleteIstmptrx(rowNum, needDelIndexCollection)
            Call changeHopType(rowNum)
            Call changeMaGrpListMain(rowNum, curActualDelFreqCollection)
            Call changeTrxBind(rowNum, needDelIndexCollection)

        End If
    Next
    
    '提示用户不存在的频点
    If notExistFreq <> "" Then
        IsHint = True
        Call WriteLogFile("Hint:" + getResByKey("deleteTrxSuccess") + getResByKey("NotExistFreq"))
        Call WriteLineLogFile("BSCName,CellName,Freq")
        Call WriteLineLogFile(notExistFreq)
    End If
       
End Sub

'获取控制器与基站名称的对应集合，collection(控制器名称，collection(基站名称))
Private Function prepareControllerAndBtsRelation(ByRef controllerAndBtsRelCollection As Collection, ByRef controllerNameString As String)
    Dim maxRow As Integer
    Dim rowNum As Integer
    Dim controlName As String
    Dim btsName As String
    
    Set controllerAndBtsRelCollection = New Collection
    
    maxRow = ThisWorkbook.Sheets(getResByKey("BaseStationSheet")).range("a1048576").End(xlUp).row
    'maxRow = ThisWorkbook.Sheets(getResByKey("BaseStationSheet")).UsedRange.Rows.count
    
    For rowNum = 3 To maxRow
        controlName = ThisWorkbook.Sheets(getResByKey("BaseStationSheet")).Cells(rowNum, getControlNameCol(ThisWorkbook.Sheets(getResByKey("BaseStationSheet")).name)).value
        btsName = ThisWorkbook.Sheets(getResByKey("BaseStationSheet")).Cells(rowNum, getBTSNameCol(ThisWorkbook.Sheets(getResByKey("BaseStationSheet")).name)).value
        
        If controlName <> "" And btsName <> "" Then
            If Contains(controllerAndBtsRelCollection, controlName) Then
                Dim btsCollection As Collection
                Set btsCollection = controllerAndBtsRelCollection(controlName)
                If Not Contains(btsCollection, btsName) Then
                    btsCollection.Add Item:=btsName
                End If
            Else
                Dim btsNameCol As Collection
                Set btsNameCol = New Collection
                btsNameCol.Add Item:=btsName
                
                controllerAndBtsRelCollection.Add Item:=btsNameCol, key:=controlName
                
                '多控制器场景，获取控制器名称列表
                If controllerNameString = "" Then
                    controllerNameString = controlName
                Else
                    controllerNameString = controllerNameString + "," + controlName
                End If
                
            End If
        End If
    Next
    
End Function
Function WorkBookExists(FileName As String) As Boolean
    Dim str As String
    On Error GoTo WorkBookExistsErr
    Set ws = Workbooks(FileName)
    WorkBookExists = True
    Exit Function
WorkBookExistsErr:
    WorkBookExists = False
End Function

'将表格中读取的数据保存在数据结构collection(控制器名称，collection(小区名称，collection(频点)))
Private Function prepareUserDataCollectionForFreq(ByRef FilePath As String, ByRef customFreqCollection As Collection, controllerAndBtsRelCollection As Collection, notExistControllerName As String, notExistCellName As String)
    Dim maxRow As Long
    Dim srcRowNum As Long
    Dim controlName As String
    Dim cellName As String
    Dim SelectedFreq As String
    

    FileName = GetInputFileName(FilePath)
    If VBA.Right(FileName, 4) = ".xls" Then
        Set srcSht = srcWorkbook.Sheets(getResByKey("SheetName"))
    ElseIf VBA.Right(FileName, 4) = ".csv" Then
        Set srcSht = srcWorkbook.Sheets(1)
    End If
    
    Set customFreqCollection = New Collection
    
    
    maxRow = srcSht.range("a1048576").End(xlUp).row
    
    '不配置控制器名称
    If maxRow < 2 Then
        Call WriteLogFile("Error:" + getResByKey("NotExistControllerName"))
        'Call WriteLineLogFile("BSCName")
        Call WriteLineLogFile("RowNo,BSCName")
        Call WriteLineLogFile("2, ''")
        WriteLogFile ("Delete GTRX end...")
        Call EndLogFile
        MsgBox getResByKey("deleteTrxFailed") & getResByKey("NotExistControllerNameFile") & vbCrLf & LogFile
        srcWorkbook.Close False
        Exit Function
    End If
    
    For srcRowNum = 2 To maxRow
        '获取Batch delete GTRX表格中数据
        controlName = srcSht.Cells(srcRowNum, 1).value
        cellName = srcSht.Cells(srcRowNum, 2).value
        SelectedFreq = srcSht.Cells(srcRowNum, 3).value

        
        If GCellName_Map.hasKey(cellName) = False Then
            If notExistCellName = "" Then
                'notExistCellName = controlName + "," + cellName + vbCrLf
                notExistCellName = str(srcRowNum) + "," + controlName + "," + cellName + vbCrLf
            Else
                'notExistCellName = notExistCellName + controlName + "," + cellName + vbCrLf
                notExistCellName = notExistCellName + str(srcRowNum) + "," + controlName + "," + cellName + vbCrLf
            End If
            
        End If
        SelectedFreqArr = Split(SelectedFreq, ",")
        Dim arrIndex As Long
        
        '判断当前用户配置控制器名称是否存在
        If Contains(controllerAndBtsRelCollection, controlName) Then
            '数据集合
            For arrIndex = LBound(SelectedFreqArr) To UBound(SelectedFreqArr)
                If Contains(customFreqCollection, controlName) Then
                    Dim cellNameCollection As Collection
                    Set cellNameCollection = customFreqCollection(controlName)
                    If Contains(cellNameCollection, cellName) Then
                        cellNameCollection(cellName).Add Item:=SelectedFreqArr(arrIndex)
                    Else
                        Dim selectedfreqCollection As Collection
                        Set selectedfreqCollection = New Collection
                        selectedfreqCollection.Add Item:=SelectedFreqArr(arrIndex)
                
                        cellNameCollection.Add Item:=selectedfreqCollection, key:=cellName
                    End If
                Else
                    Dim selectedfreqCol As Collection
                    Set selectedfreqCol = New Collection
                    selectedfreqCol.Add Item:=SelectedFreqArr(arrIndex)
            
                    Dim cellNameCol As Collection
                    Set cellNameCol = New Collection
                    cellNameCol.Add Item:=selectedfreqCol, key:=cellName
            
                    customFreqCollection.Add Item:=cellNameCol, key:=controlName
            
                End If
            Next
        Else
            '当前用户配置的控制器名称不存在
            If notExistControllerName = "" Then
                'notExistControllerName = controlName + vbCrLf
                notExistControllerName = str(srcRowNum) + "," + controlName + vbCrLf
            Else
                'notExistControllerName = notExistControllerName + controlName + vbCrLf
                notExistControllerName = notExistControllerName + str(srcRowNum) + "," + controlName + vbCrLf
            End If
        End If
        
    Next
    
    srcWorkbook.Close False
End Function

Private Function getGcellCellNameCol(shtname As String) As Long
    getGcellCellNameCol = getColNum(shtname, 2, "CELLNAME", "GCELL")
End Function
Private Function getBcchCol(shtname As String) As Long
    getBcchCol = getColNum(shtname, 2, "BCCHFREQ", "TRXINFO")
End Function
Private Function getNonBcchCol(shtname As String) As Long
    getNonBcchCol = getColNum(shtname, 2, "NONBCCHFREQLIST", "TRXINFO")
End Function
Private Function getTrxNumCol(shtname As String) As Long
    getTrxNumCol = getColNum(shtname, 2, "TRXNUM", "TRXINFO")
End Function
Private Function getHopModeCol(shtname As String) As Long
    getHopModeCol = getColNum(shtname, 2, "HOPMODE", "GCELLMAGRP")
End Function
Private Function getBrdNoCol(shtname As String) As Long
    getBrdNoCol = getColNum(shtname, 2, "BRDNO", "TRXBIND2PHYBRD")
End Function
Private Function getTrxPnCol(shtname As String) As Long
    getTrxPnCol = getColNum(shtname, 2, "TRXPN", "TRXBIND2PHYBRD")
End Function
Private Function getAntPassNoCol(shtname As String) As Long
    getAntPassNoCol = getColNum(shtname, 2, "ANTPASSNO", "TRXBIND2PHYBRD")
End Function
Private Function getAntTennaGroupIdCol(shtname As String) As Long
    getAntTennaGroupIdCol = getColNum(shtname, 2, "ANTENNAGROUPID", "TRXBIND2PHYBRD")
End Function
Private Function getMaGrpListCol(shtname As String) As Long
    getMaGrpListCol = getColNum(shtname, 2, "MAGRPFREQLIST", "TRXINFO")
End Function
Private Function getHsnCol(shtname As String) As Long
    getHsnCol = getColNum(shtname, 2, "HSN", "GCELLMAGRP")
End Function
Private Function getCellBandCol(shtname As String) As Long
    getCellBandCol = getColNum(shtname, 2, "TYPE", "GCELL")
End Function
Private Function getControlNameCol(shtname As String) As Long
    getControlNameCol = getColNum(shtname, 2, "BSCName", "BTS")
End Function
Private Function getBTSNameCol(shtname As String) As Long
    getBTSNameCol = getColNum(shtname, 2, "BTSNAME", "BTS")
End Function
Public Function WorksheetExists(wb As Workbook, strName As String) As Boolean
    Dim str As String
    On Error GoTo worksheetExistsErr
    str = wb.Worksheets(strName).name
    WorksheetExists = True
    Exit Function
worksheetExistsErr:
    WorksheetExists = False
End Function
