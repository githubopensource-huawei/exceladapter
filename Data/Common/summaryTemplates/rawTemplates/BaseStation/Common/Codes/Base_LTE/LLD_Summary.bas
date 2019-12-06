Attribute VB_Name = "LLD_Summary"
Option Explicit
Public lldModelFlag As String
Public hiddenShts As Collection

Function getLLDModelFlag() As String
    getLLDModelFlag = lldModelFlag
End Function

Private Sub collectHiddenShts()
    On Error GoTo ErrorHandler
    Set hiddenShts = Nothing
    Set hiddenShts = New Collection
    Dim sht
    For Each sht In ThisWorkbook.Sheets
        If sht.Visible = xlSheetHidden Then hiddenShts.Add item:=sht.name, key:=sht.name
    Next
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in collectHiddenShts, " & Err.Description
    Resume Next
End Sub

'LLD<-->Summary
Sub Summary2LLD()
    On Error Resume Next
    Dim sheetNum, i As Long
    Dim sht As Worksheet
    Dim sheetArray As Variant
    
    Application.ScreenUpdating = False
    DisplayMessageOnStatusbar
    
    lldModelFlag = "SUMMARY"
    If containsASheet(ThisWorkbook, getResByKey("COMMON")) Then
        lldModelFlag = "LLD"
    End If
    
    If lldModelFlag = "SUMMARY" Then Call collectHiddenShts
    
    '如存在则进行LLD->Summary的处理
    If lldModelFlag = "LLD" Then
        '显示Summary的必要sheet页
        For Each sht In ActiveWorkbook.Worksheets
            If sht.Visible <> 2 And Not Contains(hiddenShts, sht.name) Then
                sht.Visible = True
                sht.Cells.EntireColumn.Hidden = False
            End If
        Next
        
        '填写[COMMON]页采集到的信息
        setSummaryInfo
        'DSCP映射到DIFPRI
        setQOSInfo
        'DSCP和VLAN Priority映射到VlanClass
        setVlanClass
        
        '删除[COMMON]
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(getResByKey("COMMON")).Delete
        Call selectCertainSheet(ThisWorkbook, getResByKey("Comm Data"))
        
        Application.DisplayAlerts = True
        lldModelFlag = ""
        Call AddLink
    '否则进行Summary->LLD的处理
    ElseIf (lldModelFlag = "SUMMARY") And (NeedChange() <> "") Then
        '创建[COMMON]
        Sheets(getResByKey("Comm Data")).Copy Before:=Sheets(getResByKey("Comm Data"))
        Sheets(getResByKey("Comm Data") + " (2)").name = getResByKey("COMMON")
        
        '删除[COMMON]中的非LLD信息，隐藏传输和无线的非LLD信息
        setLLDInfo
        '复制QOS表格
        copyQOSInfo
        
        For Each sht In ActiveWorkbook.Worksheets
            If sht.Visible <> 2 And GetDesStr(sht.name) <> GetDesStr(getResByKey("Cover")) Then
                Sheets(sht.name).Visible = False
            End If
        Next
        
        sheetArray = Split(NeedChange(), ";")
        For i = 1 To UBound(sheetArray)
            If sheetArray(i) = getResByKey("Comm Data") Then
                sheetArray(i) = getResByKey("COMMON")
            End If
            Sheets(sheetArray(i)).Visible = True
        Next
        Sheets(getResByKey("COMMON")).Visible = True
        
        lldModelFlag = ""
        
        Call clearGrayRange(ThisWorkbook.Worksheets(getResByKey("COMMON"))) 'COMMON页无分支控制，清空这个页签的灰化单元格
        Call selectCertainSheet(ThisWorkbook, getResByKey("COMMON"))
    End If
    
    lldModelFlag = ""
    ReturnStatusbaring
    Application.ScreenUpdating = True
End Sub

Public Sub selectCertainCell(ByVal ws As Worksheet, ByVal rangeName As String)
    ws.Activate
    Application.GoTo Reference:=ws.range(rangeName), Scroll:=True
End Sub

Public Sub selectCertainSheet(ByVal wb As Workbook, ByVal sheetName As String)
    wb.Worksheets(sheetName).Activate
    Call selectCertainCell(wb.Worksheets(sheetName), "A1")
End Sub

Private Sub clearGrayRange(ByRef ws As Worksheet)
    Dim maxRowNumber As Long, maxColumnNumber As Long, rowNumber As Long, columnNumber As Long
    maxRowNumber = ws.range("A65535").End(xlUp).row
    For rowNumber = 1 To maxRowNumber
        maxColumnNumber = ws.range("IV" & rowNumber).End(xlToLeft).column
        For columnNumber = 1 To maxColumnNumber
            Call clearGrayCell(ws, rowNumber, columnNumber)
        Next columnNumber
    Next rowNumber
End Sub

Private Sub clearGrayCell(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef columnNumber As Long)
    Dim cellRange As range
    Set cellRange = ws.Cells(rowNumber, columnNumber)
    If cellIsGray(cellRange) Then
        Call setRangeNormal(cellRange)
    End If
End Sub

'Summary->LLD，删除[COMMON]中的非LLD信息，隐藏传输和无线的非LLD信息
Sub setLLDInfo()
    On Error Resume Next
    Dim sheetName As String  '[MAPPING DEF]内sheet名
    Dim groupName As String  '[MAPPING DEF]内group名
    Dim colName As String    '[MAPPING DEF]内列名
    Dim lldFlag As String    '[MAPPING DEF]内[LLD]列的值
    
    Dim m_rowNum
    Dim readRowNum As Long
    Dim rowCount As Long
    Dim readColNum As Long
    Dim readColEnd As Long
       
    Dim mainSheetName  As String, lteCellSheetName As String
    Dim mappingDef As Worksheet, mainSheet As Worksheet, lteCellSheet As Worksheet
    
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    
    mainSheetName = GetMainSheetName() '得到传输页名称
    Set mainSheet = ThisWorkbook.Worksheets(mainSheetName)
    
    lteCellSheetName = getLteSheetName() '得到LTE小区页签名称
    Set lteCellSheet = ThisWorkbook.Worksheets(lteCellSheetName)
    
    Dim findGroupParaNameType As Long
    Dim dstGroupStartColumn As Long, dstGroupEndColumn As Long, dstParaColumn As Long
    
    Dim newLldFlag As String
    
    Application.DisplayAlerts = False
    '遍历『MAPPING DEF』
    For m_rowNum = 2 To mappingDef.range("a65536").End(xlUp).row
        sheetName = mappingDef.Cells(m_rowNum, 1).value
        groupName = mappingDef.Cells(m_rowNum, 2).value
        colName = mappingDef.Cells(m_rowNum, 3).value
        newLldFlag = mappingDef.Cells(m_rowNum, 10).value
        
        If GetDesStr(sheetName) = GetDesStr(mainSheetName) Then
            'lldFlag = GetLldFlag(sheetName, groupName, colName)
            'LLD列的值非True时，隐藏
            If GetDesStr(newLldFlag) <> GetDesStr("TRUE") Then
                'readColNum = Get_Col(sheetName, 2, colName)
                findGroupParaNameType = findGroupParaNameTypeInSheet(mainSheet, groupName, colName, dstGroupStartColumn, dstGroupEndColumn, dstParaColumn)
                If findGroupParaNameType = 1 Then
                    mainSheet.Cells(2, dstParaColumn).EntireColumn.Hidden = True
                End If
            End If
        ElseIf GetDesStr(sheetName) = GetDesStr(lteCellSheetName) Then
            'LLD列的值非True时，隐藏
            If GetDesStr(newLldFlag) <> GetDesStr("TRUE") Then
                findGroupParaNameType = findGroupParaNameTypeInSheet(lteCellSheet, groupName, colName, dstGroupStartColumn, dstGroupEndColumn, dstParaColumn)
                If findGroupParaNameType = 1 Then
                    lteCellSheet.Cells(2, dstParaColumn).EntireColumn.Hidden = True
                End If
            End If
        ElseIf GetDesStr(sheetName) = GetDesStr(getResByKey("Comm Data")) Then
            'lldFlag = GetLldFlag(sheetName, groupName, colName)
            'LLD列的值非True时，从[COMMON]页删除
            If GetDesStr(newLldFlag) <> GetDesStr("TRUE") Then
                readRowNum = Get_GroupRow(getResByKey("COMMON"), groupName) + 1
                
                rowCount = getEndRow(readRowNum, getResByKey("COMMON"))
                
                readColNum = Get_Col(getResByKey("COMMON"), readRowNum, colName)
                If readColNum > 0 Then
                    
                    readColEnd = Worksheets(getResByKey("COMMON")).range("IV" + CStr(readRowNum)).End(xlToLeft).column
                    
                    '如果该列为最后一项，则关联行，整体删除
                    If readColEnd = 1 And readColNum = 1 Then
                        range("A" & CStr(readRowNum - 1) & ":A" & CStr(readRowNum + rowCount)).EntireRow.Delete
                    '否则删除该列，后续列左移一格
                    ElseIf readColNum <= readColEnd Then
                        range(Cells(readRowNum, readColNum + 1), Cells(readRowNum + rowCount - 1, readColEnd)).Cut
                        Cells(readRowNum, readColNum).Select
                        ActiveSheet.Paste
                        
                        rows(readRowNum - 1).MergeCells = False
                        range(Cells(readRowNum - 1, 1), Cells(readRowNum - 1, readColEnd - 1)).Merge
                        
                        range(Cells(readRowNum - 1, readColEnd), Cells(readRowNum + rowCount - 1, readColEnd)).Select
                        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        Selection.Borders(xlEdgeTop).LineStyle = xlNone
                        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
                        Selection.Borders(xlEdgeRight).LineStyle = xlNone
                        Selection.Borders(xlInsideVertical).LineStyle = xlNone
                        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .colorIndex = xlAutomatic
                        End With
                        Selection.Interior.colorIndex = xlNone
                        
                        Selection.ClearContents
                        Selection.clearComments
                    End If
                End If
            End If
        End If
    Next

Application.ScreenUpdating = True

End Sub

Function getEndRow(readRowNum As Long, sheetName As String) As Long
        Dim rowCount, colIndex As Long
        Dim curStatus, preStatus As Boolean
        preStatus = False
        For rowCount = 1 To Worksheets(sheetName).UsedRange.rows.count
                curStatus = True
                For colIndex = 1 To Worksheets(sheetName).range("IV2").End(xlToLeft).column
                    If Worksheets(sheetName).Cells(readRowNum + rowCount, colIndex).value <> "" Then
                        curStatus = False
                    End If
                Next
                If preStatus = True And curStatus = False Then
                    Exit For
                End If
                preStatus = curStatus
        Next
        getEndRow = rowCount - 1
End Function

'复制QOS表格
Sub copyQOSInfo()
    Dim readRowEnd
    
    readRowEnd = Worksheets(getResByKey("COMMON")).range("a65536").End(xlUp).row
    Worksheets("QoS").Visible = True
    While Worksheets(getResByKey("COMMON")).Cells(readRowEnd + 1, 1).Borders(xlEdgeRight).LineStyle <> xlNone
        readRowEnd = readRowEnd + 1
    Wend
    
    Dim baseStationVersion As String, rowMax As Long
    baseStationVersion = UCase(getBaseStationVersion)
    rowMax = 64
    If InStr(baseStationVersion, "R011C00") > 0 Then
        rowMax = 17
    ElseIf baseStationCompare(baseStationVersion, "R013") < 0 Then
        rowMax = 21
    End If
    
    Worksheets("QoS").range("A1:D" & rowMax).Copy _
        Destination:=Worksheets(getResByKey("COMMON")).Cells(readRowEnd + 2, 1)
    Worksheets("QoS").Visible = False
End Sub

Public Function baseStationCompare(ByRef basestion As String, ByRef rversion As String) As Long
    On Error Resume Next
    Dim pos As Long
    pos = InStr(basestion, "R")
    If pos < 0 Then
       baseStationCompare = 1
       Exit Function
    End If
    Dim tmpbaseVersion As String
    tmpbaseVersion = Mid(basestion, pos, Len(rversion))
    If tmpbaseVersion > rversion Then
        baseStationCompare = 1
        Exit Function
    ElseIf tmpbaseVersion = rversion Then
        baseStationCompare = 0
        Exit Function
    Else
        baseStationCompare = -1
        Exit Function
    End If
    baseStationCompare = 1
End Function

'填写[COMMON]页采集到的信息
Sub setSummaryInfo()
    Dim sheetName As String  '[MAPPING DEF]内sheet名
    Dim groupName As String  '[MAPPING DEF]内group名
    Dim colName As String    '[MAPPING DEF]内列名
    Dim lldFlag As String    '[MAPPING DEF]内[LLD]列的值
    Dim newLldFlag As String
    
    Dim m_rowNum
    Dim n_RowNum
    Dim readRowNum As Long
    Dim readColNum As Long
    Dim readRowCount As Long
    Dim writeRowNum As Long
    Dim writeColNum As Long
    Dim writeRowCount As Long
    Dim dstPasteRange As range
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    
    Dim commonSheet As Worksheet, commonDataSheet As Worksheet
    Set commonSheet = ThisWorkbook.Worksheets(getResByKey("COMMON"))
    Set commonDataSheet = ThisWorkbook.Worksheets(getResByKey("Comm Data"))
    
    '遍历『MAPPING DEF』，用[COMMON]填写[Common Data]
    For m_rowNum = 2 To mappingDef.range("a65536").End(xlUp).row
        sheetName = mappingDef.Cells(m_rowNum, 1).value
        groupName = mappingDef.Cells(m_rowNum, 2).value
        colName = mappingDef.Cells(m_rowNum, 3).value
        newLldFlag = mappingDef.Cells(m_rowNum, 10).value
        
        If GetDesStr(sheetName) = GetDesStr(getResByKey("Comm Data")) Then
            'lldFlag = GetLldFlag(sheetName, groupName, colName)
            If GetDesStr(newLldFlag) = GetDesStr("TRUE") Then
                readRowNum = Get_GroupRow(getResByKey("COMMON"), groupName) + 2
                readColNum = Get_Col(getResByKey("COMMON"), readRowNum - 1, colName)
                writeRowNum = Get_GroupRow(getResByKey("Comm Data"), groupName) + 2
                writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, colName)
                    
                For readRowCount = 1 To commonSheet.range("a65536").End(xlUp).row
                    If commonSheet.Cells(readRowNum + readRowCount, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
                        Exit For
                    End If
                Next
                
                For writeRowCount = 1 To commonDataSheet.range("a65536").End(xlUp).row
                    If commonDataSheet.Cells(writeRowNum + writeRowCount, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
                        Exit For
                    End If
                Next
                
                If readRowCount > writeRowCount Then '如果COMMON页签中的某个对象数据行大于Common Data页，则在Common Data页中新增行
                    For n_RowNum = 0 To readRowCount - writeRowCount - 1
                        With commonDataSheet
                            .rows(CStr(writeRowNum + writeRowCount + n_RowNum) & ":" & CStr(writeRowNum + writeRowCount + n_RowNum)).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                            .rows(CStr(writeRowNum + writeRowCount + n_RowNum - 1)).Copy
                            .Paste Destination:=.rows(writeRowNum + writeRowCount + n_RowNum)
                        End With
                    Next
                End If
                
                commonSheet.range(commonSheet.Cells(readRowNum, readColNum), commonSheet.Cells(readRowNum + readRowCount - 1, readColNum)).Copy
                
                Set dstPasteRange = commonDataSheet.range(commonDataSheet.Cells(writeRowNum, writeColNum), commonDataSheet.Cells(writeRowNum + readRowCount - 1, writeColNum))
                commonDataSheet.Paste Destination:=dstPasteRange
            
                Call setRangeNormal(dstPasteRange) '将新拷贝过来的单元格灰化效果清除
'                With Selection.Borders(xlEdgeRight)
'                    .LineStyle = xlContinuous
'                    .Weight = xlThin
'                    .ColorIndex = xlAutomatic
'                End With
                Call setBorders(dstPasteRange) '将新拷贝过来的单元格设置边框
            End If
        End If
    Next
    
End Sub

'DSCP映射到DIFPRI
Sub setQOSInfo()
    Dim readRowBegin As Long
    Dim readRowEnd As Long
    Dim writeRowEnd As Long
    Dim m_rowNum As Long
    Dim writeRowNum As Long
    Dim writeColNum As Long
    Dim difpriName As String
    Dim sourceRowfor65 As Long
    Dim sourceRowfor66 As Long
    Dim sourceRowfor69 As Long
    Dim sourceRowfor70 As Long
    Dim sourceRowfor75 As Long
    Dim sourceRowfor79 As Long
    
    readRowBegin = Get_GroupRow(getResByKey("COMMON"), "QOS") + 2
    readRowEnd = Worksheets(getResByKey("COMMON")).range("c65536").End(xlUp).row
    writeRowEnd = Worksheets("Qos").range("c65536").End(xlUp).row
    
    '复制之前先清除内容
    Worksheets("QoS").Visible = True
    Worksheets("QoS").Select
    range(Cells(3, 1), Cells(64, 4)).ClearContents
    
    Worksheets(getResByKey("COMMON")).Select
    range(Cells(readRowBegin, 1), Cells(readRowBegin + 64, 4)).Select
    Selection.Copy
    Worksheets("QoS").Select
    range("A3").Select
    Worksheets("QoS").Paste
    Worksheets("QoS").Visible = False
    
    difpriName = getGrpNameFromCommSht("DIFPRI")
    writeRowNum = Get_GroupRow(getResByKey("Comm Data"), difpriName) + 2
    
    Worksheets(getResByKey("Comm Data")).rows(CStr(writeRowNum) & ":" & CStr(writeRowNum)).ClearContents
    Call clearUserParaData
    For m_rowNum = readRowBegin To readRowEnd
        Select Case ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(m_rowNum, 1).value
        Case "SCTP"
           Call writeDifpriData(writeRowNum, m_rowNum, "SIGPRI")
           Call writeDifpriData(writeRowNum, m_rowNum, "PRIRULE")
        Case "OM(MML)"
           Call writeDifpriData(writeRowNum, m_rowNum, "OMHIGHPRI")
        Case "OM(FTP)"
            Call writeDifpriData(writeRowNum, m_rowNum, "OMLOWPRI")
        Case "Synchronization"
            Call writeDifpriData(writeRowNum, m_rowNum, "IPCLKPRI")
        Case "QCI1"
            sourceRowfor65 = m_rowNum
            sourceRowfor66 = m_rowNum
            sourceRowfor75 = m_rowNum
            Call writeUserParaData(m_rowNum, 1)
        Case "QCI2"
            Call writeUserParaData(m_rowNum, 2)
        Case "QCI3"
            Call writeUserParaData(m_rowNum, 3)
        Case "QCI4"
            Call writeUserParaData(m_rowNum, 4)
        Case "QCI5"
            sourceRowfor69 = m_rowNum
            Call writeUserParaData(m_rowNum, 5)
        Case "QCI6"
            sourceRowfor70 = m_rowNum
            sourceRowfor79 = m_rowNum
            Call writeUserParaData(m_rowNum, 6)
        Case "QCI7"
            Call writeUserParaData(m_rowNum, 7)
        Case "QCI8"
            Call writeUserParaData(m_rowNum, 8)
        Case "QCI9"
            Call writeUserParaData(m_rowNum, 9)
        Case "QCI65"
            Call writeUserParaData4redundant(m_rowNum, sourceRowfor65, 10)
        Case "QCI66"
            Call writeUserParaData4redundant(m_rowNum, sourceRowfor66, 11)
        Case "QCI69"
            Call writeUserParaData4redundant(m_rowNum, sourceRowfor69, 12)
        Case "QCI70"
            Call writeUserParaData4redundant(m_rowNum, sourceRowfor70, 13)
        Case "QCI75"
            Call writeUserParaData4redundant(m_rowNum, sourceRowfor75, 14)
        Case "QCI79"
            Call writeUserParaData4redundant(m_rowNum, sourceRowfor79, 15)
        Case Else

        End Select
    Next
End Sub

Private Sub copyCerainRows(ByRef ws As Worksheet, ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal rowCount As Long)
    Dim n As Long
    For n = 1 To rowCount
        ws.rows(srcRowNumber).Copy
        ws.rows(dstRowNumber).INSERT
    Next n
    Application.CutCopyMode = False
End Sub

Sub clearUserParaData()
        Dim udtName As String
        Dim writeUdtRowNum As Long
        Dim rowCount As Long
        
        udtName = getGrpNameFromCommSht("UDT")
        If udtName = "" Then Exit Sub
        writeUdtRowNum = Get_GroupRow(getResByKey("Comm Data"), udtName) + 2
        rowCount = getEndRow(writeUdtRowNum - 1, getResByKey("Comm Data"))
        If (rowCount > 4) Then
            Worksheets(getResByKey("Comm Data")).rows(CStr(writeUdtRowNum + 3) & ":" & CStr(writeUdtRowNum + rowCount - 2)).EntireRow.Delete
        Else
            Call copyCerainRows(Worksheets(getResByKey("Comm Data")), writeUdtRowNum, writeUdtRowNum + rowCount - 1, 4 - rowCount)
        End If
        Worksheets(getResByKey("Comm Data")).rows(CStr(writeUdtRowNum) & ":" & CStr(writeUdtRowNum + 2)).ClearContents
        
        Dim udtparaGroup As String
        udtparaGroup = getGrpNameFromCommSht("UDTPARAGRP")
        Dim writeUdtParaGrpRowNum As Long
        writeUdtParaGrpRowNum = Get_GroupRow(getResByKey("Comm Data"), udtparaGroup) + 2
        
        rowCount = getEndRow(writeUdtParaGrpRowNum - 1, getResByKey("Comm Data"))
        If (rowCount > 4) Then
            Worksheets(getResByKey("Comm Data")).rows(CStr(writeUdtParaGrpRowNum + 3) & ":" & CStr(writeUdtParaGrpRowNum + rowCount - 2)).EntireRow.Delete
        Else
            Call copyCerainRows(Worksheets(getResByKey("Comm Data")), writeUdtParaGrpRowNum, writeUdtParaGrpRowNum + rowCount - 1, 4 - rowCount)
        End If
        Worksheets(getResByKey("Comm Data")).rows(CStr(writeUdtParaGrpRowNum) & ":" & CStr(writeUdtParaGrpRowNum + 2)).ClearContents
        
End Sub

Sub writeUserParaData(m_rowNum As Long, index As Long)
        Dim writeColNum As Long
        Dim columnName As String
        Dim udtName As String
        udtName = getGrpNameFromCommSht("UDT")
        If udtName = "" Then Exit Sub
        Dim writeUdtRowNum As Long
        Dim rowCount As Long
        writeUdtRowNum = Get_GroupRow(getResByKey("Comm Data"), udtName) + 2
        
        If index > 2 Then
            ThisWorkbook.Sheets(getResByKey("Comm Data")).rows(writeUdtRowNum + index - 1).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
         End If
        columnName = getColumnName(getResByKey("Comm Data"), "UDT", "UDTNO")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtRowNum - 1, columnName)
        ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = index
        columnName = getColumnName(getResByKey("Comm Data"), "UDT", "UDTPARAGRPID")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtRowNum - 1, columnName)
        ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = 39 + index
        
        Dim udtparaGroup As String
        udtparaGroup = getGrpNameFromCommSht("UDTPARAGRP")
        Dim writeUdtParaGrpRowNum As Long
        writeUdtParaGrpRowNum = Get_GroupRow(getResByKey("Comm Data"), udtparaGroup) + 2
        
        If index > 2 Then
            ThisWorkbook.Sheets(getResByKey("Comm Data")).rows(writeUdtParaGrpRowNum + index - 1).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        End If
        columnName = getColumnName(getResByKey("Comm Data"), "UDTPARAGRP", "UDTPARAGRPID")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtParaGrpRowNum - 1, columnName)
        ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtParaGrpRowNum + index - 1, writeColNum).value = 39 + index
        
        columnName = getColumnName(getResByKey("Comm Data"), "UDTPARAGRP", "PRI")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtParaGrpRowNum - 1, columnName)
        ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtParaGrpRowNum + index - 1, writeColNum).value = _
                        ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(m_rowNum, 2).value
                    
        columnName = getColumnName(getResByKey("Comm Data"), "UDTPARAGRP", "PRIRULE")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtParaGrpRowNum - 1, columnName)
        ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtParaGrpRowNum + index - 1, writeColNum).value = "DSCP"
        
End Sub

Function getUDTPARAGRPID(m_rowNum As Long, source_rowNum As Long, index As Long) As Long
        Dim flag_findUDTPARAGRPID As Boolean
        Dim curDSP As Long
        Dim sourceDSP As Long
        curDSP = ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(m_rowNum, 2).value
        sourceDSP = ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(source_rowNum, 2).value
        getUDTPARAGRPID = 0
        If curDSP <> sourceDSP Then
        getUDTPARAGRPID = 49 - index
        End If
        
End Function

Sub writeUserParaData4redundant(m_rowNum As Long, source_rowNum As Long, index As Long)
        Dim writeColNum As Long
        Dim columnName As String
        Dim udtName As String
        udtName = getGrpNameFromCommSht("UDT")
        If udtName = "" Then Exit Sub
        Dim writeUdtRowNum As Long
        Dim rowCount As Long
        Dim newUDTPARAGRPID As Long
        newUDTPARAGRPID = 0
        writeUdtRowNum = Get_GroupRow(getResByKey("Comm Data"), udtName) + 2
        
        If index > 2 Then
            ThisWorkbook.Sheets(getResByKey("Comm Data")).rows(writeUdtRowNum + index - 1).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        End If
        columnName = getColumnName(getResByKey("Comm Data"), "UDT", "UDTNO")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtRowNum - 1, columnName)
        
        Dim UDTNO As Long
        Dim UDTPARAGRPID As Long
        
        If index = 10 Then
'            ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = 65
'            columnName = getColumnName(getResByKey("Comm Data"), "UDT", "UDTPARAGRPID")
'            writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtRowNum - 1, columnName)
'            newUDTPARAGRPID = getUDTPARAGRPID(m_rowNum, source_rowNum, index)
'            If newUDTPARAGRPID > 0 Then
'                ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = newUDTPARAGRPID
'            Else
'                ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = 40
'            End If
            UDTNO = 65
            UDTPARAGRPID = 40
            
        ElseIf index = 11 Then
'            ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = 66
'            columnName = getColumnName(getResByKey("Comm Data"), "UDT", "UDTPARAGRPID")
'            writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtRowNum - 1, columnName)
'            newUDTPARAGRPID = getUDTPARAGRPID(m_rowNum, source_rowNum, index)
'            If newUDTPARAGRPID > 0 Then
'                ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = newUDTPARAGRPID
'            Else
'                ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = 40
'            End If
            UDTNO = 66
            UDTPARAGRPID = 40
            
        ElseIf index = 12 Then
'            ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = 69
'            columnName = getColumnName(getResByKey("Comm Data"), "UDT", "UDTPARAGRPID")
'            writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtRowNum - 1, columnName)
'
'            newUDTPARAGRPID = getUDTPARAGRPID(m_rowNum, source_rowNum, index)
'            If newUDTPARAGRPID > 0 Then
'                ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = newUDTPARAGRPID
'            Else
'                ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = 44
'            End If
            UDTNO = 69
            UDTPARAGRPID = 44
            
        ElseIf index = 13 Then
'            ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = 70
'            columnName = getColumnName(getResByKey("Comm Data"), "UDT", "UDTPARAGRPID")
'            writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtRowNum - 1, columnName)
'            newUDTPARAGRPID = getUDTPARAGRPID(m_rowNum, source_rowNum, index)
'            If newUDTPARAGRPID > 0 Then
'                ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = newUDTPARAGRPID
'            Else
'                ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = 45
'            End If
            UDTNO = 70
            UDTPARAGRPID = 45
        ElseIf index = 14 Then
            UDTNO = 75
            UDTPARAGRPID = 40
        ElseIf index = 15 Then
            UDTNO = 79
            UDTPARAGRPID = 45
        End If
        
        ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = UDTNO
        columnName = getColumnName(getResByKey("Comm Data"), "UDT", "UDTPARAGRPID")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtRowNum - 1, columnName)
        newUDTPARAGRPID = getUDTPARAGRPID(m_rowNum, source_rowNum, index)
        If newUDTPARAGRPID > 0 Then
            ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = newUDTPARAGRPID
        Else
            ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtRowNum + index - 1, writeColNum).value = UDTPARAGRPID
        End If
        
        
        
        On Error GoTo ErrExit
        Dim udtparaGroup As String
        udtparaGroup = getGrpNameFromCommSht("UDTPARAGRP")
        Dim writeUdtParaGrpRowNum As Long
        writeUdtParaGrpRowNum = Get_GroupRow(getResByKey("Comm Data"), udtparaGroup) + 2
        
        If index > 2 And newUDTPARAGRPID > 0 Then
            ThisWorkbook.Sheets(getResByKey("Comm Data")).rows(writeUdtParaGrpRowNum + index - 1).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        End If
        If newUDTPARAGRPID > 0 Then
        columnName = getColumnName(getResByKey("Comm Data"), "UDTPARAGRP", "UDTPARAGRPID")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtParaGrpRowNum - 1, columnName)
        ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtParaGrpRowNum + index - 1, writeColNum).value = newUDTPARAGRPID
        
        columnName = getColumnName(getResByKey("Comm Data"), "UDTPARAGRP", "PRI")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtParaGrpRowNum - 1, columnName)
        ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtParaGrpRowNum + index - 1, writeColNum).value = _
                        ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(m_rowNum, 2).value
                    
        columnName = getColumnName(getResByKey("Comm Data"), "UDTPARAGRP", "PRIRULE")
        writeColNum = Get_Col(getResByKey("Comm Data"), writeUdtParaGrpRowNum - 1, columnName)
        ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeUdtParaGrpRowNum + index - 1, writeColNum).value = "DSCP"
        End If
ErrExit:
        
End Sub


Sub writeDifpriData(writeRowNum As Long, m_rowNum As Long, attrName As String)
        Dim writeColNum As Long
        Dim columnName As String
        
           columnName = getColumnName(getResByKey("Comm Data"), "DIFPRI", attrName)
           writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, columnName)
        If writeColNum <> -1 And attrName <> "PRIRULE" Then
            ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeRowNum, writeColNum).value = _
                ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(m_rowNum, 2).value
        ElseIf attrName = "PRIRULE" Then
            ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeRowNum, writeColNum).value = _
                "DSCP"
        End If
End Sub

'DSCP和VLAN Priority映射到VlanClass
Sub setVlanClass()
    Dim readRowNum As Long
    Dim writeRowNum As Long
    Dim writeRowCount As Long
    Dim vLAnRowEnd As Long
    Dim readColNum1 As Long
    Dim writeColNum1 As Long
    Dim readColNum2 As Long
    Dim writeColNum2 As Long
    Dim readColNum3 As Long
    Dim writeColNum3 As Long
    Dim readColNum4 As Long
    Dim writeColNum4 As Long
    Dim vlanSheetName As String
    Dim vlanGrpName As String
    Dim columnName As String
    
    vlanSheetName = getPatternNameByMoc("VLANCLASS")
    vlanGrpName = get_GroupName(vlanSheetName, 1)
    
    readRowNum = Get_GroupRow(getResByKey("COMMON"), "QOS") + 2
    readColNum1 = Get_Col(getResByKey("COMMON"), readRowNum - 1, "DSCP")
    columnName = getColumnName(vlanSheetName, "VLANCLASS", "SRVPRIO")
    writeColNum1 = Get_Col(vlanSheetName, 2, columnName)
    
    readColNum2 = Get_Col(getResByKey("COMMON"), readRowNum - 1, "VLAN")
    columnName = getColumnName(vlanSheetName, "VLANCLASS", "TRAFFIC")
    writeColNum2 = Get_Col(vlanSheetName, 2, columnName)
    
    readColNum3 = Get_Col(getResByKey("COMMON"), readRowNum - 1, "VLAN Pri")
    columnName = getColumnName(vlanSheetName, "VLANCLASS", "VLANPRIO")
    writeColNum3 = Get_Col(vlanSheetName, 2, columnName)
    
    readColNum4 = Get_Col(getResByKey("COMMON"), readRowNum - 1, "Service Type")
    writeColNum4 = Get_Col(vlanSheetName, 2, getResByKey("serviceType"))
    
    
    writeRowNum = 3
    For readRowNum = readRowNum To Worksheets(getResByKey("COMMON")).range("c65536").End(xlUp).row
        If Worksheets(vlanSheetName).Cells(writeRowNum + 1, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
            Worksheets(vlanSheetName).rows(CStr(writeRowNum) & ":" & CStr(writeRowNum)).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        
        If writeColNum2 <> -1 Then '防止某个参数未定义在采集表中
            ThisWorkbook.Sheets(vlanSheetName).Cells(writeRowNum, writeColNum2).value = _
                ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum2).value
        End If
        
        If writeColNum1 <> -1 Then '防止某个参数未定义在采集表中
            If ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum2).value = "USERDATA" Then
               ThisWorkbook.Sheets(vlanSheetName).Cells(writeRowNum, writeColNum1).value = _
                    ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum1).value
            Else
                ThisWorkbook.Sheets(vlanSheetName).Cells(writeRowNum, writeColNum1).value = ""
            End If
        End If
        
        If writeColNum3 <> -1 Then '防止某个参数未定义在采集表中
            ThisWorkbook.Sheets(vlanSheetName).Cells(writeRowNum, writeColNum3).value = _
                ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum3).value
        End If
        
        If writeColNum4 <> -1 Then '防止某个参数未定义在采集表中
            ThisWorkbook.Sheets(vlanSheetName).Cells(writeRowNum, writeColNum4).value = _
                ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum4).value
        End If
        writeRowNum = writeRowNum + 1
    Next
    
    For writeRowCount = 1 To Worksheets(vlanSheetName).range("a65536").End(xlUp).row
        If Worksheets(vlanSheetName).Cells(writeRowNum + writeRowCount, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
            Exit For
        End If
    Next
    
    For writeRowNum = writeRowNum To Worksheets(vlanSheetName).range("b65536").End(xlUp).row
        Worksheets(vlanSheetName).rows(CStr(writeRowNum) & ":" & CStr(writeRowNum)).ClearContents
    Next
    
    If writeRowCount > 1 Then
        Worksheets(vlanSheetName).rows(CStr(writeRowNum + 1) & ":" & CStr(writeRowNum + writeRowCount)).Delete Shift:=xlUp
    End If
    
End Sub

'将比较字符串整形
Function GetDesStr(srcStr As String) As String
    GetDesStr = UCase(Trim(srcStr))
End Function

'从指定sheet页的指定行，查找指定列，返回列号
Function Get_Col(sheetName As String, recordRow As Long, ColValue As String) As Long
    Dim m_colNum As Long
    
    Get_Col = -1
    For m_colNum = 1 To Worksheets(sheetName).range("IV" + CStr(recordRow)).End(xlToLeft).column
        If GetDesStr(ColValue) = GetDesStr(Worksheets(sheetName).Cells(recordRow, m_colNum).value) Then
                Get_Col = m_colNum
            Exit For
        End If
    Next
End Function

'从指定sheet页查找group所在行
Function Get_GroupRow(sheetName As String, groupName As String) As Long
    Dim m_rowNum As Long
    Dim f_flag As Boolean
    
    f_flag = False
    For m_rowNum = 1 To Worksheets(sheetName).range("a65536").End(xlUp).row
        If GetDesStr(groupName) = GetDesStr(Worksheets(sheetName).Cells(m_rowNum, 1).value) Then
            f_flag = True
            Exit For
        End If
    Next
    If f_flag = False Then
        'MsgBox sheetName & "缺少Group：" & groupName, vbExclamation, "Warning"
    End If
    
    Get_GroupRow = m_rowNum
    
End Function

Function NeedChange() As String
    Dim sheetName As String  '[MAPPING DEF]内sheet名
    Dim lldFlag As String    '[MAPPING DEF]内[LLD]列的值
    Dim changeObj As String
    Dim lastSheetName As String
    Dim m_rowNum As Long
    changeObj = ""
    lastSheetName = ""
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    '遍历『MAPPING DEF』
    For m_rowNum = 2 To mappingDef.range("a65536").End(xlUp).row
        sheetName = mappingDef.Cells(m_rowNum, 1).value
        lldFlag = mappingDef.Cells(m_rowNum, 10).value
        
        If GetDesStr(lldFlag) = GetDesStr("TRUE") And sheetName <> lastSheetName Then
            changeObj = changeObj + ";" + sheetName
            lastSheetName = sheetName
        End If
    Next
    If changeObj = "" Then
        MsgBox getResByKey("lldWarning"), vbExclamation, getResByKey("Warning")
    End If
    NeedChange = changeObj
    
End Function

Function GetLldFlag(sheetName As String, groupName As String, colName As String) As String
    Dim tempSheetName As String
    Dim tempGroupName As String
    Dim tempColName As String
    Dim tempFlag As String
    Dim m_rowNum As Long
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    '遍历『MAPPING DEF』
    For m_rowNum = 2 To mappingDef.range("a65536").End(xlUp).row
        tempSheetName = mappingDef.Cells(m_rowNum, 1).value
        tempGroupName = mappingDef.Cells(m_rowNum, 2).value
        tempColName = mappingDef.Cells(m_rowNum, 3).value

        If GetDesStr(tempSheetName) = GetDesStr(sheetName) _
            And GetDesStr(tempGroupName) = GetDesStr(groupName) _
            And GetDesStr(tempColName) = GetDesStr(colName) Then
            tempFlag = mappingDef.Cells(m_rowNum, 10).value
        End If
        If GetDesStr(tempFlag) = GetDesStr("TRUE") Then
            Exit For
        End If
    Next

    GetLldFlag = tempFlag
    
End Function

Private Function getGrpNameFromCommSht(mocName As String) As String
    getGrpNameFromCommSht = ""
    Dim iRow As range
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim maxRowNumber As Long
    'maxRowNumber = mappingDef.UsedRange.rows.count
    maxRowNumber = getUsedRowCount(mappingDef)
    For Each iRow In mappingDef.rows
        If iRow.row > maxRowNumber Then
            Exit Function
        End If
        If iRow.Cells(1, 1).value = getResByKey("Comm Data") And iRow.Cells(1, 4).value = mocName Then
            getGrpNameFromCommSht = iRow.Cells(1, 2).value
            Exit Function
        End If
    Next
End Function

Private Function getColumnName(sheetName As String, mocName As String, attrName As String) As String
    getColumnName = ""
    Dim iRow As range
    Dim mappingDef As Worksheet
    Dim maxRowNumber As Long
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    'maxRowNumber = mappingDef.UsedRange.rows.count
    maxRowNumber = getUsedRowCount(mappingDef)
    For Each iRow In mappingDef.rows
        If iRow.row > maxRowNumber Then
            Exit Function
        End If
        If iRow.Cells(1, 1).value = sheetName And iRow.Cells(1, 4).value = mocName And iRow.Cells(1, 5).value = attrName Then
            getColumnName = iRow.Cells(1, 3).value
            Exit Function
        End If
    Next
End Function

Private Function getPatternNameByMoc(mocName As String) As String
    Dim iRow As range
    Dim shRow As range
    Dim mappingDef As Worksheet
    Dim sheetDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For Each iRow In mappingDef.rows
        If iRow.Cells(1, 1).value = "" Then Exit For
        If iRow.Cells(1, 4).value = mocName Then
            For Each shRow In sheetDef.rows
                If shRow.Cells(1, 1).value = "" Then Exit For
                If iRow.Cells(1, 1).value = shRow.Cells(1, 1).value And shRow.Cells(1, 2).value = "Pattern" Then
                    getPatternNameByMoc = shRow.Cells(1, 1).value
                    Exit Function
                End If
            Next
        End If
    Next
End Function


