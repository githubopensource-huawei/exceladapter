Attribute VB_Name = "LLD_Summary"
Option Explicit
Public lldModelFlag As String
Function getLLDModelFlag() As String
    getLLDModelFlag = lldModelFlag
End Function
'LLD<-->Summary
Sub Summary2LLD()
    On Error Resume Next
    Dim SheetNum, i As Long
    Dim sht As Worksheet
    Dim sheetArray As Variant
    
    Application.ScreenUpdating = False
    DisplayMessageOnStatusbar
    
    lldModelFlag = "SUMMARY"
    For SheetNum = 1 To ThisWorkbook.Sheets.count
        If ThisWorkbook.Sheets(SheetNum).name = getResByKey("COMMON") Then
            lldModelFlag = "LLD"
            Exit For
        End If
    Next
    
    '����������LLD->Summary�Ĵ���
    If lldModelFlag = "LLD" Then
        '��ʾSummary�ı�Ҫsheetҳ
        For Each sht In ActiveWorkbook.Worksheets
            If sht.Visible <> 2 And GetDesStr(sht.name) <> GetDesStr("Qos") _
            And GetDesStr(sht.name) <> GetDesStr("MAPPING DEF") _
            And GetDesStr(sht.name) <> GetDesStr("CONTROL DEF") _
            And GetDesStr(sht.name) <> GetDesStr("SHEET DEF") _
            And GetDesStr(sht.name) <> GetDesStr("VALID DEF") Then
                sht.Visible = True
                sht.Cells.EntireColumn.Hidden = False
            End If
        Next
        
        '��д[COMMON]ҳ�ɼ�������Ϣ
        setSummaryInfo
        'DSCPӳ�䵽DIFPRI
        setQOSInfo
        'DSCP��VLAN Priorityӳ�䵽VlanClass
        setVlanClass
        
        'ɾ��[COMMON]
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(getResByKey("COMMON")).Delete
        Application.DisplayAlerts = True
        lldModelFlag = ""
    '�������Summary->LLD�Ĵ���
    ElseIf (lldModelFlag = "SUMMARY") And (NeedChange() <> "") Then
        '����[COMMON]
        Sheets(getResByKey("Comm Data")).Copy Before:=Sheets(getResByKey("Comm Data"))
        Sheets(getResByKey("Comm Data")).name = getResByKey("COMMON")
        
        'ɾ��[COMMON]�еķ�LLD��Ϣ�����ش�������ߵķ�LLD��Ϣ
        setLLDInfo
        '����QOS����
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
    End If
    
    lldModelFlag = ""
    ReturnStatusbaring
    Application.ScreenUpdating = True
End Sub

'Summary->LLD��ɾ��[COMMON]�еķ�LLD��Ϣ�����ش�������ߵķ�LLD��Ϣ
Sub setLLDInfo()
    Dim sheetName As String  '[MAPPING DEF]��sheet��
    Dim groupName As String  '[MAPPING DEF]��group��
    Dim colName As String    '[MAPPING DEF]������
    Dim lldFlag As String    '[MAPPING DEF]��[LLD]�е�ֵ
    
    Dim m_rowNum
    Dim readRowNum As Long
    Dim rowCount As Long
    Dim readColNum As Long
    Dim readColEnd As Long
       
    '������MAPPING DEF��
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a65536").End(xlUp).row
        sheetName = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value
        groupName = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value
        colName = Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value
        
        If GetDesStr(sheetName) = GetDesStr(getResByKey("BaseTransPort")) Or GetDesStr(sheetName) = GetDesStr(getResByKey("ENODEB_RADIO_DATA")) Then
            lldFlag = GetLldFlag(sheetName, groupName, colName)
            'LLD�е�ֵ��Trueʱ������
            If GetDesStr(lldFlag) <> GetDesStr("TRUE") Then
                readColNum = Get_Col(sheetName, 2, colName)
                Worksheets(sheetName).Cells(2, readColNum).EntireColumn.Hidden = True
            End If
        ElseIf GetDesStr(sheetName) = GetDesStr(getResByKey("Comm Data")) Then
            lldFlag = GetLldFlag(sheetName, groupName, colName)
            'LLD�е�ֵ��Trueʱ����[COMMON]ҳɾ��
            If GetDesStr(lldFlag) <> GetDesStr("TRUE") Then
                readRowNum = Get_GroupRow(getResByKey("COMMON"), groupName) + 1
                
                rowCount = getEndRow(readRowNum)
                
                readColNum = Get_Col(getResByKey("COMMON"), readRowNum, colName)
                readColEnd = Worksheets(getResByKey("COMMON")).range("IV" + CStr(readRowNum)).End(xlToLeft).column
                
                '�������Ϊ���һ�������У�����ɾ��
                If readColEnd = 1 And readColNum = 1 Then
                    range("A" & CStr(readRowNum - 1) & ":A" & CStr(readRowNum + rowCount)).EntireRow.Delete
                '����ɾ�����У�����������һ��
                ElseIf readColNum <= readColEnd Then
                    range(Cells(readRowNum, readColNum + 1), Cells(readRowNum + rowCount - 1, readColEnd)).Cut
                    Cells(readRowNum, readColNum).Select
                    ThisWorkbook.ActiveSheet.Paste
                    
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
    Next

End Sub

Function getEndRow(readRowNum As Long) As Long
        Dim rowCount, colIndex As Long
        Dim curStatus, preStatus As Boolean
        preStatus = False
        For rowCount = 1 To Worksheets(getResByKey("COMMON")).range("a65536").End(xlUp).row
                curStatus = True
                For colIndex = 1 To Worksheets(getResByKey("COMMON")).range("IV2").End(xlToLeft).column
                    If Worksheets(getResByKey("COMMON")).Cells(readRowNum + rowCount, colIndex).value <> "" Then
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

'����QOS����
Sub copyQOSInfo()
    Dim readRowEnd
    
    readRowEnd = Worksheets(getResByKey("COMMON")).range("a65536").End(xlUp).row
    Worksheets("QoS").Visible = True
    While Worksheets(getResByKey("COMMON")).Cells(readRowEnd + 1, 1).Borders(xlEdgeRight).LineStyle <> xlNone
        readRowEnd = readRowEnd + 1
    Wend
    
    Worksheets("QoS").range("A1:D64").Copy _
        Destination:=Worksheets(getResByKey("COMMON")).Cells(readRowEnd + 2, 1)
    Worksheets("QoS").Visible = False
End Sub

'��д[COMMON]ҳ�ɼ�������Ϣ
Sub setSummaryInfo()
    Dim sheetName As String  '[MAPPING DEF]��sheet��
    Dim groupName As String  '[MAPPING DEF]��group��
    Dim colName As String    '[MAPPING DEF]������
    Dim lldFlag As String    '[MAPPING DEF]��[LLD]�е�ֵ
    
    Dim m_rowNum
    Dim n_RowNum
    Dim readRowNum As Long
    Dim readColNum As Long
    Dim readRowCount As Long
    Dim writeRowNum As Long
    Dim writeColNum As Long
    Dim writeRowCount As Long
        
    '������MAPPING DEF������[COMMON]��д[Common Data]
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a65536").End(xlUp).row
        sheetName = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value
        groupName = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value
        colName = Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value
        
        If GetDesStr(sheetName) = GetDesStr(getResByKey("Comm Data")) Then
            lldFlag = GetLldFlag(sheetName, groupName, colName)
            If GetDesStr(lldFlag) = GetDesStr("TRUE") Then
                readRowNum = Get_GroupRow(getResByKey("COMMON"), groupName) + 2
                readColNum = Get_Col(getResByKey("COMMON"), readRowNum - 1, colName)
                writeRowNum = Get_GroupRow(getResByKey("Comm Data"), groupName) + 2
                writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, colName)
                    
                For readRowCount = 1 To Sheets(getResByKey("COMMON")).range("a65536").End(xlUp).row
                    If Sheets(getResByKey("COMMON")).Cells(readRowNum + readRowCount, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
                        Exit For
                    End If
                Next
                
                For writeRowCount = 1 To Sheets(getResByKey("Comm Data")).range("a65536").End(xlUp).row
                    If Sheets(getResByKey("Comm Data")).Cells(writeRowNum + writeRowCount, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
                        Exit For
                    End If
                Next
                
                If readRowCount > writeRowCount Then
                    For n_RowNum = 0 To readRowCount - writeRowCount - 1
                        Sheets(getResByKey("Comm Data")).Select
                        rows(CStr(writeRowNum + writeRowCount + n_RowNum) & ":" & CStr(writeRowNum + writeRowCount + n_RowNum)).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                        rows(CStr(writeRowNum + writeRowCount + n_RowNum - 1)).Copy
                        Cells(writeRowNum + writeRowCount + n_RowNum, 1).Select
                        ThisWorkbook.ActiveSheet.Paste
                    Next
                End If
                
                Sheets(getResByKey("COMMON")).Select
                range(Cells(readRowNum, readColNum), Cells(readRowNum + readRowCount - 1, readColNum)).Copy
                Sheets(getResByKey("Comm Data")).Select
                Cells(writeRowNum, writeColNum).Select
                ThisWorkbook.ActiveSheet.Paste
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .colorIndex = xlAutomatic
                End With
            End If
        End If
    Next
    
End Sub

'DSCPӳ�䵽DIFPRI
Sub setQOSInfo()
    Dim readRowBegin As Long
    Dim readRowEnd As Long
    Dim writeRowEnd As Long
    Dim m_rowNum As Long
    Dim writeRowNum As Long
    Dim writeColNum As Long
    
    
    readRowBegin = Get_GroupRow(getResByKey("COMMON"), "QOS") + 2
    readRowEnd = Worksheets(getResByKey("COMMON")).range("c65536").End(xlUp).row
    writeRowEnd = Worksheets("Qos").range("c65536").End(xlUp).row
    
    '����֮ǰ���������
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
    
    writeRowNum = Get_GroupRow(getResByKey("Comm Data"), getResByKey("DIFF_SERVICE")) + 2
    Worksheets(getResByKey("Comm Data")).rows(CStr(writeRowNum) & ":" & CStr(writeRowNum)).ClearContents
    
    For m_rowNum = readRowBegin To readRowEnd
        Select Case ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(m_rowNum, 1).value
        Case "SCTP"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("SIGNAL_PRI"))
        Case "OM(MML)"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("OM_HIGH_PRI"))
        Case "OM(FTP)"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("OM_LOW_PRI"))
        Case "Synchronization"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("IP_CLOCK_PRI"))
        Case "QCI1"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("USER_DATA_1_PRI"))
        Case "QCI2"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("USER_DATA_2_PRI"))
        Case "QCI3"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("USER_DATA_3_PRI"))
        Case "QCI4"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("USER_DATA_4_PRI"))
        Case "QCI5"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("USER_DATA_5_PRI"))
        Case "QCI6"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("USER_DATA_6_PRI"))
        Case "QCI7"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("USER_DATA_7_PRI"))
        Case "QCI8"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("USER_DATA_8_PRI"))
        Case "QCI9"
            writeColNum = Get_Col(getResByKey("Comm Data"), writeRowNum - 1, getResByKey("USER_DATA_9_PRI"))
        Case Else
            writeColNum = -1
        End Select
        
        If writeColNum <> -1 Then
            ThisWorkbook.Sheets(getResByKey("Comm Data")).Cells(writeRowNum, writeColNum).value = _
                ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(m_rowNum, 2).value
        End If
    Next
End Sub

'DSCP��VLAN Priorityӳ�䵽VlanClass
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
    
    readRowNum = Get_GroupRow(getResByKey("COMMON"), "QOS") + 2
    readColNum1 = Get_Col(getResByKey("COMMON"), readRowNum - 1, "DSCP")
    writeColNum1 = Get_Col(getResByKey("VLAN_PRI_MAP"), 2, getResByKey("USER_DATA_PRI"))
    readColNum2 = Get_Col(getResByKey("COMMON"), readRowNum - 1, "VLAN")
    writeColNum2 = Get_Col(getResByKey("VLAN_PRI_MAP"), 2, getResByKey("BUSINESS_TYPE"))
    readColNum3 = Get_Col(getResByKey("COMMON"), readRowNum - 1, "VLAN Pri")
    writeColNum3 = Get_Col(getResByKey("VLAN_PRI_MAP"), 2, getResByKey("VLAN_PRI"))
    readColNum4 = Get_Col(getResByKey("COMMON"), readRowNum - 1, "Service Type")
    writeColNum4 = Get_Col(getResByKey("VLAN_PRI_MAP"), 2, getResByKey("serviceType"))
    
    writeRowNum = 3
    For readRowNum = readRowNum To Worksheets(getResByKey("COMMON")).range("c65536").End(xlUp).row
        If Worksheets(getResByKey("VLAN_PRI_MAP")).Cells(writeRowNum + 1, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
            Worksheets(getResByKey("VLAN_PRI_MAP")).rows(CStr(writeRowNum) & ":" & CStr(writeRowNum)).INSERT Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        
        ThisWorkbook.Sheets(getResByKey("VLAN_PRI_MAP")).Cells(writeRowNum, writeColNum2).value = _
            ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum2).value
                
        If ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum2).value = "USERDATA" Then
           ThisWorkbook.Sheets(getResByKey("VLAN_PRI_MAP")).Cells(writeRowNum, writeColNum1).value = _
                ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum1).value
        Else
            ThisWorkbook.Sheets(getResByKey("VLAN_PRI_MAP")).Cells(writeRowNum, writeColNum1).value = ""
        End If
                
        ThisWorkbook.Sheets(getResByKey("VLAN_PRI_MAP")).Cells(writeRowNum, writeColNum3).value = _
            ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum3).value
        ThisWorkbook.Sheets(getResByKey("VLAN_PRI_MAP")).Cells(writeRowNum, writeColNum4).value = _
            ThisWorkbook.Sheets(getResByKey("COMMON")).Cells(readRowNum, readColNum4).value
        writeRowNum = writeRowNum + 1
    Next
    
    For writeRowCount = 1 To Worksheets(getResByKey("VLAN_PRI_MAP")).range("a65536").End(xlUp).row
        If Worksheets(getResByKey("VLAN_PRI_MAP")).Cells(writeRowNum + writeRowCount, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
            Exit For
        End If
    Next
    
    For writeRowNum = writeRowNum To Worksheets(getResByKey("VLAN_PRI_MAP")).range("b65536").End(xlUp).row
        Worksheets(getResByKey("VLAN_PRI_MAP")).rows(CStr(writeRowNum) & ":" & CStr(writeRowNum)).ClearContents
    Next
    
    If writeRowCount > 1 Then
        Worksheets(getResByKey("VLAN_PRI_MAP")).rows(CStr(writeRowNum + 1) & ":" & CStr(writeRowNum + writeRowCount)).Delete Shift:=xlUp
    End If
    
End Sub

'���Ƚ��ַ�������
Function GetDesStr(srcStr As String) As String
    GetDesStr = UCase(Trim(srcStr))
End Function

'��ָ��sheetҳ��ָ���У�����ָ���У������к�
Function Get_Col(sheetName As String, recordRow As Long, ColValue As String) As Long
    Get_Col = -1

    Dim targetRange As range
    Set targetRange = ThisWorkbook.Worksheets(sheetName).rows(recordRow).Find(ColValue, lookat:=xlWhole, LookIn:=xlValues)
    If Not targetRange Is Nothing Then Get_Col = targetRange.column
End Function

'��ָ��sheetҳ����group������
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
        'MsgBox sheetName & "ȱ��Group��" & groupName, vbExclamation, "����"
    End If
    
    Get_GroupRow = m_rowNum
    
End Function

Function NeedChange() As String
    Dim sheetName As String  '[MAPPING DEF]��sheet��
    Dim lldFlag As String    '[MAPPING DEF]��[LLD]�е�ֵ
    Dim changeObj As String
    Dim lastSheetName As String
    Dim m_rowNum As Long
    changeObj = ""
    lastSheetName = ""
       
    '������MAPPING DEF��
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a65536").End(xlUp).row
        sheetName = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value
        lldFlag = Worksheets("MAPPING DEF").Cells(m_rowNum, 10).value
        
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
       
    '������MAPPING DEF��
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a65536").End(xlUp).row
        tempSheetName = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value
        tempGroupName = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value
        tempColName = Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value

        If GetDesStr(tempSheetName) = GetDesStr(sheetName) _
            And GetDesStr(tempGroupName) = GetDesStr(groupName) _
            And GetDesStr(tempColName) = GetDesStr(colName) Then
            tempFlag = Worksheets("MAPPING DEF").Cells(m_rowNum, 10).value
        End If
        If GetDesStr(tempFlag) = GetDesStr("TRUE") Then
            Exit For
        End If
    Next

    GetLldFlag = tempFlag
    
End Function











