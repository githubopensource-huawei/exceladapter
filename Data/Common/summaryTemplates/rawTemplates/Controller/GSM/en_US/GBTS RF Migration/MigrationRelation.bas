Attribute VB_Name = "MigrationRelation"
Option Explicit




Public Sub AddDelMigrationSourceNe()
    Dim neType As String
    neType = getNeType()
    If neType = "GSM" Then
        ConfigMRATSrcNEForm.Show
    End If
    
End Sub


Public Function isMigrationRelationSheet(ByRef ws As Worksheet) As Boolean
    Dim sheetName As String, migrationSheetName As String
    isMigrationRelationSheet = False
'    If ws.name = "RF Migration NEs Relationship" Or ws.name = "射频搬迁网元关系" Then
'        isMigrationRelationSheet = True
'    End If
    If ws.name = getResByKey("RF Migration NEs Relationship") Then
        isMigrationRelationSheet = True
    End If
End Function

Public Sub MigrationSelectionChange(ByRef ws As Worksheet, ByRef Target As range)
    On Error GoTo ErrorHandler
    Dim groupName As String, columnName As String
    Dim rowNumber As Long, columnNumber As Long
    If Target.Rows.count <> 1 Or Target.columns.count <> 1 Then Exit Sub
    If isMigrationRelationSheet(ws) = False Then Exit Sub
    If Target.row < 4 Then Exit Sub
    
    rowNumber = Target.row
    columnNumber = Target.column
    
    Dim srcNeMap As CMap
    Set srcNeMap = initSrcGroupNameMap(ws)
    
    Dim referencedString As String
    groupName = ws.Cells(2, columnNumber).value
    columnName = ws.Cells(3, columnNumber).value
    If InStr(columnName, getResByKey("RADIO")) = 0 Then Exit Sub
    
    If columnName = getResByKey("Global Radio GBTS reference") And srcNeMap.hasKey("BTS") Then
        referencedString = srcNeMap.GetAt("BTS")
        
                If Target.Borders.LineStyle <> xlNone Then
            Call setBoardStyleListBoxRangeValidation(ws.name, groupName, columnName, referencedString, ws, Target)
        Else
            Target.Validation.Delete
        End If
    End If
    
'    If (columnName = "Global Radio NodeB reference" Or columnName = "NodeB全局无线参数取值") And srcNeMap.haskey("NodeB") Then
'        referencedString = srcNeMap.GetAt("NodeB")
'    End If
'
'    If (columnName = "Global Radio eNodeB reference" Or columnName = "eNodeB全局无线参数取值") And srcNeMap.haskey("eNodeB") Then
'        referencedString = srcNeMap.GetAt("eNodeB")
'    End If
    
    
 
ErrorHandler:
End Sub
Private Function initSrcGroupNameMap(ByRef ws As Worksheet) As CMap
    Dim m_colNum As Long
    Dim cellValue As String
    Dim btsNameStr As String
    Dim nodebNameStr As String
    Dim enodebNameStr As String
    Dim srcNeMap As CMap
    Set srcNeMap = New CMap
    
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        If btsNameStr = "" Then
            btsNameStr = getSrcNeNameCol("BTS", ws.Cells(3, m_colNum).value)
        ElseIf getSrcNeNameCol("BTS", ws.Cells(3, m_colNum).value) <> "" Then
            btsNameStr = btsNameStr + "," + getSrcNeNameCol("BTS", ws.Cells(3, m_colNum).value)
        Else
        End If
        
        If nodebNameStr = "" Then
            nodebNameStr = getSrcNeNameCol("NodeB", ws.Cells(3, m_colNum).value)
        ElseIf getSrcNeNameCol("NodeB", ws.Cells(3, m_colNum).value) <> "" Then
            nodebNameStr = nodebNameStr + "," + getSrcNeNameCol("NodeB", ws.Cells(3, m_colNum).value)
        Else
        End If
        
        If enodebNameStr = "" Then
            enodebNameStr = getSrcNeNameCol("eNodeB", ws.Cells(3, m_colNum).value)
        ElseIf getSrcNeNameCol("eNodeB", ws.Cells(3, m_colNum).value) <> "" Then
            enodebNameStr = enodebNameStr + "," + getSrcNeNameCol("eNodeB", ws.Cells(3, m_colNum).value)
        Else
        End If
    Next
    
    Dim targetNum As Long
    targetNum = getTargetNeClounmNum(ws)
    
    If btsNameStr <> "" Then
        btsNameStr = btsNameStr + "," + ws.Cells(2, targetNum).value
        Call srcNeMap.SetAt("BTS", btsNameStr)
    End If
    If nodebNameStr <> "" Then
        nodebNameStr = nodebNameStr + "," + ws.Cells(2, targetNum).value
        Call srcNeMap.SetAt("NodeB", nodebNameStr)
    End If
    If enodebNameStr <> "" Then
        enodebNameStr = enodebNameStr + "," + ws.Cells(2, targetNum).value
        Call srcNeMap.SetAt("eNodeB", enodebNameStr)
    End If
    Set initSrcGroupNameMap = srcNeMap
End Function


Private Function initSrcNeMap(ByRef ws As Worksheet, ByRef rowNumber As Long) As CMap
    Dim m_colNum As Long
    Dim cellValue As String
    Dim btsNameStr As String
    Dim nodebNameStr As String
    Dim enodebNameStr As String
    Dim srcNeMap As CMap
    Set srcNeMap = New CMap
    
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        If isSrcNameCol("BTS", ws.Cells(3, m_colNum).value) Then
            If btsNameStr = "" Then
                btsNameStr = ws.Cells(rowNumber, m_colNum).value
            ElseIf ws.Cells(rowNumber, m_colNum).value <> "" Then
                btsNameStr = btsNameStr + "," + ws.Cells(rowNumber, m_colNum).value
            Else
            End If
        End If
        
        If isSrcNameCol("NodeB", ws.Cells(3, m_colNum).value) Then
            If nodebNameStr = "" Then
                nodebNameStr = ws.Cells(rowNumber, m_colNum).value
            ElseIf ws.Cells(rowNumber, m_colNum).value <> "" Then
                nodebNameStr = nodebNameStr + "," + ws.Cells(rowNumber, m_colNum).value
            Else
            End If
        End If
        
        If isSrcNameCol("eNodeB", ws.Cells(3, m_colNum).value) Then
            If enodebNameStr = "" Then
                enodebNameStr = ws.Cells(rowNumber, m_colNum).value
            ElseIf ws.Cells(rowNumber, m_colNum).value <> "" Then
                enodebNameStr = enodebNameStr + "," + ws.Cells(rowNumber, m_colNum).value
            Else
            End If
        End If
    Next
    
    Dim targetNum As Long
    targetNum = getTargetNeClounmNum(ws)
    
    If btsNameStr <> "" Then
        btsNameStr = btsNameStr + "," + ws.Cells(rowNumber, targetNum).value
        Call srcNeMap.SetAt("BTS", btsNameStr)
    End If
    If nodebNameStr <> "" Then
        nodebNameStr = nodebNameStr + "," + ws.Cells(rowNumber, targetNum).value
        Call srcNeMap.SetAt("NodeB", nodebNameStr)
    End If
    If enodebNameStr <> "" Then
        enodebNameStr = enodebNameStr + "," + ws.Cells(rowNumber, targetNum).value
        Call srcNeMap.SetAt("eNodeB", enodebNameStr)
    End If
    Set initSrcNeMap = srcNeMap
End Function

Public Function isSrcNameCol(neType As String, cellValue As String) As Boolean
    Dim num As Long
    Dim tempValue As String
'    Dim tempCnValue As String
    isSrcNameCol = False
    
    
    
    For num = 1 To 50
'        tempEnValue = neType + CStr(num) + " NE Name"
'        tempCnValue = neType + CStr(num) + " 网元名称"
        tempValue = neType + CStr(num) + " " + getResByKey("NE_NAME")
        If tempValue = cellValue Then
            isSrcNameCol = True
            Exit Function
        End If
    Next
End Function
Private Function getSrcNeNameCol(neType As String, cellValue As String) As String
    Dim num As Long
    Dim tempValue As String
'    Dim tempCnValue As String
    getSrcNeNameCol = ""
    
    For num = 1 To 50
        tempValue = neType + CStr(num) + " " + getResByKey("NE_NAME")
'        tempCnValue = neType + CStr(num) + " 网元名称"
        If tempValue = cellValue Then
            getSrcNeNameCol = neType + CStr(num)
            Exit Function
        End If
    Next
End Function

Public Function getTargetNeClounmNum(ByRef ws As Worksheet) As Long
    Dim m_colNum As Long
    Dim tempFunctionStr As String
'    Dim tempEnFunctionStr As String
    
'    tempCnFunctionStr = "目标网元"
'    tempEnFunctionStr = "Target NE"
    tempFunctionStr = getResByKey("TARGET_NE_NAME")
    
    getTargetNeClounmNum = -1
    
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        If ws.Cells(2, m_colNum).value = tempFunctionStr Then
            getTargetNeClounmNum = m_colNum
            Exit Function
        End If
    Next
End Function

Public Function geteNodebStartClounmNum(ByRef ws As Worksheet) As Long
    Dim m_colNum As Long
    Dim tempFunctionStr As String
'    Dim tempEnFunctionStr As String
    
    tempFunctionStr = "eNodeB/ eNodeB " + getResByKey("FUNCTION")
'    tempEnFunctionStr = "eNodeB/ eNodeB Function"
    
    geteNodebStartClounmNum = -1
    
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        If ws.Cells(2, m_colNum).value = tempFunctionStr Then
            geteNodebStartClounmNum = m_colNum
            Exit Function
        End If
    Next
End Function
Public Function getNodebStartClounmNum(ByRef ws As Worksheet) As Long
    Dim m_colNum As Long
    Dim tempFunctionStr As String
'    Dim tempEnFunctionStr As String
    
    tempFunctionStr = "NodeB/ NodeB " + getResByKey("FUNCTION")
'    tempEnFunctionStr = "NodeB/ NodeB Function"
    
    getNodebStartClounmNum = -1
    
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        If ws.Cells(2, m_colNum).value = tempFunctionStr Then
            getNodebStartClounmNum = m_colNum
            Exit Function
        End If
    Next
End Function

'Public Function isCnEnv(ByRef ws As Worksheet) As Boolean
'    Dim m_colNum As Long
'    Dim tempCnFunctionStr As String
'    Dim tempEnFunctionStr As String
'
'    tempCnFunctionStr = "目标网元"
'    tempEnFunctionStr = "Target NE"
'
'    isCnEnv = False
'
'    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
'        If ws.Cells(2, m_colNum).value = tempCnFunctionStr Then
'            isCnEnv = True
'            Exit Function
'        End If
'    Next
'End Function

'Creat Tool Bar


'---------------------------------------------------------------------------------------
Public Sub delRadioColunmRecs(ByRef neType As String)
    Dim srcNeType As String
    Dim srccolumNum As Long
    Dim ws As Worksheet
    Dim srccellValue As String
    Set ws = ThisWorkbook.ActiveSheet
    Dim delCollter As String
    
    If isExistRadioGroup(neType, srccolumNum) And isExistGroup(neType) = False Then
        delCollter = getColStr(srccolumNum)
        srccellValue = ws.Cells(1, srccolumNum).value
        ws.columns(delCollter & ":" & delCollter).Delete Shift:=xlToLeft
    End If
    
    If srccellValue <> "" Then
        If isExistRadioGroup("BTS", srccolumNum) Then
            ws.Cells(1, srccolumNum).value = srccellValue
        ElseIf isExistRadioGroup("NodeB", srccolumNum) Then
            ws.Cells(1, srccolumNum).value = srccellValue
        ElseIf isExistRadioGroup("eNodeB", srccolumNum) Then
            ws.Cells(1, srccolumNum).value = srccellValue
        End If
    End If

End Sub

Public Sub delNameColunmRecs(ByRef neType As String, ByRef newColunmNum As Long)
    Dim srcColNum As Long
    Dim srcColletter As String
    Dim colNum As Long
    Dim maxNum As Long
    Dim srccellValue As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    Dim targetCellValue As String
    Dim startCol As Long, endCol As Long, totalNum As Long
    Call getStartEndSrcNameCol(neType, startCol, endCol, totalNum)
    
    If totalNum < newColunmNum Then
        If neType = "BTS" Then totalNum = totalNum + 1
        newColunmNum = totalNum
    End If
    Dim delCollter As String
    Dim startCollter As String
    startCollter = getColStr(startCol)
    
    For colNum = 1 To newColunmNum
        delCollter = getColStr(endCol + 1 - colNum)
        ws.columns(delCollter & ":" & delCollter).Delete Shift:=xlToLeft
    Next

End Sub
Private Sub getStartEndSrcNameCol(ByRef neType As String, ByRef startCol As Long, ByRef endCol As Long, ByRef totalNum As Long)
    Dim m_colNum As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column

        If InStr(ws.Cells(3, m_colNum).value, neType) <> 0 Then
           startCol = m_colNum
           endCol = getNameEndNum(m_colNum, ws, neType)
           totalNum = endCol - startCol + 1
           Exit Sub
       End If
    Next
    
End Sub

Public Sub makeNameColunmRecs(ByRef neType As String, ByRef newColunmNum As Long)
    Dim srcColNum As Long
    Dim srcColletter As String
    Dim colNum As Long
    Dim maxNum As Long
    Dim srccellValue As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    Dim targetCellValue As String
    srcColNum = getSrcNameCol(neType)
    srcColletter = getColStr(srcColNum + 1)
    srccellValue = ws.Cells(3, srcColNum).value
    maxNum = getNameNum(neType, srccellValue)
    
    For colNum = 1 To newColunmNum
        ws.columns(srcColletter & ":" & srcColletter).Insert Shift:=xlToLeft
        ws.Cells(3, srcColNum + 1).value = Replace(ws.Cells(3, srcColNum).value, CStr(maxNum), CStr(maxNum + newColunmNum + 1 - colNum))
    Next
    Dim columStartChar As String
    Dim columEndChar As String
    
    columStartChar = getcolumStartChar(neType)
    srcColNum = getSrcNameCol(neType)
    columEndChar = getColStr(srcColNum)
    
    Application.DisplayAlerts = False
    ws.range(columStartChar + "2:" + columEndChar + "2").Merge
    ws.range("A1:" + columEndChar + "1").Merge
    Application.DisplayAlerts = True
End Sub

Private Function getcolumStartChar(ByRef neType As String) As String
    Dim m_colNum As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim cellValue As String
    
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        If InStr(ws.Cells(3, m_colNum).value, neType) <> 0 Then
            If neType = "BTS" Then
                getcolumStartChar = getColStr(m_colNum - 1)
            Else
                getcolumStartChar = getColStr(m_colNum)
            End If
            Exit Function
        End If
    Next
    
End Function

Private Function getSrcNameCol(ByRef neType As String) As Long
    Dim m_colNum As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim cellValue As String
'    Dim tempCnFunctionStr As String
    Dim tempFunctionStr As String
    
    tempFunctionStr = "G" + neType + "/ " + "G" + neType + " " + getResByKey("FUNCTION")
'    tempEnFunctionStr = "G" + neType + "/ " + "G" + neType + " Function"
    
    For m_colNum = 2 To ws.range("XFD2").End(xlToLeft).column
        If ws.Cells(2, m_colNum).value = tempFunctionStr Then
            getSrcNameCol = m_colNum
        End If
        
        If InStr(ws.Cells(3, m_colNum).value, neType) <> 0 Then
           getSrcNameCol = getNameEndNum(m_colNum, ws, neType)
           'getColStr(nameEndNum)
           Exit Function
       End If
    Next
    
End Function

Private Function getNameEndNum(ByRef startColNum As Long, ByRef ws As Worksheet, ByRef neType As String) As Long
    Dim cellValue As String
    Dim m_colNum As Long
    getNameEndNum = startColNum
    
    For m_colNum = startColNum To ws.range("XFD2").End(xlToLeft).column
       If isSameNameCol(neType, ws.Cells(3, m_colNum).value) = False Then
           getNameEndNum = m_colNum - 1
           Exit Function
       End If
    Next
End Function
Private Function getNameNum(ByRef neType As String, ByRef nameValue As String) As Long
    Dim neTypeLen As Long
    Dim numStr As String
    getNameNum = 1
    
    neTypeLen = Len(neType)
    numStr = Mid(nameValue, neTypeLen + 1, 1)
    getNameNum = CLng(numStr)
    
End Function

Private Function isSameNameCol(ByRef neType As String, cellValue As String) As Boolean
    Dim num As Long
'    Dim tempEnValue As String
    Dim tempValue As String
    isSameNameCol = False
    
    For num = 1 To 50
        tempValue = neType + CStr(num) + " " + getResByKey("NE_NAME")
'        tempCnValue = neType + CStr(num) + " 网元名称"
        If tempValue = cellValue Then
            isSameNameCol = True
            Exit Function
        End If
    Next
End Function

Public Sub makeNewRadioNameColunmRec(ByRef neType As String)
    Dim srccolumNum As Long
    Dim srccolumLetter As String
    Dim getTargetNeCol As Long
    Dim getTargetNeColletter As String
    Dim srcNeType As String
    Dim ws As Worksheet
    Dim tempNeType As String
    Set ws = ThisWorkbook.ActiveSheet
    tempNeType = neType
    If neType = "BTS" Then tempNeType = "GBTS"
    If isExistRadioGroup(tempNeType, srccolumNum) Then Exit Sub
    
    getTargetNeCol = getTargetNeClounmNum(ws)

    If neType = "BTS" Then
        getTargetNeColletter = getColStr(getTargetNeCol + 1)
    ElseIf neType = "NodeB" Then
        getTargetNeColletter = getColStr(getTargetNeCol + 2)
    Else
        getTargetNeColletter = getColStr(getTargetNeCol + 3)
    End If
    
    'ws.Columns(srccolumLetter & ":" & srccolumLetter).Copy
    ws.columns(getTargetNeColletter & ":" & getTargetNeColletter).Insert Shift:=xlToRight
    
    Call getSrcColNum(srcNeType, srccolumNum)
    srccolumLetter = getColStr(srccolumNum)
    
    ws.range(getTargetNeColletter & "2").value = Replace(ws.range(srccolumLetter & "2").value, srcNeType, tempNeType)
    ws.range(getTargetNeColletter & "3").value = Replace(ws.range(srccolumLetter & "3").value, srcNeType, tempNeType)
    If ws.range(getTargetNeColletter & "1").value = "" Then
        ws.range(getTargetNeColletter & "1").value = ws.range(srccolumLetter & "1").value
        
        ws.range(srccolumLetter & "1").Copy
        ws.range(getTargetNeColletter & "1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
        ws.range(srccolumLetter & "2").Copy
        ws.range(getTargetNeColletter & "2").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
        ws.range(srccolumLetter & "3").Copy
        ws.range(getTargetNeColletter & "3").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
        Application.CutCopyMode = False
    End If

    
    Dim endNum As Long
    endNum = getTargetNeCol + 1
    If isExistRadioGroup("BTS", endNum) Or isExistRadioGroup("NodeB", endNum) Or isExistRadioGroup("eNodeB", endNum) Then
        Application.DisplayAlerts = False
        ws.range(getColStr(getTargetNeCol + 1) + "1:" + getColStr(endNum) + "1").Merge
        Application.DisplayAlerts = True
    End If
End Sub
Public Sub getSrcColNum(ByRef srcNeType As String, ByRef srcColNum As Long)
    If isExistRadioGroup("BTS", srcColNum) Then
        srcNeType = "GBTS"
    ElseIf isExistRadioGroup("NodeB", srcColNum) Then
        srcNeType = "NodeB"
    ElseIf isExistRadioGroup("eNodeB", srcColNum) Then
        srcNeType = "eNodeB"
    End If
End Sub

Public Sub makeNewGroupNameColunmRecs(ByRef neType As String, ByRef newColunmNum As Long)
    Dim startColunNum As Long
    Dim startColunLetter As String
    Dim colNum As Long
    Dim tempNeType As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    startColunNum = getInseterStartColunmNum(neType, ws)
    If startColunNum = -1 Then Exit Sub
    startColunLetter = getColStr(startColunNum)
    tempNeType = neType
    If neType = "BTS" Then
        newColunmNum = newColunmNum + 1
        tempNeType = "GBTS"
    End If
    
    Dim functionName As String
    Dim neTypeName As String
    Dim firstGroupName As String
    Dim bscname As String
    Dim nameNum As String
    
    Dim getTargetNeCol As Long
    Dim getTargetNeColletter As String
    
    getTargetNeCol = getTargetNeClounmNum(ws)
    getTargetNeColletter = getColStr(getTargetNeCol - 1)

'    If isCnEnv(ws) Then
'        functionName = tempNeType + "/ " + tempNeType + " 功能"
'        neTypeName = neType + "1 网元名称"
'        bscname = "BSC 名称"
'        firstGroupName = "源网元"
'    Else
'        functionName = tempNeType + "/ " + tempNeType + " Function"
'        neTypeName = neType + "1 NE Name"
'        bscname = "BSC Name"
'        firstGroupName = "Source NE"
'    End If
    
    firstGroupName = getResByKey("SOURCE_NE")
    functionName = tempNeType + "/ " + tempNeType + " " + getResByKey("FUNCTION")
    neTypeName = neType + getResByKey("SRCNE_NAME")
    bscname = getResByKey("BSC_NENAME")
    
    For colNum = 1 To newColunmNum
        ws.columns(startColunLetter & ":" & startColunLetter).Insert Shift:=xlToRight
        If neType = "BTS" Then
            nameNum = CStr(newColunmNum - colNum)
        Else
            nameNum = CStr(newColunmNum + 1 - colNum)
        End If
        
        If colNum = newColunmNum Then
            ws.Cells(2, startColunNum).value = functionName
            If neType = "BTS" Then
                ws.Cells(3, startColunNum).value = bscname
            Else
                ws.Cells(3, startColunNum).value = Replace(neTypeName, "1", nameNum)
            End If
        Else
            ws.Cells(3, startColunNum).value = Replace(neTypeName, "1", nameNum)
        End If
        ws.Cells(3, getTargetNeCol).Copy
        ws.Cells(3, startColunNum).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    Next
    
    Dim columStartChar As String
    Dim columEndChar As String
    Dim endColNum As Long

    columStartChar = getcolumStartChar(neType)
    endColNum = getSrcNameCol(neType)
    columEndChar = getColStr(endColNum)
    
    getTargetNeCol = getTargetNeClounmNum(ws)
    getTargetNeColletter = getColStr(getTargetNeCol - 1)
    
    If ws.Cells(1, 1).value = "" Then
        ws.Cells(1, getTargetNeCol).Copy
        ws.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ws.Cells(1, 1).value = firstGroupName
        
        ws.Cells(2, getTargetNeCol).Copy
        ws.Cells(2, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
  
        Application.CutCopyMode = False
    End If
    
    Application.DisplayAlerts = False
    ws.range(columStartChar + "2:" + columEndChar + "2").Merge
    ws.range("A1:" + getTargetNeColletter + "1").Merge
    Application.DisplayAlerts = True
End Sub

Public Function isExistGroup(ByRef neType As String) As Boolean
    Dim m_colNum As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim tempFunctionStr As String
'    Dim tempEnFunctionStr As String
    Dim tempNeType As String
    Dim getTargetNeCol As Long
    tempNeType = neType
    If neType = "BTS" Then tempNeType = "GBTS"
    
    getTargetNeCol = getTargetNeClounmNum(ws)
    
    tempFunctionStr = tempNeType + "/ " + tempNeType + " " + getResByKey("FUNCTION")
'    tempEnFunctionStr = tempNeType + "/ " + tempNeType + " Function"
    
    isExistGroup = False
    
    For m_colNum = 1 To getTargetNeCol
        If ws.Cells(2, m_colNum).value = tempFunctionStr Then
            isExistGroup = True
            Exit Function
        End If
    Next
End Function
Public Function isExistRadioGroup(ByRef neType As String, ByRef srcColNum As Long) As Boolean
    Dim m_colNum As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim tempCnFunctionStr As String
    Dim tempEnFunctionStr As String
    Dim tempNeType As String
    tempNeType = neType
    If neType = "BTS" Then tempNeType = "GBTS"
    
    tempEnFunctionStr = "Global Radio " + tempNeType + " " + getResByKey("RADIO_REFERENCE")
    tempCnFunctionStr = tempNeType + getResByKey("RADIO_REFERENCE")
    
    isExistRadioGroup = False
    
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        If ws.Cells(3, m_colNum).value = tempCnFunctionStr Or ws.Cells(3, m_colNum).value = tempEnFunctionStr Then
            isExistRadioGroup = True
            srcColNum = m_colNum
            Exit Function
        End If
    Next
End Function

Private Function getInseterStartColunmNum(ByRef neType As String, ByRef ws As Worksheet) As Long
    getInseterStartColunmNum = -1
    
    If neType = "BTS" Then
        If isExistGroup("NodeB") Then
            getInseterStartColunmNum = getNodebStartClounmNum(ws)
        ElseIf isExistGroup("eNodeB") Then
            getInseterStartColunmNum = geteNodebStartClounmNum(ws)
        Else
            getInseterStartColunmNum = getTargetNeClounmNum(ws)
        End If
    End If
    
    If neType = "eNodeB" Then
        getInseterStartColunmNum = getTargetNeClounmNum(ws)
    End If
    
    If neType = "NodeB" Then
        If isExistGroup("eNodeB") Then
            getInseterStartColunmNum = geteNodebStartClounmNum(ws)
        Else
            getInseterStartColunmNum = getTargetNeClounmNum(ws)
        End If
    End If

End Function


Public Sub getStartendColNum(ByRef neType As String, ByRef startColNum As Long, ByRef endColNum As Long)
    Dim m_colNum As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim cellValue As String
    Dim getTargetNeCol As Long
    
    getTargetNeCol = getTargetNeClounmNum(ws)
    
    For m_colNum = 1 To getTargetNeCol
        If InStr(ws.Cells(3, m_colNum).value, neType) <> 0 Then
            endColNum = getNameEndNum(m_colNum, ws, neType)
            If neType = "BTS" Then
                startColNum = m_colNum - 1
            Else
                startColNum = m_colNum
            End If
            Exit Sub
        End If
    Next
    
End Sub



Public Function initSrcTargetNeMapInfo() As CMapValueObject
    Dim srcTargetNeMap As CMapValueObject
    Dim targetColNum As Long
    Dim targetNeName As String

    Dim ws As Worksheet
    Dim index As Long
    Dim onerowSrcNeMap As CMap
    Dim temponerowSrcNeMap As CMap
    
    Set srcTargetNeMap = New CMapValueObject
    
    For Each ws In ThisWorkbook.Worksheets
        If isMigrationRelationSheet(ws) Then
            targetColNum = getTargetNeClounmNum(ws)
            For index = 4 To ws.range("a1048576").End(xlUp).row
                targetNeName = ws.Cells(index, targetColNum).value
                Set onerowSrcNeMap = initSrcNeMap(ws, index)
                
                If srcTargetNeMap.hasKey(targetNeName) Then
                
                Else
                    Call srcTargetNeMap.SetAt(targetNeName, onerowSrcNeMap)
                End If
            Next
        End If
    Next
    
    Set initSrcTargetNeMapInfo = srcTargetNeMap

End Function
Public Function getNeSationType() As String
    Dim sheetName As String
    Dim ws As Worksheet
    sheetName = "NE TYPE"
    If containsASheet(ThisWorkbook, sheetName) Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        getNeSationType = ws.Cells(2, 1).value
    Else
        getNeSationType = getNeType()
    End If
End Function


Public Sub refreshMigrationSourceData(ByRef ws As Worksheet)
    Dim sheetName As String
    Dim groupName As String
    Dim keyValue As Variant
    Dim keyValueStr As String
    Dim sheetgroupNameArry() As String
    Dim sourceRowNumber As Long
    Dim maxColNum As Long
    
    If sourceBoardStyleDataMap Is Nothing Then Exit Sub
    
    For Each keyValue In sourceBoardStyleDataMap.KeyCollection
        keyValueStr = keyValue
        maxColNum = sourceBoardStyleDataMap.GetAt(keyValueStr)
        sheetgroupNameArry = Split(keyValueStr, ",")
        sheetName = sheetgroupNameArry(0)
        groupName = sheetgroupNameArry(1)
        sourceRowNumber = sheetgroupNameArry(2)
        
        If ws.name = sheetName Then Call setMigrationRecbackColor(ws, sourceRowNumber, maxColNum)
    Next
    
    sourceBoardStyleDataMap.Clean
End Sub

Public Sub setFunctionNameBoxLst(ByVal sh As Object, ByVal Target As range)
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim tempcellRange As range
    Dim groupName As String
    Dim columnName As String
    Dim rowNum As Long
    Dim colNum As Long
    Dim m_colNum As Long
    Dim stationNeType As String
    Set sheet = sh
    stationNeType = getNeType()
    If Target.count Mod 256 = 0 Then Exit Sub

    
    Dim migrationDataManager As CMigrationDataManager
    Dim cellNeNameColumMap As CMap
    Dim baseStationNeMap As CMapValueObject
    Set migrationDataManager = New CMigrationDataManager
    Call migrationDataManager.init
    Set cellNeNameColumMap = migrationDataManager.cellNeNameColumMap
    Set baseStationNeMap = migrationDataManager.baseStationNeMap
    
    Dim groupColnameStr As String
    Dim neType As String
    Dim srcneName As String
    Dim neName As String
    
    Dim boxstr As String
    
    For Each cellRange In Target
        If findAttrName(Trim(cellRange.value)) = True Or findGroupName(Trim(cellRange.value)) = True Or cellRange.Borders.LineStyle = xlLineStyleNone Then
            Exit Sub
        End If
        
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        rowNum = cellRange.row
        colNum = cellRange.column
        groupColnameStr = groupName + "," + columnName
        
        neType = getFunctionNameColumNeType(cellNeNameColumMap, groupColnameStr)
        If columnName = "*BTS Name" Or columnName = getResByKey("COLNAME_BTSNAME") Or columnName = getResByKey("COLNAME_BTSNAME_EX") Then
            neType = "BTS"
        End If
        If neType = "" Then GoTo NextLoop
        
        '默认源站名称在目标站名称的左边一个单元格
        If isSrcNeColumn(sheet, rowNum, colNum - 1) Then
            srcneName = sheet.Cells(rowNum, colNum - 1).value '需要转为物理网元
            
            boxstr = getBoxStrWhenHasSrcNeName(migrationDataManager, srcneName, stationNeType, neType)
            Call setFunctionNameListBoxRange(sheet, groupName, columnName, cellRange, boxstr)
            
        Else '没有源基站列时
            neName = sheet.Cells(rowNum, colNum).value
            
            '如果基站名称为空，没有ListBox；如果基站名称不为空，ListBox显示当前值，及其对应的目标站和源站的集合
            If neName <> "" Then
                boxstr = getBoxStrWhenOnlyDstNeName(migrationDataManager, neName, stationNeType, neType)
                Call setFunctionNameListBoxRange(sheet, groupName, columnName, cellRange, boxstr)
            Else
                Target.Validation.Delete
            End If
        End If
NextLoop:
    Next
End Sub

Private Function getFunctionNameColumNeType(ByRef cellNeNameColumMap As CMap, ByRef groupColnameStr As String) As String
    Dim keyValueStr As Variant
    Dim valueStr As String
    getFunctionNameColumNeType = ""
    For Each keyValueStr In cellNeNameColumMap.KeyCollection
        valueStr = cellNeNameColumMap.GetAt(keyValueStr)
        If valueStr = groupColnameStr Then
            getFunctionNameColumNeType = keyValueStr
            Exit Function
        End If
    Next
End Function

'判断是否是源基站名称列
Private Function isSrcNeColumn(ByRef curSheet As Worksheet, rowNum As Long, colNum As Long) As Boolean
    Dim cellRange As range
    isSrcNeColumn = False
    
    '默认源站名称在目标站名称的左边一个单元格
    If colNum < 1 Then
        Exit Function
    End If
    
    '取列名称
    Dim groupName As String
    Dim columnName As String
    Set cellRange = curSheet.Cells(rowNum, colNum)
    Call getGroupAndColumnName(curSheet, cellRange, groupName, columnName)
    
    '根据列名称找到属性名称，判断属性名称是否为SOURCEBTSNAME
    Dim mapdef As CMappingDef
    Set mapdef = getMappingDefine(curSheet.name, groupName, columnName)
    
    If mapdef Is Nothing Or mapdef.neType = "" Or mapdef.attributeName <> "SOURCEBTSNAME" Then
        Exit Function
    End If
    
    isSrcNeColumn = True
End Function

Private Function getBoxStrWhenHasSrcNeName(ByRef migrationDataManager As CMigrationDataManager, ByRef srcneName As String, _
    ByRef stationNeType As String, ByRef neType As String) As String
    Dim pythsrcNeName As String
    Dim boxstr As String
    
    pythsrcNeName = srcneName
    If stationNeType = "MRAT" Then pythsrcNeName = getpythNeNamebyFunctionName(srcneName, neType, migrationDataManager.baseStationNeMap)

    boxstr = getTargetNeBoxStr(migrationDataManager, neType, pythsrcNeName, stationNeType)
    If boxstr <> "" Then
        boxstr = boxstr + "," + srcneName
    Else
        boxstr = srcneName
    End If
    
    getBoxStrWhenHasSrcNeName = boxstr
End Function

Private Function setFunctionNameListBoxRange(ByRef sheet As Worksheet, ByRef groupName As String, ByRef columnName As String, _
    ByRef Target As range, ByRef referencedString As String)
    
    If referencedString <> "" Then
        Call setBoardStyleListBoxRangeValidation(sheet.name, groupName, columnName, referencedString, sheet, Target)
    Else
        Target.Validation.Delete
    End If
            
End Function

Private Function getBoxStrWhenOnlyDstNeName(ByRef migrationDataManager As CMigrationDataManager, ByRef neName As String, _
    ByRef stationNeType As String, ByRef neType As String) As String
    Dim pythNeName As String
    Dim boxstr As String
    
    pythNeName = neName
    If stationNeType = "MRAT" Then pythNeName = getpythNeNamebyFunctionName(neName, neType, migrationDataManager.baseStationNeMap)
                
    '先假设此为目标站，取其源站信息
    boxstr = getSourceNeBoxStr(migrationDataManager, neType, pythNeName, stationNeType)
    
    '如果源站信息为空,假设此为源站，取其目标站信息
    If boxstr = "" Then
        boxstr = getTargetNeBoxStr(migrationDataManager, neType, pythNeName, stationNeType)
    End If
    
    If boxstr <> "" Then
        boxstr = boxstr + "," + neName
    Else
        boxstr = neName
    End If

    getBoxStrWhenOnlyDstNeName = boxstr
End Function

Private Function getpythNeNamebyFunctionName(ByRef functionName As String, ByRef neType As String, ByRef baseStationNeMap As CMapValueObject) As String
    Dim onerowRec As CMap
    Dim targetValue As Variant
    Dim tempfunctionValue As String
    getpythNeNamebyFunctionName = ""
    For Each targetValue In baseStationNeMap.KeyCollection
        Set onerowRec = baseStationNeMap.GetAt(targetValue)
        If onerowRec.hasKey(neType) Then tempfunctionValue = onerowRec.GetAt(neType)
        If functionName = tempfunctionValue Then
            getpythNeNamebyFunctionName = targetValue
            Exit Function
        End If
    Next
End Function

Private Function getTargetNeBoxStr(ByRef migrationDataManager As CMigrationDataManager, ByRef neType As String, ByRef srcneName As String, ByRef stationNeType As String) As String
    Dim targetSourceNeMap As CMapValueObject
    Dim baseStationNeMap As CMapValueObject
    Dim onerowNeMap As CMap
    Dim srcNeNameStr As String
    Dim tempNeType As String
    tempNeType = neType
    'If stationNeType <> "MRAT" Then tempNeType = "BASESTATION"
    
    getTargetNeBoxStr = ""

    Set targetSourceNeMap = migrationDataManager.targetSourceNeMap
    Set baseStationNeMap = migrationDataManager.baseStationNeMap
    
    Dim targetNeName As Variant
    For Each targetNeName In targetSourceNeMap.KeyCollection
        Set onerowNeMap = targetSourceNeMap.GetAt(targetNeName)
        If onerowNeMap.hasKey(neType) Then
            srcNeNameStr = onerowNeMap.GetAt(neType)
            If isExitSrcNe(srcneName, srcNeNameStr) Then getTargetNeBoxStr = getTargetFunctionNeName(baseStationNeMap, CStr(targetNeName), tempNeType) '单站需要处理这个neType
        End If
    Next
    
End Function

Private Function getSourceNeBoxStr(ByRef migrationDataManager As CMigrationDataManager, ByRef neType As String, ByRef dstPhyNeName As String, ByRef stationNeType As String) As String
    Dim targetSourceNeMap As CMapValueObject
    Dim baseStationNeMap As CMapValueObject
    Dim onerowNeMap As CMap
    Dim phySrcNeNameStr As String
    Dim phySrcNeNameArray() As String
    
    Dim tempNeType As String
    tempNeType = neType
    'If stationNeType <> "MRAT" Then tempNeType = "BASESTATION"
    
    getSourceNeBoxStr = ""

    Set targetSourceNeMap = migrationDataManager.targetSourceNeMap
    Set baseStationNeMap = migrationDataManager.baseStationNeMap
    
    If targetSourceNeMap.hasKey(dstPhyNeName) Then
        Set onerowNeMap = targetSourceNeMap.GetAt(dstPhyNeName)
        If onerowNeMap.hasKey(neType) Then
            phySrcNeNameStr = onerowNeMap.GetAt(neType)  '得到的是Node名称，需要再转换成RAT名称
        End If
    End If
    
    If phySrcNeNameStr = "" Then
        Exit Function
    End If
        
    phySrcNeNameArray = Split(phySrcNeNameStr, ",")
    
    Dim delimiter As String
    delimiter = ""
    
    Dim index As Long
    For index = LBound(phySrcNeNameArray) To UBound(phySrcNeNameArray)
        getSourceNeBoxStr = getSourceNeBoxStr + delimiter + getTargetFunctionNeName(baseStationNeMap, CStr(phySrcNeNameArray(index)), tempNeType)
        delimiter = ","
    Next
    
End Function

Private Function isExitSrcNe(ByRef srcneName As String, ByRef srcNeNameStr As String) As Boolean
    Dim srcNeNameArry() As String
    isExitSrcNe = False
    If srcNeNameStr <> "" Then srcNeNameArry = Split(srcNeNameStr, ",")
    
    Dim index As Long
    For index = LBound(srcNeNameArry) To UBound(srcNeNameArry)
        If srcNeNameArry(index) = srcneName Then
            isExitSrcNe = True
            Exit Function
        End If
    Next
End Function


Private Function getTargetFunctionNeName(ByRef baseStationNeMap As CMapValueObject, ByRef targetNeName As String, ByRef neType As String) As String
    Dim onerowRec As CMap
    getTargetFunctionNeName = ""
    If baseStationNeMap.hasKey(targetNeName) Then
        Set onerowRec = baseStationNeMap.GetAt(targetNeName)
    End If
    If onerowRec.hasKey(neType) Then
        getTargetFunctionNeName = onerowRec.GetAt(neType)
    End If
End Function
