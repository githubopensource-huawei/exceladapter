VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMigrationNeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private RelationNeColumnNameMap_ As CMap
Private RelationNeColumnValueMap_ As CMap
Private ws As Worksheet

Public Function RelationNeColumnNameMap(ByRef rowNumber As Long) As CMap
    Call initRelationNeColumnMap(rowNumber)
    Set RelationNeColumnNameMap = RelationNeColumnNameMap_
End Function

Public Function RelationNeColumnValueMap(ByRef rowNumber As Long) As CMap
    Call initRelationNeColumnMap(rowNumber)
    Set RelationNeColumnValueMap = RelationNeColumnValueMap_
End Function

Public Function neStartColNumber(ByRef neType As String) As Long
    neStartColNumber = getNeStartColNumber(neType)
End Function

Public Function neEndColNumber(ByRef neType As String) As Long
    neEndColNumber = getNeEndColNumber(neType)
End Function

Public Function srcNeNameColNum(ByRef neType As String) As Long
    Dim startColNum As Long
    Dim endColNum As Long
    srcNeNameColNum = 0
    startColNum = getNeStartColNumber(neType)
    endColNum = getNeEndColNumber(neType)
    If startColNum = -1 Then
        srcNeNameColNum = 0
    Else
        srcNeNameColNum = endColNum - startColNum + 1
    End If
End Function

Public Sub init()

End Sub

Private Sub initRelationNeColumnMap(ByRef rowNumber As Long)
    Dim m_colNum As Long
    Dim cellValue As String
    Dim columnName As String
    Dim btsNameStr As String
    Dim nodebNameStr As String
    Dim enodebNameStr As String
    Dim btsValueStr As String
    Dim nodebValueStr As String
    Dim enodebValueStr As String
    Dim srcNeColName As String
    
    Set RelationNeColumnNameMap_ = New CMap
    Set RelationNeColumnValueMap_ = New CMap
    
    Dim targetNum As Long
    targetNum = getTargetNeClounmNum(ws)
    
    btsNameStr = ws.Cells(2, targetNum).value
    nodebNameStr = ws.Cells(2, targetNum).value
    enodebNameStr = ws.Cells(2, targetNum).value
    
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        columnName = ws.Cells(3, m_colNum).value
        cellValue = ws.Cells(rowNumber, m_colNum).value
        
        If getSrcNeNameCol("BTS", columnName) <> "" And cellValue <> "" Then btsNameStr = btsNameStr + "," + getSrcNeNameCol("BTS", columnName)
        If getSrcNeNameCol("NodeB", columnName) <> "" And cellValue <> "" Then nodebNameStr = nodebNameStr + "," + getSrcNeNameCol("NodeB", columnName)
        If getSrcNeNameCol("eNodeB", columnName) <> "" And cellValue <> "" Then enodebNameStr = enodebNameStr + "," + getSrcNeNameCol("eNodeB", columnName)
        
        If isSrcNameCol("BTS", columnName) And cellValue <> "" Then
            If btsValueStr = "" Then
                btsValueStr = cellValue
            Else
                btsValueStr = btsValueStr + "," + cellValue
            End If
        End If
        
        If isSrcNameCol("NodeB", columnName) And cellValue <> "" Then
            If nodebValueStr = "" Then
                nodebValueStr = cellValue
            Else
                nodebValueStr = nodebValueStr + "," + cellValue
            End If
        End If
        
        If isSrcNameCol("eNodeB", columnName) And cellValue <> "" Then
            If enodebValueStr = "" Then
                enodebValueStr = cellValue
            Else
                enodebValueStr = enodebValueStr + "," + cellValue
            End If
        End If
        
    Next
    
    Call RelationNeColumnNameMap_.SetAt("BTS", btsNameStr)
    Call RelationNeColumnNameMap_.SetAt("NodeB", nodebNameStr)
    Call RelationNeColumnNameMap_.SetAt("eNodeB", enodebNameStr)
    
    Call RelationNeColumnValueMap_.SetAt("BTS", btsValueStr)
    Call RelationNeColumnValueMap_.SetAt("NodeB", nodebValueStr)
    Call RelationNeColumnValueMap_.SetAt("eNodeB", enodebValueStr)
    
End Sub

Public Sub insertSrcNeNameColumn(ByRef neType As String, ByRef newColNum As Long)
    Dim startColNum As Long
    Dim endColNum As Long
    Dim colNum As Long
    Dim insertLetter As String
    Dim insertCellValue As String
    Dim srcNeNameNum As Long
    endColNum = getNeEndColNumber(neType)
    
    If endColNum = -1 Then
        Call insertNewSrcNeNameColumn(neType, newColNum)
        Call insertNewRadioNameColumn(neType)
    Else
        insertLetter = getColStr(endColNum + 1)
        insertCellValue = ws.Cells(3, endColNum).value
        srcNeNameNum = getSrcNeNameNum(neType, insertCellValue)
        For colNum = 1 To newColNum
            ws.columns(insertLetter & ":" & insertLetter).Insert Shift:=xlToLeft
            ws.Cells(3, endColNum + 1).value = Replace(insertCellValue, CStr(srcNeNameNum), CStr(srcNeNameNum + newColNum + 1 - colNum))
        Next
    End If
    Call mergeFirandSecGroupColumn(neType)
End Sub

Public Sub delSrcNeNameColRec(ByRef neType As String, ByRef delColNum As Long)
    Call delSrcNeNameColunmRecs(neType, delColNum)
    Call delRadioColunmRecs(neType)
End Sub


Private Sub insertNewSrcNeNameColumn(ByRef neType As String, ByRef newColNum As Long)
    Dim startColunNum As Long
    Dim startColunLetter As String
    startColunNum = getNewSrcNeNameInseterStartColunmNum(neType)
    If startColunNum = -1 Then Exit Sub
    startColunLetter = getColStr(startColunNum)
    
    If neType = "BTS" Or neType = "NodeB" Then newColNum = newColNum + 1
    Dim firstGroupName As String
    Dim secondGroupName As String
    Dim srcneName As String
    Dim bscname As String
    Dim rncName As String
    Call getAllGroupName(neType, firstGroupName, secondGroupName, srcneName, bscname, rncName)

    Dim colNum As Long
    Dim nameNum As Long
    Dim getTargetNeCol As Long
    Dim getTargetNeColletter As String
    
    getTargetNeCol = getTargetNeClounmNum(ws)
    getTargetNeColletter = getColStr(getTargetNeCol - 1)

    
    For colNum = 1 To newColNum
        ws.columns(startColunLetter & ":" & startColunLetter).Insert Shift:=xlToRight
        nameNum = CStr(newColNum - colNum)
        If neType = "eNodeB" Then nameNum = CStr(newColNum + 1 - colNum)
        If colNum = newColNum Then
            ws.Cells(2, startColunNum).value = secondGroupName
            If neType = "BTS" Then ws.Cells(3, startColunNum).value = bscname
            If neType = "NodeB" Then ws.Cells(3, startColunNum).value = rncName
            If neType = "eNodeB" Then ws.Cells(3, startColunNum).value = Replace(srcneName, "1", nameNum)
        Else
            ws.Cells(3, startColunNum).value = Replace(srcneName, "1", nameNum)
        End If
        ws.Cells(3, getTargetNeCol).Copy
        ws.Cells(3, startColunNum).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    Next
    
    Call mergeFirandSecGroupColumn(neType)
End Sub

Private Sub insertNewRadioNameColumn(ByRef neType As String)
    Dim srccolumNum As Long
    Dim srccolumLetter As String
    Dim getTargetNeCol As Long
    Dim getTargetNeColletter As String
    Dim srcNeType As String
    Dim tempNeType As String
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
    
    ws.columns(getTargetNeColletter & ":" & getTargetNeColletter).Insert Shift:=xlToRight
    
    Call getSrcColNum(srcNeType, srccolumNum)
    srccolumLetter = getColStr(srccolumNum)
    
    ws.range(getTargetNeColletter & "2").value = Replace(ws.range(srccolumLetter & "2").value, srcNeType, tempNeType)
    ws.range(getTargetNeColletter & "3").value = Replace(ws.range(srccolumLetter & "3").value, srcNeType, tempNeType)
    If ws.range(getTargetNeColletter & "1").value = "" Then
        ws.range(getTargetNeColletter & "1").value = ws.range(srccolumLetter & "1").value
        
        ws.range(srccolumLetter & "1").Copy
        ws.range(getTargetNeColletter & "1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        ws.range(srccolumLetter & "2").Copy
        ws.range(getTargetNeColletter & "2").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
        ws.range(srccolumLetter & "3").Copy
        ws.range(getTargetNeColletter & "3").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
        Application.CutCopyMode = False
    End If

    
    Dim endNum As Long
    Dim colNum As Long
    Dim firstGroupName As String
    endNum = getTargetNeCol + 1
    If isExistRadioGroup("BTS", endNum) Or isExistRadioGroup("NodeB", endNum) Or isExistRadioGroup("eNodeB", endNum) Then
        For colNum = getTargetNeCol + 1 To endNum
            If ws.Cells(1, colNum).value <> "" Then
                firstGroupName = ws.Cells(1, colNum).value
                ws.Cells(1, colNum).value = ""
            End If
        Next
        ws.range(getColStr(getTargetNeCol + 1) + "1:" + getColStr(endNum) + "1").Merge
        ws.range(getColStr(getTargetNeCol + 1) + "1:" + getColStr(endNum) + "1").FormulaR1C1 = firstGroupName
    End If
End Sub

Private Sub delRadioColunmRecs(ByRef neType As String)
    Dim srcNeType As String
    Dim srccolumNum As Long
    Dim srccellValue As String
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

Private Sub delSrcNeNameColunmRecs(ByRef neType As String, ByRef delColNum As Long)
    Dim srcColNum As Long
    Dim srcColletter As String
    Dim colNum As Long
    Dim maxNum As Long
    Dim srccellValue As String

    Dim targetCellValue As String
    Dim startCol As Long, endCol As Long, totalNum As Long
    startCol = getNeStartColNumber(neType)
    endCol = getNeEndColNumber(neType)
    totalNum = endCol - startCol + 1
    
    If totalNum = delColNum Then
        If neType = "BTS" Or neType = "NodeB" Then totalNum = totalNum + 1
        delColNum = totalNum
    End If
    
    Dim delCollter As String
    Dim startCollter As String
    Dim getTargetNeCol As Long
    
    startCollter = getColStr(startCol)
    srccellValue = ws.Cells(1, 1).value

    For colNum = 1 To delColNum
        delCollter = getColStr(endCol + 1 - colNum)
        ws.columns(delCollter & ":" & delCollter).Delete Shift:=xlToLeft
    Next
    
    getTargetNeCol = getTargetNeClounmNum(ws)
    ws.range("A1:" + getColStr(getTargetNeCol - 1) + "1").FormulaR1C1 = srccellValue

End Sub


Private Sub getAllGroupName(ByRef neType As String, ByRef firstGroupName As String, ByRef secondGroupName As String, ByRef srcneName As String, ByRef bscname As String, ByRef rncName As String)
    Dim tempNeType As String
    tempNeType = neType
    If neType = "BTS" Then tempNeType = "GBTS"
'    If isCnEnv(ws) Then
'        firstGroupName = "源网元"
'        secondGroupName = tempNeType + "/ " + tempNeType + " " + "功能"
'        srcneName = neType + "1 网元名称"
'        bscName = "BSC 网元名称"
'        rncName = "RNC 网元名称"
'    Else
'        firstGroupName = "Source NE"
'        secondGroupName = tempNeType + "/ " + tempNeType + " " + "Function"
'        srcneName = neType + "1 NE Name"
'        bscName = "BSC NE Name"
'        rncName = "RNC NE Name"
'    End If
    
    firstGroupName = getResByKey("SOURCE_NE")
    secondGroupName = tempNeType + "/ " + tempNeType + " " + getResByKey("FUNCTION")
    srcneName = neType + getResByKey("SRCNE_NAME")
    bscname = getResByKey("BSC_NENAME")
    rncName = getResByKey("RNC_NENAME")
End Sub

Private Function getNewSrcNeNameInseterStartColunmNum(ByRef neType As String) As Long
    getNewSrcNeNameInseterStartColunmNum = -1
    
    If neType = "BTS" Then
        If isExistGroup("NodeB") Then
            getNewSrcNeNameInseterStartColunmNum = getNeStartColNumber("NodeB") - 1
        ElseIf isExistGroup("eNodeB") Then
            getNewSrcNeNameInseterStartColunmNum = getNeStartColNumber("eNodeB")
        Else
            getNewSrcNeNameInseterStartColunmNum = getTargetNeClounmNum(ws)
        End If
    End If
    
    If neType = "eNodeB" Then
        getNewSrcNeNameInseterStartColunmNum = getTargetNeClounmNum(ws)
    End If
    
    If neType = "NodeB" Then
        If isExistGroup("eNodeB") Then
            getNewSrcNeNameInseterStartColunmNum = getNeStartColNumber("eNodeB")
        Else
            getNewSrcNeNameInseterStartColunmNum = getTargetNeClounmNum(ws)
        End If
    End If
End Function

Private Function getSrcNeNameNum(ByRef neType As String, ByRef nameValue As String) As Long
    Dim neTypeLen As Long
    Dim numStr As String
    getSrcNeNameNum = 1
    
    neTypeLen = Len(neType)
    numStr = Mid(nameValue, neTypeLen + 1, 1)
    getSrcNeNameNum = CLng(numStr)
    
End Function

Private Function getNeStartColNumber(ByRef neType As String) As Long
    Dim m_colNum As Long
    Dim tempCnFunctionStr As String
    Dim tempEnFunctionStr As String
    
    Dim targetNum As Long
    targetNum = getTargetNeClounmNum(ws)
    
    getNeStartColNumber = -1
    
    For m_colNum = 1 To targetNum
        If isSameNameCol(neType, ws.Cells(3, m_colNum).value) Then
            getNeStartColNumber = m_colNum
            Exit Function
        End If
    Next
End Function

Private Function getNeEndColNumber(ByRef neType As String) As Long
    Dim startColNum As Long
    Dim m_colNum As Long
    startColNum = getNeStartColNumber(neType)
    getNeEndColNumber = -1
    If startColNum = -1 Then Exit Function
    
    Dim targetNum As Long
    targetNum = getTargetNeClounmNum(ws)
    
    For m_colNum = startColNum To targetNum
        If isSameNameCol(neType, ws.Cells(3, m_colNum).value) = False Then
            getNeEndColNumber = m_colNum - 1
            Exit Function
        End If
    Next
End Function

Private Function isSameNameCol(ByRef neType As String, cellValue As String) As Boolean
    Dim num As Long
    Dim tempValue As String
'    Dim tempCnValue As String
    isSameNameCol = False
    
    
    
    For num = 1 To 50
        tempValue = neType + CStr(num) + " " + getResByKey("NE_NAME")
'        tempEnValue = neType + CStr(num) + " NE Name"
'        tempCnValue = neType + CStr(num) + " 网元名称"
        If tempValue = cellValue Then
            isSameNameCol = True
            Exit Function
        End If
    Next
End Function

Private Sub mergeFirandSecGroupColumn(ByRef neType As String)
    Dim startColNum As Long
    Dim endColNum As Long
    Dim columStartChar As String
    Dim columEndChar As String
    Dim getTargetNeCol As Long
    Dim getTargetNeColletter As String
    Dim firstGroupName As String
    Dim secondGroupName As String
    Dim colNum As Long
    
    startColNum = getNeStartColNumber(neType)
    endColNum = getNeEndColNumber(neType)
    
    If neType = "BTS" Or neType = "NodeB" Then startColNum = startColNum - 1
    
    columStartChar = getColStr(startColNum)
    columEndChar = getColStr(endColNum)


    
    getTargetNeCol = getTargetNeClounmNum(ws)
    getTargetNeColletter = getColStr(getTargetNeCol - 1)
    
    If ws.Cells(1, 1).value = "" Then
        ws.Cells(1, getTargetNeCol).Copy
        ws.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ws.Cells(1, 1).value = firstGroupName
        
        ws.Cells(2, getTargetNeCol).Copy
        ws.Cells(2, 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If
    

    For colNum = startColNum To endColNum
        If ws.Cells(2, colNum).value <> "" Then
            secondGroupName = ws.Cells(2, colNum).value
            ws.Cells(2, colNum).value = ""
        End If
    Next
    
    For colNum = 1 To getTargetNeCol - 1
        If ws.Cells(1, colNum).value <> "" Then
            firstGroupName = ws.Cells(1, colNum).value
            ws.Cells(1, colNum).value = ""
        End If
    Next
    

    ws.range(columStartChar + "2:" + columEndChar + "2").Merge
    ws.range("A1:" + getTargetNeColletter + "1").Merge
    
    ws.range("A1:" + getTargetNeColletter + "1").FormulaR1C1 = firstGroupName
    ws.range(columStartChar + "2:" + columEndChar + "2").FormulaR1C1 = secondGroupName
End Sub

Private Function getSrcNeNameCol(neType As String, cellValue As String) As String
    Dim num As Long
    Dim tempValue As String
'    Dim tempCnValue As String
    getSrcNeNameCol = ""
    
    
    
    For num = 1 To 50
        tempValue = neType + CStr(num) + " " + getResByKey("NE_NAME")
'        tempEnValue = neType + CStr(num) + " NE Name"
'        tempCnValue = neType + CStr(num) + " 网元名称"
        If tempValue = cellValue Then
            getSrcNeNameCol = neType + CStr(num)
            Exit Function
        End If
    Next
End Function

Private Sub Class_Initialize()
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        If isMigrationRelationSheet(sheet) Then
            Set ws = sheet
            Exit Sub
        End If
    Next
End Sub




