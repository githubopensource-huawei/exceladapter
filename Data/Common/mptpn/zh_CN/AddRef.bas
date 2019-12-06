Attribute VB_Name = "AddRef"
Option Explicit
Public hisHyperlinkMap  As Collection
Private controlDefineMap As Collection
Public mappingDefineMap As Collection

'从指定sheet页的指定行，查找指定列，返回列号
Function Get_RefCol(sheetName As String, recordRow As Long, groupName As String, ColValue As String) As Long
    On Error GoTo ErrorHandler
    Dim m_colNum As Long
    Dim m_GroupColNum As Long
    Dim f_flag As Boolean
    f_flag = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For m_colNum = 1 To ws.Range("IV" + CStr(recordRow)).End(xlToLeft).column
        If Get_DesStr(ColValue) = Get_DesStr(ws.Cells(recordRow, m_colNum).value) Then
            If groupName = "" Then
                f_flag = True
                Exit For
            Else
                m_GroupColNum = m_colNum
                While Get_DesStr(ws.Cells(recordRow - 1, m_GroupColNum).value) = ""
                    m_GroupColNum = m_GroupColNum - 1
                Wend
                If Get_DesStr(groupName) = Get_DesStr(ws.Cells(recordRow - 1, m_GroupColNum).value) Then
                    f_flag = True
                    Exit For
                End If
            End If
        End If
    Next
    If f_flag = False Then
        'MsgBox sheetName & "缺少列：" & ColValue, vbExclamation, "Warning"
    Else
        Get_RefCol = m_colNum
    End If
    Exit Function
ErrorHandler:
    Get_RefCol = -1
End Function

'从指定sheet页查找group所在行
Function Get_GroupRefRow(sheetName As String, groupName As String) As Long
    Dim m_rowNum As Long
    Dim f_flag As Boolean
    f_flag = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For m_rowNum = 1 To ws.Range("a65536").End(xlUp).row
        If GetDesStr(groupName) = GetDesStr(ws.Cells(m_rowNum, 1).value) Then
            f_flag = True
            Exit For
        End If
    Next
    If f_flag = False Then
        'MsgBox sheetName & "缺少Group：" & groupName, vbExclamation, "Warning"
    End If
    
    Get_GroupRefRow = m_rowNum
    
End Function


'将比较字符串整形
Function Get_DesStr(srcStr As String) As String
    Get_DesStr = UCase(Trim(srcStr))
End Function

'检查是否为合法超链接格式
Function Is_Ref(celValue As String, splitStr As String) As Boolean
    Dim v
    Dim refFlag As Boolean
    refFlag = True
    
    v = Split(celValue, splitStr)
    Dim i As Long
    For i = 0 To UBound(v)
        If v(i) = "" Then
            refFlag = False
        End If
    Next i
        
    If i <> 3 Then
        refFlag = False
    End If
    
    Is_Ref = refFlag
End Function

Private Sub setRefValidation(ByVal contedType As String, ByVal boundValue As String, sheet As Worksheet, ByVal rowNum As Long, ByVal colNum As Long)
    On Error Resume Next
    Dim inputTitle As String
    inputTitle = getResByKey("Range")
    If contedType = "Enum" Then
        If boundValue <> sheet.Cells(rowNum, colNum).Validation.Formula1 Or sheet.Cells(rowNum, colNum).Validation.inputMessage = "" Then
            With sheet.Cells(rowNum, colNum).Validation
                .Add Type:=xlValidateList, Formula1:=boundValue
                .inputTitle = getResByKey("Range")
                .inputMessage = "[" + boundValue + "]"
                .ShowInput = True
                .ShowError = False
            End With
            sheet.Cells(rowNum, colNum).Validation.Modify Type:=xlValidateList, Formula1:=boundValue
        End If
    End If
End Sub

Function Contains(coll As Collection, key As String) As Boolean
        On Error GoTo NotFound
        Call coll(key)
        Contains = True
        Exit Function
NotFound:
        Contains = False
End Function


Function CheckValueInRefRange(ByVal cCtlDef As CControlDef, ByVal attrValue As String, cellRange As Range) As Boolean
    Dim attrType As String
    Dim attrRange As String
    Dim arrayList() As String
    Dim errorMsg, sItem As String
    Dim i, nResponse, nLoop As Long
    Dim min, max As Long
    
    attrType = cCtlDef.dataType
    attrRange = cCtlDef.bound
    
    If attrRange = "" Then
        CheckValueInRefRange = True
        Exit Function
    End If
    
    If attrType = "Enum" Then
        CheckValueInRefRange = False
        arrayList = Split(attrRange, ",")
        For i = 0 To UBound(arrayList)
            If Trim(attrValue) = arrayList(i) Then
                CheckValueInRefRange = True
                Exit For
            End If
        Next
        errorMsg = getResByKey("Range") + "[" + attrRange + "]"
    ElseIf attrType = "String" Or attrType = "Password" Or attrType = "ATM" Then
        min = CLng(Mid(attrRange, 2, InStr(1, attrRange, ",") - 2))
        max = CLng(Mid(attrRange, InStr(1, attrRange, ",") + 1, InStr(1, attrRange, "]") - InStr(1, attrRange, ",") - 1))
        If LenB(StrConv(attrValue, vbFromUnicode)) < min Or LenB(StrConv(attrValue, vbFromUnicode)) > max Then
            CheckValueInRefRange = False
        Else
            CheckValueInRefRange = True
        End If
        If min = max Then
            errorMsg = getResByKey("Limited Length") + "[" + CStr(min) + "]"
        Else
            errorMsg = getResByKey("Limited Length") + Replace(attrRange, ",", "~")
        End If
    ElseIf attrType = "IPV4" Or attrType = "IPV6" _
        Or attrType = "Time" Or attrType = "Date" _
        Or attrType = "DateTime" Or attrType = "Bitmap" _
        Or attrType = "Mac" Then
        CheckValueInRefRange = False
        Exit Function
    Else  '数值
        If Check_Int_Validation(attrRange, attrValue) = True Then
            CheckValueInRefRange = True
        Else
            CheckValueInRefRange = False
        End If
        errorMsg = getResByKey("Range") + formatRange(attrRange)
    End If
    
    If CheckValueInRefRange = False Then
        errorMsg = getResByKey("Referenced By") + cCtlDef.groupName + "," + cCtlDef.sheetName + "," + cCtlDef.columnName + vbCr + vbLf + errorMsg
        nResponse = MsgBox(errorMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
        If nResponse = vbRetry Then
            cellRange.Select
        End If
        cellRange.value = ""
    End If
End Function

Sub initAddRef()
    initControlDefineMap
    initmappingDefineMap
End Sub
Sub initControlDefineMap()
        Set controlDefineMap = getAllControlDefines()
End Sub
Sub initmappingDefineMap()
        Set mappingDefineMap = getAllMappingDefs()
End Sub
Function getControlDefine(sheetName As String, groupName As String, columnName As String) As CControlDef
        Dim key As String
        Dim def As CControlDef
        
        key = sheetName + "," + groupName + "," + columnName
        If controlDefineMap Is Nothing Then
            initControlDefineMap
        End If
        
        If Contains(controlDefineMap, key) Then
            Set def = controlDefineMap(key)
        End If
        Set getControlDefine = def
End Function


Function getMappingDefine(sheetName As String, groupName As String, columnName As String) As CMappingDef
        Dim key As String
        Dim def As CMappingDef
        
        key = sheetName + "," + groupName + "," + columnName
        If mappingDefineMap Is Nothing Then
            initmappingDefineMap
        End If
        
        If Contains(mappingDefineMap, key) Then
            Set def = mappingDefineMap(key)
        End If
        Set getMappingDefine = def
End Function

Function getAllMappingDefs() As Collection
        Dim mp As Collection
        Dim mpdef As CMappingDef
        Dim sheetDef As Worksheet
        Dim index As Long
        Dim defCollection As New Collection
        Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
        For index = 2 To sheetDef.Range("a65536").End(xlUp).row
                Set mpdef = New CMappingDef
                mpdef.sheetName = sheetDef.Cells(index, 1)
                mpdef.groupName = sheetDef.Cells(index, 2)
                mpdef.columnName = sheetDef.Cells(index, 3)
                mpdef.mocName = sheetDef.Cells(index, 4)
                mpdef.attributeName = sheetDef.Cells(index, 5)
                mpdef.neType = sheetDef.Cells(index, 12)
                mpdef.neVersion = sheetDef.Cells(index, 13)
                If Not Contains(defCollection, mpdef.getKey) Then
                    defCollection.Add Item:=mpdef, key:=mpdef.getKey
                End If
        Next
        Set getAllMappingDefs = defCollection
End Function

Function getAllControlDefines() As Collection
    On Error Resume Next
    Dim ctlDef As CControlDef
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    
    Set sheetDef = ThisWorkbook.Worksheets("CONTROL DEF")
    
    For index = 2 To sheetDef.Range("a65536").End(xlUp).row
            Set ctlDef = New CControlDef
            ctlDef.mocName = sheetDef.Cells(index, 1).value
            ctlDef.attributeName = sheetDef.Cells(index, 2).value
            ctlDef.dataType = sheetDef.Cells(index, 3).value
            ctlDef.bound = sheetDef.Cells(index, 4).value
            ctlDef.lstValue = sheetDef.Cells(index, 5).value
            ctlDef.controlInfo = sheetDef.Cells(index, 6).value
            ctlDef.sheetName = sheetDef.Cells(index, 7).value
            ctlDef.groupName = sheetDef.Cells(index, 8).value
            ctlDef.columnName = sheetDef.Cells(index, 9).value
            ctlDef.neType = sheetDef.Cells(index, 10).value
            If Not Contains(defCollection, ctlDef.getKey) Then
                defCollection.Add Item:=ctlDef, key:=ctlDef.getKey
            End If
    Next
    Set getAllControlDefines = defCollection
End Function








