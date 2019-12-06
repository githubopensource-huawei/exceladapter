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
    For m_colNum = 1 To ws.range("XFD" + CStr(recordRow)).End(xlToLeft).column
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
    For m_rowNum = 1 To ws.range("a1048576").End(xlUp).row
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

Sub modifyRefRange(sh As Worksheet, nowSheetHyper As Collection, hisSheetHyper As Collection, address As String)
    
    Dim nowlinkType As ClinkType
    Dim hisLinkType As ClinkType
    Set nowlinkType = nowSheetHyper(address)
    Set hisLinkType = hisSheetHyper(address)
    
    '判断被应用的是否是Moc，如果是，不用改变提示
    Dim mapdef As CMappingDef
    Set mapdef = getMappingDefine(nowlinkType.linkSheetName, nowlinkType.linkGroupName, nowlinkType.linkColumName)
    If mapdef Is Nothing Or mapdef.neType = "" Then  '增加Range提示
            Dim columnName As String
            Dim groupName As String
            Dim ctlDef As New CControlDef
            groupName = get_GroupName(sh.name, nowlinkType.colNum)
            columnName = get_ColumnName(sh.name, nowlinkType.colNum)
            Set ctlDef = getControlDefine(sh.name, groupName, columnName)
            Call deleteRang(hisLinkType)
            If Not (ctlDef Is Nothing) Then
                Call addRange(nowlinkType, ctlDef)
            End If
            Call deleteRefComment(nowSheetHyper, hisLinkType)
            Call addRefComment(nowlinkType)
    End If
    
    Call hisSheetHyper.Remove(address) '删除原来引用
    hisSheetHyper.Add Item:=nowlinkType, key:=address '增加当前的引用
    Call hisHyperlinkMap.Remove(sh.name)
    hisHyperlinkMap.Add Item:=hisSheetHyper, key:=sh.name
    
End Sub
Sub addRefRange(sh As Worksheet, linktype As ClinkType, hisSheetHyper As Collection, address As String)
    
    '判断被应用的是否是Moc，如果是，不用改变提示
    Dim mapdef As CMappingDef
    Set mapdef = getMappingDefine(linktype.linkSheetName, linktype.linkGroupName, linktype.linkColumName)
    If mapdef Is Nothing Or mapdef.neType = "" Then
        Dim columnName As String
        Dim groupName As String
        Dim ctlDef As New CControlDef
        groupName = get_GroupName(sh.name, linktype.colNum)
        columnName = get_ColumnName(sh.name, linktype.colNum)
        Set ctlDef = getControlDefine(sh.name, groupName, columnName)
        If Not (ctlDef Is Nothing) Then
                Call addRange(linktype, ctlDef)
        End If
        Call addRefComment(linktype)
    End If
    
    hisSheetHyper.Add Item:=linktype, key:=address
    Call hisHyperlinkMap.Remove(sh.name)
    hisHyperlinkMap.Add Item:=hisSheetHyper, key:=sh.name
End Sub

Private Sub addRange(linktype As ClinkType, ctlDef As CControlDef)
            Dim boundValue As String
            If ctlDef.dataType = "Enum" Then
                boundValue = ctlDef.lstValue
            Else
                boundValue = ctlDef.bound
            End If
            If linktype.linkSheetName = getResByKey("Comm Data") Then
                Call setRefValidation(ctlDef.dataType, boundValue, ThisWorkbook.Worksheets(linktype.linkSheetName), linktype.linkRowNum, linktype.linkColNum)
            Else
                Dim index As Long
                For index = 3 To ThisWorkbook.Worksheets(linktype.linkSheetName).range("a1048576").End(xlUp).row
                         Call setRefValidation(ctlDef.dataType, boundValue, ThisWorkbook.Worksheets(linktype.linkSheetName), index, linktype.linkColNum)
                Next
            End If
End Sub

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

Sub deleteRefRang(sh As Worksheet, nowSheetHyper As Collection, hisSheetHyper As Collection, address As String)
    Dim mapdef As CMappingDef
    Dim linktype As ClinkType
    
    Set linktype = hisSheetHyper(address)
    
    Set mapdef = getMappingDefine(linktype.linkSheetName, linktype.linkGroupName, linktype.linkColumName)
    If mapdef Is Nothing Or mapdef.neType = "" Then
            Call deleteRang(linktype)
            Call deleteRefComment(nowSheetHyper, linktype)
    End If
    Call hisSheetHyper.Remove(address)
    Call hisHyperlinkMap.Remove(sh.name)
    hisHyperlinkMap.Add Item:=hisSheetHyper, key:=sh.name
    'If isListSheet(Sh.name) Then
    '    Call deletListRange(linktype)
    'End If
    
End Sub

Private Sub deleteRang(linktype As ClinkType)
            If linktype.linkSheetName = getResByKey("Comm Data") Then
                ThisWorkbook.Worksheets(linktype.linkSheetName).Cells(linktype.linkRowNum, linktype.linkColNum).Validation.Delete
            Else
                Dim index As Long
                For index = 3 To ThisWorkbook.Worksheets(linktype.linkSheetName).range("a1048576").End(xlUp).row
                         ThisWorkbook.Worksheets(linktype.linkSheetName).Cells(index, linktype.linkColNum).Validation.Delete
                Next
            End If
End Sub

Private Sub addRefComment(linktype As ClinkType)
            Dim refComment As comment
            Dim textComment As String
            Dim reRange As range
            Set reRange = ThisWorkbook.Worksheets(linktype.linkSheetName).Cells(linktype.linkColumRowNum, linktype.linkColNum)
            If Not (reRange Is Nothing) Then
                Set refComment = reRange.comment
                If refComment Is Nothing Then
                    textComment = getResByKey("Referenced By") + vbCr + vbLf
                    textComment = textComment + "(" + linktype.getKey + ")"
                    With reRange.addComment
                        .Visible = False
                        .text textComment
                    End With
                Else
                    textComment = addComment(refComment.text, linktype)
                    reRange.comment.Delete
                    With reRange.addComment
                        .Visible = False
                        .text textComment
                    End With
                End If
                reRange.comment.Shape.TextFrame.AutoSize = True
                reRange.comment.Shape.TextFrame.Characters.Font.Bold = True
            End If
End Sub



Private Sub deleteRefComment(nowSheetHyper As Collection, linktype As ClinkType)
            Dim refComment As comment
            Dim textComment As String
            Dim textFinal As String
            Dim reRange As range
            Set reRange = ThisWorkbook.Worksheets(linktype.linkSheetName).Cells(linktype.linkColumRowNum, linktype.linkColNum)
            If Not (reRange Is Nothing) Then
                Set refComment = reRange.comment
                If Not (refComment Is Nothing) Then
                        textFinal = deleteComment(nowSheetHyper, refComment.text, linktype)
                        refComment.Delete
                        If textFinal <> "" Then
                            With reRange.addComment
                                .Visible = False
                                .text textFinal
                            End With
                            reRange.comment.Shape.TextFrame.AutoSize = True
                            reRange.comment.Shape.TextFrame.Characters.Font.Bold = True
                        End If
                    End If
            End If
End Sub



Private Function addComment(text As String, linktype As ClinkType) As String
       Dim comments As Variant
       Dim recomments As String
       comments = Split(text, vbCr + vbLf)
       Dim index As Long
       Dim linktxt As String
       For index = 1 To UBound(comments)
                 If comments(index) <> "" Then
                    linktxt = Mid(comments(index), 2, Len(comments(index)) - 2)
                    If Trim(linktxt) <> linktype.getKey Then
                            recomments = recomments + "(" + linktxt + ")" + vbCr + vbLf
                    End If
                 End If
       Next
       recomments = comments(0) + vbCr + vbLf + recomments + "(" + linktype.getKey + ")"
       addComment = recomments
End Function

Private Function deleteComment(nowSheetHyper As Collection, text As String, linktype As ClinkType) As String
    Dim comments As Variant
    Dim recomments As String
    comments = Split(text, vbCr + vbLf)
    
    Dim index As Long
    Dim linktxt As String
    For index = 1 To UBound(comments)
            If comments(index) <> "" Then
                linktxt = Mid(comments(index), 2, Len(comments(index)) - 2)
               If Trim(linktxt) <> linktype.getKey Or isLinked(nowSheetHyper, linktype) Then
                    recomments = recomments + "(" + linktxt + ")" + vbCr + vbLf
               End If
            End If
    Next
    If recomments <> "" Then
        recomments = comments(0) + vbCr + vbLf + recomments
    End If
    deleteComment = recomments
End Function

Private Function isLinked(nowSheetHyper As Collection, linktxt As ClinkType) As Boolean
        Dim link As ClinkType
        For Each link In nowSheetHyper
                If link.sheetName = linktxt.sheetName And link.groupName = linktxt.groupName And link.columName = linktxt.columName _
                And link.linkSheetName = linktxt.linkSheetName And link.linkGroupName = linktxt.linkGroupName And link.linkColNum = linktxt.linkColNum Then
                    isLinked = True
                    Exit Function
                End If
        Next
        isLinked = False
End Function


Function Contains(coll As Collection, key As String) As Boolean
        On Error GoTo NotFound
        Call coll(key)
        Contains = True
        Exit Function
NotFound:
        Contains = False
End Function


Function CheckValueInRefRange(ByVal cCtlDef As CControlDef, ByVal attrValue As String, cellRange As range) As Boolean
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
        For index = 2 To sheetDef.range("a1048576").End(xlUp).row
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
    
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
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








