Attribute VB_Name = "AddRef"
Option Explicit
Public hisHyperlinkMap  As Collection
Public controlDefineMap As Collection
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
    For m_colNum = 1 To ws.Range("XFD" + CStr(recordRow)).End(xlToLeft).column
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
    For m_rowNum = 1 To ws.Range("a1048576").End(xlUp).row
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
                For index = 3 To ThisWorkbook.Worksheets(linktype.linkSheetName).Range("a1048576").End(xlUp).row
                         Call setRefValidation(ctlDef.dataType, boundValue, ThisWorkbook.Worksheets(linktype.linkSheetName), index, linktype.linkColNum)
                Next
            End If
End Sub

Private Sub setRefValidation(ByVal contedType As String, ByVal boundValue As String, sheet As Worksheet, ByVal rowNum As Long, ByVal colNum As Long)
    On Error Resume Next
    Dim inputTitle As String
    inputTitle = getResByKey("Range")
    If contedType = "Enum" Then
        If boundValue <> sheet.Cells(rowNum, colNum).Validation.formula1 Or sheet.Cells(rowNum, colNum).Validation.inputMessage = "" Then
            With sheet.Cells(rowNum, colNum).Validation
                .Add Type:=xlValidateList, formula1:=boundValue
                .inputTitle = getResByKey("Range")
                .inputMessage = "[" + boundValue + "]"
                .ShowInput = True
                .ShowError = False
            End With
            sheet.Cells(rowNum, colNum).Validation.Modify Type:=xlValidateList, formula1:=boundValue
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
                For index = 3 To ThisWorkbook.Worksheets(linktype.linkSheetName).Range("a1048576").End(xlUp).row
                         ThisWorkbook.Worksheets(linktype.linkSheetName).Cells(index, linktype.linkColNum).Validation.Delete
                Next
            End If
End Sub

Private Sub addRefComment(linktype As ClinkType)
            Dim refComment As comment
            Dim textComment As String
            Dim reRange As Range
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
            Dim reRange As Range
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
        If Len(attrValue) < min Or Len(attrValue) > max Then
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
        For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
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
    
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
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

Sub addListHyperlinks()
    'Load HyperlinksForm
    ListHyperlinksForm.Show
End Sub

Sub deleteRef()

        If Not checkDeleteRange() Then
            Exit Sub
        End If
        If Not checkDeleteText() Then
            Exit Sub
        End If
        
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(GetMainSheetName())
        Dim iRange As Range
        
        For Each iRange In Selection
                iRange.Hyperlinks.Delete
                iRange.ClearContents
                Call setBorders(iRange)
                Call setHyperlinkRangeFont(iRange)
                With iRange
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .NumberFormatLocal = "@"
        End With
        Next
End Sub

Private Function checkDeleteRange() As Boolean
        Dim iRange As Range
        For Each iRange In Selection
                If iRange.Hyperlinks.count <> 1 Then
                    MsgBox getResByKey(getResByKey("deleteRefWarning")), vbExclamation + vbOKCancel, getResByKey("Warning")
                    checkDeleteRange = False
                    Exit Function
                End If
        Next
        checkDeleteRange = True
End Function

Private Function checkDeleteText() As Boolean
        Dim iRange As Range
        Dim sheetDef As Worksheet
        Dim sheetName As String
        Dim groupName As String
        Dim columnName As String
        Dim index As Long
        Dim mappingdefsheetName As String
        Dim mappingdefgroupName As String
        Dim mappingdefcolumnName As String
        Dim text As String
        Dim text1
        For Each iRange In Selection
                text = iRange.text
                If (InStr(text, "\")) Then
                    text1 = Split(text, "\")
                    sheetName = text1(0)
                    groupName = text1(1)
                    columnName = text1(2)
                    
                    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
                    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
                        mappingdefsheetName = sheetDef.Cells(index, 1)
                        mappingdefgroupName = sheetDef.Cells(index, 2)
                        mappingdefcolumnName = sheetDef.Cells(index, 3)
                        If mappingdefsheetName = sheetName And mappingdefgroupName = groupName And columnName = mappingdefcolumnName Then
                            checkDeleteText = True
                            Exit Function
                        End If
                    Next
                End If
        Next
        checkDeleteText = False
        MsgBox getResByKey(getResByKey("deleteRefWarning")), vbExclamation + vbOKCancel, getResByKey("Warning")
End Function

Public Sub getRowNumAndColumnNum(sheetName As String, groupName As String, columnName As String, rowNum As Long, columnNum As Long)
    Dim ws As Worksheet
    Dim m_rowNum As Long
    Dim m_colNum As Long
    Dim m_colNum1 As Long
    Dim columnsNum As Long
    
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If sheetName = getResByKey("Comm Data") Then
        For m_rowNum = 1 To ws.Range("a1048576").End(xlUp).row
            If groupName = ws.Cells(m_rowNum, 1).value Then
                For m_colNum = 1 To ws.Range("XFD" + CStr(m_rowNum + 1)).End(xlToLeft).column
                    If columnName = ws.Cells(m_rowNum + 1, m_colNum).value Then
                        rowNum = m_rowNum + 1
                        columnNum = m_colNum
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    Else
        For m_colNum = 1 To ws.Range("XFD1").End(xlToLeft).column
            If groupName = ws.Cells(1, m_colNum).value Then
                columnsNum = ws.Cells(1, m_colNum).MergeArea.Columns.count
                For m_colNum1 = m_colNum To m_colNum + columnsNum - 1
                    If columnName = ws.Cells(2, m_colNum1).value Then
                        rowNum = 2
                        columnNum = m_colNum1
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    End If
End Sub

Sub addGroupAndColoum(sheetName As String, groupName As String, coloumName As String, brdStyleRefRowNum As Long, brdStyleRefColNum As Long)
        Dim m_colNum As Long
        Dim groupIndex As Long
        Dim coloumStart As Long
        Dim coloumEnd As Long
        Dim columStartChar, columEndChar As String
        Dim coloumIndex As Long
        Dim refEnumStr As String
        
        Dim isfound As Boolean
        isfound = False
        groupIndex = -1
        coloumIndex = -1
        Dim index As Long
      '   For index = 0 To UBound(listRefSheet)
      '      If listRefSheet(index) = "" Then
      '          listRefSheet(index) = ActiveSheet.name
      '        Exit For
      '     End If
      ' Next
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(sheetName)
        
        For m_colNum = 1 To ws.Range("XFD" + CStr(2)).End(xlToLeft).column
            If groupName = ws.Cells(1, m_colNum).value Then
                isfound = True
                coloumStart = m_colNum
                groupIndex = m_colNum
            ElseIf ws.Cells(1, m_colNum).value <> "" And isfound = True Then
                Exit For
            End If
            coloumEnd = m_colNum
        Next
        
        Application.CutCopyMode = False
        If groupIndex > 0 Then
            For m_colNum = coloumStart To coloumEnd
                If coloumName = ws.Cells(2, m_colNum).value Then
                    coloumIndex = m_colNum
                End If
            Next
        End If
        If coloumIndex > 0 Then
            Exit Sub
        ElseIf coloumIndex <= 0 And groupIndex > 0 Then
            columEndChar = getColStr(coloumEnd + 1)
            columStartChar = getColStr(coloumStart)
            ws.Columns(columEndChar + ":" + columEndChar).Insert Shift:=xlToLeft
            
            Call clearValidationAndResetBackgroundStyle(ws, columEndChar)
            
            ws.Cells(2, coloumEnd + 1).value = coloumName
                        Application.DisplayAlerts = False
            ws.Range(columStartChar + "1:" + columEndChar + "1").Merge
                        Application.DisplayAlerts = True
            Call addGroupNameAndColoumName(sheetName, groupName, coloumName)
            refEnumStr = getRefEnumValue(brdStyleRefRowNum, brdStyleRefColNum)  '获取对应的枚举值
            If refEnumStr <> "" Then
                Call addControlDefInfo(sheetName, groupName, coloumName, refEnumStr)   '在ControlDef中新增一行记录
            End If
            setBoard (sheetName)
        ElseIf coloumIndex <= 0 And groupIndex <= 0 Then
            columEndChar = getColStr(coloumEnd + 1)
            ws.Columns(columEndChar + ":" + columEndChar).Insert Shift:=xlToLeft
            
            Call clearValidationAndResetBackgroundStyle(ws, columEndChar)

            ws.Cells(2, coloumEnd + 1).value = coloumName
            ws.Cells(1, coloumEnd + 1).value = groupName
            Call addGroupNameAndColoumName(sheetName, groupName, coloumName)
            refEnumStr = getRefEnumValue(brdStyleRefRowNum, brdStyleRefColNum)  '获取对应的枚举值
            If refEnumStr <> "" Then
                Call addControlDefInfo(sheetName, groupName, coloumName, refEnumStr)   '在ControlDef中新增一行记录
            End If
            setBoard (sheetName)
        End If
        
End Sub

Sub clearValidationAndResetBackgroundStyle(ByRef ws As Worksheet, ByRef columEndChar As String)
    Dim newColumnRange As Range
    
    Set newColumnRange = ws.Columns(columEndChar + ":" + columEndChar)
    Call clearValidation(newColumnRange)
    
    Dim maxRow As Integer
    maxRow = ws.UsedRange.Rows.count
    'Set newColumnRange = ws.range(ws.range(columEndChar & "4"), ws.range(columEndChar & "1048576"))
    Set newColumnRange = ws.Range(ws.Range(columEndChar & "4"), ws.Range(columEndChar & maxRow))
    
    Call resetBackgroundStyle(newColumnRange)
End Sub

Sub clearValidation(ByRef certainRange As Range)
    With certainRange.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Sub resetBackgroundStyle(ByRef certainRange As Range)
    certainRange.Interior.colorIndex = xlNone
    certainRange.Interior.Pattern = xlNone
    certainRange.Style = "Normal"
End Sub

Function getRefEnumValue(brdStyleRefRowNum As Long, brdStyleRefColNum As Long)
    Dim controlDef  As CControlDef
    Dim m_Str As String
    Dim groupName As String
    Dim columnName As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    getRefEnumValue = ""
    
    If getRangeGroupAndColumnName1(ws, brdStyleRefRowNum, brdStyleRefColNum, groupName, columnName) = True Then
        Set controlDef = getControlDefine(getResByKey("Board Style"), groupName, columnName)
        If controlDef Is Nothing Then
            Exit Function
        End If
        m_Str = controlDef.lstValue
        
        If Not controlDef Is Nothing Then
            On Error Resume Next
            If UCase(controlDef.dataType) = "ENUM" And controlDef.lstValue <> "" Then
                getRefEnumValue = m_Str
            End If
        End If
    End If
End Function

Sub addControlDefInfo(sheetName As String, groupName As String, coloumName As String, refEnumStr As String)
    Dim index As Long
    Dim row As Long
    Dim controlDef As Worksheet
    row = -1
    Set controlDef = ThisWorkbook.Worksheets("CONTROL DEF")
    For index = 2 To controlDef.Range("a1048576").End(xlUp).row
        row = index
        If controlDef.Cells(index, 7).value = sheetName _
        And controlDef.Cells(index, 8).value = groupName _
        And controlDef.Cells(index, 9).value = coloumName Then
            Exit For
        End If
    Next
    
    controlDef.Rows(row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    controlDef.Cells(row + 1, 1).value = groupName + "_CUSTOMIZE_REF"
    controlDef.Cells(row + 1, 2).value = coloumName + "_CUSTOMIZE_REF"
    controlDef.Cells(row + 1, 3).value = "Enum"
    controlDef.Cells(row + 1, 5).value = refEnumStr
    controlDef.Cells(row + 1, 7).value = sheetName
    controlDef.Cells(row + 1, 8).value = groupName
    controlDef.Cells(row + 1, 9).value = coloumName
    controlDef.Cells(row + 1, 10).value = "GSM Controller"
    
    Dim cndef As CControlDef
    Set cndef = New CControlDef
    cndef.sheetName = sheetName
    cndef.groupName = groupName
    cndef.columnName = coloumName
    If Not Contains(controlDefineMap, cndef.getKey) And Not controlDefineMap Is Nothing Then
        controlDefineMap.Add Item:=cndef, key:=cndef.getKey
    End If

End Sub

Sub addGroupNameAndColoumName(sheetName As String, groupName As String, coloumName As String)
        Dim index As Long
        Dim row As Long
        Dim neVersion As String
        Dim neType As String
        
        row = -1
        Dim mappingDef As Worksheet
        Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
        For index = 2 To mappingDef.Range("a1048576").End(xlUp).row
            row = index
            If mappingDef.Cells(index, 1).value = sheetName _
            And mappingDef.Cells(index, 2).value = groupName _
            And mappingDef.Cells(index, 3).value = coloumName _
            And mappingDef.Cells(index, 4).value = groupName + "_CUSTOMIZE_REF" _
            And mappingDef.Cells(index, 5).value = coloumName + "_CUSTOMIZE_REF" Then
                Exit For
            End If
        Next
        mappingDef.Rows(row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        mappingDef.Cells(row + 1, 1).value = sheetName
        mappingDef.Cells(row + 1, 2).value = groupName
        mappingDef.Cells(row + 1, 3).value = coloumName
        mappingDef.Cells(row + 1, 4).value = groupName + "_CUSTOMIZE_REF"
        mappingDef.Cells(row + 1, 5).value = coloumName + "_CUSTOMIZE_REF"
        mappingDef.Cells(row + 1, 9).value = "TRUE"
        neType = mappingDef.Cells(row, 12).value
        neVersion = mappingDef.Cells(row, 13).value
        mappingDef.Cells(row + 1, 12).value = neType
        mappingDef.Cells(row + 1, 13).value = neVersion
        Dim mpdef As CMappingDef
         Set mpdef = New CMappingDef
                mpdef.sheetName = sheetName
                mpdef.groupName = groupName
                mpdef.columnName = coloumName
            If Not Contains(mappingDefineMap, mpdef.getKey) And Not mappingDefineMap Is Nothing Then
                mappingDefineMap.Add Item:=mpdef, key:=mpdef.getKey
            End If
End Sub
Sub setBoard(sheetName As String)
    Dim maxRow As Long
    Dim maxColomn As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    maxColomn = ws.Range("XFD2").End(xlToLeft).column
    'maxRow = ws.Range("a1048576").End(xlUp).row
    maxRow = getMaxRowNumberWithBorder(ws) '获得当前有边框的最大行
    If maxRow = 0 Then Exit Sub
    Dim myRange As Range
    Set myRange = ws.Range("A1:" + getColStr(maxColomn) + CStr(maxRow))
    'Set myRange = ws.UsedRange
    Call setRangeBoard(myRange)
End Sub

Private Function getMaxRowNumberWithBorder(ByRef ws As Worksheet, Optional ByVal columnletter As String = "A") As Long
    Dim maxRowNumber As Long
    maxRowNumber = ws.UsedRange.Rows.count
    getMaxRowNumberWithBorder = maxRowNumber
    Dim rowNumber As Long
    For rowNumber = 1 To maxRowNumber
        If rangeHasBorder(ws.Range(columnletter & rowNumber)) = False Then
            getMaxRowNumberWithBorder = rowNumber - 1
            Exit Function
        End If
    Next rowNumber
End Function

Sub setRangeBoard(myRange As Range)
    With myRange
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .colorIndex = xlAutomatic
                    '.TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .colorIndex = xlAutomatic
                    '.TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .colorIndex = xlAutomatic
                    '.TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .colorIndex = xlAutomatic
                    '.TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Borders(xlInsideVertical)
                    .Weight = xlThin
                End With
               With .Borders(xlInsideHorizontal)
                    .Weight = xlThin
                End With
                .NumberFormatLocal = "@"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
    End With
    
End Sub




Public Sub setHyperlinkRangeFont(ByRef certainRange As Range)
    With certainRange.Font
        .name = "Arial"
        .Size = 10
    End With
End Sub




