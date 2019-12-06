Attribute VB_Name = "AddRef"
Option Explicit
Public hisHyperlinkMap  As Collection
Private controlDefineMap As Collection
Public mappingDefineMap As Collection

'AddLink
Sub AddLink()
    On Error Resume Next
    Dim colRef  '���ñ�־λ�����к�
    Dim m_rowNum
    Dim n_RowNum
    Dim textValue As String
    Dim v
    
    Dim sheetName As String  '��Ҫ������ӵ�sheet
    Dim groupName As String  '��Ҫ������ӵ�group
    Dim ColName As String    '��Ҫ������ӵ�����
    Dim colNum As Long    '��Ҫ������ӵ��к�
    
    Dim linkSheetName As String
    Dim linkGroupName As String
    Dim linkColName As String
    Dim linkColNum As String
    Dim linkRowNum As Long
    Dim linkRowAdd As Long
    
    Dim posStart
    Dim posEnd

    '������MAPPING DEF������ȡ��Ҫ���Reference����:Is Reference = true
    colRef = Get_RefCol("MAPPING DEF", 1, "", "Is Reference")
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim ws As Worksheet
    
    For m_rowNum = 2 To mappingDef.Range("a65536").End(xlUp).row
        textValue = mappingDef.Cells(m_rowNum, colRef).value

            
        If Get_DesStr(textValue) = Get_DesStr("TRUE") Then
            sheetName = mappingDef.Cells(m_rowNum, 1).value
            groupName = mappingDef.Cells(m_rowNum, 2).value
            ColName = mappingDef.Cells(m_rowNum, 3).value
            colNum = Get_RefCol(sheetName, 2, groupName, ColName)
            Set ws = ThisWorkbook.Worksheets(sheetName)
            
            For n_RowNum = 3 To ws.Range("a65536").End(xlUp).row
                textValue = ws.Cells(n_RowNum, colNum).value
                
                 If Is_Ref(textValue, "\") Then
                    v = Split(textValue, "\")
                    linkSheetName = v(0)
                    linkGroupName = v(1)
                    linkColName = v(2)
                    linkRowNum = 2
                    linkRowAdd = 0

                    If InStr(linkColName, "[") <> 0 Then
                        linkRowAdd = CLng(Mid(linkColName, InStr(linkColName, "[") + 1, InStr(linkColName, "]") - InStr(linkColName, "[") - 1)) + 1
                        linkColName = Mid(linkColName, 1, InStr(linkColName, "[") - 1)
                        'MsgBox linkColName & "^^^^" & linkRowNum, vbExclamation, "Warning"
                    End If
                    
                    If linkSheetName = getResByKey("Comm Data") Then
                        linkRowNum = Get_GroupRefRow(linkSheetName, linkGroupName) + 1  '��ȡ��ʼ���к�
                    End If
                    
                    linkColNum = "R" + CStr(linkRowNum + linkRowAdd) + "C" + CStr(Get_RefCol(linkSheetName, linkRowNum, linkGroupName, linkColName))
                    
                    ThisWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=ws.Cells(n_RowNum, colNum), address:="", SubAddress:="'" + linkSheetName + "'!" + linkColNum, TextToDisplay:=textValue
                ElseIf Is_Ref(textValue, ".") Then
                    v = Split(textValue, ".")
                    linkSheetName = v(0)
                    linkGroupName = v(1)
                    linkColName = v(2)
                    linkRowNum = 2
                    linkRowAdd = 0

                    If InStr(linkColName, "[") <> 0 Then
                        linkRowAdd = CLng(Mid(linkColName, InStr(linkColName, "[") + 1, InStr(linkColName, "]") - InStr(linkColName, "[") - 1)) + 1
                        linkColName = Mid(linkColName, 1, InStr(linkColName, "[") - 1)
                        'MsgBox linkColName & "^^^^" & linkRowNum, vbExclamation, "Warning"
                    End If
                    
                    If linkSheetName = getResByKey("Comm Data") Then
                        linkRowNum = Get_GroupRefRow(linkSheetName, linkGroupName) + 1
                    End If
                    
                    linkColNum = "R" + CStr(linkRowNum + linkRowAdd) + "C" + CStr(Get_RefCol(linkSheetName, linkRowNum, linkGroupName, linkColName))
                    
                   ThisWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=ws.Cells(n_RowNum, colNum), address:="", SubAddress:="'" + linkSheetName + "'!" + linkColNum, TextToDisplay:=textValue
                Else
                    ws.Cells(n_RowNum, colNum).Hyperlinks.Delete
                End If
            Next
        End If
    Next
    
End Sub

'��ָ��sheetҳ��ָ���У�����ָ���У������к�
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
        'MsgBox sheetName & "ȱ���У�" & ColValue, vbExclamation, "Warning"
    Else
        Get_RefCol = m_colNum
    End If
    Exit Function
ErrorHandler:
    Get_RefCol = -1
End Function

'��ָ��sheetҳ����group������
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
        'MsgBox sheetName & "ȱ��Group��" & groupName, vbExclamation, "Warning"
    End If
    
    Get_GroupRefRow = m_rowNum
    
End Function


'���Ƚ��ַ�������
Function Get_DesStr(srcStr As String) As String
    Get_DesStr = UCase(Trim(srcStr))
End Function

'����Ƿ�Ϊ�Ϸ������Ӹ�ʽ
Function Is_Ref(celValue As String, splitStr As String) As Boolean
    Dim v
    Dim refFlag As Boolean
    refFlag = True
    
    v = Split(celValue, splitStr)
    Dim I As Long
    For I = 0 To UBound(v)
        If v(I) = "" Then
            refFlag = False
        End If
    Next I
        
    If I <> 3 Then
        refFlag = False
    End If
    
    Is_Ref = refFlag
End Function


Function getLinkType(ByVal textValue As String) As ClinkType
        Dim linkTp As New ClinkType
        Dim linkSheetName As String
        Dim linkGroupName As String
        Dim linkColName As String
        Dim linkRowNum As Long
        Dim linkColNum, linkRowAdd As Long
        Dim v
        If Is_Ref(textValue, "\") Then
                    v = Split(textValue, "\")
                    linkSheetName = v(0)
                    linkGroupName = v(1)
                    linkColName = v(2)
        ElseIf Is_Ref(textValue, ".") Then
                    v = Split(textValue, ".")
                    linkSheetName = v(0)
                    linkGroupName = v(1)
                    linkColName = v(2)
        End If
        linkRowNum = 2
        
        If InStr(linkColName, "[") <> 0 Then
            linkRowAdd = CLng(Mid(linkColName, InStr(linkColName, "[") + 1, InStr(linkColName, "]") - InStr(linkColName, "[") - 1))
            linkColName = Mid(linkColName, 1, InStr(linkColName, "[") - 1)
            'MsgBox linkColName & "^^^^" & linkRowNum, vbExclamation, "Warning"
        End If
            
        If linkSheetName = getResByKey("Comm Data") Then
            linkRowNum = Get_GroupRefRow(linkSheetName, linkGroupName) + 1
        End If
        
        linkTp.linkColumRowNum = linkRowNum
        
        linkColNum = Get_RefCol(linkSheetName, linkRowNum, linkGroupName, linkColName)
        linkRowNum = linkRowNum + linkRowAdd + 1
        
        linkTp.linkColNum = linkColNum
        linkTp.linkColumName = linkColName
        linkTp.linkGroupName = linkGroupName
        linkTp.linkRowNum = linkRowNum
        linkTp.linkSheetName = linkSheetName
        Set getLinkType = linkTp
End Function

Public Sub setLinkRange(ByVal sh As Worksheet, ByVal target As Range)

        Dim nowHyperlinksMap As Collection '��ǰ����link
        Dim nowSheetHyper As Collection '��ǰҳ���е�link
        Dim hisSheetHyper As Collection '��ǰҳ��ʷlink
        Dim hasKey As Boolean
        Set nowHyperlinksMap = getAllHyperlinkMap()
        If Contains(nowHyperlinksMap, sh.name) Then
            Set nowSheetHyper = nowHyperlinksMap(sh.name)
        Else
            Set nowSheetHyper = New Collection
        End If
        
        If Contains(hisHyperlinkMap, sh.name) Then
            Set hisSheetHyper = hisHyperlinkMap(sh.name)
        Else
            Set hisSheetHyper = New Collection
        End If
        
        
         Dim rows As Long
         Dim columns As Long
         Dim rowstart, index As Long
         Dim colmstart, colmindex As Long
         Dim address As String
         rowstart = target.row
         colmstart = target.column
         
         If target.rows.count < sh.Range("a65536").End(xlUp).row Then
            rows = target.rows.count
        Else
            rows = sh.Range("a65536").End(xlUp).row
        End If
         If target.columns.count < sh.Range("IV2").End(xlToLeft).column Then
            columns = target.columns.count
        Else
            columns = sh.Range("IV2").End(xlToLeft).column
        End If
         
         For index = target.row To rowstart + rows - 1
                For colmindex = target.column To colmstart + columns - 1
                    '�䶯�ĵ�Ԫ��
                    address = sh.Cells(index, colmindex).address
                    'address ���¼ӵģ�����nowSheetHyper��������hisSheetHyper
                    'address ���޸ĵģ�����nowSheetHyper��Ҳ����hisSheetHyper
                    'address ��ɾ���ģ�������nowSheetHyper��Ҳ����hisSheetHyper
                    'address ǰ�󶼲���link��������nowSheetHyper��Ҳ������hisSheetHyper
                    If Contains(nowSheetHyper, address) And Not Contains(hisSheetHyper, address) Then

                        Call addRefRange(sh, nowSheetHyper(address), hisSheetHyper, address)
                    ElseIf Contains(nowSheetHyper, address) And Contains(hisSheetHyper, address) Then

                        Call modifyRefRange(sh, nowSheetHyper, hisSheetHyper, address)
                    ElseIf Not Contains(nowSheetHyper, address) And Contains(hisSheetHyper, address) Then

                        Call deleteRefRang(sh, nowSheetHyper, hisSheetHyper, address)
                    Else

                    End If
                Next
         Next
End Sub
Sub modifyRefRange(sh As Worksheet, nowSheetHyper As Collection, hisSheetHyper As Collection, address As String)
    
    Dim nowlinkType As ClinkType
    Dim hisLinkType As ClinkType
    Set nowlinkType = nowSheetHyper(address)
    Set hisLinkType = hisSheetHyper(address)
    
    '�жϱ�Ӧ�õ��Ƿ���Moc������ǣ����øı���ʾ
    Dim mapdef As CMappingDef
    Set mapdef = getMappingDefine(nowlinkType.linkSheetName, nowlinkType.linkGroupName, nowlinkType.linkColumName)
    If mapdef Is Nothing Or mapdef.neType = "" Then  '����Range��ʾ
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
    
    Call hisSheetHyper.Remove(address) 'ɾ��ԭ������
    hisSheetHyper.Add Item:=nowlinkType, key:=address '���ӵ�ǰ������
    Call hisHyperlinkMap.Remove(sh.name)
    hisHyperlinkMap.Add Item:=hisSheetHyper, key:=sh.name
    
End Sub
Sub addRefRange(sh As Worksheet, linktype As ClinkType, hisSheetHyper As Collection, address As String)
    
    '�жϱ�Ӧ�õ��Ƿ���Moc������ǣ����øı���ʾ
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
                For index = 3 To ThisWorkbook.Worksheets(linktype.linkSheetName).Range("a65536").End(xlUp).row
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
                For index = 3 To ThisWorkbook.Worksheets(linktype.linkSheetName).Range("a65536").End(xlUp).row
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
    Dim I, nResponse, nLoop As Long
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
        For I = 0 To UBound(arrayList)
            If Trim(attrValue) = arrayList(I) Then
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
    Else  '��ֵ
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

Sub CheckRefedCellValidation(sheet As Worksheet, cellRange As Range)
        On Error Resume Next
        Dim sheetMap As Collection
        Dim clink   As ClinkType
        Dim cCtlDef As CControlDef
        Dim mainSheet As String
        Dim commSheet As String
        Dim groupName As String
        Dim parentGroupName As String
        Dim columnName As String
        Dim parentColumnNam As String
        Dim isvalide As Boolean
        mainSheet = GetMainSheetName()
        commSheet = GetCommonSheetName()
        
        If (sheet.name <> mainSheet And sheet.name <> commSheet) Or cellRange.value = "" Then
            Exit Sub
        End If
        
        If hisHyperlinkMap Is Nothing Then
             Call initHisHyperlinkMap
        End If

        For Each sheetMap In hisHyperlinkMap
                For Each clink In sheetMap
                        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
                        
                        If sheet.name = mainSheet Then
                                If clink.linkSheetName = sheet.name And clink.linkGroupName = groupName And clink.linkColumName = columnName Then
                                        parentGroupName = get_GroupName(clink.sheetName, clink.colNum)
                                        parentColumnNam = get_ColumnName(clink.sheetName, clink.colNum)
                                        Set cCtlDef = getControlDefine(clink.sheetName, parentGroupName, parentColumnNam)
                                        If Not (cCtlDef Is Nothing) Then
                                                isvalide = CheckValueInRefRange(cCtlDef, cellRange.value, cellRange)
                                                If isvalide = False Then
                                                    Exit Sub
                                                End If
                                        End If
                                End If
                        Else
                                If clink.linkSheetName = sheet.name And clink.linkGroupName = groupName And clink.linkColumName = columnName _
                                And clink.linkRowNum = cellRange.row Then
                                        parentGroupName = get_GroupName(clink.sheetName, clink.colNum)
                                        parentColumnNam = get_ColumnName(clink.sheetName, clink.colNum)
                                        Set cCtlDef = getControlDefine(clink.sheetName, parentGroupName, parentColumnNam)
                                        If Not (cCtlDef Is Nothing) Then
                                                isvalide = CheckValueInRefRange(cCtlDef, cellRange.value, cellRange)
                                                If isvalide = False Then
                                                    Exit Sub
                                                End If
                                        End If
                                End If
                        End If
                Next
        Next
End Sub

Sub initAddRef()
    initHisHyperlinkMap
    initControlDefineMap
    initmappingDefineMap
End Sub
Sub initHisHyperlinkMap()
        Set hisHyperlinkMap = getAllHyperlinkMap()
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

Function getAllHyperlinkMap() As Collection
       Dim ws As Worksheet
       Dim h As Hyperlink
       Dim mp As Collection
       Dim subMap As Collection
       Dim linkTp As ClinkType
       Dim sheetType As String
        Set mp = New Collection
        
       For Each ws In ThisWorkbook.Worksheets
                Set subMap = New Collection
                'mp.SetAt strKey:=ws.name, vVal:=subMap
                mp.Add Item:=subMap, key:=ws.name
                For Each h In ws.Hyperlinks
                        Set linkTp = getLinkType(h.TextToDisplay)
                        linkTp.sheetName = ws.name
                        linkTp.groupName = get_GroupName(ws.name, h.Range.column)
                        linkTp.columName = ws.Cells(2, h.Range.column).value
                        linkTp.colNum = h.Range.column
                        linkTp.rowNum = h.Range.row
                        linkTp.address = h.Range.address
                        'Call subMap.SetAt(h.range.address, linkTp)
                        subMap.Add Item:=linkTp, key:=h.Range.address
                Next
           ''     sheetType = getSheetType(ws.name)
             ''   If sheetType = "LIST" And isCellSheet(ws.name) = False Then
            ''            Dim index as long
             ''              For index = 0 To UBound(listRefSheet)
            ''                   If listRefSheet(index) = "" Then
            ''                       listRefSheet(index) = ws.name
            ''                       Exit For
             ''                   End If
              ''             Next
                ''End If
        Next
        Set getAllHyperlinkMap = mp
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








