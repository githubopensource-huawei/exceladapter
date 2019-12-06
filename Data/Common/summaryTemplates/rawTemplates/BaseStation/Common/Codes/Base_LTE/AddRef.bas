Attribute VB_Name = "AddRef"
Option Explicit
Public hisHyperlinkMap  As Collection
Private controlDefineMap As Collection
Public mappingDefineMap As Collection


'AddLink
Public Const refColInModDiffSht As Integer = 1
Public Const mocColInModDiffSht As Integer = 2
Public Const attrColInModDiffSht As Integer = 3
Public Const StandardRowHeight As Double = 13.5
Public Const BluePrintSheetColor = 5


'AddLink
Public Sub AddLink()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Call addLink4NormalShts
    
    If existsASheet(getResByKey("ModelDiffSht")) Then
        Call addLink4MultiVer
    End If
    
    Application.ScreenUpdating = True
End Sub



Public Sub addLink4NormalShts()
    Application.ScreenUpdating = False
    On Error Resume Next
    Dim colRef  '引用标志位所在列号
    Dim m_rowNum
    Dim n_RowNum
    Dim textValue As String
    Dim v
    
    Dim sheetName As String  '需要添加链接的sheet
    Dim groupName As String  '需要添加链接的group
    Dim colName As String    '需要添加链接的列名
    Dim colNum As Long    '需要添加链接的列号
    
    Dim linkSheetName As String
    Dim linkGroupName As String
    Dim linkColName As String
    Dim linkColNum As String
    Dim linkRowNum As Long
    Dim linkRowAdd As Long
    
    Dim posStart
    Dim posEnd
    
    Call changeAlert(False)
    '遍历『MAPPING DEF』，获取需要添加Reference的列:Is Reference = true
    colRef = Get_RefCol("MAPPING DEF", 1, "", "Is Reference")
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim ws As Worksheet
    
    For m_rowNum = 2 To mappingDef.range("a65536").End(xlUp).row
        textValue = mappingDef.Cells(m_rowNum, colRef).value

            
        If Get_DesStr(textValue) = Get_DesStr("TRUE") Then
            sheetName = mappingDef.Cells(m_rowNum, 1).value
            groupName = mappingDef.Cells(m_rowNum, 2).value
            colName = mappingDef.Cells(m_rowNum, 3).value
            colNum = Get_RefCol(sheetName, 2, groupName, colName)
            Set ws = ThisWorkbook.Worksheets(sheetName)
            
            Dim count As Long
            count = getUsedRowCount(ws)
            For n_RowNum = 3 To count
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
                    End If
                    
                    If linkSheetName = getResByKey("Comm Data") Then
                        linkRowNum = Get_GroupRefRow(linkSheetName, linkGroupName) + 1  '获取开始的行号
                    End If
                    
                    linkColNum = "R" + CStr(linkRowNum + linkRowAdd) + "C" + CStr(Get_RefCol(linkSheetName, linkRowNum, linkGroupName, linkColName))
                    
                    ThisWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=ws.Cells(n_RowNum, colNum), address:="", SubAddress:="'" + linkSheetName + "'!" + linkColNum, TextToDisplay:=textValue
                    Call setHyperlinkRangeFont(ws.Cells(n_RowNum, colNum))
                Else
                    ws.Cells(n_RowNum, colNum).Hyperlinks.Delete
                End If
            Next
        End If
    Next
    
    Application.ScreenUpdating = True
    
End Sub

Public Sub addLink4MultiVer()
    Dim modelDiffSht As Worksheet
    Set modelDiffSht = ThisWorkbook.Worksheets(getResByKey("ModelDiffSht"))
    
    Dim rowIdx As Integer
    Dim refValue As String
    Dim refArray As Variant
    With modelDiffSht
        For rowIdx = 2 To .range("a65535").End(xlUp).row
            refValue = .Cells(rowIdx, refColInModDiffSht)
            If Not isValidReference(refValue, refArray) Then GoTo NextLoop
            
            Dim shtName As String
            Dim grpName As String
            Dim attrName As String
            shtName = refArray(0)
            grpName = refArray(1)
            attrName = refArray(2)
            
            Call addModDiffHyperLink(shtName, grpName, attrName, modelDiffSht, rowIdx)
NextLoop:
        Next
    End With
    
    Call setModDiffShtFont(modelDiffSht)
End Sub

Public Function existsASheet(shtName As String) As Boolean
On Error GoTo ErrorHandler:
    existsASheet = True
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(shtName)
    Exit Function
ErrorHandler:
    existsASheet = False
End Function


Private Sub setModDiffShtFont(sht As Worksheet)
    Dim maxRow As Long
    Dim maxCol As Long
    Dim dataRange As range
    Dim titleRange As range
    Dim linkRange As range
    Dim mocAttrRange As range
    Dim versRange As range

    With sht
        .Activate
        ActiveWindow.FreezePanes = False
        
        With .UsedRange
            maxRow = .rows.count
            maxCol = .columns.count
        End With
        
        Set dataRange = .range("A2:" & getColStr(maxCol) & maxRow)
        With dataRange
            .rows.EntireRow.RowHeight = 40
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        Set titleRange = .range("A1:" & getColStr(maxCol) & "1")
        With titleRange
            .rows.EntireRow.RowHeight = StandardRowHeight
            .Interior.colorIndex = 40
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
        
        Set linkRange = .range("A2:A" & maxRow)
        With linkRange
            .columns.EntireColumn.ColumnWidth = 40
            .Font.colorIndex = BluePrintSheetColor
            .WrapText = False
        End With
        
        Set mocAttrRange = .range("B2:C" & maxRow)
        With mocAttrRange
            .WrapText = False
            .columns.EntireColumn.AutoFit
        End With
        
        Set versRange = .range("D2:" & getColStr(maxCol) & maxRow)
        With versRange
            .columns.EntireColumn.ColumnWidth = 50
            .WrapText = True
        End With
        
        Call setBorders(.UsedRange)
    End With
End Sub


'检查是否为合法超链接格式
Public Function isValidReference(refAddr As String, Optional refArray As Variant, Optional delimeter As String) As Boolean
    isValidReference = False
    
    If delimeter <> "" Then
        refArray = Split(refAddr, delimeter)
        If UBound(refArray) <> 2 Then Exit Function
        If refArray(0) = "" Or refArray(1) = "" Or refArray(2) = "" Then Exit Function
        isValidReference = True
        Exit Function
    End If
    
    If isValidReference(refAddr, refArray, "\") Then
        isValidReference = True
        Exit Function
    End If
    
    If isValidReference(refAddr, refArray, ".") Then
        isValidReference = True
        Exit Function
    End If
End Function


Private Sub addModDiffHyperLink(shtName As String, grpName As String, attrName As String, modelDiffSht As Worksheet, ByVal srcRow As Integer)
    Dim linkRangeInModDiffSht As String
    linkRangeInModDiffSht = "'" & modelDiffSht.name & "'!" & "R" & srcRow & "C" & refColInModDiffSht
    
    Dim sht As Worksheet
    Dim targetRange As range
    Dim firstAddr As String
    Dim firstGrpName As String
    
    firstGrpName = ""
    Set sht = ThisWorkbook.Worksheets(shtName)
    With sht
        Set targetRange = .UsedRange.Find(attrName, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                firstGrpName = getGroupNameFromMappingDef(shtName, attrName, firstGrpName)
                If grpName = firstGrpName Then
                    .Hyperlinks.Add Anchor:=targetRange, address:="", SubAddress:=linkRangeInModDiffSht
                    With targetRange.Font
                        .name = "Arial"
                        .Size = 10
                    End With
                    Exit Do
                End If
                Set targetRange = .UsedRange.FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    Dim linkRangeInTargetSht As String
    linkRangeInTargetSht = "'" & shtName & "'!" & "R" & targetRange.row & "C" & targetRange.column
    With modelDiffSht
        .Hyperlinks.Add Anchor:=.range(getColStr(refColInModDiffSht) & srcRow), address:="", SubAddress:=linkRangeInTargetSht
    End With
End Sub

'An attribute may be in different groups
Public Function getGroupNameFromMappingDef(sheetName As String, attributeName As String, Optional excludeGrpName As String) As String
    Dim mappingDef As Worksheet
    Dim index, count As Long
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    count = mappingDef.UsedRange.rows.count
    With mappingDef
        For index = 1 To count
            If .Cells(index, 1).value = sheetName And .Cells(index, 3).value = attributeName Then
                getGroupNameFromMappingDef = .Cells(index, 2)
                If getGroupNameFromMappingDef <> excludeGrpName Then Exit Function
            End If
        Next
    End With
    getGroupNameFromMappingDef = ""
End Function


'从指定sheet页的指定行，查找指定列，返回列号
Function Get_RefCol(sheetName As String, recordRow As Long, groupName As String, ColValue As String) As Long
    On Error GoTo ErrorHandler
    Dim m_colNum As Long
    Dim m_GroupColNum As Long
    Dim f_flag As Boolean
    f_flag = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For m_colNum = 1 To ws.range("IV" + CStr(recordRow)).End(xlToLeft).column
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
    For m_rowNum = 1 To ws.range("a65536").End(xlUp).row
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

Public Sub setLinkRange(ByVal sh As Worksheet, ByVal target As range)
    Dim nowHyperlinksMap As Collection '当前所有link
    Dim nowSheetHyper As Collection '当前页现有的link
    Dim hisSheetHyper As Collection '当前页历史link
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
    
    If target.rows.count < sh.range("a65536").End(xlUp).row Then
        rows = target.rows.count
    Else
        rows = sh.range("a65536").End(xlUp).row
    End If
    If target.columns.count < sh.range("IV2").End(xlToLeft).column Then
        columns = target.columns.count
    Else
        columns = sh.range("IV2").End(xlToLeft).column
    End If
     
    For index = target.row To rowstart + rows - 1
        For colmindex = target.column To colmstart + columns - 1
            '变动的单元格
            address = sh.Cells(index, colmindex).address
            'address 是新加的，存在nowSheetHyper，不存在hisSheetHyper
            'address 是修改的，存在nowSheetHyper，也存在hisSheetHyper
            'address 是删除的，不存在nowSheetHyper，也存在hisSheetHyper
            'address 前后都不是link，不存在nowSheetHyper，也不存在hisSheetHyper
            If Contains(nowSheetHyper, address) And Not Contains(hisSheetHyper, address) Then
                Call addRefRange(sh, nowSheetHyper(address), hisSheetHyper, address)
            ElseIf Contains(nowSheetHyper, address) And Contains(hisSheetHyper, address) Then
                Call modifyRefRange(sh, nowSheetHyper, hisSheetHyper, address)
            ElseIf Not Contains(nowSheetHyper, address) And Contains(hisSheetHyper, address) Then
                Call deleteRefRang(sh, nowSheetHyper, hisSheetHyper, address)
            End If
        Next
    Next
End Sub
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
    hisSheetHyper.Add item:=nowlinkType, key:=address '增加当前的引用
    Call hisHyperlinkMap.Remove(sh.name)
    hisHyperlinkMap.Add item:=hisSheetHyper, key:=sh.name
    
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
    
    hisSheetHyper.Add item:=linktype, key:=address
    Call hisHyperlinkMap.Remove(sh.name)
    hisHyperlinkMap.Add item:=hisSheetHyper, key:=sh.name
End Sub

Private Sub addRange(linktype As ClinkType, ctlDef As CControlDef)
    Dim boundValue As String
    If ctlDef.dataType = "Enum" Then
        boundValue = ctlDef.lstvalue
    Else
        boundValue = ctlDef.bound
    End If
    If linktype.linkSheetName = getResByKey("Comm Data") Then
        Call setValidation(ctlDef.dataType, boundValue, Worksheets(linktype.linkSheetName), linktype.linkRowNum, linktype.linkColNum)
    Else
        Dim index As Long
        For index = 3 To Worksheets(linktype.linkSheetName).range("a65536").End(xlUp).row
            Call setValidation(ctlDef.dataType, boundValue, Worksheets(linktype.linkSheetName), index, linktype.linkColNum)
        Next
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
    hisHyperlinkMap.Add item:=hisSheetHyper, key:=sh.name
End Sub

Private Sub deleteRang(linktype As ClinkType)
    If linktype.linkSheetName = getResByKey("Comm Data") Then
        ThisWorkbook.Worksheets(linktype.linkSheetName).Cells(linktype.linkRowNum, linktype.linkColNum).Validation.Delete
    Else
        Dim index As Long
        For index = 3 To ThisWorkbook.Worksheets(linktype.linkSheetName).range("a65536").End(xlUp).row
            ThisWorkbook.Worksheets(linktype.linkSheetName).Cells(index, linktype.linkColNum).Validation.Delete
        Next
    End If
End Sub

Private Sub addRefComment(linktype As ClinkType)
    Dim oldComment As comment
    Dim textComment As String
    Dim finalComment As String
    Dim targetRange As range
    
    Dim targetShtName As String
    Dim targetGrpName As String
    Dim targetColName As String
    targetShtName = linktype.linkSheetName
    targetGrpName = linktype.linkGroupName
    targetColName = linktype.linkColumName
    
    Set targetRange = ThisWorkbook.Worksheets(targetShtName).Cells(linktype.linkColumRowNum, linktype.linkColNum)
    
    If Not targetRange Is Nothing Then
        Set oldComment = targetRange.comment
        If oldComment Is Nothing Then
            textComment = getResByKey("Referenced By") + vbCr + vbLf
            textComment = textComment + "(" + linktype.getKey + ")"
            finalComment = modifyComment(targetShtName, targetGrpName, targetColName, textComment)
            With targetRange.addComment
                .Visible = False
                .text finalComment
            End With
        Else
            If existsComment(targetShtName, targetGrpName, targetColName, oldComment.text) Then
                textComment = addComment(oldComment.text, linktype)
                finalComment = modifyComment(targetShtName, targetGrpName, targetColName, textComment, OVERWRITE)
            Else
                textComment = getComment(targetShtName, targetGrpName, targetColName)
                finalComment = addComment(textComment, linktype)
                Call modifyComment(targetShtName, targetGrpName, targetColName, finalComment, OVERWRITE)
            End If
            
            targetRange.comment.Delete
            With targetRange.addComment
                .Visible = False
                .text finalComment
            End With
        End If
        targetRange.comment.Shape.TextFrame.AutoSize = True
        targetRange.comment.Shape.TextFrame.Characters.Font.Bold = True
    End If
End Sub

Private Sub deleteRefComment(nowSheetHyper As Collection, linktype As ClinkType)
    Dim oldComment As comment
    Dim textComment As String
    Dim textFinal As String
    Dim targetRange As range
    
    Dim targetShtName As String
    Dim targetGrpName As String
    Dim targetColName As String
    targetShtName = linktype.linkSheetName
    targetGrpName = linktype.linkGroupName
    targetColName = linktype.linkColumName
    
    Set targetRange = ThisWorkbook.Worksheets(targetShtName).Cells(linktype.linkColumRowNum, linktype.linkColNum)
    If Not targetRange Is Nothing Then
        Set oldComment = targetRange.comment
        If Not oldComment Is Nothing Then
            textFinal = deleteComment(nowSheetHyper, oldComment.text, linktype)
            Call modifyComment(targetShtName, targetGrpName, targetColName, textFinal, OVERWRITE)
            oldComment.Delete
            If textFinal <> "" Then
                With targetRange.addComment
                    .Visible = False
                    .text textFinal
                End With
                targetRange.comment.Shape.TextFrame.AutoSize = True
                targetRange.comment.Shape.TextFrame.Characters.Font.Bold = True
            End If
        End If
    End If
End Sub

'去除字符串最后的换行符
Private Function getEraseVbLfStringFromEnd(ByRef cellComment As String) As String
    Dim singleChar As String, tailString As String, newCellCommentWithoutSpace As String
    Dim maxLength As Long, k As Long
    newCellCommentWithoutSpace = Trim(cellComment)
    maxLength = Len(newCellCommentWithoutSpace)
    For k = 1 To maxLength
        tailString = Right(newCellCommentWithoutSpace, k)
        singleChar = Left(tailString, 1)
        If singleChar <> vbLf And singleChar <> vbCr And singleChar <> vbCrLf Then Exit For
    Next k
    getEraseVbLfStringFromEnd = Left(newCellCommentWithoutSpace, maxLength - k + 1)
End Function

Private Sub getPrefixCommentAndAddedComment(ByRef text As String, ByRef prefixComment As String, ByRef addedComment As String)
    Dim prefixCommentTitle As String
    Dim prefixCommentTitleLength As Long
    prefixCommentTitle = getResByKey("CanBeReferencedBy")
    If InStr(text, prefixCommentTitle) = 0 Then
        prefixComment = ""
        addedComment = text '如果没有找到预定引用Title，则直接将addedComment赋值为text，与原来一样
    Else
        Dim referencedByPos As Long
        Dim referencedByString As String
        referencedByString = getResByKey("Referenced By")
        referencedByPos = InStr(Len(prefixCommentTitle), text, referencedByString)
        If referencedByPos = 0 Then
            '如果没有找到Referenced By，则直接将prefixComment赋值为text
            prefixComment = getEraseVbLfStringFromEnd(text)
            addedComment = referencedByString
        Else
            '找到了，则从找到的位置分别给prefixComment和addedComment赋值
            prefixComment = getEraseVbLfStringFromEnd(Mid(text, 1, referencedByPos - 1))
            addedComment = Mid(text, referencedByPos)
        End If
    End If
End Sub

Private Function addComment(text As String, linktype As ClinkType) As String
    Dim comments As Variant
    Dim recomments As String
    
    '新增需求非模型参数需要加提示使用范围批注，因此新增的批注要将之前的批注取出来
    Dim prefixComment As String, addedComment As String
    Call getPrefixCommentAndAddedComment(text, prefixComment, addedComment)
    
    comments = Split(addedComment, vbCr + vbLf)
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
    If prefixComment = "" Then
        '提示使用范围批注为空，则不加进来
        recomments = comments(0) + vbCr + vbLf + recomments + "(" + linktype.getKey + ")"
    Else
        recomments = prefixComment + vbCrLf + vbCrLf + comments(0) + vbCr + vbLf + recomments + "(" + linktype.getKey + ")"
    End If
    addComment = recomments
End Function

Private Function deleteComment(nowSheetHyper As Collection, text As String, linktype As ClinkType) As String
    Dim comments As Variant
    Dim recomments As String
    
    '新增需求非模型参数需要加提示使用范围批注，因此新增的批注要将之前的批注取出来
    Dim prefixComment As String, addedComment As String
    Call getPrefixCommentAndAddedComment(text, prefixComment, addedComment)
    
    comments = Split(addedComment, vbCr + vbLf)
    
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
    If prefixComment <> "" And recomments <> "" Then
        '两个都不为空，则正常加两个批注
        recomments = prefixComment + vbCrLf + vbCrLf + comments(0) + vbCr + vbLf + recomments
    ElseIf prefixComment <> "" And recomments = "" Then
        '提示批注不为空，重新增加的批注为空，则最终批注为提示批注
        recomments = prefixComment
    ElseIf prefixComment = "" And recomments <> "" Then
        '提示批注为空，则最终批注为重新增加的批注
        recomments = comments(0) + vbCr + vbLf + recomments
        '最后一种情况下，两个都为空，不需要赋值了
    End If
    deleteComment = getEraseVbLfStringFromEnd(recomments)
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

Private Function Contains(coll As Collection, key As String) As Boolean
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

Sub CheckRefedCellValidation(sheet As Worksheet, cellRange As range)
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
        For index = 2 To sheetDef.range("a65536").End(xlUp).row
                Set mpdef = New CMappingDef
                mpdef.sheetName = sheetDef.Cells(index, 1)
                mpdef.groupName = sheetDef.Cells(index, 2)
                mpdef.columnName = sheetDef.Cells(index, 3)
                mpdef.mocName = sheetDef.Cells(index, 4)
                mpdef.attributeName = sheetDef.Cells(index, 5)
                mpdef.neType = sheetDef.Cells(index, 12)
                mpdef.neVersion = sheetDef.Cells(index, 13)
                If Not Contains(defCollection, mpdef.getKey) Then
                    defCollection.Add item:=mpdef, key:=mpdef.getKey
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
        Set mp = New Collection
        
       For Each ws In ThisWorkbook.Worksheets
                Set subMap = New Collection
                'mp.SetAt strKey:=ws.name, vVal:=subMap
                mp.Add item:=subMap, key:=ws.name
                For Each h In ws.Hyperlinks
                        Set linkTp = getLinkType(h.TextToDisplay)
                        linkTp.sheetName = ws.name
                        linkTp.groupName = get_GroupName(ws.name, h.range.column)
                        linkTp.columName = ws.Cells(2, h.range.column).value
                        linkTp.colNum = h.range.column
                        linkTp.rowNum = h.range.row
                        linkTp.address = h.range.address
                        'Call subMap.SetAt(h.range.address, linkTp)
                        subMap.Add item:=linkTp, key:=h.range.address
                Next
        Next
        Set getAllHyperlinkMap = mp
End Function

Function getAllControlDefines() As Collection
    On Error Resume Next
    Dim ctlDef As CControlDef
    Dim ctrlDefSht As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    
    Set ctrlDefSht = ThisWorkbook.Worksheets("CONTROL DEF")

    Dim conInfor As String
    With ctrlDefSht
        For index = 2 To .range("a65536").End(xlUp).row
            Set ctlDef = New CControlDef
            Dim ctrlInfoItemsArray As Variant
            ctrlInfoItemsArray = .range("A" & index & ":J" & index).value
        
            ctlDef.mocName = Trim(CStr(ctrlInfoItemsArray(1, 1)))
            ctlDef.attributeName = Trim(CStr(ctrlInfoItemsArray(1, 2)))
            ctlDef.dataType = Trim(CStr(ctrlInfoItemsArray(1, 3)))
            ctlDef.bound = Trim(CStr(ctrlInfoItemsArray(1, 4)))
            ctlDef.lstvalue = Trim(CStr(ctrlInfoItemsArray(1, 5)))
            conInfor = Trim(CStr(ctrlInfoItemsArray(1, 6)))
            If isControlInfoRef(conInfor) Then conInfor = getRealControlInfo(conInfor)
            ctlDef.controlInfo = conInfor
            ctlDef.sheetName = Trim(CStr(ctrlInfoItemsArray(1, 7)))
            ctlDef.groupName = Trim(CStr(ctrlInfoItemsArray(1, 8)))
            ctlDef.columnName = Trim(CStr(ctrlInfoItemsArray(1, 9)))
            ctlDef.neType = Trim(CStr(ctrlInfoItemsArray(1, 10)))
            
            If Not Contains(defCollection, ctlDef.getKey) Then
                defCollection.Add item:=ctlDef, key:=ctlDef.getKey
            End If
        Next
    End With
    Set getAllControlDefines = defCollection
End Function


