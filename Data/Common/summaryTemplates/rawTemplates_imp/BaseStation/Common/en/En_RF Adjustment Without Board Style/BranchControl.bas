Attribute VB_Name = "BranchControl"
Option Explicit
Public Type controlRelation
    mocName As String
    contAttr As String
    contedAttrs(100) As String
    contedNum As Long '从0开始
    neType  As String '此参数是为了区分 是控制器的还是物理站的，现在控制器和物理站有可能Moc和参数属性名是一样的
    sheetName As String ' 此参数是为了区分，同一Moc出现在不同的Sheet，Comm Data页会出现，这时候，不能跨sheet页签控制
End Type
Public ControlRelMap(1000) As controlRelation
Public contAttrValArray(100) As String
Public ControlRelationNum As Long

Public controlRelationManager As CControlRelationManager
Public mappingNumberManager As CMappingNumberManager
Public actualBoardStyleName As String

'用以设置颜色
Public Const SolidColorIdx = 16
Public Const SolidPattern = xlGray16
Public Const NullPattern = xlNone
Public Const NormalPattern = 1
'BoardStyle页签组和列信息
Public boardStyleGroupMap As CMap
Public boardStyleColumnMap As CMap

Function isControlDefSheetExist() As Boolean
    Dim SheetNum As Long
    isControlDefSheetExist = False
    For SheetNum = 1 To ThisWorkbook.Worksheets.count
        If "CONTROL DEF" = ThisWorkbook.Worksheets(SheetNum).name Then
            isControlDefSheetExist = True
            Exit For
        End If
    Next
End Function

Sub getGroupAndColumnName(ByVal CurSheet As Worksheet, ByVal cellRange As range, ByRef groupName As String, ByRef columnName As String, Optional groupStartRow As Long = -1)
    Dim m_rowNum, m_colNum As Long
    Dim m_endRowNum As Long
    
    'DTS2017022204401
    If groupStartRow <> -1 Then
        m_endRowNum = groupStartRow
    Else
        m_endRowNum = 1
    End If
    
    If CurSheet.name = getResByKey("Comm Data") Or InStr(CurSheet.name, getResByKey("Board Style")) <> 0 Then
        For m_rowNum = cellRange.row To m_endRowNum Step -1
            If boardStyleColumnMap.hasKey(Trim(CurSheet.Cells(m_rowNum, cellRange.column).value)) Then '34是绿色
                columnName = CurSheet.Cells(m_rowNum, cellRange.column).value
                Exit For
            End If
        Next
        For m_colNum = cellRange.column To 1 Step -1
            If CurSheet.Cells(m_rowNum - 1, m_colNum).value <> "" Then
                groupName = CurSheet.Cells(m_rowNum - 1, m_colNum).value
                Exit For
            End If
        Next
    Else
        columnName = CurSheet.Cells(2, cellRange.column).value
        For m_colNum = cellRange.column To 1 Step -1
            If CurSheet.Cells(1, m_colNum).value <> "" Then
                groupName = CurSheet.Cells(1, m_colNum).value
                Exit For
            End If
        Next
    End If
End Sub

Sub Execute_Branch_Control(ByVal sheet As Worksheet, ByVal cellRange As range, contRel As controlRelation, ByRef currentNeType As String)
    On Error Resume Next
    
    Dim sheetName, groupName, columnName As String
    Dim branchInfor As String, contedType As String
    Dim boundValue As String

    Dim allBranchMatch, contedOutOfControl As Boolean
    Dim xmlObject As Object
    Dim m, conRowNum, contedColNum As Long
    Dim noUse As Long
    Dim rootNode As Variant
    Dim controlDef As Worksheet
    Dim controlledRange As range
    Set controlDef = ThisWorkbook.Worksheets("CONTROL DEF")
    '对各个被控参数进行分支控制
    For m = 0 To contRel.contedNum
        For conRowNum = 2 To controlDef.range("a1048576").End(xlUp).row
            If (contRel.mocName = controlDef.Cells(conRowNum, 1).value) _
                            And contRel.neType = controlDef.Cells(conRowNum, 10).value _
                            And (contRel.contedAttrs(m) = controlDef.Cells(conRowNum, 2).value) _
                            And (contRel.sheetName = controlDef.Cells(conRowNum, 7).value) Then
                sheetName = controlDef.Cells(conRowNum, 7).value
                'sheetName = sheet.name
                groupName = controlDef.Cells(conRowNum, 8).value
                columnName = controlDef.Cells(conRowNum, 9).value
                contedType = controlDef.Cells(conRowNum, 3).value
                contedColNum = get_colNum(sheetName, groupName, columnName, noUse)
                Set controlledRange = sheet.Cells(cellRange.row, contedColNum)
                If (Trim(cellRange.value) = "" And cellRange.Interior.colorIndex <> SolidColorIdx And cellRange.Interior.Pattern <> SolidPattern) Or UBound(Split(cellRange.value, "\")) = 2 Then '主控为空或是引用，则此时被控应变为非灰即有效，范围恢复成初始值
                    If controlledRange.Interior.colorIndex = SolidColorIdx And controlledRange.Interior.Pattern = SolidPattern Then
                        If controlledRange.Hyperlinks.count = 1 Then
                            controlledRange.Hyperlinks.Delete
                        End If
                        '如果正在进行增加单板操作，并且修改的行是增加的行，那么样式清除时应该设置为蓝底或绿底
                        If Not setControlledRangeColorAndPattern(controlledRange) Then
                            controlledRange.Interior.colorIndex = NullPattern
                            controlledRange.Interior.Pattern = NullPattern
                        End If
                        controlledRange.Validation.ShowInput = True
                    End If
                    '恢复成初始范围
                    boundValue = controlDef.Cells(conRowNum, 4).value + controlDef.Cells(conRowNum, 5).value
                    Call setValidation(contedType, boundValue, sheet, cellRange.row, contedColNum)
                    Call setEmptyValidation(sheet, cellRange.row, contedColNum)
                Else '主控不空，进行contRel.contedAttrs(m)的分支控制
                    branchInfor = controlDef.Cells(conRowNum, 6).value
                    If isControlInfoRef(branchInfor) Then branchInfor = getRealControlInfo(branchInfor)
                    Set xmlObject = CreateObject("msxml2.domdocument")
                    xmlObject.LoadXML branchInfor
                    'Set BranchNodes = xmlObject.DocumentElement.ChildNodes
                    Set rootNode = xmlObject.DocumentElement
                    contedOutOfControl = False
                    allBranchMatch = checkAllBranchMatch(rootNode, sheet, cellRange, contRel, contedType, contedOutOfControl, contedColNum, currentNeType, branchInfor)
                    '所有主控参数的值都不在分支参数规定的范围内，则被控参数被灰化（但被控参数不受控制除外）
                    If allBranchMatch = False Then
                        If contedOutOfControl = False Then
                            controlledRange.Interior.colorIndex = SolidColorIdx
                            controlledRange.Interior.Pattern = SolidPattern
                            controlledRange.value = ""
                            controlledRange.Validation.ShowInput = False
                        Else
                            controlledRange.Validation.ShowInput = True
                        End If
                        If controlledRange.Hyperlinks.count = 1 Then
                            controlledRange.Hyperlinks.Delete
                        End If
                    Else
                        '如果正在进行增加单板操作，并且修改的行是增加的行，那么样式清除时应该设置为蓝底或绿底
                        If Not setControlledRangeColorAndPattern(controlledRange) Then
                            controlledRange.Interior.colorIndex = NullPattern
                            controlledRange.Interior.Pattern = NullPattern
                        End If
                    End If
                End If
                Exit For
            End If
        Next
    Next
    
End Sub

Private Function setControlledRangeColorAndPattern(ByRef controlledRange As range) As Boolean
    '先判断是否是新增行必填项，蓝底，再判断是否是新增行普通格子
    setControlledRangeColorAndPattern = False
    If inAddProcessFlag = True And moiRowsManager.rangeInNeedFillInRange(controlledRange) Then
        controlledRange.Interior.colorIndex = NeedFillInRangeColorIndex
        controlledRange.Interior.Pattern = NormalPattern
        setControlledRangeColorAndPattern = True
    ElseIf inAddProcessFlag = True And moiRowsManager.rangeInAddingRows(controlledRange) Then
        controlledRange.Interior.colorIndex = NewMoiRangeColorIndex
        controlledRange.Interior.Pattern = NormalPattern
        setControlledRangeColorAndPattern = True
    End If
End Function

Sub deleteValidation(ByRef sheet As Worksheet, ByRef rowNumber As Long, ByRef columnNumber As Long)
    sheet.Cells(rowNumber, columnNumber).Validation.Delete
End Sub

Sub setValidation(ByVal contedType As String, ByVal boundValue As String, sheet As Worksheet, ByVal rowNum As Long, ByVal colNum As Long)
    On Error Resume Next
    
    Dim inputTitle As String
    inputTitle = getResByKey("Range")
    
    '非枚举，无Validation则加上，若有则看是否要改
    If contedType <> "Enum" And contedType <> "Bitmap" And contedType <> "IPV4" And contedType <> "IPV6" _
        And contedType <> "Time" And contedType <> "Date" And contedType <> "DateTime" Then
        If boundValue <> sheet.Cells(rowNum, colNum).Validation.inputMessage Then
            If contedType = "String" Or contedType = "Password" Then
                inputTitle = getResByKey("Length")
                boundValue = formatRange(boundValue)
            End If
            
            If isNum(contedType) Then
                 boundValue = formatRange(boundValue)
            End If
            
            With sheet.Cells(rowNum, colNum).Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
                .inputTitle = inputTitle
                .inputMessage = boundValue
                .ShowInput = True
                .ShowError = False
            End With
        End If
    '枚举
    ElseIf contedType = "Enum" Then
        If boundValue <> sheet.Cells(rowNum, colNum).Validation.formula1 Or sheet.Cells(rowNum, colNum).Validation.inputMessage = "" Then
            With sheet.Cells(rowNum, colNum).Validation
                .Delete
                .Add Type:=xlValidateList, formula1:=boundValue
                .inputTitle = getResByKey("Range")
                .inputMessage = "[" + boundValue + "]"
                .ShowInput = True
                .ShowError = True
            End With
            sheet.Cells(rowNum, colNum).Validation.Modify Type:=xlValidateList, formula1:=boundValue
        End If
    End If
    
End Sub
Function get_colNum(ByVal sheetName As String, ByVal groupName As String, ByVal columnName As String, rowNum As Long) As Long
    Dim m_colNum1, m_colNum2, m_rowNum As Long
    Dim ws As Worksheet
    Dim startColnumLetter As Long
    startColnumLetter = 1
    If isOperationExcel Then
        startColnumLetter = 2
    End If
    
    If sheetName = getResByKey("Comm Data") Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For m_rowNum = 1 To ws.range("a1048576").End(xlUp).row
            If groupName = ws.Cells(m_rowNum, startColnumLetter).value Then
                For m_colNum1 = 1 To ws.range("XFD" + CStr(m_rowNum + 1)).End(xlToLeft).column
                    If columnName = ws.Cells(m_rowNum + 1, m_colNum1).value Then
                        get_colNum = m_colNum1
                        rowNum = m_rowNum + 1
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    ElseIf InStr(sheetName, getResByKey("Board Style")) <> 0 Then
        If containsASheet(ThisWorkbook, sheetName) Then
            Set ws = ThisWorkbook.Worksheets(sheetName)
        Else
            Set ws = ThisWorkbook.Worksheets(actualBoardStyleName)
        End If
        
        get_colNum = findColNumByGrpAndColNameEx(ws, groupName, columnName)
        
    Else
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For m_colNum1 = 1 To ws.range("XFD2").End(xlToLeft).column
            If columnName = ws.Cells(2, m_colNum1).value Then
                m_colNum2 = m_colNum1
                While Trim(ws.Cells(1, m_colNum2).value) = ""
                    m_colNum2 = m_colNum2 - 1
                Wend
                If groupName = Trim(ws.Cells(1, m_colNum2).value) Then
                    get_colNum = m_colNum1
                    Exit For
                End If
            End If
        Next
    End If
End Function
Function get_MocAndAttrcolNum(ByVal mocName As String, ByVal attrName As String, ByVal sheetName As String) As Long
    Dim conRowNum, noUse As Long
    Dim groupName, columnName As String
    Dim controlDef As Worksheet
    Set controlDef = ThisWorkbook.Worksheets("CONTROL DEF")
    get_MocAndAttrcolNum = 0
    For conRowNum = 2 To controlDef.range("a1048576").End(xlUp).row
        If (mocName = controlDef.Cells(conRowNum, 1).value) _
            And (attrName = controlDef.Cells(conRowNum, 2).value) _
            And (sheetName = controlDef.Cells(conRowNum, 7).value) Then
            groupName = controlDef.Cells(conRowNum, 8).value
            columnName = controlDef.Cells(conRowNum, 9).value
            get_MocAndAttrcolNum = get_colNum(sheetName, groupName, columnName, noUse)
        Exit For
        End If
    Next
End Function

'判断参数Target指定区域的单元格是否为灰色不可用状态,是则清空该单元格输入的值
Public Function Ensure_Gray_Cell(CurRange As range) As Boolean
    If CurRange.value <> "" And CurRange.Interior.colorIndex = SolidColorIdx And CurRange.Interior.Pattern = SolidPattern Then
        MsgBox getResByKey("NoInput"), vbOKOnly + vbExclamation + vbApplicationModal + vbDefaultButton1, getResByKey("Warning")
        CurRange.value = ""
        CurRange.Select
        Ensure_Gray_Cell = True
    Else
        Ensure_Gray_Cell = False
    End If
End Function

Function formatRange(attrRange As String) As String
        Dim reRange As String
        reRange = ""
        Dim min As Double
        Dim max As Double
        
        While attrRange <> ""
            min = CDbl(Mid(attrRange, 2, InStr(1, attrRange, ",") - 2))
            max = CDbl(Mid(attrRange, InStr(1, attrRange, ",") + 1, InStr(1, attrRange, "]") - InStr(1, attrRange, ",") - 1))
            attrRange = Mid(attrRange, InStr(1, attrRange, "]") + 1)
            If min = max Then
                If reRange <> "" Then
                    reRange = reRange + ",[" + CStr(min) + "]"
                Else
                     reRange = "[" + CStr(min) + "]"
                End If
            Else
                If reRange <> "" Then
                    reRange = reRange + ",[" + CStr(min) + "~" + CStr(max) + "]"
                Else
                     reRange = "[" + CStr(min) + "~" + CStr(max) + "]"
                End If
            End If
        Wend
        formatRange = reRange
End Function

Public Sub boardStyleSheetControl(ByVal sh As Object, ByVal target As range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Set sheet = sh
    If target.count > 1 Then
        Exit Sub
    End If
    If boardStyleGroupMap Is Nothing Then
        Call initBoardStyleMap
    End If
    Dim currentNeType As String
    If sheet.name <> getResByKey("Board Style") Then actualBoardStyleName = sheet.name
    For Each cellRange In target
        If boardStyleColumnMap.hasKey(Trim(cellRange.value)) Or boardStyleGroupMap.hasKey(Trim(cellRange.value)) Or cellRange.Borders.LineStyle = xlLineStyleNone Then
            Exit Sub
        End If
        If Ensure_Gray_Cell(cellRange) = False Then
            'isConAttr表明是否是主控参数
'            If UBound(Split(cellRange.value, "\")) <> 2 And cellRange.Hyperlinks.count = 1 Then
'                cellRange.Hyperlinks.Delete
'            End If
            isConAttr = False
           If (Check_Value_Validation(sheet, cellRange, isConAttr, contRel, currentNeType) = 1) And (isConAttr = True) Then
                Call Execute_Branch_Control(sheet, cellRange, contRel, currentNeType)
            End If
        End If
    Next cellRange
End Sub
'
Public Function isGsmCellSheet(sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = getResByKey("GSMCellSheetName") Then
        isGsmCellSheet = True
        Exit Function
    End If
    isGsmCellSheet = False
End Function

Private Sub setEmptyValidation(ByRef sheet As Worksheet, ByRef rowNumber As Long, ByRef columnNumber As Long)
    On Error Resume Next
    With sheet.Cells(rowNumber, columnNumber).Validation
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
        .inputTitle = ""
        .inputMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
End Sub

Public Function sheetSelectionShouldCheck(ByRef ws As Worksheet) As Boolean
    Dim sheetName As String
    sheetName = ws.name
    If (isGsmCellSheet(sheetName) = False) And (sheetName <> "MappingSiteTemplate") And (sheetName <> "ProductType") _
        And (sheetName <> "MappingCellTemplate") And (sheetName <> "MappingRadioTemplate") _
        And (sheetName <> "MAPPING DEF") And (sheetName <> "SHEET DEF") And (sheetName <> "CONTROL DEF") And (sheetName <> "RELATION DEF") _
        And (sheetName <> "Help for USB Parameter") And (sheetName <> getResByKey("help")) _
        And (sheetName <> "IPRouteMap") And (sheetName <> getResByKey("Cover")) And (sheetName <> "COMMON") _
        And (sheetName <> "Qos") And (sheetName <> "USB Parameter for Sites") And (sheetName <> "SummaryRes") _
        And (sheetName <> getResByKey("Temp Sheet")) Then
        sheetSelectionShouldCheck = True
    Else
        sheetSelectionShouldCheck = False
    End If
End Function

Public Function getCertainControlDefine(ByRef CurSheet As Worksheet, ByRef cellRange As range, ByRef groupName As String, ByRef sheetName As String, ByRef columnName As String) As CControlDef
    Call getGroupAndColumnName(CurSheet, cellRange, groupName, columnName)
    If Not isBoardStyleSheet(CurSheet) Then
        Set getCertainControlDefine = getControlDefine(CurSheet.name, groupName, columnName)
    Else
        If CurSheet.name <> getResByKey("Board Style") Then actualBoardStyleName = CurSheet.name
        Set getCertainControlDefine = getControlDefine(getResByKey("Board Style"), groupName, columnName)
    End If
End Function

'单元格是否是黄底的单元格
Public Function cellIsNotHyperlinkColor(ByRef cellRange As range) As Boolean
    cellIsNotHyperlinkColor = True
End Function

Public Sub boardStyleSheetChangeRruChainControl(ByVal sh As Object, ByVal target As range)
    On Error Resume Next
    Call boardStyleSheetChangeRruChainControlTpn(sh, target)
    Call boardStyleSheetChangeRruChainControlHpn(sh, target)
End Sub

Private Sub boardStyleSheetChangeRruChainControlTpn(ByVal sh As Object, ByVal target As range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Dim groupName As String
    Dim columnName As String
    Dim ttValue As String
    Dim tboardValue As String
    Dim nResponse As String
    Dim branch As Boolean
    branch = False
    Set sheet = sh
    If target.count Mod 256 = 0 Then
        Exit Sub
    End If
    
    If boardStyleGroupMap Is Nothing Then
        Call initBoardStyleMap
    End If
    
    Dim currentNeType As String
    If sheet.name <> getResByKey("Board Style") Then actualBoardStyleName = sheet.name
    For Each cellRange In target
        If boardStyleGroupMap.hasKey(Trim(cellRange.value)) Or boardStyleColumnMap.hasKey(Trim(cellRange.value)) _
            Or cellRange.Borders.LineStyle = xlLineStyleNone Or isReferenceValue(cellRange.value) Then
            Exit Sub
        End If
        
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        
        If isRruChainTTColum(groupName, columnName) = True Then
            'Call getRruChainTTandTboardNoValue(sheet, ttValue, tboardValue, cellRange.row)
            Call getRruChainTsrnValue(sheet, cellRange.value, cellRange.row)
            branch = True
        End If
        
        If isRruChainTpnColum(groupName, columnName) = True And cellRange.Interior.colorIndex <> SolidColorIdx Then
        
            Call getRruChainTTandTboardNoValue(sheet, ttValue, tboardValue, cellRange.row)
            
            If (ttValue = "CHAIN" Or ttValue = "TRUNK_CHAIN") And cellRange.Interior.colorIndex <> SolidColorIdx Then
                cellRange.Interior.colorIndex = SolidColorIdx
                cellRange.Interior.Pattern = SolidPattern
                cellRange.value = ""
                cellRange.Validation.ShowInput = False
            ElseIf ttValue = "RING" And cellRange.value <> "" Then
               '0-5
               If cellRange.Interior.colorIndex = SolidColorIdx Then
                   cellRange.Interior.colorIndex = NullPattern
                   cellRange.Interior.Pattern = NullPattern
                   cellRange.value = ""
                   cellRange.Validation.ShowInput = True
               End If
               If (CStr(cellRange.value) < 0 Or CStr(cellRange.value) > 5) Then
                    nResponse = MsgBox(getResByKey("Range") & "[0~5]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                    If nResponse = vbRetry Then
                        cellRange.Select
                    End If
                    cellRange.value = ""
                    Exit Sub
               End If
            ElseIf ttValue = "LOADBALANCE" And cellRange.value <> "" Then
            'TSRN 0-1  0-5
            'TSRN 60-254 0-7
                If cellRange.Interior.colorIndex = SolidColorIdx Then
                   cellRange.Interior.colorIndex = NullPattern
                   cellRange.Interior.Pattern = NullPattern
                   cellRange.value = ""
                   cellRange.Validation.ShowInput = True
                End If
                If CStr(tboardValue) = 0 Or CStr(tboardValue) = 1 Then
                    If (CStr(cellRange.value) < 0 Or CStr(cellRange.value) > 5) Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~5]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                   End If
                ElseIf CStr(tboardValue) > 59 And CStr(tboardValue) < 255 Then
                    If (CStr(cellRange.value) < 0 Or CStr(cellRange.value) > 7) Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~7]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                   End If
                Else
                
                End If
            End If
            
        End If
    Next cellRange
End Sub
Private Sub boardStyleSheetChangeRruChainControlHpn(ByVal sh As Object, ByVal target As range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Dim groupName As String
    Dim columnName As String
    Dim ttValue As String
    Dim tboardValue As String
    Dim nResponse As String
    Dim branch As Boolean
    branch = False
    Set sheet = sh
    If target.count Mod 256 = 0 Then
        Exit Sub
    End If
    
    If boardStyleGroupMap Is Nothing Then
        Call initBoardStyleMap
    End If
    
    Dim currentNeType As String
    If sheet.name <> getResByKey("Board Style") Then actualBoardStyleName = sheet.name
    For Each cellRange In target
        If boardStyleGroupMap.hasKey(Trim(cellRange.value)) Or boardStyleColumnMap.hasKey(Trim(cellRange.value)) _
            Or cellRange.Borders.LineStyle = xlLineStyleNone Or isReferenceValue(cellRange.value) Then
            Exit Sub
        End If
        
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        
        If isRruChainTTColum(groupName, columnName) = True Then
            'Call getRruChainTTandTboardNoValue(sheet, ttValue, tboardValue, cellRange.row)
            Call getRruChainHsrnValue(sheet, cellRange.value, cellRange.row)
            branch = True
        End If
        
        If isRruChainHpnColum(groupName, columnName) = True And cellRange.Interior.colorIndex <> SolidColorIdx Then
        
            Call getRruChainTTandHboardNoValue(sheet, ttValue, tboardValue, cellRange.row)
            
            If ttValue = "TRUNK_CHAIN" And cellRange.Interior.colorIndex <> SolidColorIdx Then
                cellRange.Interior.colorIndex = SolidColorIdx
                cellRange.Interior.Pattern = SolidPattern
                cellRange.value = ""
                cellRange.Validation.ShowInput = False
            ElseIf ttValue = "RING" And cellRange.value <> "" Then
               '0-5
               If cellRange.Interior.colorIndex = SolidColorIdx Then
                   cellRange.Interior.colorIndex = NullPattern
                   cellRange.Interior.Pattern = NullPattern
                   cellRange.value = ""
                   cellRange.Validation.ShowInput = True
               End If
               If (CStr(cellRange.value) < 0 Or CStr(cellRange.value) > 5) Then
                    nResponse = MsgBox(getResByKey("Range") & "[0~5]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                    If nResponse = vbRetry Then
                        cellRange.Select
                    End If
                    cellRange.value = ""
                    Exit Sub
               End If
            ElseIf ttValue = "LOADBALANCE" And cellRange.value <> "" Then
            'TSRN 0-1  0-5
            'TSRN 60-254 0-7
                If cellRange.Interior.colorIndex = SolidColorIdx Then
                   cellRange.Interior.colorIndex = NullPattern
                   cellRange.Interior.Pattern = NullPattern
                   cellRange.value = ""
                   cellRange.Validation.ShowInput = True
                End If
                If CStr(tboardValue) = 0 Or CStr(tboardValue) = 1 Then
                    If (CStr(cellRange.value) < 0 Or CStr(cellRange.value) > 5) Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~5]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                   End If
                ElseIf CStr(tboardValue) > 59 And CStr(tboardValue) < 255 Then
                    If (CStr(cellRange.value) < 0 Or CStr(cellRange.value) > 11) Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~11]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                   End If
                Else
                End If
                
            ElseIf ttValue = "CHAIN" And cellRange.value <> "" Then
            'TSRN 0-1  0-5
            'TSRN 60-254 0-7
                If cellRange.Interior.colorIndex = SolidColorIdx Then
                   cellRange.Interior.colorIndex = NullPattern
                   cellRange.Interior.Pattern = NullPattern
                   cellRange.value = ""
                   cellRange.Validation.ShowInput = True
                End If
                If CStr(tboardValue) = 0 Or CStr(tboardValue) = 1 Then
                    If (CStr(cellRange.value) < 0 Or CStr(cellRange.value) > 5) Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~5]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                   End If
                ElseIf CStr(tboardValue) > 59 And CStr(tboardValue) < 255 Then
                    If (CStr(cellRange.value) < 0 Or CStr(cellRange.value) > 11) Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~11]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                   End If
                Else
                
                End If
            End If
        End If
    Next cellRange
End Sub

Public Sub boardStyleSheetRruChainControl(ByVal sh As Object, ByVal target As range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Dim groupName As String
    Dim columnName As String
    Dim ttValue As String
    Dim tboardValue As String
    Dim nResponse As String
    
    Set sheet = sh
    If target.count Mod 256 = 0 Then
        Exit Sub
    End If
    If boardStyleGroupMap Is Nothing Then
        Call initBoardStyleMap
    End If
    
    Dim currentNeType As String
    If sheet.name <> getResByKey("Board Style") Then actualBoardStyleName = sheet.name
    For Each cellRange In target
        If boardStyleGroupMap.hasKey(Trim(cellRange.value)) Or boardStyleColumnMap.hasKey(Trim(cellRange.value)) _
            Or cellRange.Borders.LineStyle = xlLineStyleNone Or isReferenceValue(cellRange.value) Then
            Exit Sub
        End If
        
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        
        'If isRruChainTTColum(groupName, columnName) = True Then
            'Call getRruChainTTandTboardNoValue(sheet, ttValue, tboardValue, cellRange.row)
           ' Call getRruChainTsrnValue(sheet, ttValue, cellRange.row)
        'End If
        
        If isRruChainTpnColum(groupName, columnName) = True Then
        
            Call getRruChainTTandTboardNoValue(sheet, ttValue, tboardValue, cellRange.row)
            
            If (ttValue = "CHAIN" Or ttValue = "TRUNK_CHAIN") And cellRange.Interior.colorIndex <> SolidColorIdx Then
                cellRange.Interior.colorIndex = SolidColorIdx
                cellRange.Interior.Pattern = SolidPattern
                cellRange.value = ""
                cellRange.Validation.ShowInput = False
            Else
                If cellRange.Interior.colorIndex = SolidColorIdx Then
                    cellRange.Interior.colorIndex = NullPattern
                    cellRange.Interior.Pattern = NullPattern
                    cellRange.value = ""
                    cellRange.Validation.ShowInput = True
                End If
            End If
            
        End If
        
        If isRruChainHpnColum(groupName, columnName) = True Then
        
            Call getRruChainTTandHboardNoValue(sheet, ttValue, tboardValue, cellRange.row)
            
            If ttValue = "TRUNK_CHAIN" And cellRange.Interior.colorIndex <> SolidColorIdx Then
                cellRange.Interior.colorIndex = SolidColorIdx
                cellRange.Interior.Pattern = SolidPattern
                cellRange.value = ""
                cellRange.Validation.ShowInput = False
            Else
                If cellRange.Interior.colorIndex = SolidColorIdx Then
                    cellRange.Interior.colorIndex = NullPattern
                    cellRange.Interior.Pattern = NullPattern
                    cellRange.value = ""
                    cellRange.Validation.ShowInput = True
                End If
            End If
            
        End If
    Next cellRange
End Sub

Private Function isRruChainTpnColum(ByRef groupName As String, ByRef columnName As String) As Boolean
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    isRruChainTpnColum = False
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mocName = "RRUCHAIN" And attributeName = "TPN" Then
            Exit For
        End If
    Next
    If groupName = mappingdefgroupName And columnName = mappingdefcolumnName Then
        isRruChainTpnColum = True
    End If
End Function
Private Function isRruChainHpnColum(ByRef groupName As String, ByRef columnName As String) As Boolean
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    isRruChainHpnColum = False
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mocName = "RRUCHAIN" And attributeName = "HPN" Then
            Exit For
        End If
    Next
    If groupName = mappingdefgroupName And columnName = mappingdefcolumnName Then
        isRruChainHpnColum = True
    End If
End Function
Private Function isRruChainTTColum(ByRef groupName As String, ByRef columnName As String) As Boolean
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    isRruChainTTColum = False
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mocName = "RRUCHAIN" And attributeName = "TT" Then
            Exit For
        End If
    Next
    If groupName = mappingdefgroupName And columnName = mappingdefcolumnName Then
        isRruChainTTColum = True
    End If
End Function

Private Sub getRruChainTTandTboardNoValue(ByRef sheet As Worksheet, ByRef ttValue As String, ByRef tboardValue As String, ByRef rowNumber As Long)
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim ttmappingdefcolumnName As String
    Dim tboardmappingdefcolumnName As String
    Dim ttColumnNum As Long
    Dim tboardColumnNum As Long
    Dim noUse As Long
    Dim strs() As String

    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mocName = sheetDef.Cells(index, 4)
        If mocName = "RRUCHAIN" And sheetDef.Cells(index, 5) = "TT" Then
            ttmappingdefcolumnName = sheetDef.Cells(index, 3)
            mappingdefgroupName = sheetDef.Cells(index, 2)
        End If
        If mocName = "RRUCHAIN" And sheetDef.Cells(index, 5) = "TailBoardNo" Then
            tboardmappingdefcolumnName = sheetDef.Cells(index, 3)
        End If
    Next
    
    ttColumnNum = get_colNum(sheet.name, mappingdefgroupName, ttmappingdefcolumnName, noUse)
    tboardColumnNum = get_colNum(sheet.name, mappingdefgroupName, tboardmappingdefcolumnName, noUse)
    
    ttValue = sheet.Cells(rowNumber, ttColumnNum)
    tboardValue = sheet.Cells(rowNumber, tboardColumnNum)
    
    strs = Split(tboardValue, "_")
    tboardValue = strs(1)

End Sub
Private Sub getRruChainTTandHboardNoValue(ByRef sheet As Worksheet, ByRef ttValue As String, ByRef tboardValue As String, ByRef rowNumber As Long)
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim ttmappingdefcolumnName As String
    Dim tboardmappingdefcolumnName As String
    Dim ttColumnNum As Long
    Dim tboardColumnNum As Long
    Dim noUse As Long
    Dim strs() As String

    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mocName = sheetDef.Cells(index, 4)
        If mocName = "RRUCHAIN" And sheetDef.Cells(index, 5) = "TT" Then
            ttmappingdefcolumnName = sheetDef.Cells(index, 3)
            mappingdefgroupName = sheetDef.Cells(index, 2)
        End If
        If mocName = "RRUCHAIN" And sheetDef.Cells(index, 5) = "HeadBoardNo" Then
            tboardmappingdefcolumnName = sheetDef.Cells(index, 3)
        End If
    Next
    
    ttColumnNum = get_colNum(sheet.name, mappingdefgroupName, ttmappingdefcolumnName, noUse)
    tboardColumnNum = get_colNum(sheet.name, mappingdefgroupName, tboardmappingdefcolumnName, noUse)
    
    ttValue = sheet.Cells(rowNumber, ttColumnNum)
    tboardValue = sheet.Cells(rowNumber, tboardColumnNum)
    
    strs = Split(tboardValue, "_")
    tboardValue = strs(1)

End Sub

Private Sub getRruChainTsrnValue(ByRef sheet As Worksheet, ByRef ttValue As String, ByRef rowNumber As Long)
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim ttmappingdefcolumnName As String
    Dim tboardmappingdefcolumnName As String
    Dim ttColumnNum As Long
    Dim tboardColumnNum As Long
    Dim noUse As Long
    Dim strs() As String

    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mocName = sheetDef.Cells(index, 4)
        If mocName = "RRUCHAIN" And sheetDef.Cells(index, 5) = "TPN" Then
            ttmappingdefcolumnName = sheetDef.Cells(index, 3)
            mappingdefgroupName = sheetDef.Cells(index, 2)
        End If
    Next
    
    ttColumnNum = get_colNum(sheet.name, mappingdefgroupName, ttmappingdefcolumnName, noUse)
    If ttValue = "CHAIN" Then
        sheet.Cells(rowNumber, ttColumnNum).Interior.colorIndex = SolidColorIdx
        sheet.Cells(rowNumber, ttColumnNum).Interior.Pattern = SolidPattern
        sheet.Cells(rowNumber, ttColumnNum).value = ""
        sheet.Cells(rowNumber, ttColumnNum).Validation.ShowInput = False
    Else
        If sheet.Cells(rowNumber, ttColumnNum).Interior.colorIndex = SolidColorIdx Then
            sheet.Cells(rowNumber, ttColumnNum).Interior.colorIndex = NullPattern
            sheet.Cells(rowNumber, ttColumnNum).Interior.Pattern = NullPattern
            sheet.Cells(rowNumber, ttColumnNum).value = ""
            sheet.Cells(rowNumber, ttColumnNum).Validation.ShowInput = True
        End If
    End If

End Sub

Private Sub getRruChainHsrnValue(ByRef sheet As Worksheet, ByRef ttValue As String, ByRef rowNumber As Long)
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim ttmappingdefcolumnName As String
    Dim hboardmappingdefcolumnName As String
    Dim ttColumnNum As Long
    Dim hboardColumnNum As Long
    Dim noUse As Long
    Dim strs() As String

    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mocName = sheetDef.Cells(index, 4)
        If mocName = "RRUCHAIN" And sheetDef.Cells(index, 5) = "HPN" Then
            ttmappingdefcolumnName = sheetDef.Cells(index, 3)
            mappingdefgroupName = sheetDef.Cells(index, 2)
        End If
    Next
    
    ttColumnNum = get_colNum(sheet.name, mappingdefgroupName, ttmappingdefcolumnName, noUse)
    If ttValue = "TRUNK_CHAIN" Then
        sheet.Cells(rowNumber, ttColumnNum).Interior.colorIndex = SolidColorIdx
        sheet.Cells(rowNumber, ttColumnNum).Interior.Pattern = SolidPattern
        sheet.Cells(rowNumber, ttColumnNum).value = ""
        sheet.Cells(rowNumber, ttColumnNum).Validation.ShowInput = False
    Else
        If sheet.Cells(rowNumber, ttColumnNum).Interior.colorIndex = SolidColorIdx Then
            sheet.Cells(rowNumber, ttColumnNum).Interior.colorIndex = NullPattern
            sheet.Cells(rowNumber, ttColumnNum).Interior.Pattern = NullPattern
            sheet.Cells(rowNumber, ttColumnNum).value = ""
            sheet.Cells(rowNumber, ttColumnNum).Validation.ShowInput = True
        End If
    End If

End Sub

'**********************************************************
'′óáDêyμ?μ?áD??￡o1->A￡?27->AA
'**********************************************************
Public Function getColumnNameFromColumnNum(iColumn As Long) As String
  If iColumn >= 257 Or iColumn < 0 Then
    getColumnNameFromColumnNum = ""
    Return
  End If
  
  Dim result As String
  Dim High, Low As Long
  
  High = Int((iColumn - 1) / 26)
  Low = iColumn Mod 26
  
  If High > 0 Then
    result = Chr(High + 64)
  End If

  If Low = 0 Then
    Low = 26
  End If
  
  result = result & Chr(Low + 64)
  getColumnNameFromColumnNum = result
End Function
