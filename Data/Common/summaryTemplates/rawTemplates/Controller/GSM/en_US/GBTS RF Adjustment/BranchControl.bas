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
Dim ControlRelationNum As Long

Private controlRelationManager As CControlRelationManager

Private actualBoardStyleName As String

'用以设置颜色
Public Const SolidColorIdx = 16
Public Const SolidPattern = xlGray16

Public Const NullPattern = xlNone
Const NormalPattern = 1

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
Sub buildControlRelations()
    Dim conRowNum As Long
    Dim i, j As Long
    Dim conInfor As String, mocName As String, contedName As String, contName As String
    Dim isExist, flag, isfound As Boolean
    Dim index1, index2, index3, index4 As Long
    Dim neType As String
    Dim sheetName As String
    
    Dim groupName As String, columnName As String, valueType As String
    
    Set controlRelationManager = New CControlRelationManager '当前参数范围校验管理类
    
    If isControlDefSheetExist = False Then
        Exit Sub
    End If
    
    ControlRelationNum = 0
    Dim controlDef As Worksheet
    Set controlDef = ThisWorkbook.Worksheets("CONTROL DEF")
    For conRowNum = 2 To controlDef.range("a1048576").End(xlUp).row

        mocName = Trim(controlDef.Cells(conRowNum, 1).value)
        contedName = Trim(controlDef.Cells(conRowNum, 2).value)
        conInfor = Trim(controlDef.Cells(conRowNum, 6).value)
        neType = Trim(controlDef.Cells(conRowNum, 10).value)
        sheetName = Trim(controlDef.Cells(conRowNum, 7).value)
    
        groupName = Trim(controlDef.range("H" & conRowNum).value)
        columnName = Trim(controlDef.range("I" & conRowNum).value)
        valueType = Trim(controlDef.range("C" & conRowNum).value)
        Call controlRelationManager.addNewAttributeRelation(mocName, contedName, conInfor, neType, sheetName, groupName, columnName, valueType)
        
        If conInfor <> "" Then
            While InStr(1, conInfor, "attribute", vbBinaryCompare) <> 0
                conInfor = Mid(conInfor, InStr(1, conInfor, "attribute", vbBinaryCompare) + 11)
                contName = Mid(conInfor, 1, InStr(1, conInfor, """") - 1)
                '已有主控的映射关系
                isExist = False
                If ControlRelationNum > 0 Then
                    For i = 0 To ControlRelationNum - 1
                        If (mocName = ControlRelMap(i).mocName) And (contName = ControlRelMap(i).contAttr) _
                        And (neType = ControlRelMap(i).neType) And (sheetName = ControlRelMap(i).sheetName) Then
                            flag = False
                            For j = 0 To ControlRelMap(i).contedNum
                                If ControlRelMap(i).contedAttrs(j) = contedName Then
                                    flag = True
                                    Exit For
                                End If
                            Next
                            If flag = False Then
                                '保证放在此被控参数的被控参数后面，例如A->B->C, 要保证在A的控制列表中B在C的前面
                                isfound = False
                                For index1 = 0 To ControlRelationNum - 1
                                    If (mocName = ControlRelMap(index1).mocName) And (contedName = ControlRelMap(index1).contAttr) And (neType = ControlRelMap(index1).neType) And (sheetName = ControlRelMap(index1).sheetName) Then
                                        For index2 = 0 To ControlRelMap(index1).contedNum '被控参数B的控制列表
                                            For index3 = 0 To ControlRelMap(i).contedNum    '主控参数A的控制列表
                                                If ControlRelMap(index1).contedAttrs(index2) = ControlRelMap(i).contedAttrs(index3) Then
                                                    For index4 = ControlRelMap(i).contedNum To index3 Step -1
                                                        ControlRelMap(i).contedAttrs(index4 + 1) = ControlRelMap(i).contedAttrs(index4)
                                                    Next
                                                    ControlRelMap(i).contedAttrs(index3) = contedName
                                                    isfound = True
                                                    Exit For
                                                End If
                                            Next
                                            If isfound = True Then
                                                Exit For
                                            End If
                                        Next
                                        Exit For
                                    End If
                                Next
                                ControlRelMap(i).contedNum = ControlRelMap(i).contedNum + 1
                                If isfound = False Then
                                    ControlRelMap(i).contedAttrs(ControlRelMap(i).contedNum) = contedName
                                End If
                            End If
                            isExist = True
                            Exit For
                        End If
                    Next
                End If
                '没有主控的映射关系，则新建
                If isExist = False Then
                    ControlRelMap(ControlRelationNum).mocName = mocName
                    ControlRelMap(ControlRelationNum).contAttr = contName
                    ControlRelMap(ControlRelationNum).contedAttrs(0) = contedName
                    ControlRelMap(ControlRelationNum).contedNum = 0
                    ControlRelMap(ControlRelationNum).neType = neType
                    ControlRelMap(ControlRelationNum).sheetName = sheetName
                    ControlRelationNum = ControlRelationNum + 1
                End If
            Wend
        End If
    Next

End Sub
'1表示在范围内，0表示不在范围内，-1表示未定义映射关系(即后续不用再进行处理)
Function Check_Value_Validation(ByVal CurSheet As Worksheet, ByVal cellRange As range, ByRef isConAttr As Boolean, ByRef contRel As controlRelation, ByRef currentNeType As String) As Long
    On Error Resume Next
    
    Dim groupName As String
    Dim columnName As String
    Dim mocName, attrName, attrType, attrRange As String
    Dim conRowNum, i As Long
    Dim neType As String
    Dim sheetName As String
    Dim controlDef As CControlDef
    
    Call getGroupAndColumnName(CurSheet, cellRange, groupName, columnName)
'    If Not isBoardStyleSheet(CurSheet) Then
'        Call getGroupAndColumnName(CurSheet, cellRange, groupName, columnName)
'        Set controldef = getControlDefine(CurSheet.name, groupName, columnName)
'    Else
'        Call getGroupNameShNameAndAttrName(CurSheet, cellRange, groupName, sheetName, columnName)
'        Set controldef = getControlDefine(sheetName, groupName, columnName)
'    End If
    If Not isBoardStyleSheet(CurSheet) Then
        Set controlDef = getControlDefine(CurSheet.name, groupName, columnName)
    Else
        Set controlDef = getControlDefine(getResByKey("Board Style"), groupName, columnName)
    End If
    
    If controlDef Is Nothing Then
        Check_Value_Validation = -1
        Exit Function
    End If
    
    mocName = controlDef.mocName
    attrName = controlDef.attributeName
    neType = controlDef.neType
    sheetName = controlDef.sheetName
    
    currentNeType = neType
    '判断是否是主控参数
    For i = 0 To ControlRelationNum - 1
        If (ControlRelMap(i).mocName = mocName) And (ControlRelMap(i).contAttr = attrName) _
        And (ControlRelMap(i).neType = neType) And (ControlRelMap(i).sheetName = sheetName) Then
            If Moc_Appear_In_SameSheet(ControlRelMap(i).mocName, attrName) = False Then
                isConAttr = True
                contRel = ControlRelMap(i)
             End If
            Exit For
        End If
    Next
    If Trim(cellRange.value) <> "" And UBound(Split(cellRange.value, "\")) <> 2 Then '值非空，则判断输入值是否在范围内
        attrType = controlDef.dataType
        attrRange = controlDef.bound + controlDef.lstValue
        If Trim(controlDef.controlInfo) <> "" Then
            '有分支参数控制信息的范围可能发生变化,Tip中显示的就是最新范围
            If attrType = "Enum" And cellRange.Validation.inputMessage <> "" Then
                attrRange = Mid(cellRange.Validation.inputMessage, 2, Len(cellRange.Validation.inputMessage) - 2)
            ElseIf cellRange.Validation.inputMessage <> "" Then
                attrRange = cellRange.Validation.inputMessage
            End If
        End If
        
        Dim alreadyCheckFlag As Boolean
        alreadyCheckFlag = False '参数值是否已经校验的标志
        
        If Check_Value_In_Range(attrType, attrRange, cellRange.value, cellRange, alreadyCheckFlag) = True Then
            Check_Value_Validation = 1
        Else
            If alreadyCheckFlag = False Then '说明在Check_Value_In_Range出异常了，需要用新的被控参数校验方法进行校验
                If controlledAttrValidationCheck(CurSheet, controlDef, cellRange) = True Then '校验的参数在范围内
                    Check_Value_Validation = 1
                Else '校验的参数不在范围内，返回0，主要为了不进行后续的被控参数分支了
                    Check_Value_Validation = 0
                End If
            Else '这是正常的不在范围内
                Check_Value_Validation = 0
            End If
        End If
        Exit Function
    Else  '值为空，由于还要进行分支控制，所以判定为在范围内
        Check_Value_Validation = 1
        Exit Function
    End If
    
    Check_Value_Validation = -1
    
End Function

'新增的被控参数校验，在Tip提示的基础上再做校验
Private Function controlledAttrValidationCheck(ByRef ws As Worksheet, ByRef controlDef As CControlDef, ByRef cellRange As range) As Boolean
    On Error GoTo ErrorHandler
    controlledAttrValidationCheck = False
    
    Dim mocName As String, attrName As String, neType As String, controlInfo As String, sheetName As String
    Call getControlAttrInfo(controlDef, mocName, attrName, neType, sheetName, controlInfo)
    
     '如果该参数不是被控参数，则退出
    If Not controlRelationManager.containsControlledAttributeRelation(mocName, attrName, neType, sheetName) Then Exit Function
    
    '找到该参数的所有主控参数类
    Dim controlRelation As CControlRelation
    Set controlRelation = controlRelationManager.getControlRelation(mocName, attrName, neType, sheetName)
    
    Dim controlAttrValueManager As New CControlAttrValueManager '管理所有主控参数的管理类，后续用这个主控参数管理类与被参数的分支控制信息作校验
    Dim mainControlAttrReturnedValue As Long
    mainControlAttrReturnedValue = makeControlAttrValueCol(ws, mocName, attrName, neType, sheetName, controlRelation, controlAttrValueManager, cellRange)
    
    If mainControlAttrReturnedValue = 1 Then Exit Function '不需要对当前被控参数进行控制，直接退出
    
    Dim oneBranchMatchFlag As Boolean
    Dim branchMatchChecker As New CBranchMatchChecker '分支匹配校验类
    Call branchMatchChecker.init(controlAttrValueManager, controlInfo)
    oneBranchMatchFlag = branchMatchChecker.getOneBranchMatchFlag
    
    If oneBranchMatchFlag = False Then Exit Function '如果没找到匹配分支，则退出
    
    Dim matchBranchNode As Variant
    Set matchBranchNode = branchMatchChecker.getMatchBranchNode
    
    Dim controlAttrValue As New CControlAttrValue '声明一个控制属性类对象
    Call controlAttrValue.init("", "", cellRange.value, controlDef.dataType, False, False, False)
    Dim branchAttrValueValidationFlag As Boolean
    branchAttrValueValidationFlag = controlAttrValue.checkABranchAttrValues(matchBranchNode)
    
    
    If branchAttrValueValidationFlag = False Then
        Dim nResponse As Long
        Dim errorMsg As String
        With cellRange.Validation
            errorMsg = .inputTitle & .inputMessage
        End With
        nResponse = MsgBox(errorMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
        If nResponse = vbRetry Then
            cellRange.Select
        End If
        cellRange.value = ""
    End If
    
    controlledAttrValidationCheck = branchAttrValueValidationFlag '这是最后范围校验的结果
    Exit Function
ErrorHandler:
    controlledAttrValidationCheck = True '如果出异常了，则认为在范围内，与以前的流程保持一致
End Function


Function Check_Value_In_Range(ByVal attrType As String, ByVal attrRange As String, ByVal attrValue As String, cellRange As range, ByRef alreadyCheckFlag As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim arrayList() As String
    Dim errorMsg, sItem As String
    Dim i, nResponse, nLoop As Long
    Dim min, max As Double
    
    '当参数范围为空，不需要校验 DTS2012111306136
    If attrRange = "" Then
        Check_Value_In_Range = True
        alreadyCheckFlag = True
        Exit Function
    End If
    
    If attrType = "Enum" Then
        Check_Value_In_Range = False
        arrayList = Split(attrRange, ",")
        For i = 0 To UBound(arrayList)
            If Trim(attrValue) = arrayList(i) Then
                Check_Value_In_Range = True
                Exit For
            End If
        Next
        errorMsg = getResByKey("Range") + "[" + attrRange + "]"
    ElseIf attrType = "String" Or attrType = "Password" Or attrType = "ATM" Then
        min = CDbl(Mid(attrRange, 2, InStr(1, attrRange, ",") - 2))
        max = CDbl(Mid(attrRange, InStr(1, attrRange, ",") + 1, InStr(1, attrRange, "]") - InStr(1, attrRange, ",") - 1))
        If Len(attrValue) < min Or Len(attrValue) > max Then
            Check_Value_In_Range = False
        Else
            Check_Value_In_Range = True
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
        Check_Value_In_Range = False
        alreadyCheckFlag = True
        Exit Function
    Else  '数值
        If Check_Int_Validation(attrRange, attrValue) = True Then
            Check_Value_In_Range = True
        Else
            Check_Value_In_Range = False
        End If
        errorMsg = getResByKey("Range") + formatRange(attrRange)
    End If
    
    If Check_Value_In_Range = False Then
        nResponse = MsgBox(errorMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
        If nResponse = vbRetry Then
            cellRange.Select
        End If
        cellRange.value = ""
    End If
    alreadyCheckFlag = True
    Exit Function
ErrorHandler:
    Check_Value_In_Range = False '出异常，说明校验出错，需要在外层进行被控参数的校验
End Function
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
'    For m = 0 To contRel.contedNum - 1
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
                            '如果正在进行增加单板操作，并且修改的行是增加的行，那么样式清除时应该设置为蓝底或绿底
                            If Not setControlledRangeColorAndPattern(controlledRange) Then
                                controlledRange.Interior.colorIndex = NullPattern
                                controlledRange.Interior.Pattern = NullPattern
                            End If
                            controlledRange.Validation.ShowInput = True
                        End If
                        If controlledRange.Hyperlinks.count = 1 Then
                            controlledRange.Hyperlinks.Delete
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

Private Sub makeControlAttrValueManager(ByRef sheet As Worksheet, ByRef dstRowNumber As Long, ByRef mainControlMocName As String, ByRef mainControlAttrName As String, _
    ByRef neType As String, ByRef virtualSheetName As String, ByRef oneMainControlAttrNotExist As Boolean, ByRef controlAttrValueManager As CControlAttrValueManager)
    
    '如果已经有了该主控参数，就无需再添加了，直接退出
    If controlAttrValueManager.hasControlAttr(mainControlAttrName) Then Exit Sub
    
    Dim dstColumnNumber As Long
    dstColumnNumber = get_MocAndAttrcolNum(mainControlMocName, mainControlAttrName, virtualSheetName) '得到主控参数的列号
    If dstColumnNumber = 0 Then '等于0，说明主控参数不存在，则退出，并将flag置为True
        oneMainControlAttrNotExist = True
        Exit Sub
    End If

    '如果在当前页签没有找到该主控参数的控制信息，说明主控参数缺少，则该被控参数不控制
    If Not controlRelationManager.containsAttributeRelation(mainControlMocName, mainControlAttrName, neType, virtualSheetName) Then
        oneMainControlAttrNotExist = True
        Exit Sub
    End If
    
    Dim mainControlRelation As CControlRelation '这是主控参数的Control Def管理类
    Dim mainControlAttrValue As CControlAttrValue '管理一个主控参数值，类型，长短名称的类
    Dim mainControlGroupName As String, mainControlColumnName As String, mainValueType As String
    
    Set mainControlRelation = controlRelationManager.getControlRelation(mainControlMocName, mainControlAttrName, neType, virtualSheetName)
    
    mainControlGroupName = mainControlRelation.groupName
    mainControlColumnName = mainControlRelation.columnName
    mainValueType = mainControlRelation.valueType
    
    Dim mainAttrCell As range
    Dim mainValue As String
    Set mainAttrCell = sheet.Cells(dstRowNumber, dstColumnNumber)
    
    mainValue = mainAttrCell.value
        
    Dim valueEmptyFlag As Boolean, valueReferenceFlag As Boolean, valueCellGrayFlag As Boolean
    valueEmptyFlag = False
    valueReferenceFlag = False
    valueCellGrayFlag = False
    
    If UBound(Split(mainValue, "\")) = 2 Then  '如果其中一个主控参数为引用，标志设为True，由对象自己在类中check
        valueReferenceFlag = True '主控为引用标志
    ElseIf mainValue = "" Then '此时主控参数为空，有可能该主控参数也是无效的，则需要先对该主控单元格进行校验
        Call currentParameterBranchCheck(sheet, mainAttrCell)
'        If cellIsGray(cellRange) Then
'            makeControlAttrValueCol = 1 '可能由于主控参数的灰化改变已经使当前单元格灰化了，则不需要判断了，直接退出
'            Exit Sub
'        End If
    End If
    
    If cellIsGray(mainAttrCell) Then '如果某个主控参数灰化，标志设为True，由对象自己在类中check
        valueCellGrayFlag = True '主控为灰化标志
        valueEmptyFlag = True '主控为灰化，则肯定为空
    ElseIf mainValue = "" Then '如果主控单元格未灰化，说明是主控单元格是有效分支，只是未填写值，空标志设为True，由对象自己在类中check
        valueEmptyFlag = True
    End If
    
    Set mainControlAttrValue = New CControlAttrValue
    Call mainControlAttrValue.init(mainControlAttrName, mainControlColumnName, mainValue, mainValueType, valueEmptyFlag, valueReferenceFlag, valueCellGrayFlag)
    
    Call controlAttrValueManager.addNewControlAttrValue(mainControlAttrValue)
End Sub

Private Function getMainControlAttrCol(ByRef root As Variant) As Collection
    Dim controlAttrCol As New Collection
    Dim controlAttributeNode As Variant
    Dim controlAttributeName As String
    For Each controlAttributeNode In root.GetElementsByTagName("Equals")
        controlAttributeName = controlAttributeNode.getAttribute("attribute")
        If Not IsNull(controlAttributeName) Then
            If Not Contains(controlAttrCol, controlAttributeName) Then
                controlAttrCol.Add Item:=controlAttributeName, key:=controlAttributeName '将各主控参数加入容器中
            End If
        End If
    Next controlAttributeNode
    Set getMainControlAttrCol = controlAttrCol
End Function

Function checkAllBranchMatch(rootNode As Variant, sheet As Worksheet, cellRange As range, contRel As controlRelation, contedType As String, contedOutOfControl As Boolean, contedColNum As Long, ByRef currentNeType As String, ByRef controlInfo As String) As Boolean
    On Error Resume Next
    
    Dim matchBranchNode As Variant '匹配的分支节点
    
    Dim i, j, colNum As Long
    Dim contAttrNum As Long
    Dim equalsNodes, boundNodes
    Dim eachContAttr As String, eachContAttrVal As String, boundValue As String
    Dim oneBranchMatch As Boolean, oneContNotExist As Boolean, oneContGray As Boolean, oneContNull As Boolean
    Dim valIsRight As Boolean
    
    checkAllBranchMatch = False

    Dim oneMainControlAttrNotExist As Boolean
    oneMainControlAttrNotExist = False
    
    Dim cellRowNumber As Long
    cellRowNumber = cellRange.row
    
    Dim controlAttrValueManager As New CControlAttrValueManager
    
    Dim mainControlAttrCol As Collection
    Set mainControlAttrCol = getMainControlAttrCol(rootNode) '得到主控参数容器
    
    Dim eachMainControlAttr As Variant
    For Each eachMainControlAttr In mainControlAttrCol
        Call makeControlAttrValueManager(sheet, cellRowNumber, contRel.mocName, CStr(eachMainControlAttr), currentNeType, contRel.sheetName, oneMainControlAttrNotExist, controlAttrValueManager)
    Next eachMainControlAttr

    If oneMainControlAttrNotExist = True Then '主控参数缺少，则认定在范围内，直接退出
        contedOutOfControl = True
        Exit Function
    End If
    
    oneBranchMatch = newCheckBranchMatch(controlAttrValueManager, controlInfo, matchBranchNode)
    If oneBranchMatch = False Then '未找到匹配分支
        contedOutOfControl = False
    Else '此分支中各主控参数匹配成功，则进行分支控制
        Set boundNodes = matchBranchNode.childNodes
        '获得被控参数的范围
        boundValue = getContedAttrBoundValue(boundNodes, valIsRight, sheet, cellRange, contedColNum)
        '进行分支控制
        If sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = SolidColorIdx And sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = SolidPattern Then
            If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
                sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
            End If
            sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = NullPattern
            sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = NullPattern
            sheet.Cells(cellRange.row, contedColNum).Validation.ShowInput = True
        End If
        '不在范围内时要清空
        If valIsRight = False And Trim(sheet.Cells(cellRange.row, contedColNum).value) <> "" _
            And contedType <> "IPV4" And contedType <> "IPV6" Then
            If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
                sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
            End If
            sheet.Cells(cellRange.row, contedColNum).value = ""
        End If
        '设置被控参数的范围
        Call setValidation(contedType, boundValue, sheet, cellRange.row, contedColNum)
        checkAllBranchMatch = True
    End If
    
'    '遍历各个控制分支，看满足哪一个（进行这个的分支控制），或者全不满足（灰化）
'    For I = 0 To BranchNodes.Length - 1
'        oneBranchMatch = True
'        oneContGray = False
'        oneContNotExist = False
'        oneContNull = False
'        contAttrNum = 0
'        Set equalsNodes = BranchNodes(I).GetElementsByTagName("Equals")
'        '看此分支中的各个主控参数值是否在范围内
'        For J = 0 To equalsNodes.Length - 1
'            eachContAttr = equalsNodes(J).getAttributeNode("attribute").NodeValue
'
'            Call makeControlAttrValueManager(sheet, cellRowNumber, contRel.mocName, eachContAttr, currentNeType, contRel.sheetName, oneMainControlAttrNotExist, controlAttrValueManager)
'
'            If eachContAttr = contRel.contAttr Then
'                If cellRange.Interior.colorIndex = SolidColorIdx And cellRange.Interior.Pattern = SolidPattern Then
'                    oneContGray = True
'                    oneBranchMatch = False
'                    eachContAttrVal = ""
'                Else
'                    eachContAttrVal = Trim(cellRange.value)
'                End If
'            Else
'                eachContAttrVal = ""
'                colNum = get_MocAndAttrcolNum(contRel.mocName, eachContAttr, contRel.sheetName)
'                If colNum <> 0 Then
'                    If sheet.Cells(cellRange.row, colNum).Interior.colorIndex = SolidColorIdx And sheet.Cells(cellRange.row, colNum).Interior.Pattern = SolidPattern Then
'                        oneContGray = True
'                        oneBranchMatch = False
'                        eachContAttrVal = ""
'                    ElseIf UBound(Split(sheet.Cells(cellRange.row, colNum), "\")) = 2 Then ' 主控是引用，被控应不受控制，等价于主控为空
'                        oneContNull = True
'                        oneBranchMatch = False
'                        eachContAttrVal = ""
'                    Else
'                        eachContAttrVal = Trim(sheet.Cells(cellRange.row, colNum).value)
'                    End If
'                Else
'                    oneContNotExist = True
'                    oneBranchMatch = False
'                    Exit For
'                End If
'            End If
'            If oneBranchMatch = True And eachContAttrVal <> "" Then
'                '保存到数组中，后面用checkBranchMatch判断各个主控值是否匹配
'                contAttrValArray(contAttrNum) = eachContAttrVal
'                contAttrNum = contAttrNum + 1
'            ElseIf oneContGray = False Then
'                oneContNull = True
'                oneBranchMatch = False
'            End If
'        Next
'
'        '不存在灰化，空值，参数不存在等特殊情况，则判断各个参数值是否匹配
'        If contAttrNum = equalsNodes.Length Then
'            'oneBranchMatch = checkBranchMatch(equalsNodes)
'            oneBranchMatch = newCheckBranchMatch(controlAttrValueManager, controlInfo)
'        End If
'
'        If oneBranchMatch = False Then
'            If oneContNotExist = True Then
'                contedOutOfControl = True
'                Exit For
'            ElseIf oneContGray = True Then
'                contedOutOfControl = False
'                'Exit For
'            ElseIf oneContNull = True Then
'                contedOutOfControl = True
'                Exit For
'            End If
'        Else '此分支中各主控参数匹配成功，则进行分支控制
'            Set boundNodes = BranchNodes(I).ChildNodes
'            '获得被控参数的范围
'            boundValue = getContedAttrBoundValue(boundNodes, valIsRight, sheet, cellRange, contedColNum)
'            '进行分支控制
'            If sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = SolidColorIdx And sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = SolidPattern Then
'                If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
'                    sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
'                End If
'                sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = NullPattern
'                sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = NullPattern
'                sheet.Cells(cellRange.row, contedColNum).Validation.ShowInput = True
'            End If
'            '不在范围内时要清空
'            If valIsRight = False And Trim(sheet.Cells(cellRange.row, contedColNum).value) <> "" _
'                And contedType <> "IPV4" And contedType <> "IPV6" Then
'                If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
'                    sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
'                End If
'                sheet.Cells(cellRange.row, contedColNum).value = ""
'            End If
'            '设置被控参数的范围
'            Call setValidation(contedType, boundValue, sheet, cellRange.row, contedColNum)
'            checkAllBranchMatch = True
'            Exit For
'        End If
'    Next
End Function

Private Function newCheckBranchMatch(ByRef controlAttrValueManager As CControlAttrValueManager, ByRef controlInfo As String, ByRef matchBranchNode As Variant) As Boolean
    Dim branchMatchChecker As New CBranchMatchChecker
    Call branchMatchChecker.init(controlAttrValueManager, controlInfo)
    newCheckBranchMatch = branchMatchChecker.getOneBranchMatchFlag
    Set matchBranchNode = branchMatchChecker.getMatchBranchNode
End Function

Function getContedAttrBoundValue(boundNodes, valIsRight As Boolean, sheet As Worksheet, cellRange As range, contedColNum As Long) As String
    Dim k As Long
    Dim minValue, maxValue, tmp As Variant
    Dim boundValue As String
    Dim isfound As Boolean
     
    boundValue = ""
    valIsRight = False
    For k = 0 To boundNodes.length - 1
        If (boundNodes(k).nodeName = "EnumItem") Or (boundNodes(k).nodeName = "BitEnumItem") Then
            isfound = True
            If UBound(Split(sheet.Cells(cellRange.row, contedColNum).value, "\")) = 2 Or _
                Trim(sheet.Cells(cellRange.row, contedColNum).value) = boundNodes(k).getAttributeNode("name").NodeValue Then
                valIsRight = True
            End If
            If boundValue <> "" Then
                boundValue = boundValue + "," + boundNodes(k).getAttributeNode("name").NodeValue
            Else
                boundValue = boundNodes(k).getAttributeNode("name").NodeValue
            End If
        ElseIf boundNodes(k).nodeName = "NumBoundTrait" Or boundNodes(k).nodeName = "StringLenTrait" Then
            isfound = True
            minValue = boundNodes(k).getAttributeNode("min").NodeValue
            maxValue = boundNodes(k).getAttributeNode("max").NodeValue
            boundValue = boundValue + "[" + minValue + "," + maxValue + "]"
            If sheet.Cells(cellRange.row, contedColNum).value <> "" And _
                UBound(Split(sheet.Cells(cellRange.row, contedColNum).value, "\")) <> 2 Then
                If boundNodes(k).nodeName = "NumBoundTrait" Then
                    tmp = CDbl(sheet.Cells(cellRange.row, contedColNum).value)
                Else
                    tmp = Len(sheet.Cells(cellRange.row, contedColNum).value)
                End If
                If (tmp >= CDbl(minValue)) And (tmp <= CDbl(maxValue)) Then
                    valIsRight = True
                End If
            ElseIf UBound(Split(sheet.Cells(cellRange.row, contedColNum).value, "\")) = 2 Then
                valIsRight = True
            End If
        End If
    Next
    getContedAttrBoundValue = boundValue
        '不是枚举和数字类型时
    If isfound = False Then
        valIsRight = True
    End If
End Function
Function checkBranchMatch(equalsNodes) As Boolean
    Dim i As Long
    Dim j, k As Long
    Dim hasExist, isMatch1, isMatch2 As Boolean
        
    checkBranchMatch = True
    For i = 0 To equalsNodes.length - 1
        hasExist = False
        For j = i + 1 To equalsNodes.length - 1
            If i <> j And equalsNodes(i).getAttributeNode("attribute").NodeValue = equalsNodes(j).getAttributeNode("attribute").NodeValue Then
                hasExist = True
                isMatch1 = True
                isMatch2 = True
                For k = i To j - 1
                    If checkContValEquals(equalsNodes, k) = False Then
                        isMatch1 = False
                        Exit For
                    End If
                Next
                For k = j To equalsNodes.length - 1
                    If checkContValEquals(equalsNodes, k) = False Then
                        isMatch2 = False
                        Exit For
                    End If
                Next
                If isMatch1 = False And isMatch2 = False Then
                    checkBranchMatch = False
                End If
                Exit For
            End If
        Next
        If hasExist = True Then
            Exit For
        Else
            If checkContValEquals(equalsNodes, i) = False Then
                checkBranchMatch = False
            End If
        End If
    Next

End Function
Function checkContValEquals(equalsNodes, index As Long) As Boolean
    Dim minValue, maxValue As String
    
    checkContValEquals = True
    If equalsNodes(index).FirstChild.nodeName = "EnumItem" Then
        If contAttrValArray(index) <> equalsNodes(index).FirstChild.getAttributeNode("name").NodeValue Then
            checkContValEquals = False
        End If
    ElseIf equalsNodes(index).FirstChild.nodeName = "NumBoundTrait" Then
        minValue = equalsNodes(index).FirstChild.getAttributeNode("min").NodeValue
        maxValue = equalsNodes(index).FirstChild.getAttributeNode("max").NodeValue
        If (CDbl(contAttrValArray(index)) < CDbl(minValue)) Or (CDbl(contAttrValArray(index)) > CDbl(maxValue)) Then
            checkContValEquals = False
        End If
    End If
End Function
Sub getGroupAndColumnName(ByVal CurSheet As Worksheet, ByVal cellRange As range, groupName As String, columnName As String)
    Dim m_rowNum, m_colNum As Long
    
    If boardStyleColumnMap Is Nothing Then
        Call initBoardStyleMap
    End If
    
    If CurSheet.name = getResByKey("Comm Data") Or InStr(CurSheet.name, getResByKey("Board Style")) <> 0 Then
        For m_rowNum = cellRange.row To 1 Step -1
            'If CurSheet.Cells(m_rowNum, cellRange.column).Interior.colorIndex = 40 Then '34是绿色
            If boardStyleColumnMap.hasKey(Trim(CurSheet.Cells(m_rowNum, cellRange.column).value)) Then
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

Public Function getLongEnumValueFromValidDef(ByRef mocSheetName As String, ByRef mocGroupName As String, ByRef mocColumnName As String) As String
    On Error GoTo ErrorHandler
    Dim mappingDef As CMappingDef

    Set mappingDef = getMappingDefine(mocSheetName, mocGroupName, mocColumnName)
    Dim currmocName As String, currattrname As String
    currmocName = mappingDef.mocName
    currattrname = mappingDef.attributeName
    If VBA.Trim(currmocName) = "" Or VBA.Trim(currattrname) = "" Then
        getLongEnumValueFromValidDef = ""
        Exit Function
    End If
    Dim laststr As String

    Dim validSheet As Worksheet
    Set validSheet = ThisWorkbook.Sheets("VALID DEF")
    Dim index As Long
    For index = 2 To validSheet.range("a1048576").End(xlUp).row
        If GetDesStr(validSheet.Cells(index, 1).value) = GetDesStr(currmocName) And GetDesStr(validSheet.Cells(index, 2).value) = GetDesStr(currattrname) Then
            Dim maxColnum As Long
            maxColnum = validSheet.range("XFD" + CStr(index)).End(xlToLeft).column
            getLongEnumValueFromValidDef = "=INDIRECT(""'VALID DEF'!C" & CStr(index) & ":" & getColStr(maxColnum) & CStr(index) & """)"
        End If
    Next
    Exit Function
ErrorHandler:
    getLongEnumValueFromValidDef = ""
End Function

'**********************************************************
'从列数得到列名：1->A，27->AA
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
            If Len(boundValue) > 255 Then
                Dim groupName As String
                Dim columnName As String
                Dim cellRange As range
                Set cellRange = sheet.range(getColumnNameFromColumnNum(colNum) + CStr(rowNum))
                Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
                If isBoardStyleSheet(sheet) Then
                    boundValue = getLongEnumValueFromValidDef(getResByKey("Board Style"), groupName, columnName)
                Else
                    boundValue = getLongEnumValueFromValidDef(sheet.name, groupName, columnName)
                End If
                With sheet.Cells(rowNum, colNum).Validation
                    .Delete
                    .Add Type:=xlValidateList, formula1:=boundValue
                End With
            Else
                With sheet.Cells(rowNum, colNum).Validation
                    .Delete
                    .Add Type:=xlValidateList, formula1:=boundValue
                    .inputTitle = getResByKey("Range")
                    .inputMessage = "[" + boundValue + "]"
                    .ShowInput = True
                    .ShowError = True
                End With
            End If
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
'精确检查输入整数是否在离散范围内[1，2][3，4]
Function Check_Int_Validation(ByVal attrRange As String, ByVal attrValue As String) As Boolean
    Dim sItem As String
    Dim min, max As Double
    Dim nLoop As Long
    Dim isInFlag, bFlag As Boolean

    sItem = Right(Left(Trim(attrValue), 1), 1)
    If sItem = "#" And Len(Trim(attrValue)) > 1 Then
        For nLoop = 2 To Len(Trim(attrValue))
            sItem = Right(Left(Trim(attrValue), nLoop), 1)
            If sItem < "0" Or sItem > "9" Then
                Check_Int_Validation = False
                Exit Function
            End If
        Next
        Check_Int_Validation = True
        Exit Function
    End If
    
    isInFlag = False
    bFlag = True
    For nLoop = 1 To Len(Trim(attrValue))
        sItem = Right(Left(Trim(attrValue), nLoop), 1)
        If sItem < "0" Or sItem > "9" Then
            If nLoop = 1 And sItem = "-" Then
                bFlag = True
            Else
                bFlag = False
                Check_Int_Validation = False
                Exit Function
            End If
        End If
    Next
    
    If bFlag = True Then
        While attrRange <> ""
            min = CDbl(Mid(attrRange, 2, InStr(1, attrRange, ",") - 2))
            max = CDbl(Mid(attrRange, InStr(1, attrRange, ",") + 1, InStr(1, attrRange, "]") - InStr(1, attrRange, ",") - 1))
            If CDbl(attrValue) >= min And CDbl(attrValue) <= max Then
                Check_Int_Validation = True
                Exit Function
            End If
            attrRange = Mid(attrRange, InStr(1, attrRange, "]") + 1)
        Wend
    End If
    Check_Int_Validation = False
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

Public Function Moc_Appear_In_SameSheet(ByVal mocName As String, ByVal attrName As String) As Boolean
        
        Dim rowIndex As Long
        Dim found As Boolean
        found = False
        Dim mappingDef As Worksheet
        Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
        For rowIndex = 2 To mappingDef.range("a1048576").End(xlUp).row
                If (mappingDef.Cells(rowIndex, 1).value = ThisWorkbook.ActiveSheet.name _
                    And mappingDef.Cells(rowIndex, 4).value = mocName _
                    And mappingDef.Cells(rowIndex, 5).value = attrName) Then
                    If found Then
                            Moc_Appear_In_SameSheet = True
                            Exit Function
                    Else
                        found = True
                    End If
                End If
        Next
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

Public Sub boardStyleSheetControl(ByVal sh As Object, ByVal Target As range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Set sheet = sh
    If Target.count Mod 256 = 0 Then
        Exit Sub
    End If
    
    If boardStyleColumnMap Is Nothing Then
        Call initBoardStyleMap
    End If
    
    Dim currentNeType As String
    If sheet.name <> getResByKey("Board Style") Then actualBoardStyleName = sheet.name
    For Each cellRange In Target
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

Public Function isGTRXSheet(sheetName As String) As Boolean
    If sheetName = "GTRX" Or sheetName = getResByKey("GTRX_ZH") Then
        isGTRXSheet = True
        Exit Function
    End If
    isGTRXSheet = False
End Function

Public Function isGsmCellSheet(sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = getResByKey("GSM Logic Cell") Then
        isGsmCellSheet = True
        Exit Function
    End If
    isGsmCellSheet = False
End Function

Public Function isBaseStationSheet(sheetName As String) As Boolean
    If sheetName = "Base Station Transport Data" Or sheetName = getResByKey("BaseTransPort") Then
        isBaseStationSheet = True
        Exit Function
    End If
    isBaseStationSheet = False
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

Public Function cellIsGray(ByRef certainCell As range) As Boolean
    If certainCell.Interior.colorIndex = SolidColorIdx And certainCell.Interior.Pattern = SolidPattern Then
        cellIsGray = True
    Else
        cellIsGray = False
    End If
End Function

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

Public Function selectionIsValid(ByRef ws As Worksheet, ByRef cellRange As range) As Boolean
    
    If boardStyleColumnMap Is Nothing Then
        Call initBoardStyleMap
    End If
    
    If cellRange.value = "" And cellRange.Borders.LineStyle <> xlLineStyleNone Then
        selectionIsValid = True
    '校验用户选中是否为组或者列，如果不是则认为单元格合入
    ElseIf boardStyleGroupMap.hasKey(Trim(cellRange.value)) Or boardStyleColumnMap.hasKey(Trim(cellRange.value)) Or cellRange.Borders.LineStyle = xlLineStyleNone Then
        selectionIsValid = False
    Else
        selectionIsValid = True
    End If
End Function

Public Sub currentParameterBranchCheck(ByRef ws As Worksheet, ByRef cellRange As range)
    On Error GoTo ErrorHandler
    If cellRange.count > 1 Then Exit Sub '选择的单元格大于1，则退出
    
    If selectionIsValid(ws, cellRange) = False Then Exit Sub '如果选择的单元格非法，直接退出
    
    '如果已经灰化了，则不需要控制了，退出
    If cellIsGray(cellRange) Then Exit Sub
    
    Dim controlDef As CControlDef
    Dim groupName As String, columnName As String, sheetName As String
    
    Call getGroupAndColumnName(ws, cellRange, groupName, columnName)
    
    If Not isBoardStyleSheet(ws) Then
        Set controlDef = getControlDefine(ws.name, groupName, columnName)
    Else
        If ws.name <> getResByKey("Board Style") Then actualBoardStyleName = ws.name
        Set controlDef = getControlDefine(getResByKey("Board Style"), groupName, columnName)
    End If
    
    '未找到Control Def控制信息，退出
    If controlDef Is Nothing Then Exit Sub
    
    Dim mocName As String, attrName As String, neType As String, controlInfo As String
    Call getControlAttrInfo(controlDef, mocName, attrName, neType, sheetName, controlInfo)
    
     '如果该参数不是被控参数，则退出
    If Not controlRelationManager.containsControlledAttributeRelation(mocName, attrName, neType, sheetName) Then Exit Sub
    
    '找到该参数的所有主控参数类
    Dim controlRelation As CControlRelation
    Set controlRelation = controlRelationManager.getControlRelation(mocName, attrName, neType, sheetName)
    
    Dim controlAttrValueManager As New CControlAttrValueManager '管理所有主控参数的管理类，后续用这个主控参数管理类与被参数的分支控制信息作校验
    Dim mainControlAttrReturnedValue As Long
    mainControlAttrReturnedValue = makeControlAttrValueCol(ws, mocName, attrName, neType, sheetName, controlRelation, controlAttrValueManager, cellRange)
    
    '如果返回值是1，要么主控参数不全，要么主控参数值为空或为引用，则不需要分支校验，直接退出
    If mainControlAttrReturnedValue = 1 Then
        Exit Sub
    ElseIf mainControlAttrReturnedValue = 3 Then '如果返回值是3，说明其中一个主控参数灰化，则当前分支无效，灰化退出
        If cellIsGray(cellRange) Then Exit Sub '此时有可能已经由主控参数的灰化导致该被控参数灰化了，那么就不需要再次灰化当前单元格了，直接退出，提高效率
        Call setRangeGray(cellRange)
        Exit Sub
    End If
    
    Dim oneBranchMatchFlag As Boolean
    Dim branchMatchChecker As New CBranchMatchChecker
    Call branchMatchChecker.init(controlAttrValueManager, controlInfo)
    oneBranchMatchFlag = branchMatchChecker.getOneBranchMatchFlag
    
    If oneBranchMatchFlag = False Then '说明未找到匹配分支，则将当前单元灰化
        Call setRangeGray(cellRange)
    Else '如果不需要灰化，那么需要增加被控分支的Tip和下拉框有效性
        Call setControlledParameterTipAndValidation(ws, cellRange, controlDef.dataType, branchMatchChecker)
    End If
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub setRangeGray(ByRef certainRange As range)
    On Error Resume Next
    certainRange.Interior.colorIndex = SolidColorIdx
    certainRange.Interior.Pattern = SolidPattern
    certainRange.value = ""
    certainRange.Validation.ShowInput = False
End Sub

Private Sub setControlledParameterTipAndValidation(ByRef ws As Worksheet, ByRef cellRange As range, ByRef valueType As String, ByRef branchMatchChecker As CBranchMatchChecker)
    If targetHasInputMessage(cellRange) Then Exit Sub  '如果有了InputMessage的Tip，则退出
    
    If branchMatchChecker.getMatchBranchAttrEmptyFlag = True Then Exit Sub '如果某个主控参数为空，则不需要加批注和下拉框，退出
    
    Dim matchBranchNode As Variant, boundNodes As Variant
    Set matchBranchNode = branchMatchChecker.getMatchBranchNode
    Set boundNodes = matchBranchNode.childNodes
    
    Dim boundValue As String
    Dim valIsRight As Boolean '没什么用，只是为了调用之前的函数getContedAttrBoundValue
    '获得被控参数的范围
    boundValue = getContedAttrBoundValue(boundNodes, valIsRight, ws, cellRange, cellRange.column)
    Call setValidation(valueType, boundValue, ws, cellRange.row, cellRange.column)
End Sub

Private Function targetHasInputMessage(ByRef Target As range) As Boolean
    On Error GoTo ErrorHandler
    targetHasInputMessage = True
    If Target.Validation Is Nothing Then '没有有效性，则没有InputMessage
        targetHasInputMessage = False
        Exit Function
    End If
    
    Dim inputMessage As String
    inputMessage = Target.Validation.inputMessage '如果有InputMessage，则赋值成功，如果没有，则赋值出错，进入ErrorHandler
    If inputMessage = "" Then targetHasInputMessage = False '如果InputMessage为空，则认为没有Tip
    Exit Function
ErrorHandler:
    targetHasInputMessage = False
End Function

Private Function makeControlAttrValueCol(ByRef ws As Worksheet, ByRef mocName As String, ByRef attributeName As String, ByRef neType As String, ByRef sheetName As String, ByRef controlRelation As CControlRelation, ByRef controlAttrValueManager As CControlAttrValueManager, ByRef cellRange As range) As Long
    makeControlAttrValueCol = -1
    Dim cellRow As Long, cellColumn As Long, noUse As Long
    cellRow = cellRange.row
    
    Dim eachControlAttr As Variant
    Dim mainControlAttr As String, mainControlGroupName As String, mainControlColumnName As String, mainValue As String, mainValueType As String
    Dim mainControlColumnNumber As Long
    
    Dim controlAttrCol As Collection
    Set controlAttrCol = controlRelation.controlAttrCol
    
    Dim mainControlRelation As CControlRelation '这是主控参数的Control Def管理类
    Dim mainControlAttrValue As CControlAttrValue '管理一个主控参数值，类型，长短名称的类
    
    For Each eachControlAttr In controlAttrCol
        mainControlAttr = CStr(eachControlAttr)
        '如果在当前页签没有找到该主控参数的控制信息，说明主控参数缺少，则该被控参数不控制
        If Not controlRelationManager.containsAttributeRelation(mocName, mainControlAttr, neType, sheetName) Then
            makeControlAttrValueCol = 1 '1表示在控制范围内
            Exit Function
        End If
        
        Set mainControlRelation = controlRelationManager.getControlRelation(mocName, mainControlAttr, neType, sheetName)
        
        mainControlGroupName = mainControlRelation.groupName
        mainControlColumnName = mainControlRelation.columnName
        mainValueType = mainControlRelation.valueType
                
        mainControlColumnNumber = get_colNum(sheetName, mainControlGroupName, mainControlColumnName, noUse)
        '该主控参数未在页签中找到，说明Control Def有冗余信息，不做校验退出
        If mainControlColumnNumber = 0 Then
            makeControlAttrValueCol = 1 '1表示在控制范围内
            Exit Function
        End If
        
        Dim mainAttrCell As range
        Set mainAttrCell = ws.Cells(cellRow, mainControlColumnNumber)
        
        mainValue = mainAttrCell.value
        
        Dim valueEmptyFlag As Boolean, valueReferenceFlag As Boolean, valueCellGrayFlag As Boolean
        valueEmptyFlag = False
        valueReferenceFlag = False
        valueCellGrayFlag = False
        
        If UBound(Split(mainValue, "\")) = 2 Then  '如果其中一个主控参数为引用，标志设为True，由对象自己在类中check
            valueReferenceFlag = True '主控为引用标志
        ElseIf mainValue = "" Then '此时主控参数为空，有可能该主控参数也是无效的，则需要先对该主控单元格进行校验
            Call currentParameterBranchCheck(ws, mainAttrCell)
            If cellIsGray(cellRange) Then
                makeControlAttrValueCol = 1 '可能由于主控参数的灰化改变已经使当前单元格灰化了，则不需要判断了，直接退出
                Exit Function
            End If
        End If
        
        If cellIsGray(mainAttrCell) Then '如果某个主控参数灰化，标志设为True，由对象自己在类中check
            valueCellGrayFlag = True '主控为灰化标志
            valueEmptyFlag = True '主控为灰化，则肯定为空
        ElseIf mainValue = "" Then '如果主控单元格未灰化，说明是主控单元格是有效分支，只是未填写值，空标志设为True，由对象自己在类中check
            valueEmptyFlag = True
        End If
        
        Set mainControlAttrValue = New CControlAttrValue
        Call mainControlAttrValue.init(mainControlAttr, mainControlColumnName, mainValue, mainValueType, valueEmptyFlag, valueReferenceFlag, valueCellGrayFlag)
        
        Call controlAttrValueManager.addNewControlAttrValue(mainControlAttrValue)
    Next eachControlAttr
    makeControlAttrValueCol = 2 '2表格各主控参数都是有值的，需要进行分支控制判断
End Function

Private Sub getControlAttrInfo(ByRef controlDef As CControlDef, ByRef mocName As String, ByRef attrName As String, ByRef neType As String, ByRef sheetName As String, ByRef controlInfo As String)
    mocName = controlDef.mocName
    attrName = controlDef.attributeName
    neType = controlDef.neType
    sheetName = controlDef.sheetName
    controlInfo = controlDef.controlInfo
End Sub

