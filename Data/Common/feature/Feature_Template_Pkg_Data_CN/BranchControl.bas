Attribute VB_Name = "BranchControl"
Option Explicit
Public Type controlRelation
    mocName As String
    contAttr As String
    contedAttrs(100) As String
    contedNum As Integer '从0开始
    neType  As String '此参数是为了区分 是控制器的还是物理站的，现在控制器和物理站有可能Moc和参数属性名是一样的
    sheetName As String ' 此参数是为了区分，同一Moc出现在不同的Sheet，Comm Data页会出现，这时候，不能跨sheet页签控制
End Type
Public ControlRelMap(1000) As controlRelation
Public contAttrValArray(100) As String
Dim ControlRelationNum As Integer

Private controlRelationManager As CControlRelationManager

'用以设置颜色
Const MustGiveColorIdx = 46
Const SolidColorIdx = 16
Const SolidPattern = xlGray16
Const NullPattern = xlNone
Function isControlDefSheetExist() As Boolean
    Dim sheetNum As Integer
    isControlDefSheetExist = False
    For sheetNum = 1 To ThisWorkbook.Worksheets.count
        If ControllSheetName = ThisWorkbook.Worksheets(sheetNum).name Then
            isControlDefSheetExist = True
            Exit For
        End If
    Next
End Function
Sub buildControlRelations()
    Dim conRowNum As Integer
    Dim i, j As Integer
    Dim conInfor As String, mocName As String, contedName As String, contName As String
    Dim isExist, flag, isFound As Boolean
    Dim index1, index2, index3, index4 As Integer
    Dim neType As String
    Dim sheetName As String
    
    Dim GroupName As String, columnName As String, valueType As String
    
    Set controlRelationManager = New CControlRelationManager '当前参数范围校验管理类
    
    If isControlDefSheetExist = False Then
        Exit Sub
    End If
    
    ControlRelationNum = 0
    Dim controldef As Worksheet
    Set controldef = ThisWorkbook.Worksheets(ControllSheetName)
    For conRowNum = HiddenSheetValidRowBegin To controldef.Range("a65536").End(xlUp).row

        mocName = Trim(controldef.Cells(conRowNum, 1).value)
        contedName = Trim(controldef.Cells(conRowNum, 2).value)
        conInfor = Trim(controldef.Cells(conRowNum, 6).value)
        neType = Trim(controldef.Cells(conRowNum, 10).value)
        sheetName = Trim(controldef.Cells(conRowNum, 7).value)
    
        GroupName = Trim(controldef.Range("H" & conRowNum).value)
        columnName = Trim(controldef.Range("I" & conRowNum).value)
        valueType = Trim(controldef.Range("C" & conRowNum).value)
        Call controlRelationManager.addNewAttributeRelation(mocName, contedName, conInfor, neType, sheetName, GroupName, columnName, valueType)
        
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
                                isFound = False
                                For index1 = 0 To ControlRelationNum - 1
                                    If (mocName = ControlRelMap(index1).mocName) And (contedName = ControlRelMap(index1).contAttr) And (neType = ControlRelMap(index1).neType) And (sheetName = ControlRelMap(index1).sheetName) Then
                                        For index2 = 0 To ControlRelMap(index1).contedNum '被控参数B的控制列表
                                            For index3 = 0 To ControlRelMap(i).contedNum    '主控参数A的控制列表
                                                If ControlRelMap(index1).contedAttrs(index2) = ControlRelMap(i).contedAttrs(index3) Then
                                                    For index4 = ControlRelMap(i).contedNum To index3 Step -1
                                                        ControlRelMap(i).contedAttrs(index4 + 1) = ControlRelMap(i).contedAttrs(index4)
                                                    Next
                                                    ControlRelMap(i).contedAttrs(index3) = contedName
                                                    isFound = True
                                                    Exit For
                                                End If
                                            Next
                                            If isFound = True Then
                                                Exit For
                                            End If
                                        Next
                                        Exit For
                                    End If
                                Next
                                ControlRelMap(i).contedNum = ControlRelMap(i).contedNum + 1
                                If isFound = False Then
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
Function Check_Value_Validation(ByVal curSheet As Worksheet, ByVal cellRange As Range, ByRef isConAttr As Boolean, ByRef contRel As controlRelation, ByRef currentNeType As String) As Integer
    On Error Resume Next
    
    Dim GroupName As String
    Dim columnName As String
    Dim mocName, attrName, attrType, attrRange As String
    Dim conRowNum, i As Integer
    Dim neType As String
    Dim sheetName As String
    Dim controldef As CControlDef

    Call getGroupAndColumnName(curSheet, cellRange, GroupName, columnName)
    Set controldef = getControlDefine(curSheet.name, GroupName, columnName)
    
    If controldef Is Nothing Then
        Check_Value_Validation = -1
        Exit Function
    End If
    
    mocName = controldef.mocName
    attrName = controldef.attributeName
    neType = controldef.neType
    sheetName = controldef.sheetName
    
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
        attrType = controldef.dataType
        attrRange = controldef.bound + controldef.lstValue
        If Trim(controldef.controlInfo) <> "" Then
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
                If controlledAttrValidationCheck(curSheet, controldef, cellRange) = True Then '校验的参数在范围内
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
Private Function controlledAttrValidationCheck(ByRef ws As Worksheet, ByRef controldef As CControlDef, ByRef cellRange As Range) As Boolean
    On Error GoTo ErrorHandler
    controlledAttrValidationCheck = False
    
    Dim mocName As String, attrName As String, neType As String, controlInfo As String, sheetName As String
    Call getControlAttrInfo(controldef, mocName, attrName, neType, sheetName, controlInfo)
    
     '如果该参数不是被控参数，则退出
    If Not controlRelationManager.containsControlledAttributeRelation(mocName, attrName, neType, sheetName) Then Exit Function
    
    '找到该参数的所有主控参数类
    Dim controlRelation As CControlRelation
    Set controlRelation = controlRelationManager.getControlRelation(mocName, attrName, neType, sheetName)
    
    Dim controlAttrValueManager As New CControlAttrValueManager '管理所有主控参数的管理类，后续用这个主控参数管理类与被参数的分支控制信息作校验
    Dim mainControlAttrReturnedValue As Integer
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
    Call controlAttrValue.init("", "", cellRange.value, controldef.dataType, False, False, False, True)
    Dim branchAttrValueValidationFlag As Boolean
    branchAttrValueValidationFlag = controlAttrValue.checkABranchAttrValues(matchBranchNode)
    
    If branchAttrValueValidationFlag = False Then
        Dim nResponse As Integer
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
Sub getGroupAndColumnName(ByVal curSheet As Worksheet, ByVal cellRange As Range, GroupName As String, columnName As String)
    Dim m_rowNum, m_colNum As Integer
    
    columnName = curSheet.Cells(DataSheetAttrRow, cellRange.Column).value
    For m_colNum = cellRange.Column To 1 Step -1
        If curSheet.Cells(DataSheetMocRow, m_colNum).value <> "" Then
            GroupName = curSheet.Cells(DataSheetMocRow, m_colNum).value
            Exit For
        End If
    Next
End Sub

Function Check_Value_In_Range(ByVal attrType As String, ByVal attrRange As String, ByVal attrValue As String, cellRange As Range, ByRef alreadyCheckFlag As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim arrayList() As String
    Dim errorMsg, sItem As String
    Dim i, nResponse, nLoop As Integer
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
Sub Execute_Branch_Control(ByVal sheet As Worksheet, ByVal cellRange As Range, contRel As controlRelation, ByRef currentNeType As String)
    On Error Resume Next
    
    If cellRange.row < DataSheetDataRowBegin Then
        Exit Sub
    End If
    
    Dim sheetName, GroupName, columnName, attrName As String
    Dim branchInfor As String, contedType As String
    Dim boundValue As String

    Dim allBranchMatch, contedOutOfControl As Boolean
    Dim xmlObject As Object
    Dim m, conRowNum, contedColNum As Integer
    Dim noUse As Integer
    Dim rootNode As Variant
    Dim controldef As Worksheet
    Set controldef = ThisWorkbook.Worksheets(ControllSheetName)
    '对各个被控参数进行分支控制
    For m = 0 To contRel.contedNum
        For conRowNum = HiddenSheetValidRowBegin To controldef.Range("a65536").End(xlUp).row
            If (contRel.mocName = controldef.Cells(conRowNum, 1).value) _
                            And contRel.neType = controldef.Cells(conRowNum, 10).value _
                            And (contRel.contedAttrs(m) = controldef.Cells(conRowNum, 2).value) _
                            And (contRel.sheetName = controldef.Cells(conRowNum, 7).value) Then
                sheetName = controldef.Cells(conRowNum, 7).value
                GroupName = controldef.Cells(conRowNum, 8).value
                columnName = controldef.Cells(conRowNum, 9).value
                contedType = controldef.Cells(conRowNum, 3).value
                attrName = controldef.Cells(conRowNum, 2).value
                contedColNum = get_colNum(sheetName, GroupName, columnName, noUse)
                If (Trim(cellRange.value) = "" And cellRange.Interior.ColorIndex <> SolidColorIdx And cellRange.Interior.Pattern <> SolidPattern) Or UBound(Split(cellRange.value, "\")) = 2 Then '主控为空或是引用，则此时被控应变为非灰即有效，范围恢复成初始值
                    If sheet.Cells(cellRange.row, contedColNum).Interior.ColorIndex = SolidColorIdx And sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = SolidPattern Then
                        If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
                            sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
                        End If
                        'Call setRangeNormal(sheet.Cells(cellRange.row, contedColNum))
                        Call setRangeColor(sheetName, attrName, sheet.Cells(cellRange.row, contedColNum))
                    End If
                    '恢复成初始范围
                    boundValue = controldef.Cells(conRowNum, 4).value + controldef.Cells(conRowNum, 5).value
                    Call setValidation(contedType, boundValue, sheet, cellRange.row, contedColNum)
                    Call setEmptyValidation(sheet, cellRange.row, contedColNum)
                Else '主控不空，进行contRel.contedAttrs(m)的分支控制
                    branchInfor = controldef.Cells(conRowNum, 6).value
                    Set xmlObject = CreateObject("msxml2.domdocument")
                    xmlObject.LoadXML branchInfor
                    'Set BranchNodes = xmlObject.DocumentElement.ChildNodes
                    Set rootNode = xmlObject.DocumentElement
                    contedOutOfControl = False
                    allBranchMatch = checkAllBranchMatch(rootNode, sheet, cellRange, contRel, contedType, contedOutOfControl, contedColNum, currentNeType, branchInfor, sheetName, attrName)
                    '所有主控参数的值都不在分支参数规定的范围内，则被控参数被灰化（但被控参数不受控制除外）
                    If allBranchMatch = False Then
                        If contedOutOfControl = False Then
                            Call setRangeGray(sheet.Cells(cellRange.row, contedColNum))
                        Else
                            'Call setRangeNormal(sheet.Cells(cellRange.row, contedColNum))
                            Call setRangeColor(sheetName, attrName, sheet.Cells(cellRange.row, contedColNum))
                        End If
                        If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
                            sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
                        End If
                    End If
                End If
                Exit For
            End If
        Next
    Next
    
End Sub

Sub deleteValidation(ByRef sheet As Worksheet, ByRef rowNumber As Integer, ByRef columnNumber As Integer)
    sheet.Cells(rowNumber, columnNumber).Validation.Delete
End Sub

Private Sub makeControlAttrValueManager(ByRef sheet As Worksheet, ByRef dstRowNumber As Integer, ByRef mainControlMocName As String, ByRef mainControlAttrName As String, _
    ByRef neType As String, ByRef virtualSheetName As String, ByRef oneMainControlAttrNotExist As Boolean, ByRef controlAttrValueManager As CControlAttrValueManager)
    
    '如果已经有了该主控参数，就无需再添加了，直接退出
    If controlAttrValueManager.hasControlAttr(mainControlAttrName) Then Exit Sub
    
    Dim dstColumnNumber As Integer
    dstColumnNumber = get_MocAndAttrcolNum(mainControlMocName, mainControlAttrName, virtualSheetName) '得到主控参数的列号
    If dstColumnNumber = 0 Then '等于0，说明主控参数不存在，则退出，并将flag置为True
        'oneMainControlAttrNotExist = True
        Call addNoneExistMainAttr(mainControlAttrName, controlAttrValueManager)
        Exit Sub
    End If

    '如果在当前页签没有找到该主控参数的控制信息，说明主控参数缺少，则该被控参数不控制
    If Not controlRelationManager.containsAttributeRelation(mainControlMocName, mainControlAttrName, neType, virtualSheetName) Then
        'oneMainControlAttrNotExist = True
        Call addNoneExistMainAttr(mainControlAttrName, controlAttrValueManager)
        Exit Sub
    End If
    
    Dim mainControlRelation As CControlRelation '这是主控参数的Control Def管理类
    Dim mainControlAttrValue As CControlAttrValue '管理一个主控参数值，类型，长短名称的类
    Dim mainControlGroupName As String, mainControlColumnName As String, mainValueType As String
    
    Set mainControlRelation = controlRelationManager.getControlRelation(mainControlMocName, mainControlAttrName, neType, virtualSheetName)
    
    mainControlGroupName = mainControlRelation.GroupName
    mainControlColumnName = mainControlRelation.columnName
    mainValueType = mainControlRelation.valueType
    
    Dim mainAttrCell As Range
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
    Call mainControlAttrValue.init(mainControlAttrName, mainControlColumnName, mainValue, mainValueType, valueEmptyFlag, valueReferenceFlag, valueCellGrayFlag, True)
    
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

Function checkAllBranchMatch(rootNode As Variant, sheet As Worksheet, cellRange As Range, contRel As controlRelation, contedType As String, contedOutOfControl As Boolean, contedColNum As Integer, ByRef currentNeType As String, ByRef controlInfo As String, _
    ByVal sheetName As String, ByVal attrName As String) As Boolean
    
    On Error Resume Next
    
    Dim matchBranchNode As Variant '匹配的分支节点
    
    Dim i, j, colNum As Integer
    Dim contAttrNum As Integer
    Dim equalsNodes, boundNodes
    Dim eachContAttr As String, eachContAttrVal As String, boundValue As String
    Dim oneBranchMatch As Boolean, oneContNotExist As Boolean, oneContGray As Boolean, oneContNull As Boolean
    Dim valIsRight As Boolean
    
    checkAllBranchMatch = False

    Dim oneMainControlAttrNotExist As Boolean
    oneMainControlAttrNotExist = False
    
    Dim cellRowNumber As Integer
    cellRowNumber = cellRange.row
    
    Dim controlAttrValueManager As New CControlAttrValueManager
    
    Dim mainControlAttrCol As Collection
    Set mainControlAttrCol = getMainControlAttrCol(rootNode) '得到主控参数容器
    
    Dim eachMainControlAttr As Variant
    For Each eachMainControlAttr In mainControlAttrCol
        Call makeControlAttrValueManager(sheet, cellRowNumber, contRel.mocName, CStr(eachMainControlAttr), currentNeType, contRel.sheetName, oneMainControlAttrNotExist, controlAttrValueManager)
    Next eachMainControlAttr

'    If oneMainControlAttrNotExist = True Then '主控参数缺少，则认定在范围内，直接退出
'        contedOutOfControl = True
'        Exit Function
'    End If
    
    oneBranchMatch = newCheckBranchMatch(controlAttrValueManager, controlInfo, matchBranchNode)
    If oneBranchMatch = False Then '未找到匹配分支
        contedOutOfControl = False
    Else '此分支中各主控参数匹配成功，则进行分支控制
        Set boundNodes = matchBranchNode.childNodes
        '获得被控参数的范围
        boundValue = getContedAttrBoundValue(boundNodes, valIsRight, sheet, cellRange, contedColNum)
        '进行分支控制
        If sheet.Cells(cellRange.row, contedColNum).Interior.ColorIndex = SolidColorIdx And sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = SolidPattern Then
            If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
                sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
            End If
            Call setRangeColor(sheetName, attrName, sheet.Cells(cellRange.row, contedColNum))
            'Call setRangeNormal(sheet.Cells(cellRange.row, contedColNum))
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
    
End Function

Private Function newCheckBranchMatch(ByRef controlAttrValueManager As CControlAttrValueManager, ByRef controlInfo As String, ByRef matchBranchNode As Variant) As Boolean
    Dim branchMatchChecker As New CBranchMatchChecker
    Call branchMatchChecker.init(controlAttrValueManager, controlInfo)
    newCheckBranchMatch = branchMatchChecker.getOneBranchMatchFlag
    Set matchBranchNode = branchMatchChecker.getMatchBranchNode
End Function

Function getContedAttrBoundValue(boundNodes, valIsRight As Boolean, sheet As Worksheet, cellRange As Range, contedColNum As Integer) As String
    Dim k As Integer
    Dim minValue, maxValue, tmp As Variant
    Dim boundValue As String
    Dim isFound As Boolean
     
    boundValue = ""
    valIsRight = False
    For k = 0 To boundNodes.Length - 1
        If (boundNodes(k).nodeName = "EnumItem") Or (boundNodes(k).nodeName = "BitEnumItem") Then
            isFound = True
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
            isFound = True
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
    If isFound = False Then
        valIsRight = True
    End If
End Function
Function checkBranchMatch(equalsNodes) As Boolean
    Dim i As Integer
    Dim j, k As Integer
    Dim hasExist, isMatch1, isMatch2 As Boolean
        
    checkBranchMatch = True
    For i = 0 To equalsNodes.Length - 1
        hasExist = False
        For j = i + 1 To equalsNodes.Length - 1
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
                For k = j To equalsNodes.Length - 1
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
Function checkContValEquals(equalsNodes, index As Integer) As Boolean
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
Sub setValidation(ByVal contedType As String, ByVal boundValue As String, sheet As Worksheet, ByVal rowNum As Integer, ByVal colNum As Integer)
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
        If boundValue <> sheet.Cells(rowNum, colNum).Validation.Formula1 Or sheet.Cells(rowNum, colNum).Validation.inputMessage = "" Then
            With sheet.Cells(rowNum, colNum).Validation
                .Add Type:=xlValidateList, Formula1:=boundValue
                .inputTitle = getResByKey("Range")
                .inputMessage = "[" + boundValue + "]"
                .ShowInput = True
                .ShowError = True
            End With
            sheet.Cells(rowNum, colNum).Validation.Modify Type:=xlValidateList, Formula1:=boundValue
        End If
    End If
    
End Sub
Function get_colNum(ByVal sheetName As String, ByVal GroupName As String, ByVal columnName As String, rowNum As Integer) As Integer
    Dim m_colNum1, m_colNum2, m_rowNum As Integer
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(sheetName)
    For m_colNum1 = 1 To ws.Range("IV" + CStr(DataSheetAttrRow)).End(xlToLeft).Column
        If columnName = ws.Cells(DataSheetAttrRow, m_colNum1).value Then
            m_colNum2 = m_colNum1
            While Trim(ws.Cells(DataSheetMocRow, m_colNum2).value) = ""
                m_colNum2 = m_colNum2 - 1
            Wend
            If GroupName = Trim(ws.Cells(DataSheetMocRow, m_colNum2).value) Then
                get_colNum = m_colNum1
                Exit For
            End If
        End If
    Next
End Function
Function get_MocAndAttrcolNum(ByVal mocName As String, ByVal attrName As String, ByVal sheetName As String) As Integer
    Dim conRowNum, noUse As Integer
    Dim GroupName, columnName As String
    Dim controldef As Worksheet
    Set controldef = ThisWorkbook.Worksheets(ControllSheetName)
    get_MocAndAttrcolNum = 0
    For conRowNum = HiddenSheetValidRowBegin To controldef.Range("a65536").End(xlUp).row
        If (mocName = controldef.Cells(conRowNum, 1).value) _
            And (attrName = controldef.Cells(conRowNum, 2).value) _
            And (sheetName = controldef.Cells(conRowNum, 7).value) Then
            GroupName = controldef.Cells(conRowNum, 8).value
            columnName = controldef.Cells(conRowNum, 9).value
            get_MocAndAttrcolNum = get_colNum(sheetName, GroupName, columnName, noUse)
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
Public Function Ensure_Gray_Cell(curRange As Range) As Boolean
    If curRange.value <> "" And curRange.Interior.ColorIndex = SolidColorIdx And curRange.Interior.Pattern = SolidPattern Then
        MsgBox getResByKey("NoInput"), vbOKOnly + vbExclamation + vbApplicationModal + vbDefaultButton1, getResByKey("Warning")
        curRange.Select
        Ensure_Gray_Cell = True
    Else
        Ensure_Gray_Cell = False
    End If
End Function

Public Function Moc_Appear_In_SameSheet(ByVal mocName As String, ByVal attrName As String) As Boolean
        
        Dim rowIndex As Integer
        Dim found As Boolean
        found = False
        Dim mappingDef As Worksheet
        Set mappingDef = ThisWorkbook.Worksheets(MappingSheetName)
        For rowIndex = HiddenSheetValidRowBegin To mappingDef.Range("a65536").End(xlUp).row
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

Private Sub setEmptyValidation(ByRef sheet As Worksheet, ByRef rowNumber As Integer, ByRef columnNumber As Integer)
    On Error Resume Next
    With sheet.Cells(rowNumber, columnNumber).Validation
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
        .inputTitle = ""
        .inputMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
End Sub

Public Function cellIsGray(ByRef certainCell As Range) As Boolean
    If certainCell.Interior.ColorIndex = SolidColorIdx And certainCell.Interior.Pattern = SolidPattern Then
        cellIsGray = True
    Else
        cellIsGray = False
    End If
End Function

Public Function sheetSelectionShouldCheck(ByRef ws As Worksheet) As Boolean
    Dim sheetName As String
    sheetName = ws.name
    If (sheetName <> MappingSheetName) And (sheetName <> SheetDefName) And (sheetName <> ControllSheetName) And (sheetName <> FeatureListSheetName) _
        And (sheetName <> getResByKey("help")) And (sheetName <> getResByKey("Cover")) And (sheetName <> "SummaryRes") Then
        sheetSelectionShouldCheck = True
    Else
        sheetSelectionShouldCheck = False
    End If
End Function

Public Function selectionIsValid(ByRef ws As Worksheet, ByRef cellRange As Range) As Boolean
    If cellRange.Interior.ColorIndex = 34 Or cellRange.Interior.ColorIndex = 40 Or cellRange.Borders.LineStyle = xlLineStyleNone Then
        selectionIsValid = False
    Else
        selectionIsValid = True
    End If
End Function

Public Sub currentParameterBranchCheck(ByRef ws As Worksheet, ByRef cellRange As Range)
    On Error GoTo ErrorHandler
    If cellRange.count > 1 Then Exit Sub '选择的单元格大于1，则退出
    
    If selectionIsValid(ws, cellRange) = False Then Exit Sub '如果选择的单元格非法，直接退出
    
    '如果已经灰化了，则不需要控制了，退出
    If cellIsGray(cellRange) Then Exit Sub
    
    Dim controldef As CControlDef
    Dim GroupName As String, columnName As String, sheetName As String

'    If Not isIubStyleWorkSheet(ws.name) Then
    '不需要进行IUB表格的判断
    Call getGroupAndColumnName(ws, cellRange, GroupName, columnName)
    Set controldef = getControlDefine(ws.name, GroupName, columnName)
'    Else
'        Call getGroupNameShNameAndAttrName(ws, cellRange, groupName, sheetName, columnName)
'        Set controldef = getControlDefine(sheetName, groupName, columnName)
'    End If
    
    '未找到Control Def控制信息，退出
    If controldef Is Nothing Then Exit Sub
    
    Dim mocName As String, attrName As String, neType As String, controlInfo As String
    Call getControlAttrInfo(controldef, mocName, attrName, neType, sheetName, controlInfo)
    
     '如果该参数不是被控参数，则退出
    If Not controlRelationManager.containsControlledAttributeRelation(mocName, attrName, neType, sheetName) Then Exit Sub
    
    '找到该参数的所有主控参数类
    Dim controlRelation As CControlRelation
    Set controlRelation = controlRelationManager.getControlRelation(mocName, attrName, neType, sheetName)
    
    Dim controlAttrValueManager As New CControlAttrValueManager '管理所有主控参数的管理类，后续用这个主控参数管理类与被参数的分支控制信息作校验
    Dim mainControlAttrReturnedValue As Integer
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
        Call setControlledParameterTipAndValidation(ws, cellRange, controldef.dataType, branchMatchChecker)
    End If
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub setRangeGray(ByRef certainRange As Range)
    On Error Resume Next
    certainRange.Interior.ColorIndex = SolidColorIdx
    certainRange.Interior.Pattern = SolidPattern
    certainRange.value = ""
    certainRange.Validation.ShowInput = False
End Sub

Public Sub setRangeNormal(ByRef certainRange As Range)
    On Error Resume Next
    certainRange.Interior.ColorIndex = NullPattern
    certainRange.Interior.Pattern = NullPattern
    certainRange.Validation.ShowInput = True
End Sub

Public Sub setRangeMustGive(ByRef certainRange As Range)
    On Error Resume Next
    certainRange.Interior.Pattern = NullPattern
    certainRange.Interior.ColorIndex = MustGiveColorIdx
    certainRange.Validation.ShowInput = True
End Sub

Private Sub setControlledParameterTipAndValidation(ByRef ws As Worksheet, ByRef cellRange As Range, ByRef valueType As String, ByRef branchMatchChecker As CBranchMatchChecker)
    If targetHasInputMessage(cellRange) Then Exit Sub  '如果有了InputMessage的Tip，则退出
    
    If branchMatchChecker.getMatchBranchAttrEmptyFlag = True Then Exit Sub '如果某个主控参数为空，则不需要加批注和下拉框，退出
    
    Dim matchBranchNode As Variant, boundNodes As Variant
    Set matchBranchNode = branchMatchChecker.getMatchBranchNode
    Set boundNodes = matchBranchNode.childNodes
    
    Dim boundValue As String
    Dim valIsRight As Boolean '没什么用，只是为了调用之前的函数getContedAttrBoundValue
    '获得被控参数的范围
    boundValue = getContedAttrBoundValue(boundNodes, valIsRight, ws, cellRange, cellRange.Column)
    Call setValidation(valueType, boundValue, ws, cellRange.row, cellRange.Column)
End Sub

Private Function targetHasInputMessage(ByRef target As Range) As Boolean
    On Error GoTo ErrorHandler
    targetHasInputMessage = True
    If target.Validation Is Nothing Then '没有有效性，则没有InputMessage
        targetHasInputMessage = False
        Exit Function
    End If
    
    Dim inputMessage As String
    inputMessage = target.Validation.inputMessage '如果有InputMessage，则赋值成功，如果没有，则赋值出错，进入ErrorHandler
    If inputMessage = "" Then targetHasInputMessage = False '如果InputMessage为空，则认为没有Tip
    Exit Function
ErrorHandler:
    targetHasInputMessage = False
End Function

Private Sub addNoneExistMainAttr(ByRef mainControlAttr As String, ByRef controlAttrValueManager As CControlAttrValueManager)
    Dim mainControlAttrValue As CControlAttrValue '管理一个主控参数值，类型，长短名称的类
    Set mainControlAttrValue = New CControlAttrValue
    Call mainControlAttrValue.init(mainControlAttr, "", "", "", False, False, False, False)
        
    Call controlAttrValueManager.addNewControlAttrValue(mainControlAttrValue)
End Sub

Private Function makeControlAttrValueCol(ByRef ws As Worksheet, ByRef mocName As String, ByRef attributeName As String, ByRef neType As String, ByRef sheetName As String, ByRef controlRelation As CControlRelation, ByRef controlAttrValueManager As CControlAttrValueManager, ByRef cellRange As Range) As Integer
    makeControlAttrValueCol = -1
    Dim cellRow As Integer, cellColumn As Integer, noUse As Integer
    cellRow = cellRange.row
    
    Dim eachControlAttr As Variant
    Dim mainControlAttr As String, mainControlGroupName As String, mainControlColumnName As String, mainValue As String, mainValueType As String
    Dim mainControlColumnNumber As Integer
    
    Dim controlAttrCol As Collection
    Set controlAttrCol = controlRelation.controlAttrCol
    
    Dim mainControlRelation As CControlRelation '这是主控参数的Control Def管理类
    Dim mainControlAttrValue As CControlAttrValue '管理一个主控参数值，类型，长短名称的类
    
    Dim mainAttrExistFlag As Boolean '某个主控参数是否存在的标志
    mainAttrExistFlag = True
    
    For Each eachControlAttr In controlAttrCol
        mainControlAttr = CStr(eachControlAttr)
        '如果在当前页签没有找到该主控参数的控制信息，说明主控参数缺少，则该被控参数不控制
        If Not controlRelationManager.containsAttributeRelation(mocName, mainControlAttr, neType, sheetName) Then
            'makeControlAttrValueCol = 1 '1表示在控制范围内
            'mainAttrExistFlag = False
            Call addNoneExistMainAttr(mainControlAttr, controlAttrValueManager)
            GoTo NextLoop
        End If
        
        Set mainControlRelation = controlRelationManager.getControlRelation(mocName, mainControlAttr, neType, sheetName)
        
        mainControlGroupName = mainControlRelation.GroupName
        mainControlColumnName = mainControlRelation.columnName
        mainValueType = mainControlRelation.valueType
                
        mainControlColumnNumber = get_colNum(sheetName, mainControlGroupName, mainControlColumnName, noUse)
        '该主控参数未在页签中找到，说明Control Def有冗余信息，不做校验退出
        If mainControlColumnNumber = 0 Then
            'makeControlAttrValueCol = 1 '1表示在控制范围内
            'mainAttrExistFlag = False
            Call addNoneExistMainAttr(mainControlAttr, controlAttrValueManager)
            GoTo NextLoop
        End If
        
        Dim mainAttrCell As Range
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
        Call mainControlAttrValue.init(mainControlAttr, mainControlColumnName, mainValue, mainValueType, valueEmptyFlag, valueReferenceFlag, valueCellGrayFlag, mainAttrExistFlag)
        
        Call controlAttrValueManager.addNewControlAttrValue(mainControlAttrValue)
NextLoop:
    Next eachControlAttr
    makeControlAttrValueCol = 2 '2表格各主控参数都是有值的，需要进行分支控制判断
End Function

Private Sub getControlAttrInfo(ByRef controldef As CControlDef, ByRef mocName As String, ByRef attrName As String, ByRef neType As String, ByRef sheetName As String, ByRef controlInfo As String)
    mocName = controldef.mocName
    attrName = controldef.attributeName
    neType = controldef.neType
    sheetName = controldef.sheetName
    controlInfo = controldef.controlInfo
End Sub


Private Sub setRangeColor(ByVal sheetName As String, ByVal attrName As String, ByRef certainRange As Range)
    Dim mustgivedef As Worksheet
    Set mustgivedef = ThisWorkbook.Worksheets(MustGiveSheetName)
    Dim rowNum As Integer
    For rowNum = HiddenSheetValidRowBegin To mustgivedef.Range("a65536").End(xlUp).row
        If (sheetName = mustgivedef.Cells(rowNum, 1).value) And (attrName = mustgivedef.Cells(rowNum, 2).value) Then
            Call setRangeMustGive(certainRange)
            Exit Sub
        End If
    Next
    Call setRangeNormal(certainRange)
End Sub

Sub executeTemplateBranchControlAll()
    Dim sheet As Worksheet
    Dim mocName As String
    Dim attrName As String
    Dim value As String
    Dim index As Integer
    Set sheet = ThisWorkbook.Worksheets(getResByKey("PackageCustomTemplate"))
    For index = 2 To sheet.Range("a65536").End(xlUp).row
        mocName = sheet.Cells(index, 2).value
        attrName = sheet.Cells(index, 3).value
        value = sheet.Cells(index, 4).value
        If value = "YES" Then
            Call doTemplateBranchControl(mocName, attrName, value)
        End If
    Next
End Sub

Sub executeTemplateBranchControl(sh As Worksheet, target As Range)
    Dim col, row As Integer
    Dim mocName As String
    Dim attrName As String
    Dim value As String
    row = target.row
    col = target.Column
    If col = 4 Then
        mocName = sh.Cells(row, 2).value
        attrName = sh.Cells(row, 3).value
        value = target.value
        Call doTemplateBranchControl(mocName, attrName, value)
    End If
End Sub

Sub doTemplateBranchControl(mocName As String, attrName As String, value As String)
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim index As Integer
    For index = 2 To mappingDef.Range("a65536").End(xlUp).row
        If LCase(mappingDef.Cells(index, 4).value) = LCase(mocName) And LCase(mappingDef.Cells(index, 5).value) = LCase(attrName) Then
            Dim r As Range
            Set r = getRangeByMappingDef(mappingDef.Cells(index, 1).value, mappingDef.Cells(index, 5).value)
            If (r Is Nothing) = False Then
                If value = "YES" Then
                    r.Interior.ColorIndex = SolidColorIdx
                    r.Interior.Pattern = SolidPattern
                Else
                    r.Interior.Pattern = NullPattern
                    If mappingDef.Cells(index, 7) = "1" Or LCase(mappingDef.Cells(index, 7)) = "true" Then
                        r.Interior.ColorIndex = MustGiveColorIdx
                    Else
                        r.Interior.ColorIndex = NullPattern
                    End If
                End If
            End If
        End If
    Next
End Sub



Function getRangeByMappingDef(sheetName As String, attrName As String) As Range
        Dim sheet As Worksheet
        Dim attrStartRow, endRow As Integer
        attrStartRow = 4
        Set sheet = ThisWorkbook.Worksheets(sheetName)
        If (sheet Is Nothing) = False Then
                endRow = getEndRow(sheet)
                Dim tmpCol As Integer
                tmpCol = 0
                Do
                    tmpCol = tmpCol + 1
                Loop Until sheet.Cells(attrStartRow, tmpCol).value = attrName Or (sheet.Cells(attrStartRow, tmpCol).value = "" And sheet.Cells(attrStartRow - 1, tmpCol).value <> "")
                If sheet.Cells(attrStartRow, tmpCol).value <> "" Then
                    Set getRangeByMappingDef = sheet.Range(sheet.Cells(attrStartRow + 2, tmpCol), sheet.Cells(endRow, tmpCol))
                End If
        End If
End Function

Function getEndRow(sheet As Worksheet) As Integer
    Dim index As Integer
    Dim endRow As Integer
    Dim cell As Range
    endRow = 6
    For index = 6 To sheet.UsedRange.Rows.count
        Set cell = sheet.Cells(index, 1)
        If cell.value = "" Or Application.WorksheetFunction.CountBlank(cell) = 1 Then
            Exit For
        End If
        endRow = index
    Next index
    getEndRow = endRow
End Function







