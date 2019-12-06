Attribute VB_Name = "BranchControl"
Option Explicit
Public Type controlRelation
    mocName As String
    contAttr As String
    contedAttrs(100) As String
    contedNum As Long '��0��ʼ
    neType  As String '�˲�����Ϊ������ �ǿ������Ļ�������վ�ģ����ڿ�����������վ�п���Moc�Ͳ�����������һ����
    sheetName As String ' �˲�����Ϊ�����֣�ͬһMoc�����ڲ�ͬ��Sheet��Comm Dataҳ����֣���ʱ�򣬲��ܿ�sheetҳǩ����
End Type
Public ControlRelMap(1000) As controlRelation
Public contAttrValArray(100) As String
Dim ControlRelationNum As Long

Private controlRelationManager As CControlRelationManager

Private actualBoardStyleName As String

'����������ɫ
Public Const SolidColorIdx = 16
Public Const SolidPattern = xlGray16
Public Const NullPattern = xlNone
Public Const NormalPattern = 1

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
    Dim isExist, flag, isFound As Boolean
    Dim index1, index2, index3, index4 As Long
    Dim neType As String
    Dim sheetName As String
    
    Dim groupName As String, columnName As String, valueType As String
    
    Set controlRelationManager = New CControlRelationManager '��ǰ������ΧУ�������
    
    If isControlDefSheetExist = False Then
        Exit Sub
    End If
    
    ControlRelationNum = 0
    Dim controldef As Worksheet
    Set controldef = ThisWorkbook.Worksheets("CONTROL DEF")
    For conRowNum = 2 To controldef.range("a1048576").End(xlUp).row

        mocName = Trim(controldef.Cells(conRowNum, 1).value)
        contedName = Trim(controldef.Cells(conRowNum, 2).value)
        conInfor = Trim(controldef.Cells(conRowNum, 6).value)
        neType = Trim(controldef.Cells(conRowNum, 10).value)
        sheetName = Trim(controldef.Cells(conRowNum, 7).value)
    
        groupName = Trim(controldef.range("H" & conRowNum).value)
        columnName = Trim(controldef.range("I" & conRowNum).value)
        valueType = Trim(controldef.range("C" & conRowNum).value)
        Call controlRelationManager.addNewAttributeRelation(mocName, contedName, conInfor, neType, sheetName, groupName, columnName, valueType)
        
        If conInfor <> "" Then
            While InStr(1, conInfor, "attribute", vbBinaryCompare) <> 0
                conInfor = Mid(conInfor, InStr(1, conInfor, "attribute", vbBinaryCompare) + 11)
                contName = Mid(conInfor, 1, InStr(1, conInfor, """") - 1)
                '�������ص�ӳ���ϵ
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
                                '��֤���ڴ˱��ز����ı��ز������棬����A->B->C, Ҫ��֤��A�Ŀ����б���B��C��ǰ��
                                isFound = False
                                For index1 = 0 To ControlRelationNum - 1
                                    If (mocName = ControlRelMap(index1).mocName) And (contedName = ControlRelMap(index1).contAttr) And (neType = ControlRelMap(index1).neType) And (sheetName = ControlRelMap(index1).sheetName) Then
                                        For index2 = 0 To ControlRelMap(index1).contedNum '���ز���B�Ŀ����б�
                                            For index3 = 0 To ControlRelMap(i).contedNum    '���ز���A�Ŀ����б�
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
                'û�����ص�ӳ���ϵ�����½�
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
'1��ʾ�ڷ�Χ�ڣ�0��ʾ���ڷ�Χ�ڣ�-1��ʾδ����ӳ���ϵ(�����������ٽ��д���)
Function Check_Value_Validation(ByVal CurSheet As Worksheet, ByVal cellRange As range, ByRef isConAttr As Boolean, ByRef contRel As controlRelation, ByRef currentNeType As String) As Long
    On Error Resume Next
    
    Dim groupName As String
    Dim columnName As String
    Dim mocName, attrName, attrType, attrRange As String
    Dim conRowNum, i As Long
    Dim neType As String
    Dim sheetName As String
    Dim controldef As CControlDef
    
    Call getGroupAndColumnName(CurSheet, cellRange, groupName, columnName)
'    If Not isBoardStyleSheet(CurSheet) Then
'        Call getGroupAndColumnName(CurSheet, cellRange, groupName, columnName)
'        Set controldef = getControlDefine(CurSheet.name, groupName, columnName)
'    Else
'        Call getGroupNameShNameAndAttrName(CurSheet, cellRange, groupName, sheetName, columnName)
'        Set controldef = getControlDefine(sheetName, groupName, columnName)
'    End If
    If Not isBoardStyleSheet(CurSheet) Then
        Set controldef = getControlDefine(CurSheet.name, groupName, columnName)
    Else
        Set controldef = getControlDefine(getResByKey("Board Style"), groupName, columnName)
    End If
    
    If controldef Is Nothing Then
        Check_Value_Validation = -1
        Exit Function
    End If
    
    mocName = controldef.mocName
    attrName = controldef.attributeName
    neType = controldef.neType
    sheetName = controldef.sheetName
    
    currentNeType = neType
    '�ж��Ƿ������ز���
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
    If Trim(cellRange.value) <> "" And UBound(Split(cellRange.value, "\")) <> 2 Then 'ֵ�ǿգ����ж�����ֵ�Ƿ��ڷ�Χ��
        attrType = controldef.dataType
        attrRange = controldef.bound + controldef.lstValue
        If Trim(controldef.controlInfo) <> "" Then
            '�з�֧����������Ϣ�ķ�Χ���ܷ����仯,Tip����ʾ�ľ������·�Χ
            If attrType = "Enum" And cellRange.Validation.inputMessage <> "" Then
                attrRange = Mid(cellRange.Validation.inputMessage, 2, Len(cellRange.Validation.inputMessage) - 2)
            ElseIf cellRange.Validation.inputMessage <> "" Then
                attrRange = cellRange.Validation.inputMessage
            End If
        End If
        
        Dim alreadyCheckFlag As Boolean
        alreadyCheckFlag = False '����ֵ�Ƿ��Ѿ�У��ı�־
        
        If Check_Value_In_Range(attrType, attrRange, cellRange.value, cellRange, alreadyCheckFlag) = True Then
            Check_Value_Validation = 1
        Else
            If alreadyCheckFlag = False Then '˵����Check_Value_In_Range���쳣�ˣ���Ҫ���µı��ز���У�鷽������У��
                If controlledAttrValidationCheck(CurSheet, controldef, cellRange) = True Then 'У��Ĳ����ڷ�Χ��
                    Check_Value_Validation = 1
                Else 'У��Ĳ������ڷ�Χ�ڣ�����0����ҪΪ�˲����к����ı��ز�����֧��
                    Check_Value_Validation = 0
                End If
            Else '���������Ĳ��ڷ�Χ��
                Check_Value_Validation = 0
            End If
        End If
        Exit Function
    Else  'ֵΪ�գ����ڻ�Ҫ���з�֧���ƣ������ж�Ϊ�ڷ�Χ��
        Check_Value_Validation = 1
        Exit Function
    End If
    
    Check_Value_Validation = -1
    
End Function

'�����ı��ز���У�飬��Tip��ʾ�Ļ���������У��
Private Function controlledAttrValidationCheck(ByRef ws As Worksheet, ByRef controldef As CControlDef, ByRef cellRange As range) As Boolean
    On Error GoTo ErrorHandler
    controlledAttrValidationCheck = False
    
    Dim mocName As String, attrName As String, neType As String, controlInfo As String, sheetName As String
    Call getControlAttrInfo(controldef, mocName, attrName, neType, sheetName, controlInfo)
    
     '����ò������Ǳ��ز��������˳�
    If Not controlRelationManager.containsControlledAttributeRelation(mocName, attrName, neType, sheetName) Then Exit Function
    
    '�ҵ��ò������������ز�����
    Dim controlRelation As CControlRelation
    Set controlRelation = controlRelationManager.getControlRelation(mocName, attrName, neType, sheetName)
    
    Dim controlAttrValueManager As New CControlAttrValueManager '�����������ز����Ĺ����࣬������������ز����������뱻�����ķ�֧������Ϣ��У��
    Dim mainControlAttrReturnedValue As Long
    mainControlAttrReturnedValue = makeControlAttrValueCol(ws, mocName, attrName, neType, sheetName, controlRelation, controlAttrValueManager, cellRange)
    
    If mainControlAttrReturnedValue = 1 Then Exit Function '����Ҫ�Ե�ǰ���ز������п��ƣ�ֱ���˳�
    
    Dim oneBranchMatchFlag As Boolean
    Dim branchMatchChecker As New CBranchMatchChecker '��֧ƥ��У����
    Call branchMatchChecker.init(controlAttrValueManager, controlInfo)
    oneBranchMatchFlag = branchMatchChecker.getOneBranchMatchFlag
    
    If oneBranchMatchFlag = False Then Exit Function '���û�ҵ�ƥ���֧�����˳�
    
    Dim matchBranchNode As Variant
    Set matchBranchNode = branchMatchChecker.getMatchBranchNode
    
    Dim controlAttrValue As New CControlAttrValue '����һ���������������
    Call controlAttrValue.init("", "", cellRange.value, controldef.dataType, False, False, False)
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
    
    controlledAttrValidationCheck = branchAttrValueValidationFlag '�������ΧУ��Ľ��
    Exit Function
ErrorHandler:
    controlledAttrValidationCheck = True '������쳣�ˣ�����Ϊ�ڷ�Χ�ڣ�����ǰ�����̱���һ��
End Function

Sub getGroupAndColumnName(ByVal CurSheet As Worksheet, ByVal cellRange As range, groupName As String, columnName As String)
    Dim m_rowNum, m_colNum As Long
    
    If CurSheet.name = getResByKey("Comm Data") Or InStr(CurSheet.name, getResByKey("Board Style")) <> 0 Then
        For m_rowNum = cellRange.row To 1 Step -1
            If findAttrName(Trim(CurSheet.Cells(m_rowNum, cellRange.column).value)) = True Then
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

Function Check_Value_In_Range(ByVal attrType As String, ByVal attrRange As String, ByVal attrValue As String, cellRange As range, ByRef alreadyCheckFlag As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim arrayList() As String
    Dim errorMsg, sItem As String
    Dim i, nResponse, nLoop As Long
    Dim min, max As Double
    
    '��������ΧΪ�գ�����ҪУ�� DTS2012111306136
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
    Else  '��ֵ
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
    Check_Value_In_Range = False '���쳣��˵��У���������Ҫ�������б��ز�����У��
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
    Dim controldef As Worksheet
    Dim controlledRange As range
    Set controldef = ThisWorkbook.Worksheets("CONTROL DEF")
    '�Ը������ز������з�֧����
'    For m = 0 To contRel.contedNum - 1
    For m = 0 To contRel.contedNum
        For conRowNum = 2 To controldef.range("a1048576").End(xlUp).row
            If (contRel.mocName = controldef.Cells(conRowNum, 1).value) _
                            And contRel.neType = controldef.Cells(conRowNum, 10).value _
                            And (contRel.contedAttrs(m) = controldef.Cells(conRowNum, 2).value) _
                            And (contRel.sheetName = controldef.Cells(conRowNum, 7).value) Then
                sheetName = controldef.Cells(conRowNum, 7).value
                'sheetName = sheet.name
                groupName = controldef.Cells(conRowNum, 8).value
                columnName = controldef.Cells(conRowNum, 9).value
                contedType = controldef.Cells(conRowNum, 3).value
                contedColNum = get_colNum(sheetName, groupName, columnName, noUse)
                Set controlledRange = sheet.Cells(cellRange.row, contedColNum)
                If (Trim(cellRange.value) = "" And cellRange.Interior.colorIndex <> SolidColorIdx And cellRange.Interior.Pattern <> SolidPattern) Or UBound(Split(cellRange.value, "\")) = 2 Then '����Ϊ�ջ������ã����ʱ����Ӧ��Ϊ�ǻҼ���Ч����Χ�ָ��ɳ�ʼֵ
                    If controlledRange.Interior.colorIndex = SolidColorIdx And controlledRange.Interior.Pattern = SolidPattern Then
                        If controlledRange.Hyperlinks.count = 1 Then
                            controlledRange.Hyperlinks.Delete
                        End If
                        '������ڽ������ӵ�������������޸ĵ��������ӵ��У���ô��ʽ���ʱӦ������Ϊ���׻��̵�
                        If Not setControlledRangeColorAndPattern(controlledRange) Then
                            controlledRange.Interior.colorIndex = NullPattern
                            controlledRange.Interior.Pattern = NullPattern
                        End If
                        controlledRange.Validation.ShowInput = True
                    End If
                    '�ָ��ɳ�ʼ��Χ
                    boundValue = controldef.Cells(conRowNum, 4).value + controldef.Cells(conRowNum, 5).value
                    Call setValidation(contedType, boundValue, sheet, cellRange.row, contedColNum)
                    Call setEmptyValidation(sheet, cellRange.row, contedColNum)
                Else '���ز��գ�����contRel.contedAttrs(m)�ķ�֧����
                    branchInfor = controldef.Cells(conRowNum, 6).value
                    Set xmlObject = CreateObject("msxml2.domdocument")
                    xmlObject.LoadXML branchInfor
                    'Set BranchNodes = xmlObject.DocumentElement.ChildNodes
                    Set rootNode = xmlObject.DocumentElement
                    contedOutOfControl = False
                    allBranchMatch = checkAllBranchMatch(rootNode, sheet, cellRange, contRel, contedType, contedOutOfControl, contedColNum, currentNeType, branchInfor)
                    '�������ز�����ֵ�����ڷ�֧�����涨�ķ�Χ�ڣ��򱻿ز������һ��������ز������ܿ��Ƴ��⣩
                    If allBranchMatch = False Then
                        If contedOutOfControl = False Then
                            controlledRange.Interior.colorIndex = SolidColorIdx
                            controlledRange.Interior.Pattern = SolidPattern
                            controlledRange.value = ""
                            controlledRange.Validation.ShowInput = False
                        Else
                            '������ڽ������ӵ�������������޸ĵ��������ӵ��У���ô��ʽ���ʱӦ������Ϊ���׻��̵�
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
    '���ж��Ƿ��������б�������ף����ж��Ƿ�����������ͨ����
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
    
    '����Ѿ����˸����ز������������������ˣ�ֱ���˳�
    If controlAttrValueManager.hasControlAttr(mainControlAttrName) Then Exit Sub
    
    Dim dstColumnNumber As Long
    dstColumnNumber = get_MocAndAttrcolNum(mainControlMocName, mainControlAttrName, virtualSheetName) '�õ����ز������к�
    If dstColumnNumber = 0 Then '����0��˵�����ز��������ڣ����˳�������flag��ΪTrue
        oneMainControlAttrNotExist = True
        Exit Sub
    End If

    '����ڵ�ǰҳǩû���ҵ������ز����Ŀ�����Ϣ��˵�����ز���ȱ�٣���ñ��ز���������
    If Not controlRelationManager.containsAttributeRelation(mainControlMocName, mainControlAttrName, neType, virtualSheetName) Then
        oneMainControlAttrNotExist = True
        Exit Sub
    End If
    
    Dim mainControlRelation As CControlRelation '�������ز�����Control Def������
    Dim mainControlAttrValue As CControlAttrValue '����һ�����ز���ֵ�����ͣ��������Ƶ���
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
    
    If UBound(Split(mainValue, "\")) = 2 Then  '�������һ�����ز���Ϊ���ã���־��ΪTrue���ɶ����Լ�������check
        valueReferenceFlag = True '����Ϊ���ñ�־
    ElseIf mainValue = "" Then '��ʱ���ز���Ϊ�գ��п��ܸ����ز���Ҳ����Ч�ģ�����Ҫ�ȶԸ����ص�Ԫ�����У��
        Call currentParameterBranchCheck(sheet, mainAttrCell)
'        If cellIsGray(cellRange) Then
'            makeControlAttrValueCol = 1 '�����������ز����Ļһ��ı��Ѿ�ʹ��ǰ��Ԫ��һ��ˣ�����Ҫ�ж��ˣ�ֱ���˳�
'            Exit Sub
'        End If
    End If
    
    If cellIsGray(mainAttrCell) Then '���ĳ�����ز����һ�����־��ΪTrue���ɶ����Լ�������check
        valueCellGrayFlag = True '����Ϊ�һ���־
        valueEmptyFlag = True '����Ϊ�һ�����϶�Ϊ��
    ElseIf mainValue = "" Then '������ص�Ԫ��δ�һ���˵�������ص�Ԫ������Ч��֧��ֻ��δ��дֵ���ձ�־��ΪTrue���ɶ����Լ�������check
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
                controlAttrCol.Add Item:=controlAttributeName, key:=controlAttributeName '�������ز�������������
            End If
        End If
    Next controlAttributeNode
    Set getMainControlAttrCol = controlAttrCol
End Function

Function checkAllBranchMatch(rootNode As Variant, sheet As Worksheet, cellRange As range, contRel As controlRelation, contedType As String, contedOutOfControl As Boolean, contedColNum As Long, ByRef currentNeType As String, ByRef controlInfo As String) As Boolean
    On Error Resume Next
    
    Dim matchBranchNode As Variant 'ƥ��ķ�֧�ڵ�
    
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
    Set mainControlAttrCol = getMainControlAttrCol(rootNode) '�õ����ز�������
    
    Dim eachMainControlAttr As Variant
    For Each eachMainControlAttr In mainControlAttrCol
        Call makeControlAttrValueManager(sheet, cellRowNumber, contRel.mocName, CStr(eachMainControlAttr), currentNeType, contRel.sheetName, oneMainControlAttrNotExist, controlAttrValueManager)
    Next eachMainControlAttr

    If oneMainControlAttrNotExist = True Then '���ز���ȱ�٣����϶��ڷ�Χ�ڣ�ֱ���˳�
        contedOutOfControl = True
        Exit Function
    End If
    
    oneBranchMatch = newCheckBranchMatch(controlAttrValueManager, controlInfo, matchBranchNode)
    If oneBranchMatch = False Then 'δ�ҵ�ƥ���֧
        contedOutOfControl = False
    Else '�˷�֧�и����ز���ƥ��ɹ�������з�֧����
        Set boundNodes = matchBranchNode.childNodes
        '��ñ��ز����ķ�Χ
        boundValue = getContedAttrBoundValue(boundNodes, valIsRight, sheet, cellRange, contedColNum)
        '���з�֧����
        If sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = SolidColorIdx And sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = SolidPattern Then
            If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
                sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
            End If
            sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = NullPattern
            sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = NullPattern
            sheet.Cells(cellRange.row, contedColNum).Validation.ShowInput = True
        End If
        '���ڷ�Χ��ʱҪ���
        If valIsRight = False And Trim(sheet.Cells(cellRange.row, contedColNum).value) <> "" _
            And contedType <> "IPV4" And contedType <> "IPV6" Then
            If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
                sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
            End If
            sheet.Cells(cellRange.row, contedColNum).value = ""
        End If
        '���ñ��ز����ķ�Χ
        Call setValidation(contedType, boundValue, sheet, cellRange.row, contedColNum)
        checkAllBranchMatch = True
    End If
    
'    '�����������Ʒ�֧����������һ������������ķ�֧���ƣ�������ȫ�����㣨�һ���
'    For I = 0 To BranchNodes.Length - 1
'        oneBranchMatch = True
'        oneContGray = False
'        oneContNotExist = False
'        oneContNull = False
'        contAttrNum = 0
'        Set equalsNodes = BranchNodes(I).GetElementsByTagName("Equals")
'        '���˷�֧�еĸ������ز���ֵ�Ƿ��ڷ�Χ��
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
'                    ElseIf UBound(Split(sheet.Cells(cellRange.row, colNum), "\")) = 2 Then ' ���������ã�����Ӧ���ܿ��ƣ��ȼ�������Ϊ��
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
'                '���浽�����У�������checkBranchMatch�жϸ�������ֵ�Ƿ�ƥ��
'                contAttrValArray(contAttrNum) = eachContAttrVal
'                contAttrNum = contAttrNum + 1
'            ElseIf oneContGray = False Then
'                oneContNull = True
'                oneBranchMatch = False
'            End If
'        Next
'
'        '�����ڻһ�����ֵ�����������ڵ�������������жϸ�������ֵ�Ƿ�ƥ��
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
'        Else '�˷�֧�и����ز���ƥ��ɹ�������з�֧����
'            Set boundNodes = BranchNodes(I).ChildNodes
'            '��ñ��ز����ķ�Χ
'            boundValue = getContedAttrBoundValue(boundNodes, valIsRight, sheet, cellRange, contedColNum)
'            '���з�֧����
'            If sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = SolidColorIdx And sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = SolidPattern Then
'                If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
'                    sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
'                End If
'                sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = NullPattern
'                sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = NullPattern
'                sheet.Cells(cellRange.row, contedColNum).Validation.ShowInput = True
'            End If
'            '���ڷ�Χ��ʱҪ���
'            If valIsRight = False And Trim(sheet.Cells(cellRange.row, contedColNum).value) <> "" _
'                And contedType <> "IPV4" And contedType <> "IPV6" Then
'                If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
'                    sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
'                End If
'                sheet.Cells(cellRange.row, contedColNum).value = ""
'            End If
'            '���ñ��ز����ķ�Χ
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
        '����ö�ٺ���������ʱ
    If isFound = False Then
        valIsRight = True
    End If
End Function
Function checkBranchMatch(equalsNodes) As Boolean
    Dim i As Long
    Dim j, k As Long
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
Sub setValidation(ByVal contedType As String, ByVal boundValue As String, sheet As Worksheet, ByVal rowNum As Long, ByVal colNum As Long)
    On Error Resume Next
    
    Dim inputTitle As String
    inputTitle = getResByKey("Range")
    
    '��ö�٣���Validation����ϣ��������Ƿ�Ҫ��
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
    'ö��
    ElseIf contedType = "Enum" Then
        If boundValue <> sheet.Cells(rowNum, colNum).Validation.Formula1 Or sheet.Cells(rowNum, colNum).Validation.inputMessage = "" Then
            With sheet.Cells(rowNum, colNum).Validation
                .Delete
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
Function get_colNum(ByVal sheetName As String, ByVal groupName As String, ByVal columnName As String, rowNum As Long) As Long
    Dim m_colNum1, m_colNum2, m_rowNum As Long
    Dim ws As Worksheet
    If sheetName = getResByKey("Comm Data") Or InStr(sheetName, getResByKey("Board Style")) <> 0 Then
        If containsASheet(ThisWorkbook, sheetName) Then
            Set ws = ThisWorkbook.Worksheets(sheetName)
        Else
            Set ws = ThisWorkbook.Worksheets(actualBoardStyleName)
        End If
        For m_rowNum = 1 To ws.range("a1048576").End(xlUp).row
            If groupName = ws.Cells(m_rowNum, 1).value Then
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
    Dim controldef As Worksheet
    Set controldef = ThisWorkbook.Worksheets("CONTROL DEF")
    get_MocAndAttrcolNum = 0
    For conRowNum = 2 To controldef.range("a1048576").End(xlUp).row
        If (mocName = controldef.Cells(conRowNum, 1).value) _
            And (attrName = controldef.Cells(conRowNum, 2).value) _
            And (sheetName = controldef.Cells(conRowNum, 7).value) Then
            groupName = controldef.Cells(conRowNum, 8).value
            columnName = controldef.Cells(conRowNum, 9).value
            get_MocAndAttrcolNum = get_colNum(sheetName, groupName, columnName, noUse)
        Exit For
        End If
    Next
End Function
'��ȷ������������Ƿ�����ɢ��Χ��[1��2][3��4]
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
'�жϲ���Targetָ������ĵ�Ԫ���Ƿ�Ϊ��ɫ������״̬,������ոõ�Ԫ�������ֵ
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
    Dim currentNeType As String
    If sheet.name <> getResByKey("Board Style") Then actualBoardStyleName = sheet.name
    For Each cellRange In Target
        If findAttrName(Trim(cellRange.value)) = True Or findGroupName(Trim(cellRange.value)) = True Or cellRange.Borders.LineStyle = xlLineStyleNone Then
            Exit Sub
        End If
        If Ensure_Gray_Cell(cellRange) = False Then
            'isConAttr�����Ƿ������ز���
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
        And (sheetName <> getResByKey("Temp Sheet")) And (sheetName <> "GTRX") Then
        sheetSelectionShouldCheck = True
    Else
        sheetSelectionShouldCheck = False
    End If
End Function

Private Function selectionIsValid(ByRef ws As Worksheet, ByRef cellRange As range) As Boolean
    '�����Ԫ����û��ֵ��ֵ���ҵ�Ԫ���б߿�����Ϊ��Ԫ��Ϸ�������True
    If cellRange.value = "" And cellRange.Borders.LineStyle <> xlLineStyleNone Then
        selectionIsValid = True
    'У���û�ѡ���Ƿ�Ϊ������У������������Ϊ��Ԫ�����
    ElseIf findAttrName(Trim(cellRange.value)) = True Or findGroupName(Trim(cellRange.value)) = True Or cellRange.Borders.LineStyle = xlLineStyleNone Then
        selectionIsValid = False
    Else
        selectionIsValid = True
    End If
End Function

Public Sub currentParameterBranchCheck(ByRef ws As Worksheet, ByRef cellRange As range)
    On Error GoTo ErrorHandler
    If cellRange.count > 1 Then Exit Sub 'ѡ��ĵ�Ԫ�����1�����˳�
    
    If selectionIsValid(ws, cellRange) = False Then Exit Sub '���ѡ��ĵ�Ԫ��Ƿ���ֱ���˳�
    
    '����Ѿ��һ��ˣ�����Ҫ�����ˣ��˳�
    If cellIsGray(cellRange) Then Exit Sub
    
    Dim controldef As CControlDef
    Dim groupName As String, columnName As String, sheetName As String
    
    Call getGroupAndColumnName(ws, cellRange, groupName, columnName)
    
    If Not isBoardStyleSheet(ws) Then
        Set controldef = getControlDefine(ws.name, groupName, columnName)
    Else
        If ws.name <> getResByKey("Board Style") Then actualBoardStyleName = ws.name
        Set controldef = getControlDefine(getResByKey("Board Style"), groupName, columnName)
    End If
    
    'δ�ҵ�Control Def������Ϣ���˳�
    If controldef Is Nothing Then Exit Sub
    
    Dim mocName As String, attrName As String, neType As String, controlInfo As String
    Call getControlAttrInfo(controldef, mocName, attrName, neType, sheetName, controlInfo)
    
     '����ò������Ǳ��ز��������˳�
    If Not controlRelationManager.containsControlledAttributeRelation(mocName, attrName, neType, sheetName) Then Exit Sub
    
    '�ҵ��ò������������ز�����
    Dim controlRelation As CControlRelation
    Set controlRelation = controlRelationManager.getControlRelation(mocName, attrName, neType, sheetName)
    
    Dim controlAttrValueManager As New CControlAttrValueManager '�����������ز����Ĺ����࣬������������ز����������뱻�����ķ�֧������Ϣ��У��
    Dim mainControlAttrReturnedValue As Long
    mainControlAttrReturnedValue = makeControlAttrValueCol(ws, mocName, attrName, neType, sheetName, controlRelation, controlAttrValueManager, cellRange)
    
    '�������ֵ��1��Ҫô���ز�����ȫ��Ҫô���ز���ֵΪ�ջ�Ϊ���ã�����Ҫ��֧У�飬ֱ���˳�
    If mainControlAttrReturnedValue = 1 Then
        Exit Sub
    ElseIf mainControlAttrReturnedValue = 3 Then '�������ֵ��3��˵������һ�����ز����һ�����ǰ��֧��Ч���һ��˳�
        If cellIsGray(cellRange) Then Exit Sub '��ʱ�п����Ѿ������ز����Ļһ����¸ñ��ز����һ��ˣ���ô�Ͳ���Ҫ�ٴλһ���ǰ��Ԫ���ˣ�ֱ���˳������Ч��
        Call setRangeGray(cellRange)
        Exit Sub
    End If
    
    Dim oneBranchMatchFlag As Boolean
    Dim branchMatchChecker As New CBranchMatchChecker
    Call branchMatchChecker.init(controlAttrValueManager, controlInfo)
    oneBranchMatchFlag = branchMatchChecker.getOneBranchMatchFlag
    
    If oneBranchMatchFlag = False Then '˵��δ�ҵ�ƥ���֧���򽫵�ǰ��Ԫ�һ�
        Call setRangeGray(cellRange)
    Else '�������Ҫ�һ�����ô��Ҫ���ӱ��ط�֧��Tip����������Ч��
        Call setControlledParameterTipAndValidation(ws, cellRange, controldef.dataType, branchMatchChecker)
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
    If targetHasInputMessage(cellRange) Then Exit Sub  '�������InputMessage��Tip�����˳�
    
    If branchMatchChecker.getMatchBranchAttrEmptyFlag = True Then Exit Sub '���ĳ�����ز���Ϊ�գ�����Ҫ����ע���������˳�
    
    Dim matchBranchNode As Variant, boundNodes As Variant
    Set matchBranchNode = branchMatchChecker.getMatchBranchNode
    Set boundNodes = matchBranchNode.childNodes
    
    Dim boundValue As String
    Dim valIsRight As Boolean 'ûʲô�ã�ֻ��Ϊ�˵���֮ǰ�ĺ���getContedAttrBoundValue
    '��ñ��ز����ķ�Χ
    boundValue = getContedAttrBoundValue(boundNodes, valIsRight, ws, cellRange, cellRange.column)
    Call setValidation(valueType, boundValue, ws, cellRange.row, cellRange.column)
End Sub

Private Function targetHasInputMessage(ByRef Target As range) As Boolean
    On Error GoTo ErrorHandler
    targetHasInputMessage = True
    If Target.Validation Is Nothing Then 'û����Ч�ԣ���û��InputMessage
        targetHasInputMessage = False
        Exit Function
    End If
    
    Dim inputMessage As String
    inputMessage = Target.Validation.inputMessage '�����InputMessage����ֵ�ɹ������û�У���ֵ����������ErrorHandler
    If inputMessage = "" Then targetHasInputMessage = False '���InputMessageΪ�գ�����Ϊû��Tip
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
    
    Dim mainControlRelation As CControlRelation '�������ز�����Control Def������
    Dim mainControlAttrValue As CControlAttrValue '����һ�����ز���ֵ�����ͣ��������Ƶ���
    
    For Each eachControlAttr In controlAttrCol
        mainControlAttr = CStr(eachControlAttr)
        '����ڵ�ǰҳǩû���ҵ������ز����Ŀ�����Ϣ��˵�����ز���ȱ�٣���ñ��ز���������
        If Not controlRelationManager.containsAttributeRelation(mocName, mainControlAttr, neType, sheetName) Then
            makeControlAttrValueCol = 1 '1��ʾ�ڿ��Ʒ�Χ��
            Exit Function
        End If
        
        Set mainControlRelation = controlRelationManager.getControlRelation(mocName, mainControlAttr, neType, sheetName)
        
        mainControlGroupName = mainControlRelation.groupName
        mainControlColumnName = mainControlRelation.columnName
        mainValueType = mainControlRelation.valueType
                
        mainControlColumnNumber = get_colNum(sheetName, mainControlGroupName, mainControlColumnName, noUse)
        '�����ز���δ��ҳǩ���ҵ���˵��Control Def��������Ϣ������У���˳�
        If mainControlColumnNumber = 0 Then
            makeControlAttrValueCol = 1 '1��ʾ�ڿ��Ʒ�Χ��
            Exit Function
        End If
        
        Dim mainAttrCell As range
        Set mainAttrCell = ws.Cells(cellRow, mainControlColumnNumber)
        
        mainValue = mainAttrCell.value
        
        Dim valueEmptyFlag As Boolean, valueReferenceFlag As Boolean, valueCellGrayFlag As Boolean
        valueEmptyFlag = False
        valueReferenceFlag = False
        valueCellGrayFlag = False
        
        If UBound(Split(mainValue, "\")) = 2 Then  '�������һ�����ز���Ϊ���ã���־��ΪTrue���ɶ����Լ�������check
            valueReferenceFlag = True '����Ϊ���ñ�־
        ElseIf mainValue = "" Then '��ʱ���ز���Ϊ�գ��п��ܸ����ز���Ҳ����Ч�ģ�����Ҫ�ȶԸ����ص�Ԫ�����У��
            Call currentParameterBranchCheck(ws, mainAttrCell)
            If cellIsGray(cellRange) Then
                makeControlAttrValueCol = 1 '�����������ز����Ļһ��ı��Ѿ�ʹ��ǰ��Ԫ��һ��ˣ�����Ҫ�ж��ˣ�ֱ���˳�
                Exit Function
            End If
        End If
        
        If cellIsGray(mainAttrCell) Then '���ĳ�����ز����һ�����־��ΪTrue���ɶ����Լ�������check
            valueCellGrayFlag = True '����Ϊ�һ���־
            valueEmptyFlag = True '����Ϊ�һ�����϶�Ϊ��
        ElseIf mainValue = "" Then '������ص�Ԫ��δ�һ���˵�������ص�Ԫ������Ч��֧��ֻ��δ��дֵ���ձ�־��ΪTrue���ɶ����Լ�������check
            valueEmptyFlag = True
        End If
        
        Set mainControlAttrValue = New CControlAttrValue
        Call mainControlAttrValue.init(mainControlAttr, mainControlColumnName, mainValue, mainValueType, valueEmptyFlag, valueReferenceFlag, valueCellGrayFlag)
        
        Call controlAttrValueManager.addNewControlAttrValue(mainControlAttrValue)
    Next eachControlAttr
    makeControlAttrValueCol = 2 '2��������ز���������ֵ�ģ���Ҫ���з�֧�����ж�
End Function

Private Sub getControlAttrInfo(ByRef controldef As CControlDef, ByRef mocName As String, ByRef attrName As String, ByRef neType As String, ByRef sheetName As String, ByRef controlInfo As String)
    mocName = controldef.mocName
    attrName = controldef.attributeName
    neType = controldef.neType
    sheetName = controldef.sheetName
    controlInfo = controldef.controlInfo
End Sub

Public Sub setCellSheetFunctionNameBoxLst(ByVal sh As Object, ByVal Target As range)
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim tempcellRange As range
    Dim groupName As String
    Dim columnName As String
    Dim rowNum As Long
    Dim colNum As Long
    Dim m_colNum As Long
    Set sheet = sh
    If Target.count Mod 256 = 0 Then
        Exit Sub
    End If
    Dim neType As String
    neType = getNeType()
    
    Dim tempbaseStationData As CBaseStationData
    If tempbaseStationData Is Nothing Then Set tempbaseStationData = New CBaseStationData
    Call tempbaseStationData.init
    
    Dim functionNeMap As CMap
    Set functionNeMap = tempbaseStationData.functionNeMap
    Dim functionNameStr As String

    
    For Each cellRange In Target
        If findAttrName(Trim(cellRange.value)) = True Or findGroupName(Trim(cellRange.value)) = True Or cellRange.Borders.LineStyle = xlLineStyleNone Then
            Exit Sub
        End If
        
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        rowNum = cellRange.row
        colNum = cellRange.column
        
        If isFunctionNameColum(groupName, columnName) Then
            If neType = "GSM" Then
                If (sheet.name = "GSM Cell" Or sheet.name = getResByKey("PageGSMCell") Or sheet.name = "GTRX" Or sheet.name = getResByKey("GTRX_ZH")) And functionNeMap.hasKey("GBTSFUNCTION") Then
                    functionNameStr = functionNeMap.GetAt("GBTSFUNCTION")
                    Call setBoardStyleListBoxRangeValidation(sheet.name, groupName, columnName, functionNameStr, sheet, cellRange)
                End If
            Else
                If functionNeMap.hasKey("BASESTATION") Then
                    functionNameStr = functionNeMap.GetAt("BASESTATION")
                    Call setBoardStyleListBoxRangeValidation(sheet.name, groupName, columnName, functionNameStr, sheet, cellRange)
                End If
            End If
            If sheet.Cells(rowNum, colNum).value <> "" And (sheet.name = "GSM Cell" Or sheet.name = getResByKey("PageGSMCell")) Then
                If sheet.Cells(rowNum, colNum - 1).value <> sheet.Cells(rowNum, colNum).value Then
                    Call refreshcellRec(sheet, rowNum, True)
                ElseIf sheet.Cells(rowNum, colNum - 1).value = sheet.Cells(rowNum, colNum).value And cellRange.Interior.ThemeColor = xlThemeColorAccent4 Then
                    Call refreshcellRec(sheet, rowNum, False)
                End If
            End If
        End If
    Next
End Sub


Private Function isFunctionNameColum(ByRef groupName As String, ByRef columnName As String) As Boolean
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    Dim sheetName As String
    
    isFunctionNameColum = False
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        sheetName = sheetDef.Cells(index, 1)
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If sheetName = "GSM Cell" Or sheetName = getResByKey("PageGSMCell") Or sheetName = "GTRX" Or sheetName = getResByKey("GTRX_ZH") Then
            If groupName = mappingdefgroupName And columnName = mappingdefcolumnName And mocName = "GCELL" And attributeName = "BTSNAME" Then
                isFunctionNameColum = True
                Exit For
            End If
            If groupName = mappingdefgroupName And columnName = mappingdefcolumnName And mocName = "GTRX" And attributeName = "BTSNAME" Then
                isFunctionNameColum = True
                Exit For
            End If
        End If
    Next
End Function

Private Sub refreshcellRec(ByRef sheet As Worksheet, ByRef rowNum As Long, clearflag As Boolean)
    Dim tempcellRange As range
    Dim m_colNum As Long
    For m_colNum = 1 To sheet.range("XFD2").End(xlToLeft).column
        Set tempcellRange = sheet.range(getColStr(m_colNum) & rowNum)
        If tempcellRange.Interior.colorIndex <> SolidColorIdx And tempcellRange.Interior.Pattern <> SolidPattern Then
            If clearflag Then
            With tempcellRange.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            Else
                tempcellRange.Interior.colorIndex = NormalRangeColorIndex
            End If

        End If
    Next
End Sub
