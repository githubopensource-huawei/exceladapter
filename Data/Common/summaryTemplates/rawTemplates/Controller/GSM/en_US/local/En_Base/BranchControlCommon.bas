Attribute VB_Name = "BranchControlCommon"
Option Explicit
Public Const CustomizationSceneMoc As String = "Customization_CME"  '��������MOC

Sub buildControlRelations()
    On Error GoTo ErrorHandler
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
    With controldef
        For conRowNum = 2 To controldef.range("a1048576").End(xlUp).row
            Dim ctrlInfoItemsArray As Variant
            ctrlInfoItemsArray = .range("A" & conRowNum & ":J" & conRowNum).value
            mocName = Trim(CStr(ctrlInfoItemsArray(1, 1)))
            contedName = Trim(CStr(ctrlInfoItemsArray(1, 2)))
            valueType = Trim(CStr(ctrlInfoItemsArray(1, 3)))
            conInfor = Trim(CStr(ctrlInfoItemsArray(1, 6)))
            sheetName = Trim(CStr(ctrlInfoItemsArray(1, 7)))
            groupName = Trim(CStr(ctrlInfoItemsArray(1, 8)))
            columnName = Trim(CStr(ctrlInfoItemsArray(1, 9)))
            neType = Trim(CStr(ctrlInfoItemsArray(1, 10)))
            If isControlInfoRef(conInfor) Then conInfor = getRealControlInfo(conInfor)
            
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
    End With
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in buildControlRelations, " & Err.Description
End Sub

'�����Ľ���MappingDefҳǩ��MOC��Attr�����ĺ������������Moc_Appear_In_SameSheet����
Public Sub buildMappingNumberRelations()
    Set mappingNumberManager = New CMappingNumberManager '��¼MappingDef��MocAttrNetype�����Ĺ�����
    
    Dim mappingDefSheet As Worksheet
    Set mappingDefSheet = ThisWorkbook.Worksheets("MAPPING DEF")
    
    Dim rowIndex As Long
    Dim mocName As String, attrName As String, neType As String, sheetName As String
    For rowIndex = 2 To mappingDefSheet.range("a1048576").End(xlUp).row
        sheetName = mappingDefSheet.range("A" & rowIndex).value
        mocName = mappingDefSheet.range("D" & rowIndex).value
        attrName = mappingDefSheet.range("E" & rowIndex).value
        neType = mappingDefSheet.range("L" & rowIndex).value
        
        Call mappingNumberManager.addMocAttrNetype(sheetName, mocName, attrName, neType)
    Next rowIndex
End Sub

Public Sub CheckCustomizedScene(ByRef ws As Worksheet, ByRef cellRange As range)
On Error GoTo ErrorHandler

Dim controldef As CControlDef
Dim groupName As String, columnName As String, sheetName As String, transBook As Worksheet, errorMsg As String
If cellRange.row < 3 Then
    Exit Sub
End If

'Set controldef = getCertainControlDefine(ws, cellRange, groupName, sheetName, columnName)
 Call getGroupAndColumnName(ws, cellRange, groupName, columnName)
 
 Dim mocName As String, attrName As String, neType As String, controlInfo As String, lstvalue As String, index As Long, index2 As Long, nResponse As Long
 Set controldef = getControlDefine(ws.name, groupName, columnName)
 
 Call getControlAttrInfo(controldef, mocName, attrName, neType, sheetName, controlInfo)
 If mocName <> CustomizationSceneMoc Then
    Exit Sub
 End If
 
 Set controldef = getControlDefine(getResByKey("BaseTransPort"), groupName, columnName)

 Call getControlAttrInfo(controldef, mocName, attrName, neType, sheetName, controlInfo)
 lstvalue = controldef.lstvalue
 errorMsg = getResByKey("Range") + "[" + lstvalue + "]"
 Dim valideValues As Collection
 Set valideValues = SplitEx(lstvalue, ",")
 Dim currvalues() As String
 currvalues = Split(cellRange.value, ",")
 For index = LBound(currvalues) To UBound(currvalues)
     If InCollection(valideValues, currvalues(index)) = False Then
        nResponse = MsgBox(errorMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
        If nResponse = vbRetry Then
            cellRange.Select
        End If
        cellRange.value = ""
     End If
 Next

Exit Sub
ErrorHandler:
    Exit Sub
End Sub


'1��ʾ�ڷ�Χ�ڣ�0��ʾ���ڷ�Χ�ڣ�-1��ʾδ����ӳ���ϵ(�����������ٽ��д���)
Function Check_Value_Validation(ByVal CurSheet As Worksheet, ByVal cellRange As range, ByRef isConAttr As Boolean, ByRef contRel As controlRelation, ByRef currentNeType As String) As Long
    On Error Resume Next
    
    Dim groupName As String
    Dim columnName As String
    Dim mocName As String, attrName As String, attrType As String, attrRange As String
    Dim conRowNum, i As Long
    Dim neType As String
    Dim sheetName As String
    Dim controldef As CControlDef

    Set controldef = getCertainControlDefine(CurSheet, cellRange, groupName, sheetName, columnName)
    
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
            'If Moc_Appear_In_SameSheet(ControlRelMap(i).mocName, attrName) = False Then
            If mappingNumberManager.hasOneMocAttrNetypeRecord(sheetName, ControlRelMap(i).mocName, attrName, neType) Then
                isConAttr = True
                contRel = ControlRelMap(i)
             End If
            Exit For
        End If
    Next
    If Trim(cellRange.value) <> "" And UBound(Split(cellRange.value, "\")) <> 2 Then 'ֵ�ǿգ����ж�����ֵ�Ƿ��ڷ�Χ��
        attrType = controldef.dataType
        attrRange = controldef.bound + controldef.lstvalue
        If Trim(controldef.controlInfo) <> "" And cellIsNotHyperlinkColor(cellRange) Then
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
    
    If attrType = "Enum" And InStr(attrValue, ",") = 0 Then
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
        If LenB(StrConv(attrValue, vbFromUnicode)) < min Or LenB(StrConv(attrValue, vbFromUnicode)) > max Then
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
            errorMsg = getResByKey("Range") + formatRange(attrRange)
        End If
        'errorMsg = getResByKey("Range") + formatRange(attrRange)
    End If
    
    If Check_Value_In_Range = False Then
        nResponse = MsgBox(errorMsg, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
        If nResponse = vbRetry Then
            cellRange.Select
        End If
        cellRange.value = ""
    ElseIf Left(cellRange.value, 1) <> "#" Then
        '��д�����ڷ�Χ�ڣ�Ҫ����ֵ�൥Ԫ��������ȥ��0�Ĳ���
        Call removePrefixZeros(cellRange, attrValue, attrType)
    End If
    alreadyCheckFlag = True
    Exit Function
ErrorHandler:
    Check_Value_In_Range = False '���쳣��˵��У�������Ҫ�������б��ز�����У��
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
    Else
        '��д�����ڷ�Χ�ڣ�Ҫ����ֵ�൥Ԫ��������ȥ��0�Ĳ���
        Call removePrefixZeros(cellRange, cellRange.value, controldef.dataType)
    End If
    
    controlledAttrValidationCheck = branchAttrValueValidationFlag '�������ΧУ��Ľ��
    Exit Function
ErrorHandler:
    controlledAttrValidationCheck = True '������쳣�ˣ�����Ϊ�ڷ�Χ�ڣ�����ǰ�����̱���һ��
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
        
        If Not mappingNumberManager.hasOneMocAttrNetypeRecord(sheetName, mocName, mainControlAttr, neType) Then
            makeControlAttrValueCol = 1 '�ڱ�����ҵ�ͬ�����ز����Ķ�����¼������̫���˿ڵ�����˿����ԣ���ʱû�취���п��ƣ��򲻿��ƣ�ֱ���˳�
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
    makeControlAttrValueCol = 2 '2�������ز���������ֵ�ģ���Ҫ���з�֧�����ж�
End Function

Public Sub currentParameterBranchCheck(ByRef ws As Worksheet, ByRef cellRange As range)
    On Error GoTo ErrorHandler
    If cellRange.count > 1 Then Exit Sub 'ѡ��ĵ�Ԫ�����1�����˳�
    
    If selectionIsValid(ws, cellRange) = False Then Exit Sub '���ѡ��ĵ�Ԫ��Ƿ���ֱ���˳�
    
    '����Ѿ��һ��ˣ�����Ҫ�����ˣ��˳�
    If cellIsGray(cellRange) Then Exit Sub
    
    Dim controldef As CControlDef
    Dim groupName As String, columnName As String, sheetName As String
    Set controldef = getCertainControlDefine(ws, cellRange, groupName, sheetName, columnName)
    
    If cellRange.value = groupName Or cellRange.value = columnName Then Exit Sub '���ѡ��ĵ�Ԫ����������������ֱ���˳�
    
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
    
    '�������ֵ��1��Ҫô���ز�����ȫ��Ҫô���ز���ֵΪ�ջ�Ϊ���ã��������ز�����¼Ϊ����������Ҫ��֧У�飬ֱ���˳�
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

Private Sub makeControlAttrValueManager(ByRef sheet As Worksheet, ByRef dstRowNumber As Long, ByRef mainControlMocName As String, ByRef mainControlAttrName As String, _
    ByRef neType As String, ByRef virtualSheetName As String, ByRef oneMainControlAttrNotExist As Boolean, ByRef controlAttrValueManager As CControlAttrValueManager)
    
    '����Ѿ����˸����ز�����������������ˣ�ֱ���˳�
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
    
    Dim matchBranchAttrEmptyFlag As Boolean '�����ز����Ƿ�Ϊ�ջ���Ϊ���õı�־
    oneBranchMatch = newCheckBranchMatch(controlAttrValueManager, controlInfo, matchBranchNode, matchBranchAttrEmptyFlag)
    If oneBranchMatch = False Then 'δ�ҵ�ƥ���֧
        contedOutOfControl = False
    Else '�˷�֧�и����ز���ƥ��ɹ�������з�֧����
        checkAllBranchMatch = True
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
        '������ز���֮һΪ�ջ���Ϊ���ã��򲻽��п����ˣ�ֱ���˳�
        If matchBranchAttrEmptyFlag = True Then
            Exit Function
        End If
        '���ڷ�Χ��ʱҪ���
        If valIsRight = False And Trim(sheet.Cells(cellRange.row, contedColNum).value) <> "" _
            And contedType <> "IPV4" And contedType <> "IPV6" And contedType <> "Bitmap" Then
            If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
                sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
            End If
            sheet.Cells(cellRange.row, contedColNum).value = ""
        End If
        '���ñ��ز����ķ�Χ
        Call setValidation(contedType, boundValue, sheet, cellRange.row, contedColNum)
    End If
End Function

Private Function newCheckBranchMatch(ByRef controlAttrValueManager As CControlAttrValueManager, ByRef controlInfo As String, ByRef matchBranchNode As Variant, ByRef matchBranchAttrEmptyFlag As Boolean) As Boolean
    Dim branchMatchChecker As New CBranchMatchChecker
    Call branchMatchChecker.init(controlAttrValueManager, controlInfo)
    newCheckBranchMatch = branchMatchChecker.getOneBranchMatchFlag
    Set matchBranchNode = branchMatchChecker.getMatchBranchNode
    matchBranchAttrEmptyFlag = branchMatchChecker.getMatchBranchAttrEmptyFlag
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

Private Function selectionIsValid(ByRef ws As Worksheet, ByRef cellRange As range) As Boolean
    If findGroupName(Trim(cellRange.value)) = True Or findAttrName(Trim(cellRange.value)) = True Or cellRange.Borders.LineStyle = xlLineStyleNone Then
        selectionIsValid = False
    Else
        selectionIsValid = True
    End If
End Function

Private Sub getControlAttrInfo(ByRef controldef As CControlDef, ByRef mocName As String, ByRef attrName As String, ByRef neType As String, ByRef sheetName As String, ByRef controlInfo As String)
    mocName = controldef.mocName
    attrName = controldef.attributeName
    neType = controldef.neType
    sheetName = controldef.sheetName
    controlInfo = controldef.controlInfo
End Sub

Private Sub setRangeGray(ByRef certainRange As range)
    On Error Resume Next
    certainRange.Interior.colorIndex = SolidColorIdx
    certainRange.Interior.Pattern = SolidPattern
    certainRange.value = ""
    certainRange.Validation.ShowInput = False
End Sub

Private Sub removePrefixZeros(ByRef cellRange As range, ByRef cellValue As String, ByRef valueType As String)
    If isNum(valueType) Then
        Dim newValue As String
        newValue = CStr(CDbl(cellValue)) '�µ���ֵ
        '����µ���ֵ�͵�ǰֵ��һ������ȥ��0������ֵ��д����Ԫ����
        If cellValue <> newValue Then cellRange.value = CStr(CDbl(cellValue))
    End If
End Sub

Private Sub setControlledParameterTipAndValidation(ByRef ws As Worksheet, ByRef cellRange As range, ByRef valueType As String, ByRef branchMatchChecker As CBranchMatchChecker)
    If targetHasInputMessage(cellRange) Then Exit Sub  '�������InputMessage��Tip�����˳�
    
    If branchMatchChecker.getMatchBranchAttrEmptyFlag = True Then Exit Sub '���ĳ�����ز���Ϊ�ջ�Ϊ���ã�����Ҫ����ע���������˳�
    
    Dim matchBranchNode As Variant, boundNodes As Variant
    Set matchBranchNode = branchMatchChecker.getMatchBranchNode
    Set boundNodes = matchBranchNode.childNodes
    
    Dim boundValue As String
    Dim valIsRight As Boolean 'ûʲô�ã�ֻ��Ϊ�˵���֮ǰ�ĺ���getContedAttrBoundValue
    '��ñ��ز����ķ�Χ
    boundValue = getContedAttrBoundValue(boundNodes, valIsRight, ws, cellRange, cellRange.column)
    Call setValidation(valueType, boundValue, ws, cellRange.row, cellRange.column)
End Sub

Private Function targetHasInputMessage(ByRef target As range) As Boolean
    On Error GoTo ErrorHandler
    targetHasInputMessage = True
    If target.Validation Is Nothing Then 'û����Ч�ԣ���û��InputMessage
        targetHasInputMessage = False
        Exit Function
    End If
    
    Dim inputMessage As String
    inputMessage = target.Validation.inputMessage '�����InputMessage����ֵ�ɹ������û�У���ֵ��������ErrorHandler
    If inputMessage = "" Then targetHasInputMessage = False '���InputMessageΪ�գ�����Ϊû��Tip
    Exit Function
ErrorHandler:
    targetHasInputMessage = False
End Function

Public Function cellIsGray(ByRef certainCell As range) As Boolean
    If certainCell.Interior.colorIndex = SolidColorIdx And certainCell.Interior.Pattern = SolidPattern Then
        cellIsGray = True
    Else
        cellIsGray = False
    End If
End Function


Public Function isControlInfoRef(conInfor As String) As Boolean
    On Error GoTo ErrorHandler
    isControlInfoRef = False
    If Left(conInfor, 7) = "Control" Then isControlInfoRef = True
    Exit Function
ErrorHandler:
    isControlInfoRef = False
    Debug.Print "some exception in isControlInfoRef, " & Err.Description
End Function

Public Function getRealControlInfo(conInfor As String) As String
    On Error GoTo ErrorHandler
    Dim controlInforSht As Worksheet
    Set controlInforSht = ThisWorkbook.Worksheets("Control Infor")
    
    Dim infos As Variant
    infos = Split(conInfor, ",")
    
    Dim targetRow As Integer
    targetRow = CInt(infos(1))
    
    Dim maxCol As Integer
    
    With controlInforSht
        maxCol = .range("XFD" & targetRow).End(xlToLeft).column
        Dim cell As range
        For Each cell In .range("A" & targetRow & ":" & getColStr(maxCol) & targetRow)
            getRealControlInfo = getRealControlInfo & cell.value
        Next
    End With
    Exit Function
ErrorHandler:
    getRealControlInfo = conInfor
    Debug.Print "some exception in getRealControlInfo, " & Err.Description
End Function
