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

'用以设置颜色
Public Const SolidColorIdx = 16
Public Const SolidPattern = xlGray16
Public Const NullPattern = xlNone
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

Sub getGroupAndColumnName(ByVal CurSheet As Worksheet, ByVal cellRange As range, groupName As String, columnName As String)
    Dim m_rowNum, m_colNum As Long
    
    If CurSheet.name = getResByKey("Comm Data") Then
        For m_rowNum = cellRange.row To 1 Step -1
            If findAttrName(Trim(CurSheet.Cells(m_rowNum, cellRange.column).value)) = True Then '34是绿色
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
    
    Dim mocName As String, attrName As String, sheetName As String, groupName As String, columnName As String, neType As String
    Dim branchInfor As String, contedType As String
    Dim boundValue As String

    Dim allBranchMatch, contedOutOfControl As Boolean
    Dim xmlObject As Object
    Dim m, conRowNum, contedColNum As Long
    Dim noUse As Long
    Dim rootNode As Variant
    Dim controlDef As Worksheet
    Set controlDef = ThisWorkbook.Worksheets("CONTROL DEF")
    
    Dim ctrlInfoItemsArray As Variant
    '对各个被控参数进行分支控制
    For m = 0 To contRel.contedNum
        For conRowNum = 2 To controlDef.range("a65536").End(xlUp).row
            ctrlInfoItemsArray = controlDef.range("A" & conRowNum & ":J" & conRowNum).value
            mocName = ctrlInfoItemsArray(1, 1)
            attrName = ctrlInfoItemsArray(1, 2)
            sheetName = ctrlInfoItemsArray(1, 7)
            neType = ctrlInfoItemsArray(1, 10)
            If contRel.mocName = mocName And contRel.neType = neType And contRel.contedAttrs(m) = attrName And contRel.sheetName = sheetName Then
                groupName = ctrlInfoItemsArray(1, 8)
                columnName = ctrlInfoItemsArray(1, 9)
                contedType = ctrlInfoItemsArray(1, 3)
                contedColNum = get_colNum(sheetName, groupName, columnName, noUse)
                With sheet.Cells(cellRange.row, contedColNum)
                    If (Trim(cellRange.value) = "" And cellRange.Interior.colorIndex <> SolidColorIdx And cellRange.Interior.Pattern <> SolidPattern) Or UBound(Split(cellRange.value, "\")) = 2 Then '主控为空或是引用，则此时被控应变为非灰即有效，范围恢复成初始值
                        If .Interior.colorIndex = SolidColorIdx And .Interior.Pattern = SolidPattern Then
                            If .Hyperlinks.count = 1 Then
                                .Hyperlinks.Delete
                            End If
                            .Interior.colorIndex = NullPattern
                            .Interior.Pattern = NullPattern
                            .Validation.ShowInput = True
                        End If
                        '恢复成初始范围
                        boundValue = ctrlInfoItemsArray(1, 4) + ctrlInfoItemsArray(1, 5)
                        Call setValidation(contedType, boundValue, sheet, cellRange.row, contedColNum)
                        Call setEmptyValidation(sheet, cellRange.row, contedColNum)
                    Else '主控不空，进行contRel.contedAttrs(m)的分支控制
                        branchInfor = ctrlInfoItemsArray(1, 6)
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
                                .Interior.colorIndex = SolidColorIdx
                                .Interior.Pattern = SolidPattern
                                .value = ""
                                .Validation.ShowInput = False
                            Else
                                .Interior.colorIndex = NullPattern
                                .Interior.Pattern = NullPattern
                                .Validation.ShowInput = True
                            End If
                            If .Hyperlinks.count = 1 Then
                                .Hyperlinks.Delete
                            End If
                        End If
                    End If
                End With
                Exit For
            End If
        Next
    Next
End Sub

Sub deleteValidation(ByRef sheet As Worksheet, ByRef rowNumber As Long, ByRef columnNumber As Long)
    sheet.Cells(rowNumber, columnNumber).Validation.Delete
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
    For index = 2 To validSheet.range("a65536").End(xlUp).row
        If GetDesStr(validSheet.Cells(index, 1).value) = GetDesStr(currmocName) And GetDesStr(validSheet.Cells(index, 2).value) = GetDesStr(currattrname) Then
            Dim maxColNum As Long
            maxColNum = validSheet.range("IV" + CStr(index)).End(xlToLeft).column
            getLongEnumValueFromValidDef = "=INDIRECT(""'VALID DEF'!C" & CStr(index) & ":" & getColStr(maxColNum) & CStr(index) & """)"
        End If
    Next
    Exit Function
ErrorHandler:
    getLongEnumValueFromValidDef = ""
End Function

Sub setValidation(ByVal contedType As String, ByVal boundValue As String, sheet As Worksheet, ByVal rowNum As Long, ByVal colnum As Long)
    On Error Resume Next
    
    Dim inputTitle As String
    inputTitle = getResByKey("Range")
    
    '非枚举，无Validation则加上，若有则看是否要改
    If contedType <> "Enum" And contedType <> "Bitmap" And contedType <> "IPV4" And contedType <> "IPV6" _
        And contedType <> "Time" And contedType <> "Date" And contedType <> "DateTime" Then
        If boundValue <> sheet.Cells(rowNum, colnum).Validation.inputmessage Then
            If contedType = "String" Or contedType = "Password" Then
                inputTitle = getResByKey("Length")
                boundValue = formatRange(boundValue)
            End If
            
            If isNum(contedType) Then
                 boundValue = formatRange(boundValue)
            End If
            
            With sheet.Cells(rowNum, colnum).Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
                .inputTitle = inputTitle
                .inputmessage = boundValue
                .ShowInput = True
                .ShowError = False
            End With
        End If
    '枚举
    ElseIf contedType = "Enum" Then
        If boundValue <> sheet.Cells(rowNum, colnum).Validation.formula1 Or sheet.Cells(rowNum, colnum).Validation.inputmessage = "" Then
            If InStr(boundValue, "/") <> 0 Then
                Exit Sub
            End If
            
            If Len(boundValue) > 255 Then
                Dim groupName As String
                Dim columnName As String
                Dim cellRange As range
                Set cellRange = sheet.range(C(colnum) + CStr(rowNum))
                Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
                boundValue = getLongEnumValueFromValidDef(sheet.name, groupName, columnName)
                With sheet.Cells(rowNum, colnum).Validation
                    .Delete
                    .Add Type:=xlValidateList, formula1:=boundValue
                End With
            Else
                With sheet.Cells(rowNum, colnum).Validation
                    .Delete
                    .Add Type:=xlValidateList, formula1:=boundValue
                    .inputTitle = getResByKey("Range")
                    .inputmessage = "[" + boundValue + "]"
                    .ShowInput = True
                    .ShowError = True
                End With
            End If
            sheet.Cells(rowNum, colnum).Validation.Modify Type:=xlValidateList, formula1:=boundValue
        End If
    End If
    
End Sub
Function get_colNum(ByVal sheetName As String, ByVal groupName As String, ByVal columnName As String, rowNum As Long) As Long
    Dim m_colNum1, m_colNum2, m_rowNum As Long
    Dim ws As Worksheet
    If sheetName = getResByKey("Comm Data") Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For m_rowNum = 1 To ws.range("a65536").End(xlUp).row
            If groupName = ws.Cells(m_rowNum, 1).value Then
                For m_colNum1 = 1 To ws.range("IV" + CStr(m_rowNum + 1)).End(xlToLeft).column
                    If columnName = ws.Cells(m_rowNum + 1, m_colNum1).value Then
                        get_colNum = m_colNum1
                        rowNum = m_rowNum + 1
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    
    ElseIf isIubStyleWorkSheet(ThisWorkbook.ActiveSheet.name) Then
        Set ws = ThisWorkbook.ActiveSheet
        For m_rowNum = 1 To ws.range("a65536").End(xlUp).row
                If ws.Cells(m_rowNum, 1).value = sheetName Then
                    For m_colNum1 = 1 To ws.range("IV" + CStr(m_rowNum)).End(xlToLeft).column
                        If columnName = ws.Cells(m_rowNum, m_colNum1).value Then
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
        For m_colNum1 = 1 To ws.range("IV2").End(xlToLeft).column
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
    For conRowNum = 2 To controlDef.range("a65536").End(xlUp).row
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
Public Function Ensure_Gray_Cell(curRange As range) As Boolean
    If curRange.value <> "" And curRange.Interior.colorIndex = SolidColorIdx And curRange.Interior.Pattern = SolidPattern Then
        MsgBox getResByKey("NoInput"), vbOKOnly + vbExclamation + vbApplicationModal + vbDefaultButton1, getResByKey("Warning")
        curRange.value = ""
        curRange.Select
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

Private Sub setEmptyValidation(ByRef sheet As Worksheet, ByRef rowNumber As Long, ByRef columnNumber As Long)
    On Error Resume Next
    With sheet.Cells(rowNumber, columnNumber).Validation
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
        .inputTitle = ""
        .inputmessage = ""
        .ShowInput = True
        .ShowError = False
    End With
End Sub

Public Function sheetSelectionShouldCheck(ByRef ws As Worksheet) As Boolean
    Dim sheetName As String
    sheetName = ws.name
    If (isGsmCellSheet(sheetName) = False) And (sheetName <> "MappingSiteTemplate") And (sheetName <> "ProductType") _
        And (sheetName <> "MappingCellTemplate") And (sheetName <> "MappingRadioTemplate") _
        And (sheetName <> "MAPPING DEF") And (sheetName <> "SHEET DEF") And (sheetName <> "CONTROL DEF") _
        And (sheetName <> "Help for USB Parameter") And (sheetName <> getResByKey("help")) _
        And (sheetName <> "IPRouteMap") And (sheetName <> getResByKey("Cover")) And (sheetName <> "COMMON") _
        And (sheetName <> "Qos") And (sheetName <> "USB Parameter for Sites") And (sheetName <> "SummaryRes") Then
        sheetSelectionShouldCheck = True
    Else
        sheetSelectionShouldCheck = False
    End If
End Function

Public Function getCertainControlDefine(ByRef ws As Worksheet, ByRef cellRange As range, ByRef groupName As String, ByRef sheetName As String, ByRef columnName As String) As CControlDef
    If Not isIubStyleWorkSheet(ws.name) Then
        Call getGroupAndColumnName(ws, cellRange, groupName, columnName)
        Set getCertainControlDefine = getControlDefine(ws.name, groupName, columnName)
    Else
        Call getGroupNameShNameAndAttrName(ws, cellRange, groupName, sheetName, columnName)
        Set getCertainControlDefine = getControlDefine(sheetName, groupName, columnName)
    End If
End Function

'单元格是否是黄底的单元格
Public Function cellIsNotHyperlinkColor(ByRef cellRange As range) As Boolean
    If cellRange.Interior.colorIndex <> HyperLinkColorIndex Then
        cellIsNotHyperlinkColor = True
    Else
        cellIsNotHyperlinkColor = False
    End If
End Function
