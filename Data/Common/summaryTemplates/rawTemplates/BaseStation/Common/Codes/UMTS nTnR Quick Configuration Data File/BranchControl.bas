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

Sub getGroupAndColumnName(ByVal curSheet As Worksheet, ByVal cellRange As range, groupName As String, columnName As String)
    Dim m_rowNum, m_colNum As Long
    
    If Not isBoardStyleSheet(curSheet) Then
        columnName = curSheet.Cells(2, cellRange.column).value
        For m_colNum = cellRange.column To 1 Step -1
            If curSheet.Cells(1, m_colNum).value <> "" Then
                groupName = curSheet.Cells(1, m_colNum).value
                Exit For
            End If
        Next
    Else
        Dim pos As Long
        For pos = cellRange.row To 1 Step -1
            If pos = 1 Then Exit For
            If rowIsBlank(curSheet, pos - 1) = True And rowIsBlank(curSheet, pos) = False Then Exit For
        Next
        
        groupName = curSheet.Cells(pos, 1).value
        columnName = curSheet.Cells(pos + 1, cellRange.column).value
    End If
End Sub

Public Sub boardStyleSheetControl(ByVal sh As Object, ByVal target As range)
    On Error Resume Next
    Dim sheet As New Worksheet
    Dim cellRange As range
    Dim isConAttr As Boolean
    Dim rowNum As Long
    Dim contRel As controlRelation
    Set sheet = sh
    If target.count Mod 256 = 0 Then
        Exit Sub
    End If
    Dim currentNeType As String
    If sheet.name <> getResByKey("Board Style") Then actualBoardStyleName = sheet.name
    For Each cellRange In target
        If cellRange.Interior.colorIndex = 34 Or cellRange.Interior.colorIndex = 40 Or cellRange.Borders.LineStyle = xlLineStyleNone Then
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

Public Function isOperationExcel() As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.name, getResByKey("Board Style")) <> 0 And InStr(ws.Cells(1, 1), getResByKey("Operation")) <> 0 Then
            isOperationExcel = True
            Exit Function
        End If
    Next ws
    isOperationExcel = False
End Function

Function get_colNum(ByVal sheetName As String, ByVal groupName As String, ByVal columnName As String, rowNum As Long) As Long
    Dim m_colNum1, m_colNum2, m_rowNum As Long
    Dim ws As Worksheet
    Dim startColnumLetter As Long
    startColnumLetter = 1
    If isOperationExcel Then
        startColnumLetter = 2
    End If
    If sheetName = getResByKey("Comm Data") Or InStr(sheetName, getResByKey("Board Style")) <> 0 Then
        If containsASheet(ThisWorkbook, sheetName) Then
            Set ws = ThisWorkbook.Worksheets(sheetName)
        Else
            Set ws = ThisWorkbook.Worksheets(actualBoardStyleName)
        End If
        For m_rowNum = 1 To ws.range("a65536").End(xlUp).row
            If groupName = ws.Cells(m_rowNum, startColnumLetter).value Then
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

'Public Function isGsmCellSheet(sheetName As String) As Boolean
'    If sheetName = "GSM Cell" Or sheetName = getResByKey("A183") Then
'        isGsmCellSheet = True
'        Exit Function
'    End If
'    isGsmCellSheet = False
'End Function

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
    If (sheetName <> "MAPPING DEF") And (sheetName <> "SHEET DEF") And (sheetName <> "CONTROL DEF") _
        And (sheetName <> getResByKey("help")) And (sheetName <> getResByKey("Cover")) And (sheetName <> "SummaryRes") Then
        sheetSelectionShouldCheck = True
    Else
        sheetSelectionShouldCheck = False
    End If
End Function

Public Function getCertainControlDefine(ByRef curSheet As Worksheet, ByRef cellRange As range, ByRef groupName As String, ByRef sheetName As String, ByRef columnName As String) As CControlDef
    Call getGroupAndColumnName(curSheet, cellRange, groupName, columnName)
    If Not isBoardStyleSheet(curSheet) Then
        Set getCertainControlDefine = getControlDefine(curSheet.name, groupName, columnName)
    Else
    If curSheet.name <> getResByKey("Board Style") Then actualBoardStyleName = curSheet.name
        Set getCertainControlDefine = getControlDefine(getResByKey("Board Style"), groupName, columnName)
    End If
End Function

'单元格是否是黄底的单元格
Public Function cellIsNotHyperlinkColor(ByRef cellRange As range) As Boolean
    cellIsNotHyperlinkColor = True
End Function


