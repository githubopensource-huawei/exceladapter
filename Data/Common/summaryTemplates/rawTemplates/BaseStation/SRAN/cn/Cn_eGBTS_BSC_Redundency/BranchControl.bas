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

Sub getGroupAndColumnName(ByVal CurSheet As Worksheet, ByVal cellRange As range, GroupName As String, columnName As String)
    Dim m_rowNum, m_colNum As Long
    
    If CurSheet.name = getResByKey("Comm Data") Then
        For m_rowNum = cellRange.row To 1 Step -1
            If CurSheet.Cells(m_rowNum, cellRange.column).Interior.colorIndex = 40 Then '34是绿色
                columnName = CurSheet.Cells(m_rowNum, cellRange.column).value
                Exit For
            End If
        Next
        For m_colNum = cellRange.column To 1 Step -1
            If CurSheet.Cells(m_rowNum - 1, m_colNum).value <> "" Then
                GroupName = CurSheet.Cells(m_rowNum - 1, m_colNum).value
                Exit For
            End If
        Next
    Else
        columnName = CurSheet.Cells(2, cellRange.column).value
        For m_colNum = cellRange.column To 1 Step -1
            If CurSheet.Cells(1, m_colNum).value <> "" Then
                GroupName = CurSheet.Cells(1, m_colNum).value
                Exit For
            End If
        Next
    End If
End Sub

Sub Execute_Branch_Control(ByVal sheet As Worksheet, ByVal cellRange As range, contRel As controlRelation, ByRef currentNeType As String)
    On Error Resume Next
    
    Dim sheetName, GroupName, columnName As String
    Dim branchInfor As String, contedType As String
    Dim boundValue As String

    Dim allBranchMatch, contedOutOfControl As Boolean
    Dim xmlObject As Object
    Dim m, conRowNum, contedColNum As Long
    Dim noUse As Long
    Dim rootNode As Variant
    Dim controldef As Worksheet
    Set controldef = ThisWorkbook.Worksheets("CONTROL DEF")
    '对各个被控参数进行分支控制
    For m = 0 To contRel.contedNum
        For conRowNum = 2 To controldef.range("a65536").End(xlUp).row
            If (contRel.mocName = controldef.Cells(conRowNum, 1).value) _
                            And contRel.neType = controldef.Cells(conRowNum, 10).value _
                            And (contRel.contedAttrs(m) = controldef.Cells(conRowNum, 2).value) _
                            And (contRel.sheetName = controldef.Cells(conRowNum, 7).value) Then
                sheetName = controldef.Cells(conRowNum, 7).value
                GroupName = controldef.Cells(conRowNum, 8).value
                columnName = controldef.Cells(conRowNum, 9).value
                contedType = controldef.Cells(conRowNum, 3).value
                contedColNum = get_colNum(sheetName, GroupName, columnName, noUse)
                If (Trim(cellRange.value) = "" And cellRange.Interior.colorIndex <> SolidColorIdx And cellRange.Interior.Pattern <> SolidPattern) Or UBound(Split(cellRange.value, "\")) = 2 Then '主控为空或是引用，则此时被控应变为非灰即有效，范围恢复成初始值
                    If sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = SolidColorIdx And sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = SolidPattern Then
                        If sheet.Cells(cellRange.row, contedColNum).Hyperlinks.count = 1 Then
                            sheet.Cells(cellRange.row, contedColNum).Hyperlinks.Delete
                        End If
                        sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = NullPattern
                        sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = NullPattern
                        sheet.Cells(cellRange.row, contedColNum).Validation.ShowInput = True
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
                    allBranchMatch = checkAllBranchMatch(rootNode, sheet, cellRange, contRel, contedType, contedOutOfControl, contedColNum, currentNeType, branchInfor)
                    '所有主控参数的值都不在分支参数规定的范围内，则被控参数被灰化（但被控参数不受控制除外）
                    If allBranchMatch = False Then
                        If contedOutOfControl = False Then
                            sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = SolidColorIdx
                            sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = SolidPattern
                            sheet.Cells(cellRange.row, contedColNum).value = ""
                            sheet.Cells(cellRange.row, contedColNum).Validation.ShowInput = False
                        Else
                            sheet.Cells(cellRange.row, contedColNum).Interior.colorIndex = NullPattern
                            sheet.Cells(cellRange.row, contedColNum).Interior.Pattern = NullPattern
                            sheet.Cells(cellRange.row, contedColNum).Validation.ShowInput = True
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
            If InStr(boundValue, "/") <> 0 Then
                Exit Sub
            End If
            
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
Function get_colNum(ByVal sheetName As String, ByVal GroupName As String, ByVal columnName As String, rowNum As Long) As Long
    Dim m_colNum1, m_colNum2, m_rowNum As Long
    Dim ws As Worksheet
    If sheetName = getResByKey("Comm Data") Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For m_rowNum = 1 To ws.range("a65536").End(xlUp).row
            If GroupName = ws.Cells(m_rowNum, 1).value Then
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
                If GroupName = Trim(ws.Cells(1, m_colNum2).value) Then
                    get_colNum = m_colNum1
                    Exit For
                End If
            End If
        Next
    End If
End Function

Function get_MocAndAttrcolNum(ByVal mocName As String, ByVal attrName As String, ByVal sheetName As String) As Long
    Dim conRowNum, noUse As Long
    Dim GroupName, columnName As String
    Dim controldef As Worksheet
    Set controldef = ThisWorkbook.Worksheets("CONTROL DEF")
    get_MocAndAttrcolNum = 0
    For conRowNum = 2 To controldef.range("a65536").End(xlUp).row
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
        .inputMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
End Sub

Public Function sheetSelectionShouldCheck(ByRef ws As Worksheet) As Boolean
    Dim sheetName As String
    sheetName = ws.name
    If (sheetName <> "MappingSiteTemplate") And (sheetName <> "ProductType") _
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

Public Function getCertainControlDefine(ByRef ws As Worksheet, ByRef cellRange As range, ByRef GroupName As String, ByRef sheetName As String, ByRef columnName As String) As CControlDef
    Call getGroupAndColumnName(ws, cellRange, GroupName, columnName)
    Set getCertainControlDefine = getControlDefine(ws.name, GroupName, columnName)
'    If Not isIubStyleWorkSheet(ws.name) Then
'
'    Else
'        Call getGroupNameShNameAndAttrName(ws, cellRange, GroupName, sheetName, columnName)
'        Set getCertainControlDefine = getControlDefine(sheetName, GroupName, columnName)
'    End If
End Function

'单元格是否是黄底的单元格
Public Function cellIsNotHyperlinkColor(ByRef cellRange As range) As Boolean
    cellIsNotHyperlinkColor = True
'    If cellRange.Interior.colorIndex <> HyperLinkColorIndex Then
'        cellIsNotHyperlinkColor = True
'    Else
'        cellIsNotHyperlinkColor = False
'    End If
End Function
