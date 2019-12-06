Attribute VB_Name = "AutoDeployment"
'即插即用数据特殊处理
'用以设置颜色
Const SolidColorIdx = 16
Const SolidPattern = xlGray16
Const NullPattern = xlNone
Const ESN_ValidLen = 20
Const String_ValidLen = 64
'新增一个可选起始搜索行或列参数
Public Function findCertainValColumnNumber(ByRef ws As Worksheet, ByVal rowNumber As Long, ByRef cellVal As Variant, Optional ByVal startColumn As Long = 1) As Long
    Dim currentCellVal As Variant
    Dim maxColumnNumber As Long, k As Long
    maxColumnNumber = ws.UsedRange.columns.count
    findCertainValColumnNumber = -1
    For k = startColumn To maxColumnNumber
        currentCellVal = ws.Cells(rowNumber, k).value
        If currentCellVal = cellVal Then
            findCertainValColumnNumber = k
            Exit For
        End If
    Next
End Function
Public Sub AutoDeploySheetChange(ByVal sheet As Object, ByVal target As Range)
    Dim connTypeCol As Long
    Dim authenticationTypeCol As Long
    Dim subAreaCol As Long
    Dim esnCol As Long
    Dim nameCol As Long
    
    connTypeCol = findCertainValColumnNumber(sheet, 2, getResByKey("connType"))
    authenticationTypeCol = findCertainValColumnNumber(sheet, 2, getResByKey("authenticationType"))
    
    esnCol = 2
    subAreaCol = 4
    nameCol = 1
    
    If (target.Interior.colorIndex = SolidColorIdx) And target.value <> "" Then
        target.value = ""
        MsgBox getResByKey("NoInput")
        Exit Sub
    End If
    
    Dim controldef As Worksheet
    Dim subAreaValue As String
    Dim esnValue As String
    Dim nameValue As String
    
    Dim regEx
    Set regEx = CreateObject("VBSCRIPT.REGEXP")
    
    Dim specialChar As String
    Dim specialArray() As String
    Dim subAreaArray() As String
    
    specialChar = "~,!,@,#,$,%,^,&,*,{,},[,],+,-,<,>,?"
    specialArray = Split(specialChar, ",")
    'Unknownstring = Array("~", "!", "@", "#", "$", "%", "^", "&", "*", "{", "}", "[", "]", "+", "-", "<", ">", "?")
    'regEx.Pattern = "~|!|@|#|$|%|^|&|\*|\{|\}|\[|\]|\+|-|<|>|\?"

    Dim columnRange As Range, cellRange As Range
    
    For Each columnRange In target.columns
        If columnRange.column = nameCol Then
            For Each cellRange In columnRange
                nameValue = Trim(cellRange.value)
                If Len(nameValue) <> 0 Then
                    If ((String_ValidLen < LenB(StrConv(nameValue, vbFromUnicode)))) Or (1 > LenB(StrConv(nameValue, vbFromUnicode))) Then
                        nResponse = MsgBox(getResByKey("Limited Length") & "[1~64]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        'MsgBox (getResByKey("Length") & ESN_ValidLen)   '字符串待定义
                        cellRange.value = ""
                        Exit Sub
                    End If
                End If
            Next cellRange
             
         ElseIf columnRange.column = esnCol Then
            For Each cellRange In columnRange
                esnValue = Trim(cellRange.value)
                If Len(esnValue) <> 0 Then
                    If (ESN_ValidLen <> Len(esnValue)) Then
                        nResponse = MsgBox(getResByKey("Limited Length") & "0/20", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        'MsgBox (getResByKey("Length") & ESN_ValidLen)   '字符串待定义
                        cellRange.value = ""
                        Exit Sub
                    End If
                End If
            Next cellRange
        
        End If
    Next columnRange
    
    For Each columnRange In target.columns
         For subrowNum = 2 To connTypeCol
            If columnRange.column = subrowNum Then
                For Each cellRange In columnRange
                    subAreaValue = Trim(cellRange.value)
                    'If regEx.test(subAreaValue) Then
                        For I = 0 To UBound(specialArray)
                            If (InStr(subAreaValue, specialArray(I)) > 0) Then
                                nResponse = MsgBox(getResByKey("InvalidCharacter"), vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                                If nResponse = vbRetry Then
                                    cellRange.Select
                                End If
                                'MsgBox getResByKey("InvalidCharacter")  '字符串待定义
                                cellRange.value = ""
                                Exit Sub
                            End If
                        Next
                Next cellRange
            End If
         Next
    Next columnRange
    
    
    If target.row > 2 And target.column = connTypeCol And target.value = getResByKey("commConn") Then
                sheet.Cells(target.row, authenticationTypeCol).Interior.colorIndex = SolidColorIdx
                sheet.Cells(target.row, authenticationTypeCol).Interior.Pattern = SolidPattern
                sheet.Cells(target.row, authenticationTypeCol).value = ""
                sheet.Cells(target.row, authenticationTypeCol).Validation.ShowInput = False
    ElseIf target.row > 2 And target.column = connTypeCol And (target.value = getResByKey("sslConn") Or target.value = "") Then
                sheet.Cells(target.row, authenticationTypeCol).Interior.colorIndex = NullPattern
                sheet.Cells(target.row, authenticationTypeCol).Interior.Pattern = NullPattern
                sheet.Cells(target.row, authenticationTypeCol).Validation.ShowInput = True
    End If
    
End Sub









