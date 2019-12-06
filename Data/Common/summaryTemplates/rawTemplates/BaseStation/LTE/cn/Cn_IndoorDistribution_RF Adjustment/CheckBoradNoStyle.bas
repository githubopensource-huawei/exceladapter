Attribute VB_Name = "CheckBoradNoStyle"
Option Explicit
Public Sub CheckBoardNoStyle(ByVal sheet As Object, ByVal target As range)
    Dim loccellIdCol As Long
    Dim sectoIdCol As Long
    Dim basebandeqmIdCol As Long
    Dim sheetName As String
    sheetName = sheet.name
    
    If InStr(sheetName, "调整样式") Then
        
    
    End If
    
    
    loccellIdCol = findCertainValColumnNumber(sheet, 1, getResByKey("locellId"))
    sectoIdCol = findCertainValColumnNumber(sheet, 1, getResByKey("sectorId"))
    basebandeqmIdCol = findCertainValColumnNumber(sheet, 1, getResByKey("basebandeqmId"))
    
    Dim loccellIdValue As String
    Dim sectoIdValue As String
    Dim basebandeqmIdValue As String
    Dim nResponse As String
    

    Dim columnRange As range, cellRange As range
    
    For Each columnRange In target.columns
        If columnRange.column = loccellIdCol Then
            For Each cellRange In columnRange
                loccellIdValue = Trim(cellRange.value)
                If target.row > 1 And loccellIdValue <> "" Then
                    If CStr(loccellIdValue) < 0 Or CStr(loccellIdValue) > 255 Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~255]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                    End If
                End If
            Next cellRange
        ElseIf columnRange.column = sectoIdCol Then
            For Each cellRange In columnRange
                sectoIdValue = Trim(cellRange.value)
                If target.row > 1 And sectoIdValue <> "" Then
                    If CStr(sectoIdValue) < 0 Or CStr(sectoIdValue) > 255 Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~1048576]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                    End If
                End If
            Next cellRange
            
         ElseIf columnRange.column = basebandeqmIdCol Then
            For Each cellRange In columnRange
                basebandeqmIdValue = Trim(cellRange.value)
                If basebandeqmIdValue <> "" And target.row > 1 Then
                    If CStr(basebandeqmIdValue) < 0 Or (CStr(basebandeqmIdValue) > 23 And CStr(basebandeqmIdValue) <> 255) Then
                        nResponse = MsgBox(getResByKey("Range") & "[0~23],[255]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
                        If nResponse = vbRetry Then
                            cellRange.Select
                        End If
                        cellRange.value = ""
                        Exit Sub
                    End If
                End If
            Next cellRange
            
        End If
    Next columnRange
End Sub
