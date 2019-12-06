Attribute VB_Name = "initDefaultDataSub"
Private valideCollection As Collection

Sub initInnerValideDef()
        If valideCollection Is Nothing Then
            Set valideCollection = New Collection
        End If
        Dim rowCount As Long
        Dim innerSheet As Worksheet
        Set innerSheet = ThisWorkbook.Worksheets("InnerValideDef")
        rowCount = innerSheet.Range("a1048576").End(xlUp).row
        Dim index As Long
        Dim valideDef As CValideDef
        
        For index = 2 To rowCount
                Set valideDef = New CValideDef
                valideDef.sheetName = innerSheet.Cells(index, 1).value
                valideDef.groupName = innerSheet.Cells(index, 2).value
                valideDef.columnName = innerSheet.Cells(index, 3).value
                valideDef.row = index
                valideDef.column = innerSheet.Range("XFD" + CStr(index)).End(xlToLeft).column
                Call valideCollection.Add(valideDef, valideDef.getKey)
        Next
End Sub

Function getInnerValideDef(key As String) As CValideDef
        If valideCollection Is Nothing Then
            Call initInnerValideDef
        End If
        Dim isExist As Boolean
        isExist = Contains(valideCollection, key)
        If isExist Then
                Set getInnerValideDef = valideCollection.Item(key)
        End If
End Function

Function modiflyInnerValideDef(sheetName As String, groupName As String, columnName As String, value As String, ByRef valideDef As CValideDef)
        Dim innerSheet As Worksheet
        Dim rowRange As Range
        Set innerSheet = ThisWorkbook.Worksheets("InnerValideDef")
        Set rowRange = innerSheet.rows(valideDef.row)
        rowRange.EntireRow.ClearContents
        rowRange.NumberFormatLocal = "@" '设置单元格格式为文本
        rowRange.Cells(1, 1).value = sheetName
        rowRange.Cells(1, 2).value = groupName
        rowRange.Cells(1, 3).value = columnName
        
        Dim values() As String
        Dim column As Long
        Dim index As Long
        values = Split(value, ",")
        column = 3
        For index = LBound(values) To UBound(values)
                column = column + 1
                rowRange.Cells(1, column).value = values(index)
        Next
        valideDef.column = column
        Call valideCollection.Remove(valideDef.getKey)
        Call valideCollection.Add(valideDef, valideDef.getKey)
        
End Function

Function addInnerValideDef(sheetName As String, groupName As String, columnName As String, value As String) As CValideDef
        Dim key As String
        key = sheetName + "," + groupName + "," + columnName
        Dim isExist As Boolean
        isExist = Contains(valideCollection, key)
        If isExist Then
            Exit Function
        End If
        Dim innerSheet As Worksheet
        Set innerSheet = ThisWorkbook.Worksheets("InnerValideDef")
        Dim row As Long
        Dim column As Long
        Dim index As Long
        Dim values() As String
        
        values = Split(value, ",")
        row = innerSheet.Range("a1048576").End(xlUp).row + 1
        column = 3
        innerSheet.rows(row).NumberFormatLocal = "@" '设置单元格格式为文本
        innerSheet.Cells(row, 1).value = sheetName
        innerSheet.Cells(row, 2).value = groupName
        innerSheet.Cells(row, 3).value = columnName
        
        For index = LBound(values) To UBound(values)
                column = column + 1
                innerSheet.Cells(row, column).value = values(index)
        Next
        
        Dim valideDef As CValideDef
        Set valideDef = New CValideDef
        valideDef.sheetName = sheetName
        valideDef.groupName = groupName
        valideDef.columnName = columnName
        valideDef.row = row
        valideDef.column = column
        Call valideCollection.Add(valideDef, valideDef.getKey)
        Set addInnerValideDef = valideDef
End Function





