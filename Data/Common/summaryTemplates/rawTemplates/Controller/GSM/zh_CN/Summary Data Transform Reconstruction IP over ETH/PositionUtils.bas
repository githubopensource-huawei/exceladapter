Attribute VB_Name = "PositionUtils"
Public Const functionShtTitleRow = 1
Public Const listShtAttrRow = 2

'=================================================
'list and main sheets
'=================================================
Public Function colNumByAttr(ByRef ws As Worksheet, ByRef attrName As String) As Long
    colNumByAttr = -1
    
    Dim colIdx As Long
    For colIdx = 1 To ws.Range("XFD2").End(xlToLeft).column
        If ws.Cells(listShtAttrRow, colIdx) = attrName Then
            colNumByAttr = colIdx
            Exit Function
        End If
    Next
End Function

Public Function colNumByGrpAndAttr(ByRef ws As Worksheet, ByRef grpName As String, ByRef attrName As String) As Long
    colNumByGrpAndAttr = -1
    
    Dim colIdx As Long
    For colIdx = 1 To ws.Range("XFD2").End(xlToLeft).column
        If ws.Cells(listShtAttrRow, colIdx) = attrName And grpName = getGroupNameFromMappingDef(ws.name, attrName) Then
            colNumByGrpAndAttr = colIdx
            Exit Function
        End If
    Next
End Function

Public Function siteNameColNum(ByRef ws As Worksheet) As Long
    siteNameColNum = -1
    
    Dim colIdx As Long
    For colIdx = 1 To ws.Range("XFD2").End(xlToLeft).column
        If is_Site(ws.Cells(listShtAttrRow, colIdx)) Then
            siteNameColNum = colIdx
            Exit Function
        End If
    Next
End Function

Public Function controllerNameColNum(ByRef ws As Worksheet) As Long
    controllerNameColNum = -1
    
    Dim colIdx As Long
    For colIdx = 1 To ws.Range("XFD2").End(xlToLeft).column
        If is_Controller(ws.Cells(listShtAttrRow, colIdx)) Then
            controllerNameColNum = colIdx
            Exit Function
        End If
    Next
End Function

Public Function operationColNum(ByRef ws As Worksheet) As Long
    operationColNum = -1
    
    Dim colIdx As Long
    For colIdx = 1 To ws.Range("XFD2").End(xlToLeft).column
        If ws.Cells(listShtAttrRow, colIdx) = getResByKey("OPERATION") Then
            operationColNum = colIdx
            Exit Function
        End If
    Next
End Function

'=================================================
'SHEET DEF
'=================================================
Public Function shtNameColNumInShtDef() As Long
    shtNameColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    Dim colIdx As Long
    For colIdx = 1 To sheetDef.Range("XFD1").End(xlToLeft).column
        If UCase(sheetDef.Cells(functionShtTitleRow, colIdx)) = UCase("Sheet Name") Then
            shtNameColNumInShtDef = colIdx
            Exit Function
        End If
    Next
End Function

Public Function shtTypeColNumInShtDef() As Long
    shtTypeColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    Dim colIdx As Long
    For colIdx = 1 To sheetDef.Range("XFD1").End(xlToLeft).column
        If UCase(sheetDef.Cells(functionShtTitleRow, colIdx)) = UCase("Sheet Type") Then
            shtTypeColNumInShtDef = colIdx
            Exit Function
        End If
    Next
End Function

Public Function startRowColNumInShtDef() As Long
    startRowColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    Dim colIdx As Long
    For colIdx = 1 To sheetDef.Range("XFD1").End(xlToLeft).column
        If UCase(sheetDef.Cells(functionShtTitleRow, colIdx)) = UCase("StartRow") Then
            startRowColNumInShtDef = colIdx
            Exit Function
        End If
    Next
End Function

Public Function endRowColNumInShtDef() As Long
    endRowColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    Dim colIdx As Long
    For colIdx = 1 To sheetDef.Range("XFD1").End(xlToLeft).column
        If UCase(sheetDef.Cells(functionShtTitleRow, colIdx)) = UCase("EndRow") Then
            endRowColNumInShtDef = colIdx
            Exit Function
        End If
    Next
End Function

'=================================================
'MAPPING DEF
'=================================================
Public Function shtNameColNumInMappingDef() As Long
    shtNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim colIdx As Long
    For colIdx = 1 To mappingDef.Range("XFD1").End(xlToLeft).column
        If UCase(mappingDef.Cells(functionShtTitleRow, colIdx)) = UCase("Sheet Name") Then
            shtNameColNumInMappingDef = colIdx
            Exit Function
        End If
    Next
End Function

Public Function grpNameColNumInMappingDef() As Long
    grpNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim colIdx As Long
    For colIdx = 1 To mappingDef.Range("XFD1").End(xlToLeft).column
        If UCase(mappingDef.Cells(functionShtTitleRow, colIdx)) = UCase("Group Name") Then
            grpNameColNumInMappingDef = colIdx
            Exit Function
        End If
    Next
End Function

Public Function colNameColNumInMappingDef() As Long
    colNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim colIdx As Long
    For colIdx = 1 To mappingDef.Range("XFD1").End(xlToLeft).column
        If UCase(mappingDef.Cells(functionShtTitleRow, colIdx)) = UCase("Column Name") Then
            colNameColNumInMappingDef = colIdx
            Exit Function
        End If
    Next
End Function

Public Function mocNameColNumInMappingDef() As Long
    mocNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim colIdx As Long
    For colIdx = 1 To mappingDef.Range("XFD1").End(xlToLeft).column
        If UCase(mappingDef.Cells(functionShtTitleRow, colIdx)) = UCase("MOC Name") Then
            mocNameColNumInMappingDef = colIdx
            Exit Function
        End If
    Next
End Function

Public Function attrNameColNumInMappingDef() As Long
    attrNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim colIdx As Long
    For colIdx = 1 To mappingDef.Range("XFD1").End(xlToLeft).column
        If UCase(mappingDef.Cells(functionShtTitleRow, colIdx)) = UCase("Attribute Name") Then
            attrNameColNumInMappingDef = colIdx
            Exit Function
        End If
    Next
End Function

Public Function neTypeColNumInMappingDef() As Long
    neTypeColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim colIdx As Long
    For colIdx = 1 To mappingDef.Range("XFD1").End(xlToLeft).column
        If UCase(mappingDef.Cells(functionShtTitleRow, colIdx)) = UCase("Ne Type") Then
            neTypeColNumInMappingDef = colIdx
            Exit Function
        End If
    Next
End Function

'=================================================
'CONTROL DEF
'=================================================
Public Function mocNameColNumInCtrlDef() As Long
    mocNameColNumInCtrlDef = -1
    
    Dim controlDef As Worksheet
    Set controlDef = ThisWorkbook.Worksheets("CONTROL DEF")
    Dim colIdx As Long
    For colIdx = 1 To controlDef.Range("XFD1").End(xlToLeft).column
        If UCase(controlDef.Cells(functionShtTitleRow, colIdx)) = UCase("MOC Name") Then
            mocNameColNumInCtrlDef = colIdx
            Exit Function
        End If
    Next
End Function

Public Function attrNameColNumInCtrlDef() As Long
    attrNameColNumInCtrlDef = -1
    
    Dim controlDef As Worksheet
    Set controlDef = ThisWorkbook.Worksheets("CONTROL DEF")
    Dim colIdx As Long
    For colIdx = 1 To controlDef.Range("XFD1").End(xlToLeft).column
        If UCase(controlDef.Cells(functionShtTitleRow, colIdx)) = UCase("Attribute Name") Then
            attrNameColNumInCtrlDef = colIdx
            Exit Function
        End If
    Next
End Function


'=================================================
'auxiliary functions
'=================================================
Public Function getGroupNameFromMappingDef(sheetName As String, attributeName As String) As String
    Dim mappingDef As Worksheet
    Dim index, count As Long
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    count = mappingDef.UsedRange.rows.count
    For index = 1 To count
        If mappingDef.Cells(index, 1).value = sheetName And mappingDef.Cells(index, 3).value = attributeName Then
            getGroupNameFromMappingDef = mappingDef.Cells(index, 2)
            Exit Function
        End If
    Next
    getGroupNameFromMappingDef = ""
End Function

'=================================================
'从列数得到列名：1->A，27->AA
'=================================================
Public Function C(iColumn As Long) As String
  If iColumn >= 257 Or iColumn < 0 Then
    C = ""
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
  C = result
End Function

'=================================================
'从列名得到列数：A->1，AA->27
'=================================================
Public Function D(ColumnStr As String) As Long
  If Len(ColumnStr) = 1 Then
    D = Int(ColumnStr) - 64
  ElseIf Len(ColumnStr) = 2 Then
    D = (Int(Left(ColumnStr, 1)) - 64) * 26 + (Int(Left(ColumnStr, 1)) - 64)
  End If
End Function
