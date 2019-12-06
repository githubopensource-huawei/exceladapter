Attribute VB_Name = "PositionUtils"
Public Const functionShtTitleRow As Integer = 1
Public Const listShtGrpRow As Integer = 1
Public Const listShtAttrRow As Integer = 2

'=================================================
'list and main sheets
'=================================================
Public Function colNumByAttr(ByRef ws As Worksheet, ByRef attrName As String) As Long
    colNumByAttr = -1
    
    Dim targetRange As Range
    Set targetRange = ws.Rows(listShtAttrRow).Find(attrName, LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then colNumByAttr = targetRange.column
End Function

'An attribute may be in different groups
Public Function colNumByGrpAndAttr(ByRef ws As Worksheet, ByRef grpName As String, ByRef attrName As String) As Long
    colNumByGrpAndAttr = -1
    
    Dim visitedGrpNames As New Collection
    Dim firstAddr As String
    
    Dim tmpGrpName As String
    tmpGrpName = ""
    
    Dim shtname As String
    shtname = ws.name
    
    Dim targetRange As Range
    With ws.Rows(listShtAttrRow)
        Set targetRange = .Find(attrName, LookIn:=xlValues, lookat:=xlWhole)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                tmpGrpName = getGroupNameFromMappingDef(shtname, attrName, visitedGrpNames)
                If grpName = tmpGrpName Then
                    colNumByGrpAndAttr = targetRange.column
                    Exit Function
                End If
                Set targetRange = .FindNext(targetRange)
                If Not Contains(visitedGrpNames, tmpGrpName) Then visitedGrpNames.Add Item:=tmpGrpName, key:=tmpGrpName
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
End Function

'Public Function siteNameColNum(ByRef ws As Worksheet) As Long
'    siteNameColNum = -1
'
'    Dim colIdx As Long
'    With ws
'        For colIdx = 1 To .range("XFD2").End(xlToLeft).column
'            If is_Site(.Cells(listShtAttrRow, colIdx)) Then
'                siteNameColNum = colIdx
'                Exit Function
'            End If
'        Next
'    End With
'End Function

'Public Function controllerNameColNum(ByRef ws As Worksheet) As Long
'    controllerNameColNum = -1
'
'    Dim colIdx As Long
'    With ws
'        For colIdx = 1 To .range("XFD2").End(xlToLeft).column
'            If is_Controller(.Cells(listShtAttrRow, colIdx)) Then
'                controllerNameColNum = colIdx
'                Exit Function
'            End If
'        Next
'    End With
'End Function

Public Function operationColNum(ByRef ws As Worksheet) As Long
    operationColNum = -1

    Dim targetRange As Range
    Set targetRange = ws.Rows(listShtAttrRow).Find(getResByKey("Operation_Group"), LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then operationColNum = targetRange.column
End Function

'=================================================
'SHEET DEF
'=================================================
Public Function shtNameColNumInShtDef() As Long
    shtNameColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")

    Dim targetRange As Range
    Set targetRange = sheetDef.Rows(functionShtTitleRow).Find("Sheet Name", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then shtNameColNumInShtDef = targetRange.column
End Function

Public Function shtTypeColNumInShtDef() As Long
    shtTypeColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")

    Dim targetRange As Range
    Set targetRange = sheetDef.Rows(functionShtTitleRow).Find("Sheet Type", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then shtTypeColNumInShtDef = targetRange.column
End Function

Public Function startRowColNumInShtDef() As Long
    startRowColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")

    Dim targetRange As Range
    Set targetRange = sheetDef.Rows(functionShtTitleRow).Find("StartRow", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then startRowColNumInShtDef = targetRange.column
End Function

Public Function endRowColNumInShtDef() As Long
    endRowColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")

    Dim targetRange As Range
    Set targetRange = sheetDef.Rows(functionShtTitleRow).Find("EndRow", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then endRowColNumInShtDef = targetRange.column
End Function

Public Function baseShtNameColNumInShtDef() As Long
    baseShtNameColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")

    Dim targetRange As Range
    Set targetRange = sheetDef.Rows(functionShtTitleRow).Find("Base Sheet Name", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then baseShtNameColNumInShtDef = targetRange.column
End Function

Public Function mappingTypeColNumInShtDef() As Long
    mappingTypeColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")

    Dim targetRange As Range
    Set targetRange = sheetDef.Rows(functionShtTitleRow).Find("Mapping Type", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then mappingTypeColNumInShtDef = targetRange.column
End Function

Public Function ratTypeColNumInShtDef() As Long
    ratTypeColNumInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")

    Dim targetRange As Range
    Set targetRange = sheetDef.Rows(functionShtTitleRow).Find("Rat Type", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then ratTypeColNumInShtDef = targetRange.column
End Function

Public Function shtRowInShtDef(shtname As String) As Integer
    shtRowInShtDef = -1
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    
    Dim targetRange As Range
    Set targetRange = sheetDef.Columns(shtNameColNumInShtDef).Find(shtname, LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then shtRowInShtDef = targetRange.row
End Function

'=================================================
'MAPPING DEF
'=================================================
Public Function shtNameColNumInMappingDef() As Long
    shtNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Sheet Name", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then shtNameColNumInMappingDef = targetRange.column
End Function

Public Function grpNameColNumInMappingDef() As Long
    grpNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Group Name", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then grpNameColNumInMappingDef = targetRange.column
End Function

Public Function colNameColNumInMappingDef() As Long
    colNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Column Name", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then colNameColNumInMappingDef = targetRange.column
End Function

Public Function mocNameColNumInMappingDef() As Long
    mocNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("MOC Name", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then mocNameColNumInMappingDef = targetRange.column
End Function

Public Function attrNameColNumInMappingDef() As Long
    attrNameColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Attribute Name", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then attrNameColNumInMappingDef = targetRange.column
End Function

Public Function isRefColNumInMappingDef() As Long
    isRefColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Is Reference", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then isRefColNumInMappingDef = targetRange.column
End Function

Public Function isKeyColNumInMappingDef() As Long
    isKeyColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Is Key", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then isKeyColNumInMappingDef = targetRange.column
End Function

Public Function innerKeyColNumInMappingDef() As Long
    innerKeyColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Inner Key", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then innerKeyColNumInMappingDef = targetRange.column
End Function

Public Function readOnlyColNumInMappingDef() As Long
    readOnlyColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("ReadOnly", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then readOnlyColNumInMappingDef = targetRange.column
End Function

Public Function lldColNumInMappingDef() As Long
    lldColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("LLD", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then lldColNumInMappingDef = targetRange.column
End Function

Public Function lineMappingColNumInMappingDef() As Long
    lineMappingColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Line Mapping", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then lineMappingColNumInMappingDef = targetRange.column
End Function

Public Function neTypeColNumInMappingDef() As Long
    neTypeColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Ne Type", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then neTypeColNumInMappingDef = targetRange.column
End Function

Public Function neVersionColNumInMappingDef() As Long
    neVersionColNumInMappingDef = -1
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")

    Dim targetRange As Range
    Set targetRange = mappingDef.Rows(functionShtTitleRow).Find("Ne Version", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then neVersionColNumInMappingDef = targetRange.column
End Function

'=================================================
'CONTROL DEF
'=================================================
Public Function mocNameColNumInCtrlDef() As Long
    mocNameColNumInCtrlDef = -1
    
    Dim controlDef As Worksheet
    Set controlDef = ThisWorkbook.Worksheets("CONTROL DEF")

    Dim targetRange As Range
    Set targetRange = controlDef.Rows(functionShtTitleRow).Find("MOC Name", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then mocNameColNumInCtrlDef = targetRange.column
End Function

Public Function attrNameColNumInCtrlDef() As Long
    attrNameColNumInCtrlDef = -1
    
    Dim controlDef As Worksheet
    Set controlDef = ThisWorkbook.Worksheets("CONTROL DEF")

    Dim targetRange As Range
    Set targetRange = controlDef.Rows(functionShtTitleRow).Find("Attribute Name", LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then attrNameColNumInCtrlDef = targetRange.column
End Function


'=================================================
'auxiliary functions
'=================================================
'An attribute may be in different groups
Public Function getGroupNameFromMappingDef(shtname As String, attrName As String, Optional excludeGrpNames As Collection) As String
    getGroupNameFromMappingDef = ""
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    
    Dim shtNameCol As Long
    Dim colNameCol As Long
    Dim grpNameCol As Long
    shtNameCol = shtNameColNumInMappingDef
    colNameCol = colNameColNumInMappingDef
    grpNameCol = grpNameColNumInMappingDef
    
    Dim tmpShtName As String
    Dim grpName As String
    Dim targetRange As Range
    Dim firstAddr As String
    
    With mappingDef.Columns(colNameCol)
        Set targetRange = .Find(attrName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                tmpShtName = targetRange.Offset(0, shtNameCol - colNameCol).value
                If tmpShtName = shtname Then
                    grpName = targetRange.Offset(0, grpNameCol - colNameCol).value
                    If Not excludeGrpNames Is Nothing Then
                        If Not Contains(excludeGrpNames, grpName) Then
                            getGroupNameFromMappingDef = grpName
                            Exit Function
                        End If
                    Else
                        getGroupNameFromMappingDef = grpName
                        Exit Function
                    End If
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
End Function

'=================================================
'从列数得到列名：1->A，27->AA
'=================================================
Public Function C(ByVal iColumn As Long) As String
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
    D = -1
    D = Range(ColumnStr & "1").column
End Function



