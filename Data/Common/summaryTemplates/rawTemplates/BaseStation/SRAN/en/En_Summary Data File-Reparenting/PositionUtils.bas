Attribute VB_Name = "PositionUtils"
Public Const functionShtTitleRow As Integer = 1
Public Const listShtGrpRow As Integer = 1
Public Const listShtAttrRow As Integer = 2

Public innerPositionMgr As CInnerPositionManager

Public Sub loadInnerPositions()
    Set innerPositionMgr = New CInnerPositionManager
End Sub

'=================================================
'list and main sheets
'=================================================
Public Function colNumByAttr(ByRef ws As Worksheet, ByRef attrName As String) As Long
    colNumByAttr = -1
    
    Dim targetRange As range
    Set targetRange = ws.rows(listShtAttrRow).Find(attrName, LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then colNumByAttr = targetRange.column
End Function

'An attribute may be in different groups
Public Function colNumByGrpAndAttr(ByRef ws As Worksheet, ByRef grpName As String, ByRef attrName As String) As Long
    colNumByGrpAndAttr = -1
    
    Dim visitedGrpNames As New Collection
    Dim firstAddr As String
    
    Dim tmpGrpName As String
    tmpGrpName = ""
    
    Dim shtName As String
    shtName = ws.name
    
    Dim targetRange As range
    With ws.rows(listShtAttrRow)
        Set targetRange = .Find(attrName, LookIn:=xlValues, lookat:=xlWhole)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                tmpGrpName = getGroupNameFromMappingDef(shtName, attrName, visitedGrpNames)
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

Public Function siteNameColNum(ByRef ws As Worksheet) As Long
    siteNameColNum = -1
    
    Dim colIdx As Long
    With ws
        For colIdx = 1 To .range("IV2").End(xlToLeft).column
            If is_Site(.Cells(listShtAttrRow, colIdx)) Then
                siteNameColNum = colIdx
                Exit Function
            End If
        Next
    End With
End Function

Public Function controllerNameColNum(ByRef ws As Worksheet) As Long
    controllerNameColNum = -1
    
    Dim colIdx As Long
    With ws
        For colIdx = 1 To .range("IV2").End(xlToLeft).column
            If is_Controller(.Cells(listShtAttrRow, colIdx)) Then
                controllerNameColNum = colIdx
                Exit Function
            End If
        Next
    End With
End Function

Public Function operationColNum(ByRef ws As Worksheet) As Long
    operationColNum = -1

    Dim targetRange As range
    Set targetRange = ws.rows(listShtAttrRow).Find(getResByKey("OPERATION"), LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then operationColNum = targetRange.column
End Function

Public Function shtRowInShtDef(shtName As String) As Integer
    shtRowInShtDef = -1
    
    If innerPositionMgr Is Nothing Then loadInnerPositions
    Dim targetRange As range
    Set targetRange = ThisWorkbook.Worksheets("SHEET DEF").columns(innerPositionMgr.sheetDef_shtNameColNo).Find(shtName, LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then shtRowInShtDef = targetRange.row
End Function


'=================================================
'auxiliary functions
'=================================================
'An attribute may be in different groups
Public Function getGroupNameFromMappingDef(shtName As String, attrName As String, Optional excludeGrpNames As Collection) As String
    getGroupNameFromMappingDef = ""
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim shtNameCol As Long
    Dim colNameCol As Long
    Dim grpNameCol As Long
    shtNameCol = innerPositionMgr.mappingDef_shtNameColNo
    colNameCol = innerPositionMgr.mappingDef_colNameColNo
    grpNameCol = innerPositionMgr.mappingDef_grpNameColNo
    
    Dim tmpShtName As String
    Dim grpName As String
    Dim targetRange As range
    Dim firstAddr As String
    
    With mappingDef.columns(colNameCol)
        Set targetRange = .Find(attrName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                tmpShtName = targetRange.Offset(0, shtNameCol - colNameCol).value
                If tmpShtName = shtName Then
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
    D = range(ColumnStr & "1").column
End Function



