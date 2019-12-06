VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBoardStylePositionMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Private groupNameMap_ As CMapValueObject
Private nonModelGroupNameMap_ As CMapValueObject


Public Sub init()
    Call initGroupNameMap
End Sub


Public Sub reset()
    Call groupNameMap_.Clean
    Call nonModelGroupNameMap_.Clean
    Call initGroupNameMap
End Sub


Public Sub getGroupStartAndEndRowByRowNum(ByRef ws As Worksheet, ByVal rowNumber As Long, _
    ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
    
    groupNameStartRowNumber = getGroupRowNumByRow(ws, rowNumber)
    
    groupNameEndRowNumber = getNextGroupRowNumByRow(ws, groupNameStartRowNumber + 2) - 2
    
End Sub

Public Sub getGroupStartAndEndRowByGroupName(ByRef ws As Worksheet, ByRef groupName As String, _
    ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)

    groupNameStartRowNumber = -1
    groupNameEndRowNumber = -1

    If Not groupNameMap_.hasKey(groupName) Then Exit Sub

    groupNameStartRowNumber = getGroupRowNumByGroupName(ws, groupName)

    If groupNameStartRowNumber <> -1 Then
        groupNameEndRowNumber = getNextGroupRowNumByRow(ws, groupNameStartRowNumber + 2) - 2
    End If

End Sub


Public Function getGroupRowNumByRow(ByRef ws As Worksheet, ByVal rowNumber) As Long
    Dim maxRowNumber As Long
    maxRowNumber = ws.UsedRange.Rows.count
    If rowNumber > maxRowNumber Then rowNumber = maxRowNumber

    Dim k As Long
    For k = rowNumber To 1 Step -1
        If isGroupAndColumnName(ws.Cells(k, 1), ws.Cells(k + 1, 1)) Then Exit For
    Next k
    
    getGroupRowNumByRow = k
End Function


Private Function getNextGroupRowNumByRow(ByRef ws As Worksheet, ByVal rowNumber) As Long
    Dim nextGroupRowNumber As Long
    nextGroupRowNumber = -1
    
    Dim maxRowNumber As Long, k As Long
    maxRowNumber = ws.UsedRange.Rows.count
    
    For k = rowNumber To maxRowNumber
        If isGroupAndColumnName(ws.Cells(k, 1), ws.Cells(k + 1, 1)) Then
            nextGroupRowNumber = k
            Exit For
        End If
    Next k
    
    '如果是-1，说明是最后一个分组，只能用单元格是否有边框来判断最大行了
    If nextGroupRowNumber = -1 Then
        '防止最后一个对象的边框一直设置到1048576，设置一个2000的行数最大限制
'        Dim predefinedMaxRowNumber As Long
'        predefinedMaxRowNumber = Application.WorksheetFunction.min(rowNumber + 2000, maxRowNumber)
'
'        If rowNumber < predefinedMaxRowNumber Then
'            nextGroupRowNumber = predefinedMaxRowNumber
'        End If
        nextGroupRowNumber = maxRowNumber + 2  '为了与正常分组的最大值一致，这里加2
    End If
    
    getNextGroupRowNumByRow = nextGroupRowNumber
End Function


Public Function getGroupRowNumByGroupName(ByRef ws As Worksheet, ByRef groupName As String) As Long
    
    getGroupRowNumByGroupName = -1
    If Not isGroupName(groupName) Then Exit Function
    
    Dim maxRowNumber As Long, k As Long
    maxRowNumber = ws.UsedRange.Rows.count
    
    For k = 1 To maxRowNumber
        If Not isGroupAndColumnName(ws.Cells(k, 1), ws.Cells(k + 1, 1)) Then GoTo NextLoop
        
        If findColNumByRowAndValue(ws, k, groupName) <> -1 Then
            getGroupRowNumByGroupName = k
            Exit For
        End If
NextLoop:
    Next k
    
End Function



Public Function getModelGroupAndColumnNameByRow(ByRef ws As Worksheet, ByVal rowNumber, ByRef groupName As String, ByRef columnName As String) As Long
    
    Dim groupNumber As Long
    groupNumber = getGroupRowNumByRow(ws, rowNumber)
    
    Dim maxColumnNumber As Long
    maxColumnNumber = getMaxCol(ws, groupNumber + 1)
    
    Dim colIndex As Long
    For colIndex = 1 To maxColumnNumber
        '使用getGroupAndColumnName时，必须从数据区域开始，否则会有异常，所以取groupNumber + 2
        Call getGroupAndColumnName(ws, ws.Cells(groupNumber + 2, colIndex), groupName, columnName)
        
        If Not isNonModelGroupAndColumnName(groupName, columnName) Then
            getModelGroupAndColumnNameByRow = colIndex
            Exit Function
        End If
    Next colIndex
    
    getModelStartColumnNumByRow = -1
End Function


Private Function isGroupName(ByRef groupName As String) As Boolean
    isGroupName = False
    
    If Trim(groupName) = "" Then Exit Function
    
    If groupNameMap_.hasKey(Trim(groupName)) Then
        isGroupName = True
    End If
    
End Function

Private Function isGroupAndColumnName(ByRef groupName As String, ByRef columnName As String) As Boolean
    isGroupAndColumnName = isInCMapValueObject(groupNameMap_, groupName, columnName)
End Function

Private Function isNonModelGroupAndColumnName(ByRef groupName As String, ByRef columnName As String) As Boolean
    isNonModelGroupAndColumnName = isInCMapValueObject(nonModelGroupNameMap_, groupName, columnName)
End Function



Private Sub initGroupNameMap()
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    
    Dim sheetName As String, groupName As String, columnName As String, mocName As String, oldColumnName As String
    
    
    For index = 2 To sheetDef.UsedRange.Rows.count
        sheetName = sheetDef.Cells(index, 1)
        groupName = sheetDef.Cells(index, 2)
        columnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)

        If InStr(sheetName, getResByKey("Board Style")) = 0 Then GoTo NextLoop
        
        Call insert2CMapValueObject(groupNameMap_, groupName, columnName)
        
        If InStr(groupName, getResByKey("Operation")) <> 0 Or InStr(mocName, "Customization") <> 0 Then
            Call insert2CMapValueObject(nonModelGroupNameMap_, groupName, columnName)
        End If
        
NextLoop:
    Next
End Sub

Private Sub insert2CMapValueObject(ByRef destMap As CMapValueObject, ByRef key As String, ByRef value As String)
    Dim columnNameList As Collection
    
    If Not destMap.hasKey(key) Then
        Set columnNameList = New Collection
        columnNameList.Add (value)
        Call destMap.SetAt(key, columnNameList)
    Else
        Set columnNameList = destMap.GetAt(key)
        columnNameList.Add (value)
    End If
End Sub


Private Function isInCMapValueObject(ByRef destMap As CMapValueObject, groupName As String, ByRef columnName As String) As Boolean
    isInCMapValueObject = False
    
    If Trim(groupName) = "" Or Trim(columnName) = "" Then Exit Function
    
    If Not destMap.hasKey(Trim(groupName)) Then Exit Function
    
    Dim columnNameList As Collection
    Set columnNameList = destMap.GetAt(groupName)
    
    If isInCollection(columnNameList, columnName) Then
        isInCMapValueObject = True
    End If
    
End Function



Private Sub Class_Initialize()
    Set groupNameMap_ = New CMapValueObject
    Set nonModelGroupNameMap_ = New CMapValueObject
    
End Sub


