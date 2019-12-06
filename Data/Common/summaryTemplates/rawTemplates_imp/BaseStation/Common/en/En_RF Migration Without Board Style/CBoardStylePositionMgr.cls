VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBoardStylePositionMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Private modelGroupNameMap_ As CMapValueObject
Private nonModelGroupNameMap_ As CMapValueObject


Public Sub init()
    Call initGroupNameMap
End Sub


Public Sub reset()
    Call modelGroupNameMap_.Clean
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

    If Not isGroupName(groupName) Then Exit Sub

    groupNameStartRowNumber = getGroupRowNumByGroupName(ws, groupName)

    If groupNameStartRowNumber <> -1 Then
        groupNameEndRowNumber = getNextGroupRowNumByRow(ws, groupNameStartRowNumber + 2) - 2
    End If

End Sub



'根据字段内容判断是否Group行:
' 1、首列为空，不是Group行
' 2、根据上下两个单元格，判断是否GroupName+AttrName组合，如果不是，不是Group行
' 3、首次判断到模型字段的GroupName+AttrName，则认为是Group行
'不能用非模型字段判断，因为定制字段的GroupName、AttrName、AttrValue都是用户定制，可能会影响判断结果

Public Function isGroupRow(ByRef ws As Worksheet, ByVal row As Long) As Long
    Dim maxCol As Long
    Dim col As Long
    Dim groupName As String
    Dim attrName As String

    isGroupRow = False

    maxCol = getMaxCol(ws, row)
    For col = 1 To maxCol
        groupName = Trim(ws.Cells(row, col))
        attrName = Trim(ws.Cells(row + 1, col))

        If groupName = "" And col = 1 Then
            isGroupRow = False
            Exit Function
        ElseIf groupName = "" Then
            GoTo NextLoop
        End If

        If Not isGroupAndColumnName(groupName, attrName) Then
            isGroupRow = False
            Exit Function
        End If

        If isModelGroupAndColumnName(groupName, attrName) Then
            isGroupRow = True
            Exit Function
        End If
NextLoop:
    Next col
End Function

Public Function isColumnRow(ByRef ws As Worksheet, ByVal row As Long) As Boolean
    isColumnRow = False
    
    If row <= 1 Then
        Exit Function
    End If
    
    isColumnRow = boardStylePosMgr.isGroupRow(ws, row - 1)
End Function

Public Function isGroupRowEx(ByRef ws As Worksheet, ByVal row As Long, ByRef groupName As String, ByRef columnName As String) As Long
    Dim maxCol As Long

End Function

Public Function getGroupRowNumByRow(ByRef ws As Worksheet, ByVal rowNumber) As Long
    Dim maxRowNumber As Long
    maxRowNumber = ws.UsedRange.Rows.count
    If rowNumber > maxRowNumber Then rowNumber = maxRowNumber

    Dim k As Long
    For k = rowNumber To 1 Step -1
        If isGroupRow(ws, k) Then Exit For
    Next k
    
    getGroupRowNumByRow = k
End Function


Private Function getNextGroupRowNumByRow(ByRef ws As Worksheet, ByVal rowNumber) As Long
    Dim nextGroupRowNumber As Long
    nextGroupRowNumber = -1
    
    Dim maxRowNumber As Long, k As Long
    maxRowNumber = ws.UsedRange.Rows.count
    
    For k = rowNumber To maxRowNumber
        If isGroupRow(ws, k) Then
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
        If Not isGroupRow(ws, k) Then GoTo NextLoop
        
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
    
    Dim rawGroupName As String
    rawGroupName = Trim(groupName)
    
    If Len(rawGroupName) = 0 Then
        isGroupName = False
        Exit Function
    End If
        
    If modelGroupNameMap_.hasKey(rawGroupName) Then
        isGroupName = True
        Exit Function
    End If
    
    If nonModelGroupNameMap_.hasKey(rawGroupName) Then
        isGroupName = True
        Exit Function
    End If
    
End Function

Private Function isGroupAndColumnName(ByRef groupName As String, ByRef columnName As String) As Boolean
    isGroupAndColumnName = False
    
    If isModelGroupAndColumnName(groupName, columnName) Then
        isGroupAndColumnName = True
        Exit Function
    End If
    
    If isNonModelGroupAndColumnName(groupName, columnName) Then
        isGroupAndColumnName = True
        Exit Function
    End If
End Function

Private Function isModelGroupAndColumnName(ByRef groupName As String, ByRef columnName As String) As Boolean
    If Len(groupName) = 0 Or Len(columnName) = 0 Then
        isModelGroupNameAndColumn = False
        Exit Function
    End If
    
    isModelGroupAndColumnName = isInCMapValueObject(modelGroupNameMap_, groupName, columnName)
End Function

Private Function isNonModelGroupAndColumnName(ByRef groupName As String, ByRef columnName As String) As Boolean
    If Len(groupName) = 0 Or Len(columnName) = 0 Then
        isNonModelGroupAndColumnName = False
        Exit Function
    End If
    
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
        
        If InStr(groupName, getResByKey("Operation")) <> 0 Or InStr(mocName, "Customization") <> 0 Then
            Call insert2CMapValueObject(nonModelGroupNameMap_, groupName, columnName)
        Else
            Call insert2CMapValueObject(modelGroupNameMap_, groupName, columnName)
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
    Set modelGroupNameMap_ = New CMapValueObject
    Set nonModelGroupNameMap_ = New CMapValueObject
    
End Sub




