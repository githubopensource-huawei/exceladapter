VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAllBoardStyleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private allNeNamesStr_ As String
Private addDeleboxListFlag_ As Boolean
Private allBoardStyleDataMap_ As CMapValueObject

Public Sub initBoardStyleDataMap()

    If baseStationData Is Nothing Then Call initBaseStationDataPublic
    Set neBoardStyleMap = baseStationData.neBoardStyleMap
    
    If boardStyleData Is Nothing Then Call initBoardStyleMappingDataPublic
    Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    
    Dim neNames As Collection
    Dim neName As Variant
    Dim neBoardStyleDataMap As CMap
    
    Set neNames = getKeyValueCollection(neBoardStyleMap.KeyCollection)
    
    allNeNamesStr_ = ""

    For Each neName In neNames
        Set neBoardStyleDataMap = New CMap
        
        Call setNeBoardStyleDataMap(CStr(neName), neBoardStyleDataMap)
        Call allBoardStyleDataMap_.SetAt(CStr(neName), neBoardStyleDataMap)
        
        If allNeNamesStr_ = "" Then
            allNeNamesStr_ = CStr(neName)
        Else
            allNeNamesStr_ = allNeNamesStr_ + "," + CStr(neName)
        End If
    Next neName
    
End Sub
Public Function allBoardStyleDataMap() As CMapValueObject
    Set allBoardStyleDataMap = allBoardStyleDataMap_
End Function

Public Sub clearSourceColumBoxList(ByRef ws As Worksheet, ByRef groupName As String, ByRef startRowNumber As Long, ByRef endRowNumber As Long)

    addDeleboxListFlag_ = False
    Call initSourceColnumBoxLst(ws, groupName, startRowNumber, endRowNumber)
    
End Sub
Public Sub addSourceColumBoxList(ByRef ws As Worksheet, ByRef groupName As String, ByRef startRowNumber As Long, ByRef endRowNumber As Long)

    addDeleboxListFlag_ = True
    Call initSourceColnumBoxLst(ws, groupName, startRowNumber, endRowNumber)
    
End Sub

Private Sub setNeBoardStyleDataMap(ByRef neName As String, ByRef neBoardStyleDataMap As CMap)
    Dim boardStylegroupNameCol As Collection
    Set boardStylegroupNameCol = getKeyValueCollection(boardStyleMappingDefMap.KeyCollection)
    
    Dim boardStylegroupName As Variant
    Dim boardStyleSheetName As String
    Dim boardStyleSheet As Worksheet
    
    boardStyleSheetName = neBoardStyleMap.GetAt(neName)
    If containsASheet(ThisWorkbook, boardStyleSheetName) = False Then
        Exit Sub
    End If
    Set boardStyleSheet = ThisWorkbook.Worksheets(boardStyleSheetName)
    
    For Each boardStylegroupName In boardStylegroupNameCol
        Call neBoardStyleGroupDataMap(boardStyleSheet, neName, CStr(boardStylegroupName), neBoardStyleDataMap)
    Next boardStylegroupName
    
End Sub

Private Sub neBoardStyleGroupDataMap(ByRef ws As Worksheet, ByRef neName As String, ByRef groupName As String, ByRef neBoardStyleDataMap As CMap)
    Dim sourcecopyBoardNoNameCol As Collection
    Dim i As Long
    i = 0

    Set tempselectedGroupMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
    Set sourcecopyBoardNoNameCol = tempselectedGroupMappingDefData.copyFillInSourceColumnName
    
    
    
    Dim sourcecopyBoardNoName As Variant
    Dim autoCopyName As String
    
    For Each sourcecopyBoardNoName In sourcecopyBoardNoNameCol
        autoCopyName = autoFillinCopySourceColumnName(i)
        Call neBoardStyleGroupColNameDataMap(ws, neName, groupName, CStr(sourcecopyBoardNoName), autoCopyName, neBoardStyleDataMap)
        i = i + 1
    Next sourcecopyBoardNoName
    
End Sub

Private Function autoFillinCopySourceColumnName(ByRef copyNamePos As Long) As String
    Dim i As Long
    Dim autoCopySourceName As Variant
    Dim autoCopySourceNameCol As Collection
    i = 0
    
    Set autoCopySourceNameCol = tempselectedGroupMappingDefData.autoFillInSourceColumnName
    
    autoFillinCopySourceColumnName = ""
    
    For Each autoCopySourceName In autoCopySourceNameCol
        If copyNamePos = i Then
            autoFillinCopySourceColumnName = CStr(autoCopySourceName)
            Exit Function
        End If
        i = i + 1
    Next autoCopySourceName
    
End Function

Private Sub neBoardStyleGroupColNameDataMap(ByRef ws As Worksheet, ByRef neName As String, ByRef groupName As String, ByRef sourcecopyBoardNoName As String, ByRef autoCopyName As String, ByRef neBoardStyleDataMap As CMap)
    Dim columnNamePositionLetterMap As CMap
    Dim rowNumber As Long
    Dim columnNamePositionLetter As String
    Dim keyStr As String
    Dim valueStr As String
    
    Set columnNamePositionLetterMap = tempselectedGroupMappingDefData.columnNamePositionLetterMap
    
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    Call getBoardStyleGroupNameStartAndEndRowNumber(ws, groupName, groupNameStartRowNumber, groupNameEndRowNumber)
    
    columnNamePositionLetter = columnNamePositionLetterMap.GetAt(sourcecopyBoardNoName)
    keyStr = neName + "_" + groupName + "_" + autoCopyName
    valueStr = ""
    
    For rowNumber = groupNameStartRowNumber + 2 To groupNameEndRowNumber
        If valueStr = "" Then
            valueStr = ws.range(columnNamePositionLetter & rowNumber).value
        Else
            valueStr = valueStr + "," + ws.range(columnNamePositionLetter & rowNumber).value
        End If
  
    Next rowNumber
    
    Call neBoardStyleDataMap.SetAt(keyStr, valueStr)
End Sub


Private Sub initSourceColnumBoxLst(ByRef ws As Worksheet, ByRef groupName As String, ByRef startRowNumber As Long, ByRef endRowNumber As Long)
    Dim autocopyBoardNoNameCol As Collection
    Dim autocopyBoardNoNameVar As Variant
    
    Set tempselectedGroupMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
    Set autocopyBoardNoNameCol = tempselectedGroupMappingDefData.autoFillInSourceColumnName
    
    For Each autocopyBoardNoNameVar In autocopyBoardNoNameCol
        Call neSourceGroupColBoxList(ws, groupName, CStr(autocopyBoardNoNameVar), startRowNumber, endRowNumber)
    Next autocopyBoardNoNameVar
    
End Sub

Private Sub neSourceGroupColBoxList(ByRef ws As Worksheet, ByRef groupName As String, ByRef autocopyBoardNoName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
    Dim rowNumber As Long
    Dim columnNamePositionLetter As String
    Dim keyStr As String
    Dim valueStr As String
    Dim sourceNeName As String
    Dim sourceNeNameCol As String
    Dim sourceNeNameLetter As String '无线搬迁表格中，认为有效数据从这一列开始
    sourceNeNameCol = getResByKey("SOURCENENAME")
    sourceNeNameLetter = getSrcNeNameColWithGrpName(ws, groupNameStartRowNumber, groupName)
    
    Dim columnNamePositionLetterMap As CMap
    Set columnNamePositionLetterMap = tempselectedGroupMappingDefData.columnNamePositionLetterMap
    columnNamePositionLetter = columnNamePositionLetterMap.GetAt(autocopyBoardNoName)
    
    valueStr = ""
    
    Dim boardStyleDataMap As CMap
    
    For rowNumber = groupNameStartRowNumber To groupNameEndRowNumber
        If addDeleboxListFlag_ Then
            Call setBoardStyleListBoxRangeValidation(ws.name, groupName, sourceNeNameCol, allNeNamesStr_, ws, ws.range(sourceNeNameLetter & rowNumber))
            If allBoardStyleDataMap_.hasKey(sourceNeName) Then
                If sourceNeNameLetter <> "" Then sourceNeName = ws.range(sourceNeNameLetter & rowNumber).value
                
                Set boardStyleDataMap = allBoardStyleDataMap_.GetAt(sourceNeName)
                keyStr = sourceNeName + "_" + groupName + "_" + autocopyBoardNoName
                If boardStyleDataMap.hasKey(keyStr) Then
                    valueStr = boardStyleDataMap.GetAt(keyStr)
                    Call setBoardStyleListBoxRangeValidation(ws.name, groupName, autocopyBoardNoName, valueStr, ws, ws.range(columnNamePositionLetter & rowNumber))
                End If
            Else
                With ws.range(columnNamePositionLetter & rowNumber).Validation
                    .Delete
                End With
            End If
        Else
            With ws.range(columnNamePositionLetter & rowNumber).Validation
                .Delete
            End With
            
            With ws.range(sourceNeNameLetter & rowNumber).Validation
                .Delete
            End With
            
        End If
    Next rowNumber
End Sub

Private Sub Class_Initialize()
    allNeNamesStr_ = ""
    addDeleboxListFlag_ = False
    Set allBoardStyleDataMap_ = New CMapValueObject
End Sub




