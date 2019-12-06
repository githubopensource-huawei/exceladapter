VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CReferenceRangeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private referenceAddressBoardNoMap_ As CMap '引用的地址和单板编号的Map
Private boardNoReferenceAddressMap_ As CMap '单板编号和地址的Map
Private boardStyleMappingDefMap_ As CMapValueObject
Private groupNameWithListBoxColumnsMap_ As CMap '有下拉框的组名Map，key/value都是groupName

Private multiListReferenceAddressBoardNoMap_ As CMap '多List引用的地址和多List单板编号的Map
Private groupNameWithMultiListColumnsMap_ As CMap '有多List引用单板编号的组名Map，key/value都是GroupName，当前只有BASEBANDEQM

Private Sub init()
    If boardStyleData Is Nothing Then Call initBoardStyleMappingDataPublic
    Set boardStyleMappingDefMap_ = boardStyleData.getBoardStyleMappingDefMap
    
    Dim groupName As Variant
    For Each groupName In boardStyleMappingDefMap_.KeyCollection
        Call initGroupNameBoardStyleDataMap(CStr(groupName))
        'Call updateGroupNameBoardNoMap(CStr(groupName))
    Next groupName
End Sub

Public Sub setReferenceAddressHyperlinks(ByRef ws As Worksheet, ByRef boardNoManager As CBoardNoManager)
    Dim srcReferenceAddress As Variant '引用单板编号的地址
    'srcBoardNo为引用单板编号单元格的实际单板编号值，referencedBoardNoAddress为该单板编号所在地址
    Dim srcBoardNo As String, referencedBoardNoAddress As String
    For Each srcReferenceAddress In referenceAddressBoardNoMap_.KeyCollection
        srcBoardNo = referenceAddressBoardNoMap_.GetAt(srcReferenceAddress)
        referencedBoardNoAddress = boardNoManager.getBoardNoAddress(srcBoardNo)
        If referencedBoardNoAddress <> "" Then '如果找到了地址，则给地址添加超链接
            Call addHyperlink(ws, CStr(srcReferenceAddress), ws, referencedBoardNoAddress)
        Else '如果没找到对应单板编号，则将超链接删除
            Call deleteHyperlink(ws, CStr(srcReferenceAddress))
        End If
    Next srcReferenceAddress
End Sub

Public Function getReferenceAddressString(ByRef boardNo As String) As String
    If boardNoReferenceAddressMap_.hasKey(boardNo) Then
        getReferenceAddressString = boardNoReferenceAddressMap_.GetAt(boardNo)
    Else
        getReferenceAddressString = ""
    End If
End Function

Private Sub initGroupNameBoardStyleDataMap(ByRef groupName As String)
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
    If boardStyleMappingDefData.hasListBoxColumns Then
        Call groupNameWithListBoxColumnsMap_.SetAt(groupName, groupName)
    ElseIf boardStyleMappingDefData.hasMultiListReferenceColumns Then
        '把有多List的分组名加入到映射中
        Call groupNameWithMultiListColumnsMap_.SetAt(groupName, groupName)
    End If
End Sub

Public Sub generateBoardNoReferenceAddressMap()
    Call init
    
    Dim groupName As Variant
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    '引用下拉框的更新
    For Each groupName In groupNameWithListBoxColumnsMap_.KeyCollection
        Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
        Call generateEachColumnOfListBoxReference(CStr(groupName), boardStyleMappingDefData)
    Next groupName
    
    '多List引用单板编号的更新
    For Each groupName In groupNameWithMultiListColumnsMap_.KeyCollection
        Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
        Call generateEachColumnOfMultiListReference(CStr(groupName), boardStyleMappingDefData)
    Next groupName
End Sub

Private Sub generateEachColumnOfListBoxReference(ByRef groupName As String, ByRef boardStyleMappingDefData As CBoardStyleMappingDefData)
    '引用下拉框的更新
    Dim listBoxColumnsMap As CMap
    Set listBoxColumnsMap = boardStyleMappingDefData.getListBoxReferencKeyMap
    
    Dim startRowNumber As Long, endRowNumber As Long
    Call getGroupNameStartAndEndRowNumber(groupName, startRowNumber, endRowNumber)
    
    Dim columnName As Variant
    Dim columnLetter As String
    Dim rowIndex As Long
    Dim boardNo As String
    Dim boardNoRange As range
    
    For Each columnName In listBoxColumnsMap.KeyCollection
        columnLetter = boardStyleMappingDefData.getColumnNamePositionLetter(CStr(columnName))
        For rowIndex = startRowNumber + 2 To endRowNumber
            Set boardNoRange = g_CurrentSheet.range(columnLetter & rowIndex)
            boardNo = boardNoRange.value
            If boardNo <> "" Then
                Call referenceAddressBoardNoMap_.SetAt(boardNoRange.address(False, False), boardNo)
                Call updateBoardNoReferenceAddressMap(boardNoRange.address(False, False), boardNo)
            End If
        Next rowIndex
    Next columnName
End Sub

Private Sub generateEachColumnOfMultiListReference(ByRef groupName As String, ByRef boardStyleMappingDefData As CBoardStyleMappingDefData)
    '多List引用的更新
    Dim multiListColumnsMap As CMap
    Set multiListColumnsMap = boardStyleMappingDefData.getMultiListReferenceKeyMap
    
    Dim startRowNumber As Long, endRowNumber As Long
    Call getGroupNameStartAndEndRowNumber(groupName, startRowNumber, endRowNumber)
    
    Dim columnName As Variant
    Dim columnLetter As String
    Dim rowIndex As Long
    Dim boardNo As String
    Dim boardNoRange As range
    Dim boardNoCol As Variant, eachBoardNo As Variant
    
    For Each columnName In multiListColumnsMap.KeyCollection
        columnLetter = boardStyleMappingDefData.getColumnNamePositionLetter(CStr(columnName))
        For rowIndex = startRowNumber + 2 To endRowNumber
            Set boardNoRange = g_CurrentSheet.range(columnLetter & rowIndex)
            boardNo = boardNoRange.value
            If boardNo <> "" Then
                Call multiListReferenceAddressBoardNoMap_.SetAt(boardNoRange.address(False, False), boardNo)
                '将每一个单板编号字符串分隔出来再进行更新地址
                boardNoCol = Split(boardNo, BasebandReferenceBoardNoDelimeter)
                For Each eachBoardNo In boardNoCol
                    Call updateBoardNoReferenceAddressMap(boardNoRange.address(False, False), CStr(eachBoardNo))
                Next eachBoardNo
            End If
        Next rowIndex
    Next columnName
End Sub

'传入一个地址，看该地址所对应的单元格是否是多List引用的单元格
Public Function isMultiListReferenceAddress(ByRef address As String) As Boolean
    If multiListReferenceAddressBoardNoMap_.hasKey(address) Then
        isMultiListReferenceAddress = True
    Else
        isMultiListReferenceAddress = False
    End If
End Function

Private Sub updateBoardNoReferenceAddressMap(ByRef boardNoRange As String, ByRef boardNo As String)
    Dim addressString As String
    If boardNoReferenceAddressMap_.hasKey(boardNo) Then
        addressString = boardNoReferenceAddressMap_.GetAt(boardNo)
        Call addNewAddress(boardNoRange, addressString)
        Call boardNoReferenceAddressMap_.SetAt(boardNo, addressString)
    Else
        Call boardNoReferenceAddressMap_.SetAt(boardNo, boardNoRange)
    End If
End Sub

Private Sub addNewAddress(ByRef boardNoRange As String, ByRef addressString As String)
    If addressString = "" Then
        addressString = addressString & boardNoRange
    Else
        addressString = addressString & "," & boardNoRange
    End If
End Sub

Private Sub Class_Initialize()
    Set referenceAddressBoardNoMap_ = New CMap
    Set boardNoReferenceAddressMap_ = New CMap
    Set groupNameWithListBoxColumnsMap_ = New CMap
    Set multiListReferenceAddressBoardNoMap_ = New CMap
    Set groupNameWithMultiListColumnsMap_ = New CMap
End Sub
