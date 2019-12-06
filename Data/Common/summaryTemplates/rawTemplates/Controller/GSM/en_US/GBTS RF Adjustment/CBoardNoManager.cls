VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBoardNoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private boardNoMap_ As CMap 'groupName为Key，BoardNoString为Value
Private boardNoAddressMap_ As CMap '单板编号为key，该单板编号所在的地址为Value
Private boardStyleMappingDefMap_ As CMapValueObject
Private Const delimiter As String = ","
Private Const MaxBoardNoIndex_ As Long = 1000

Private Sub init()
    Set boardNoMap_ = New CMap
    Set boardNoAddressMap_ = New CMap
    If boardStyleData Is Nothing Then Call initBoardStyleMappingDataPublic
    Set boardStyleMappingDefMap_ = boardStyleData.getBoardStyleMappingDefMap
    
    Dim groupName As Variant
    For Each groupName In boardStyleMappingDefMap_.KeyCollection
        Call initGroupNameBoardNoMap(CStr(groupName))
        'Call updateGroupNameBoardNoMap(CStr(groupName))
    Next groupName
End Sub

Public Function getBoardNoColumnLetterByGroupName(ByRef groupName As String) As String
    Dim boardStyleMappindDefData As CBoardStyleMappingDefData
    Dim boardNo As String
    If boardNoMap_.hasKey(groupName) Then
        Set boardStyleMappindDefData = boardStyleMappingDefMap_.GetAt(groupName)
        boardNo = boardStyleMappindDefData.autoFillInColumnName
        getBoardNoColumnLetterByGroupName = boardStyleMappindDefData.getColumnNamePositionLetter(boardNo)
    Else
        getBoardNoColumnLetterByGroupName = ""
    End If
End Function

Public Function getBoardNoStringByGroupName(ByRef groupName As String) As String
    Dim boardNoString As String
    boardNoString = ""
    If boardNoMap_.hasKey(groupName) Then
        boardNoString = boardNoMap_.GetAt(groupName)
    End If
    getBoardNoStringByGroupName = boardNoString
End Function

Private Sub initGroupNameBoardNoMap(ByRef groupName As String)
    Dim autoFillInColumnName As String
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    
    Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
    autoFillInColumnName = boardStyleMappingDefData.autoFillInColumnName
    
    If autoFillInColumnName <> "" Then
        Call boardNoMap_.SetAt(groupName, "")
    End If
End Sub

Public Sub generateCurrentGroupNameBoardNoMap()
    'If boardNoMap_ Is Nothing Then Call init
    Call init '每次都init,以防已有的单板编号被清空后该单板留在不合理数据
    Dim groupName As Variant
    For Each groupName In boardNoMap_.KeyCollection
        Call updateGroupNameBoardNoMap(CStr(groupName))
    Next groupName
End Sub

Private Sub updateGroupNameBoardNoMap(ByRef groupName As String)
    Dim autoFillInColumnName As String
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    Dim boardNoColumLetter As String
    Dim boardNoString As String
    
    Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
    autoFillInColumnName = boardStyleMappingDefData.autoFillInColumnName
    
    If autoFillInColumnName <> "" Then
        Call getGroupNameStartAndEndRowNumber(groupName, groupNameStartRowNumber, groupNameEndRowNumber)
        boardNoColumLetter = boardStyleMappingDefData.getColumnNamePositionLetter(autoFillInColumnName)
        boardNoString = getBoardNoString(boardNoColumLetter, groupNameStartRowNumber + 2, groupNameEndRowNumber)
        Call boardNoMap_.SetAt(groupName, boardNoString)
    End If
End Sub

Private Function getBoardNoString(ByVal columnletter As String, ByVal startRowNumber As Long, ByVal endRowNumber As Long)
    Dim rowNumber As Long
    Dim boardNo As String, eachBoardNo As String
    Dim boardNoRange As range
    boardNo = ""
    For rowNumber = startRowNumber To endRowNumber
        Set boardNoRange = currentSheet.range(columnletter & rowNumber)
        eachBoardNo = Trim(boardNoRange.value)
        If eachBoardNo <> "" Then
            boardNo = boardNo & eachBoardNo & delimiter
            Call boardNoAddressMap_.SetAt(eachBoardNo, boardNoRange.address(False, False)) '将单板编号和对应地址更新进Map中
        End If
    Next rowNumber
    If boardNo <> "" Then boardNo = Left(boardNo, Len(boardNo) - 1)
    getBoardNoString = boardNo
End Function

Public Sub getNewBoardNo(ByRef inputboardNo As String)
    Dim boardNoStringCollection As Collection
    Set boardNoStringCollection = boardNoMap_.ValueCollection
    Dim boardNoString As Variant
    Dim boardNoIndexUsableFlag As Boolean
    Dim tempBoardNo As String
    Dim boardNoIndex As Long
    
    For boardNoIndex = 1 To MaxBoardNoIndex_
        boardNoIndexUsableFlag = True
        tempBoardNo = inputboardNo & boardNoIndex
        For Each boardNoString In boardNoStringCollection
            If InStr(boardNoString, tempBoardNo) <> 0 Then
                'maxBoardNoIndex = getMaxBoardNoIndex(CStr(boardNoString), tempBoardNo, maxBoardNoIndex)
                boardNoIndexUsableFlag = False
                Exit For
            End If
        Next boardNoString
        
        If boardNoIndexUsableFlag = True Then
            inputboardNo = tempBoardNo
            Exit Sub
        End If
    Next boardNoIndex
End Sub

Public Sub addNewBoardNo(ByRef groupName As String, ByRef newBoardNo As String)
    Dim boardNoString As String
    boardNoString = boardNoMap_.GetAt(groupName)
    If boardNoString = "" Then
        boardNoString = boardNoString & newBoardNo
    Else
        boardNoString = boardNoString & delimiter & newBoardNo
    End If
    'Call eraseLastChar(boardNoString)
    Call boardNoMap_.SetAt(groupName, boardNoString)
End Sub

'得到单板编号所在单元格地址
Public Function getBoardNoAddress(ByRef boardNo As String) As String
    getBoardNoAddress = ""
    If boardNoAddressMap_.hasKey(boardNo) Then
        getBoardNoAddress = boardNoAddressMap_.GetAt(boardNo)
    End If
End Function
