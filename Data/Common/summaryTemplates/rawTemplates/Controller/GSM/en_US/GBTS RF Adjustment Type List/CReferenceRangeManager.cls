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

Private referenceAddressBoardNoMap_ As CMap '���õĵ�ַ�͵����ŵ�Map
Private boardNoReferenceAddressMap_ As CMap '�����ź͵�ַ��Map
Private boardStyleMappingDefMap_ As CMapValueObject
Private groupNameWithListBoxColumnsMap_ As CMapValueObject '�������������Map��key/value����groupName

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
    Dim srcReferenceAddress As Variant '���õ����ŵĵ�ַ
    'srcBoardNoΪ���õ����ŵ�Ԫ���ʵ�ʵ�����ֵ��referencedBoardNoAddressΪ�õ��������ڵ�ַ
    Dim srcBoardNo As String, referencedBoardNoAddress As String
    For Each srcReferenceAddress In referenceAddressBoardNoMap_.KeyCollection
        srcBoardNo = referenceAddressBoardNoMap_.GetAt(srcReferenceAddress)
        referencedBoardNoAddress = boardNoManager.getBoardNoAddress(srcBoardNo)
        If referencedBoardNoAddress <> "" Then '����ҵ��˵�ַ�������ַ���ӳ�����
            Call addHyperlink(ws, CStr(srcReferenceAddress), ws, referencedBoardNoAddress)
        Else '���û�ҵ���Ӧ�����ţ��򽫳�����ɾ��
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
    End If
End Sub

Public Sub generateBoardNoReferenceAddressMap()
    Call init
    
    Dim groupName As Variant
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    For Each groupName In groupNameWithListBoxColumnsMap_.KeyCollection
        Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
        Call generateEachColumn(CStr(groupName), boardStyleMappingDefData)
    Next groupName
End Sub

Private Sub generateEachColumn(ByRef groupName As String, ByRef boardStyleMappingDefData As CBoardStyleMappingDefData)
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
            Set boardNoRange = currentSheet.range(columnLetter & rowIndex)
            boardNo = boardNoRange.value
            If boardNo <> "" Then
                Call referenceAddressBoardNoMap_.SetAt(boardNoRange.address(False, False), boardNo)
                Call updateBoardNoReferenceAddressMap(boardNoRange.address(False, False), boardNo)
            End If
        Next rowIndex
    Next columnName
End Sub

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
    Set groupNameWithListBoxColumnsMap_ = New CMapValueObject
End Sub