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
Private groupNameWithListBoxColumnsMap_ As CMap '�������������Map��key/value����groupName

Private multiListReferenceAddressBoardNoMap_ As CMap '��List���õĵ�ַ�Ͷ�List�����ŵ�Map
Private groupNameWithMultiListColumnsMap_ As CMap '�ж�List���õ����ŵ�����Map��key/value����GroupName����ǰֻ��BASEBANDEQM

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
        
        '��������ú����ӵ����ݣ�������BoardNo�ķ�ʽ���������ӡ����������趨���ù�ϵʱ�趨
        If isReferenceValue(srcBoardNo) Then GoTo NextLoop
        
        If referencedBoardNoAddress <> "" Then '����ҵ��˵�ַ�������ַ���ӳ�����
            Call addHyperlink(ws, CStr(srcReferenceAddress), ws, referencedBoardNoAddress)
        Else '���û�ҵ���Ӧ�����ţ��򽫳�����ɾ��
            Call deleteHyperlink(ws, CStr(srcReferenceAddress))
        End If
NextLoop:
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
        '���ж�List�ķ��������뵽ӳ����
        Call groupNameWithMultiListColumnsMap_.SetAt(groupName, groupName)
    End If
End Sub

Public Sub generateBoardNoReferenceAddressMap()
    Call init
    
    Dim groupName As Variant
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    '����������ĸ���
    For Each groupName In groupNameWithListBoxColumnsMap_.KeyCollection
        Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
        Call generateEachColumnOfListBoxReference(CStr(groupName), boardStyleMappingDefData)
    Next groupName
    
    '��List���õ����ŵĸ���
    For Each groupName In groupNameWithMultiListColumnsMap_.KeyCollection
        Set boardStyleMappingDefData = boardStyleMappingDefMap_.GetAt(groupName)
        Call generateEachColumnOfMultiListReference(CStr(groupName), boardStyleMappingDefData)
    Next groupName
End Sub

Private Sub generateEachColumnOfListBoxReference(ByRef groupName As String, ByRef boardStyleMappingDefData As CBoardStyleMappingDefData)
    '����������ĸ���
    Dim listBoxColumnsMap As CMap
    Set listBoxColumnsMap = boardStyleMappingDefData.getListBoxReferencKeyMap
    
    Dim startRowNumber As Long, endRowNumber As Long
    Call getGroupNameStartAndEndRowNumber(groupName, startRowNumber, endRowNumber)
    
    Dim columnName As Variant
    Dim columnLetter As String
    Dim rowIndex As Long
    Dim boardNo As String
    Dim boardNoRange As Range
    
    For Each columnName In listBoxColumnsMap.KeyCollection
        If Not boardStyleMappingDefData.columnNamePositionLetterMap.hasKey(CStr(columnName)) Then GoTo NextLoop
        
        columnLetter = boardStyleMappingDefData.getColumnNamePositionLetter(CStr(columnName))
        For rowIndex = startRowNumber + 2 To endRowNumber
            Set boardNoRange = currentSheet.Range(columnLetter & rowIndex)
            boardNo = boardNoRange.value
            If boardNo <> "" Then
                Call referenceAddressBoardNoMap_.SetAt(boardNoRange.address(False, False), boardNo)
                Call updateBoardNoReferenceAddressMap(boardNoRange.address(False, False), boardNo)
            End If
        Next rowIndex
NextLoop:
    Next columnName
End Sub

Private Sub generateEachColumnOfMultiListReference(ByRef groupName As String, ByRef boardStyleMappingDefData As CBoardStyleMappingDefData)
    '��List���õĸ���
    Dim multiListColumnsMap As CMap
    Set multiListColumnsMap = boardStyleMappingDefData.getMultiListReferenceKeyMap
    
    Dim startRowNumber As Long, endRowNumber As Long
    Call getGroupNameStartAndEndRowNumber(groupName, startRowNumber, endRowNumber)
    
    Dim columnName As Variant
    Dim columnLetter As String
    Dim rowIndex As Long
    Dim boardNo As String
    Dim boardNoRange As Range
    Dim boardNoCol As Variant, eachBoardNo As Variant
    
    For Each columnName In multiListColumnsMap.KeyCollection
        If Not boardStyleMappingDefData.columnNamePositionLetterMap.hasKey(CStr(columnName)) Then GoTo NextLoop
        
        columnLetter = boardStyleMappingDefData.getColumnNamePositionLetter(CStr(columnName))
        For rowIndex = startRowNumber + 2 To endRowNumber
            Set boardNoRange = currentSheet.Range(columnLetter & rowIndex)
            boardNo = boardNoRange.value
            If boardNo <> "" Then
                Call multiListReferenceAddressBoardNoMap_.SetAt(boardNoRange.address(False, False), boardNo)
                '��ÿһ���������ַ����ָ������ٽ��и��µ�ַ
                boardNoCol = Split(boardNo, BasebandReferenceBoardNoDelimeter)
                For Each eachBoardNo In boardNoCol
                    Call updateBoardNoReferenceAddressMap(boardNoRange.address(False, False), CStr(eachBoardNo))
                Next eachBoardNo
            End If
        Next rowIndex
NextLoop:
    Next columnName
End Sub

'����һ����ַ�����õ�ַ����Ӧ�ĵ�Ԫ���Ƿ��Ƕ�List���õĵ�Ԫ��
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