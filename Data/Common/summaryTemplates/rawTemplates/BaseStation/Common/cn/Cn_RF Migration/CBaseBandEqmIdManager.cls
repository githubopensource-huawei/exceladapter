VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBaseBandEqmIdManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Delimeter_ As String = ";"
Private Const EqmIdTypeDelimter_ As String = "_"

Private Const BasebandEqmIdAttributeName_ As String = "BASEBANDEQMID" '�����豸����������ƣ�������������ColumnName
Private basebandEqmIdColummName_ As String 'BASEBANDEQMID��Ӧ������
Private basebandEqmIdColumnLetter_ As String 'BASEBANDEQMID���ڵ��к�

Private Const BasebandEqmTypeAttributeName_ As String = "BASEBANDEQMTYPE" '�����豸���Ͷ����ƣ�������������ColumnName
Private basebandEqmTypeColumnName_ As String 'BASEBANDEQMTYPE��Ӧ������������Ҫ������ֵ��EqmIdƴ��������ΪΨһ���ֶ�
Private basebandEqmTypeColumnLetter_ As String 'BASEBANDEQMTYPE��Ӧ���к�

Private Const BasebandEqmBoardNoAttributeName_ As String = "BasebandeqmBoardNo" '�������������� �������ƣ�������������ColumnName
Private basebandEqmBoardNoColummName_ As String 'BasebandeqmBoardNo��Ӧ������
Private basebandEqmBoardNoColumnLetter_ As String 'BasebandeqmBoardNo���ڵ��к�


Private basebandEqmIdToBoardNoColMap_ As CMapValueObject '�����豸���ֵ������豸����ƴ�ӵ��ַ��� �� ���õ������ַ��� ��ӳ�䣬��0_UL-0_0_1_0;0_0_1_2

Private basebandBoardStyleMappingDefData_ As CBoardStyleMappingDefData '�����豸������

Private startRowNumber_ As Long 'BASEBANDEQM������ʼ��
Private endRowNumber_ As Long 'BASEBANDEQM���������

Private allReferenceBoardNoCol_ As Collection '���п����õĵ�����������key-value���ǵ�����ֵ


Public Function init(ByRef startRowNumber As Long, ByRef endRowNumber As Long, Optional ByRef basebandBoardStyleMappingDefData As CBoardStyleMappingDefData) As Boolean
    init = True
    startRowNumber_ = startRowNumber
    endRowNumber_ = endRowNumber
    
    If Not basebandBoardStyleMappingDefData Is Nothing Then
        Set basebandBoardStyleMappingDefData_ = basebandBoardStyleMappingDefData
        '�õ�BASEBANDEQMID��BasebandeqmBoardNo���������к�
        Call getColumnNameAndColumnLetterByAttributeName(basebandBoardStyleMappingDefData_, BasebandEqmIdAttributeName_, basebandEqmIdColummName_, basebandEqmIdColumnLetter_)
        Call getColumnNameAndColumnLetterByAttributeName(basebandBoardStyleMappingDefData_, BasebandEqmBoardNoAttributeName_, basebandEqmBoardNoColummName_, basebandEqmBoardNoColumnLetter_)
        Call getColumnNameAndColumnLetterByAttributeName(basebandBoardStyleMappingDefData_, BasebandEqmTypeAttributeName_, basebandEqmTypeColumnName_, basebandEqmTypeColumnLetter_)
    End If
    
    '���û���ҵ�EqmId�����������У����˳�
    If basebandEqmIdColummName_ = "" Or basebandEqmBoardNoColummName_ = "" Or basebandEqmBoardNoColummName_ = "" Then
        Call MsgBox(getResByKey("EqmIdOrBoardNoNotFound"), vbInformation + vbOKOnly, getResByKey("ErrorInfo"))
        init = False
        Exit Function
    End If
    '����EqmId��BoardNoCol���б�
    If makeBasebandEqmIdToBoardNoColMap = False Then
        init = False
        Exit Function
    End If
    
    '�����п����õĵĵ�����������allReferenceBoardNoCol_
    Call makeAllReferencedCol
End Function

'�����п����õĵ������ַ�����ȡ
Private Sub makeAllReferencedCol()
    Dim referencedString As String
    referencedString = basebandBoardStyleMappingDefData_.getMultiListReferenceValue(basebandEqmBoardNoColummName_)
    If referencedString = "" Then
        Exit Sub
    End If
    referencedString = getReferencedStringByInputString(referencedString)
    
    Set allReferenceBoardNoCol_ = New Collection
    '�ѵõ����п����õĵ�����ֵ�ָ���������
    Set allReferenceBoardNoCol_ = getKeyValueCollectionFromString(referencedString, ",")
End Sub

Private Function makeBasebandEqmIdToBoardNoColMap() As Boolean
    makeBasebandEqmIdToBoardNoColMap = True
    Dim rowIndex As Long
    Dim eachBasebandEqmId As String, eachReferenceBoardNoString As String, eachBasebandEqmType As String, key As String
    Set basebandEqmIdToBoardNoColMap_ = New CMapValueObject
    
    Dim referenceBoardNoCol As Collection
    '�ֱ�õ�ÿһ�е�BASEBANDEQMID��BasebandeqmBoardNo
    For rowIndex = startRowNumber_ + 2 To endRowNumber_
        eachBasebandEqmId = currentSheet.Range(basebandEqmIdColumnLetter_ & rowIndex).value
        eachBasebandEqmType = currentSheet.Range(basebandEqmTypeColumnLetter_ & rowIndex).value
        key = getKey(eachBasebandEqmId, eachBasebandEqmType) '��EqmId��EqmTypeƴ�ӳ������ַ�����ΪKeyֵ
        If basebandEqmIdToBoardNoColMap_.hasKey(key) Then
            Call MsgBox(getResByKey("SameEqmId"), vbInformation + vbOKOnly, getResByKey("ErrorInfo"))
            makeBasebandEqmIdToBoardNoColMap = False
            Exit Function
        End If
        eachReferenceBoardNoString = currentSheet.Range(basebandEqmBoardNoColumnLetter_ & rowIndex).value
        If eachBasebandEqmId <> "" Then
            '��BASEBANDEQMID��BasebandeqmBoardNo��ӳ�����
            Set referenceBoardNoCol = getKeyValueCollectionFromString(eachReferenceBoardNoString, Delimeter_)
            Call basebandEqmIdToBoardNoColMap_.SetAt(key, referenceBoardNoCol)
        End If
    Next rowIndex
End Function

Private Function getKey(ByRef eqmId As String, ByRef eqmType As String) As String
    If eqmType <> "" Then
        getKey = eqmId & EqmIdTypeDelimter_ & eqmType
    Else
        getKey = eqmId
    End If
End Function

'�õ������豸�������
Public Function getBasebandEqmIdCol() As Collection
    Set getBasebandEqmIdCol = basebandEqmIdToBoardNoColMap_.KeyCollection
End Function

'����һ���������Ƶõ������������ʾ���ƺ��к�
Private Sub getColumnNameAndColumnLetterByAttributeName(ByRef boardStyleMappingDefData As CBoardStyleMappingDefData, _
    ByRef attributeName As String, ByRef columnName As String, ByRef columnLetter As String)
    columnName = boardStyleMappingDefData.getColumnNameByAttributeName(attributeName)
    If columnName = "" Then
        columnLetter = ""
        Exit Sub
    End If
    columnLetter = boardStyleMappingDefData.getColumnNamePositionLetter(columnName)
End Sub

'��ɾһ��EqmId��Ӧ�����õ���������
Public Sub modifyReferencedBoardNoOfEqmId(ByRef eqmId As String, ByRef referencedBoardNo As String, ByRef addOrDelete As String)
    Dim referencedBoardNoCol As Collection
    
    If Not basebandEqmIdToBoardNoColMap_.hasKey(eqmId) Then Exit Sub '���û�����eqmId�����˳�
    Set referencedBoardNoCol = basebandEqmIdToBoardNoColMap_.GetAt(eqmId)
    
    If addOrDelete = "+" Then
        '����һ��BoardNo
        If Not Contains(referencedBoardNoCol, referencedBoardNo) Then
           referencedBoardNoCol.Add Item:=referencedBoardNo, key:=referencedBoardNo
        End If
    ElseIf addOrDelete = "-" Then
        'ɾ��һ��BoardNo
        If Contains(referencedBoardNoCol, referencedBoardNo) Then
            referencedBoardNoCol.Remove (referencedBoardNo)
        End If
    End If
    
    '����ɾ���������������ø���Ӧ��eqmId
    Call basebandEqmIdToBoardNoColMap_.SetAt(eqmId, referencedBoardNoCol)
End Sub

'����һ��EqmId���õ���ʣ��Ŀ�ѡ���õ����������͵�ǰ�����õ���������
Public Sub getOptionalAndCurrentBoardNoCols(ByRef eqmId As String, ByRef optionalCol As Collection, ByRef currentBoardNoCol As Collection)
    If Not basebandEqmIdToBoardNoColMap_.hasKey(eqmId) Then Exit Sub 'û�������ŵ�EqmId�����˳�
    
    Set currentBoardNoCol = basebandEqmIdToBoardNoColMap_.GetAt(eqmId)
    
    '�������п����õĵ����ţ�������ڵ�ǰ�����õ�boardNo�����У������optionalCol�У����ͻ�ѡ��
    Dim eachBoardNo As Variant
    For Each eachBoardNo In allReferenceBoardNoCol_
        If Not Contains(currentBoardNoCol, CStr(eachBoardNo)) Then
            optionalCol.Add Item:=eachBoardNo, key:=eachBoardNo
        End If
    Next eachBoardNo
End Sub

Public Sub writeNewBasebandEqmBoardNo()
    Dim rowIndex As Long
    Dim eachEqmId As String, eachEqmType As String, key As String
    Dim eachReferenceBoardNoCol As Collection
    Dim newBoardNoString As String
    
    For rowIndex = startRowNumber_ + 2 To endRowNumber_
        '�õ�ÿһ�е�EqmIdֵ
        eachEqmId = currentSheet.Range(basebandEqmIdColumnLetter_ & rowIndex).value
        eachEqmType = currentSheet.Range(basebandEqmTypeColumnLetter_ & rowIndex).value
        key = getKey(eachEqmId, eachEqmType)
        If eachEqmId = "" Or (Not basebandEqmIdToBoardNoColMap_.hasKey(key)) Then GoTo NextLoop
        
        '�õ���Ӧ��BoardNo�б�����
        Set eachReferenceBoardNoCol = basebandEqmIdToBoardNoColMap_.GetAt(key)
        
        '�õ�ƴ�ӵ�BoardNo�ַ�����д�������������������
        newBoardNoString = getConnectedStringFromCol(eachReferenceBoardNoCol, Delimeter_)
        currentSheet.Range(basebandEqmBoardNoColumnLetter_ & rowIndex).value = newBoardNoString
NextLoop:
    Next rowIndex
    
    '�����и�
    currentSheet.Rows(startRowNumber_ + 2 & ":" & endRowNumber_).AutoFit
End Sub