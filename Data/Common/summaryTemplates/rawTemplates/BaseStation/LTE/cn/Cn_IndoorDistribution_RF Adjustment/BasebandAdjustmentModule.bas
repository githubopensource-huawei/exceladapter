Attribute VB_Name = "BasebandAdjustmentModule"
Option Explicit
Private Const BasebandEqmMocName As String = "BASEBANDEQM" '�����豸MOC����
Public Const BasebandReferenceBoardNoDelimeter As String = ";" '�������ŵķָ���

Private basebandBoardStyleMappingDefData As CBoardStyleMappingDefData '�����豸�������
Public basebandEqmIdManager As CBaseBandEqmIdManager '�����豸��ź͵�����б�Ĺ�����

'���������������Ŵ���
Public Sub AdjustBasebandEqmBoardNo()
    BasebandEqmAdjustmentForm.Show
End Sub

Public Function initCurrentSheetBaseBandEqmIdManager() As Boolean
    Dim flag As Boolean
    flag = True
    '�õ�BASEBANDEQM������
    Dim baseBandEqmGroupName As String
    baseBandEqmGroupName = boardStyleData.getGroupNamebyMocName(BasebandEqmMocName)
    
    '���û�л����豸��������ʾ���˳�
    If baseBandEqmGroupName = "" Then
        Call MsgBox(getResByKey("NoBasebandEqmObject"), vbInformation + vbOKOnly, getResByKey("ErrorInfo"))
        flag = False
        Exit Function
    End If
    
    If boardStyleMappingDefMap Is Nothing Then
        Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    End If
    
    '��ȡBaseBand�����
    Set basebandBoardStyleMappingDefData = boardStyleMappingDefMap.GetAt(baseBandEqmGroupName)
    
    '�õ�BaseBand�������ʼ�кͽ�����
    Dim startRow As Long, endRow As Long
    Call getGroupNameStartAndEndRowNumber(baseBandEqmGroupName, startRow, endRow)
    
    '��ʼ����Ź�����
    Set basebandEqmIdManager = New CBaseBandEqmIdManager
    '���EqmId�л�������û���ҵ������˳�
    If basebandEqmIdManager.init(startRow, endRow, basebandBoardStyleMappingDefData) = False Then
        flag = False
        Exit Function
    End If
    
    initCurrentSheetBaseBandEqmIdManager = flag
End Function
