Attribute VB_Name = "IubMocFilterModule"
Option Explicit

Public btsNameManager As CBTSNameManager '��������IUBҳǩMOC�Ĺ�����
Private nextStepFlag As Boolean '��ѡ��MOC��������Ƿ������һ���ı�־

Public Sub initBTSNameManager()
    Set btsNameManager = New CBTSNameManager
    Call btsNameManager.init
End Sub

'չʾ���幩�ͻ�ѡ�񣬽���MOC�Ƿ�ѡ���Ĳ���
Public Function getIubUnselectMocCollection(ByRef iubUnselectMocCollection As Collection) As Boolean
    Call displayMocFilterForm
    Set iubUnselectMocCollection = btsNameManager.getUnselectedMocCollection
    '���ѡ����ȡ�����򲻽��к���ת��
    getIubUnselectMocCollection = nextStepFlag
End Function

Public Sub setNextStepFlag(ByRef flag As Boolean)
    nextStepFlag = flag
End Sub

Public Function getNextStepFlag()
    getNextStepFlag = nextStepFlag
End Function

Public Sub displayMocFilterForm()
    MuliBtsFilterForm.Show
    'LampsiteMuliBtsFilterForm.Show
End Sub
'
'Private Sub test()
'    Call displayMocFilterForm
'End Sub


