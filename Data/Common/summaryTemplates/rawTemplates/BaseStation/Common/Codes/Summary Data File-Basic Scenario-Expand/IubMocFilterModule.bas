Attribute VB_Name = "IubMocFilterModule"
Option Explicit

Public iubMocManager As CIubMocManager '��������IUBҳǩMOC�Ĺ�����
Private nextStepFlag As Boolean '��ѡ��MOC��������Ƿ������һ���ı�־

Public Sub initIubMocManager()
    Set iubMocManager = New CIubMocManager
    Call iubMocManager.init
End Sub

'չʾ���幩�ͻ�ѡ�񣬽���MOC�Ƿ�ѡ���Ĳ���
Public Function getIubUnselectMocCollection(ByRef iubUnselectMocCollection As Collection) As Boolean
    Call displayMocFilterForm
    Set iubUnselectMocCollection = iubMocManager.getUnselectedMocCollection
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
    MocFilterForm.Show
End Sub
'
'Private Sub test()
'    Call displayMocFilterForm
'End Sub


