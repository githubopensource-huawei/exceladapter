VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIubMoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mocName_ As String '��IUBҳǩ�е�A�����ƣ���SheetDef�л�ȡ
Private mocSelectedFlag_ As String '�Ƿ�ѡ���ı�־��δѡ�����ַ���N����ʶ����SheetDef�еĶ���һ��
Private mocRowNumber_ As Long 'moc��������
Private mocStartRowNumber_ As Long '��IUBҳǩ����ʼ��
Private mocEndRowNumber_ As Long '��IUBҳǩ�Ľ�����

Public Sub init(ByRef mocName As String, ByRef selectedString As String, ByRef rowNumber As Long, ByRef startRow As Long, ByRef endRow As Long)
    mocName_ = mocName
    mocSelectedFlag_ = selectedString
    mocRowNumber_ = rowNumber
    mocStartRowNumber_ = startRow
    mocEndRowNumber_ = endRow
End Sub

Public Function isMocSelected() As Boolean
    If mocSelectedFlag_ = "N" Then
        isMocSelected = False
    Else
        isMocSelected = True
    End If
End Function

Public Function getRowNumber() As Long
    getRowNumber = mocRowNumber_
End Function

Property Get startRow() As Long
    startRow = mocStartRowNumber_
End Property

Property Get endRow() As Long
    endRow = mocEndRowNumber_
End Property

