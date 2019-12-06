VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BasebandEqmAdjustmentForm 
   Caption         =   "����������������"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   OleObjectBlob   =   "BasebandEqmAdjustmentForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "BasebandEqmAdjustmentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub AddCommandButton_Click()
    Dim basebandEqmId As String
    With Me.BasebandEqmIdComboBox
        If .ListCount = 0 Then Exit Sub
        basebandEqmId = .List(.listIndex)
    End With
    
    Dim eachListIndex As Long
    Dim newAddedReferenceBoardNo As String
    With Me.OptionalBoardNoListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex�����ѡ���ˣ������
            If .Selected(eachListIndex) = True Then
                newAddedReferenceBoardNo = .List(eachListIndex)
                '�������ĵ����ż��뵽�������ӳ����
                Call basebandEqmIdManager.modifyReferencedBoardNoOfEqmId(basebandEqmId, newAddedReferenceBoardNo, "+")
            End If
        Next eachListIndex
    End With
    
    '������ʾ
    Call BasebandEqmIdComboBox_Change
End Sub

Private Sub AddAllCommandButton_Click()
    Dim basebandEqmId As String
    With Me.BasebandEqmIdComboBox
        If .ListCount = 0 Then Exit Sub
        basebandEqmId = .List(.listIndex)
    End With
    
    Dim eachListIndex As Long
    Dim newAddedReferenceBoardNo As String
    With Me.OptionalBoardNoListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex��ȫ������
            newAddedReferenceBoardNo = .List(eachListIndex)
            '�������ĵ����ż��뵽�������ӳ����
            Call basebandEqmIdManager.modifyReferencedBoardNoOfEqmId(basebandEqmId, newAddedReferenceBoardNo, "+")
        Next eachListIndex
    End With
    
    '������ʾ
    Call BasebandEqmIdComboBox_Change
End Sub


Private Sub DeleteCommandButton_Click()
    Dim basebandEqmId As String
    With Me.BasebandEqmIdComboBox
        If .ListCount = 0 Then Exit Sub
        basebandEqmId = .List(.listIndex)
    End With
    
    Dim eachListIndex As Long
    Dim deletedAddedReferenceBoardNo As String
    With Me.CurrentBoardNoListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex�����ѡ���ˣ������
            If .Selected(eachListIndex) = True Then
                deletedAddedReferenceBoardNo = .List(eachListIndex)
                '��Ҫɾ���ĵ����ż��뵽�������ӳ����
                Call basebandEqmIdManager.modifyReferencedBoardNoOfEqmId(basebandEqmId, deletedAddedReferenceBoardNo, "-")
            End If
        Next eachListIndex
    End With
    
    '������ʾ
    Call BasebandEqmIdComboBox_Change
End Sub

Private Sub DeleteAllCommandButton_Click()
    Dim basebandEqmId As String
    With Me.BasebandEqmIdComboBox
        If .ListCount = 0 Then Exit Sub
        basebandEqmId = .List(.listIndex)
    End With
    
    Dim eachListIndex As Long
    Dim deletedAddedReferenceBoardNo As String
    With Me.CurrentBoardNoListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex��ȫ������
            deletedAddedReferenceBoardNo = .List(eachListIndex)
            '��Ҫɾ���ĵ����ż��뵽�������ӳ����
            Call basebandEqmIdManager.modifyReferencedBoardNoOfEqmId(basebandEqmId, deletedAddedReferenceBoardNo, "-")
        Next eachListIndex
    End With
    
    '������ʾ
    Call BasebandEqmIdComboBox_Change
End Sub

Private Sub FinishCommandButton_Click()
    basebandEqmIdManager.writeNewBasebandEqmBoardNo
    Unload Me
End Sub

Private Sub CancelCommandButton_Click()
    Unload Me
End Sub

'���ݿ�ѡ��BoardNo�͵�ǰ��BoardNo������������ɾ����ť����ʾ��һ�
Private Sub controlAddDeleteButtons()
    If Me.OptionalBoardNoListBox.ListCount = 0 Then
        Call setAddButtonsFlag(False)
    Else
        Call setAddButtonsFlag(True)
    End If
    
    If Me.CurrentBoardNoListBox.ListCount = 0 Then
        Call setDeleteButtonsFlag(False)
    Else
        Call setDeleteButtonsFlag(True)
    End If
End Sub

Private Sub setAddButtonsFlag(ByRef flag As Boolean)
    Me.AddCommandButton.Enabled = flag
    Me.AddAllCommandButton.Enabled = flag
End Sub

Private Sub setDeleteButtonsFlag(ByRef flag As Boolean)
    Me.DeleteCommandButton.Enabled = flag
    Me.DeleteAllCommandButton.Enabled = flag
End Sub

Private Sub UserForm_Activate()
    On Error GoTo ErrorHandler
    If boardStyleData Is Nothing Then
        Call initBoardStyleMappingDataPublic
    End If

    Call initBasebandEqmIdComboBox
    Exit Sub
ErrorHandler:
End Sub

Private Sub initBasebandEqmIdComboBox()
    '��ʼ��������
    If initCurrentSheetBaseBandEqmIdManager = False Then
        Unload Me
        Exit Sub
    End If
    
    Dim basebandeqmIdCol As Collection
    Set basebandeqmIdCol = basebandEqmIdManager.getBasebandEqmIdCol
    '���EqmId
    Call addItemOfListBox(Me.BasebandEqmIdComboBox, basebandeqmIdCol)
End Sub

Private Sub BasebandEqmIdComboBox_Change()
    Call initReferenceBoardNoListBox
    Call controlAddDeleteButtons
End Sub

Private Sub initReferenceBoardNoListBox()
    '�ȵõ�EqmIdֵ
    Dim basebandEqmId As String
    basebandEqmId = Me.BasebandEqmIdComboBox.List(Me.BasebandEqmIdComboBox.listIndex)
    
    Dim optionalBoardNoCol As New Collection, currentBoardNoCol As New Collection
    Call basebandEqmIdManager.getOptionalAndCurrentBoardNoCols(basebandEqmId, optionalBoardNoCol, currentBoardNoCol)
    
    '��ӿ�ѡ��BoardNo
    Call addItemOfListBox(Me.OptionalBoardNoListBox, optionalBoardNoCol)
    '��ӵ�ǰ��BoardNo
    Call addItemOfListBox(Me.CurrentBoardNoListBox, currentBoardNoCol)
End Sub

'��һ��ListBox����һ��Col����Col�е�������ӵ�ListBox��
Private Sub addItemOfListBox(ByRef lb As Variant, ByRef col As Collection)
    '����б�
    lb.Clear
    
    '���û������ֵΪ�գ���ֱ���˳�
    If col.count = 0 Then Exit Sub
        
    Dim eachItem As Variant
    For Each eachItem In col
        lb.AddItem eachItem
    Next eachItem
    
    lb.listIndex = 0
End Sub
