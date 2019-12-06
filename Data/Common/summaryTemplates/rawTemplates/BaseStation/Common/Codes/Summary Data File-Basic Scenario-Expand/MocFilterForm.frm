VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MocFilterForm 
   Caption         =   "Moc����ѡ��"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   OleObjectBlob   =   "MocFilterForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "MocFilterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub AddCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.UnselectedMocListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex�����ѡ���ˣ������
            If .Selected(eachListIndex) = True Then
                mocName = .List(eachListIndex)
                '��������MOC���뵽�������ӳ����
                Call iubMocManager.addMocToSelected(mocName)
            End If
        Next eachListIndex
    End With
    
    '������ʾ
    Call initListboxAndButtons
End Sub

Private Sub AddAllCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.UnselectedMocListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex������
            mocName = .List(eachListIndex)
            '��������MOC���뵽�������ӳ����
            Call iubMocManager.addMocToSelected(mocName)
        Next eachListIndex
    End With
    
    '������ʾ
    Call initListboxAndButtons
End Sub

Private Sub DeleteCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.SelectedMocListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex�����ѡ���ˣ������
            If .Selected(eachListIndex) = True Then
                mocName = .List(eachListIndex)
                '��������MOC���뵽�������ӳ����
                Call iubMocManager.addMocToUnselected(mocName)
            End If
        Next eachListIndex
    End With
    
    '������ʾ
    Call initListboxAndButtons
End Sub

Private Sub DeleteAllCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.SelectedMocListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex������
            mocName = .List(eachListIndex)
            '��������MOC���뵽�������ӳ����
            Call iubMocManager.addMocToUnselected(mocName)
        Next eachListIndex
    End With
    
    '������ʾ
    Call initListboxAndButtons
End Sub

Private Sub CancelCommandButton_Click()
    Call setNextStepFlag(False)
    Unload Me
End Sub

Private Sub NextCommandButton_Click()
    Call setNextStepFlag(True)
    Call iubMocManager.setUnselectedMocFlag
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    Call initIubMocManager
    
    Call initListboxAndButtons
ErrorHandler:

End Sub

Private Sub initListboxAndButtons()
    Call initListBox
    Call controlAddDeleteButtons
End Sub

Private Sub initListBox()
    '�õ�ѡ����δѡ����moc����
    Dim selectedMocCol As Collection, unselectedMocCol As Collection 'ѡ��MOC��δѡ��MOC����
    Call iubMocManager.getCollections(selectedMocCol, unselectedMocCol)
    
    '�������е�ֵ���뵽����ListBox��
    Call addItemOfListBox(Me.SelectedMocListBox, selectedMocCol)
    Call addItemOfListBox(Me.UnselectedMocListBox, unselectedMocCol)
End Sub

'������ѡ��δѡ��moc��ListBox��������ɾ����ť�Ļһ���ʾ
Private Sub controlAddDeleteButtons()
    If Me.UnselectedMocListBox.ListCount = 0 Then
        Call setAddButtonsFlag(False)
    Else
        Call setAddButtonsFlag(True)
    End If
    
    If Me.SelectedMocListBox.ListCount = 0 Then
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
    
    lb.ListIndex = 0
End Sub

