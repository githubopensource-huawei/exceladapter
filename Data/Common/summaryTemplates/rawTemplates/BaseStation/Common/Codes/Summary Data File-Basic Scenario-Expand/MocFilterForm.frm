VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MocFilterForm 
   Caption         =   "Moc过滤选择"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   OleObjectBlob   =   "MocFilterForm.frx":0000
   StartUpPosition =   1  '所有者中心
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
        '如果没有可填加的，则直接退出
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '依次遍历所有ListIndex，如果选定了，则加入
            If .Selected(eachListIndex) = True Then
                mocName = .List(eachListIndex)
                '将新增的MOC加入到管理类的映射中
                Call iubMocManager.addMocToSelected(mocName)
            End If
        Next eachListIndex
    End With
    
    '重新显示
    Call initListboxAndButtons
End Sub

Private Sub AddAllCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.UnselectedMocListBox
        '如果没有可填加的，则直接退出
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '依次遍历所有ListIndex，加入
            mocName = .List(eachListIndex)
            '将新增的MOC加入到管理类的映射中
            Call iubMocManager.addMocToSelected(mocName)
        Next eachListIndex
    End With
    
    '重新显示
    Call initListboxAndButtons
End Sub

Private Sub DeleteCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.SelectedMocListBox
        '如果没有可填加的，则直接退出
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '依次遍历所有ListIndex，如果选定了，则加入
            If .Selected(eachListIndex) = True Then
                mocName = .List(eachListIndex)
                '将新增的MOC加入到管理类的映射中
                Call iubMocManager.addMocToUnselected(mocName)
            End If
        Next eachListIndex
    End With
    
    '重新显示
    Call initListboxAndButtons
End Sub

Private Sub DeleteAllCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.SelectedMocListBox
        '如果没有可填加的，则直接退出
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '依次遍历所有ListIndex，加入
            mocName = .List(eachListIndex)
            '将新增的MOC加入到管理类的映射中
            Call iubMocManager.addMocToUnselected(mocName)
        Next eachListIndex
    End With
    
    '重新显示
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
    '得到选定和未选定的moc容器
    Dim selectedMocCol As Collection, unselectedMocCol As Collection '选定MOC和未选定MOC容器
    Call iubMocManager.getCollections(selectedMocCol, unselectedMocCol)
    
    '将容器中的值加入到两个ListBox中
    Call addItemOfListBox(Me.SelectedMocListBox, selectedMocCol)
    Call addItemOfListBox(Me.UnselectedMocListBox, unselectedMocCol)
End Sub

'根据已选和未选的moc的ListBox进行增加删除按钮的灰化显示
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

'给一个ListBox传入一个Col，将Col中的内容添加到ListBox中
Private Sub addItemOfListBox(ByRef lb As Variant, ByRef col As Collection)
    '清空列表
    lb.Clear
    
    '如果没有容器值为空，则直接退出
    If col.count = 0 Then Exit Sub
        
    Dim eachItem As Variant
    For Each eachItem In col
        lb.AddItem eachItem
    Next eachItem
    
    lb.ListIndex = 0
End Sub

