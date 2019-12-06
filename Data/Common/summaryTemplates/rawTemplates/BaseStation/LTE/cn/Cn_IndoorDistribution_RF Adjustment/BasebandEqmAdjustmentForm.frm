VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BasebandEqmAdjustmentForm 
   Caption         =   "调整基带处理单板编号"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   OleObjectBlob   =   "BasebandEqmAdjustmentForm.frx":0000
   StartUpPosition =   1  '所有者中心
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
        '如果没有可填加的，则直接退出
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '依次遍历所有ListIndex，如果选定了，则加入
            If .Selected(eachListIndex) = True Then
                newAddedReferenceBoardNo = .List(eachListIndex)
                '将新增的单板编号加入到管理类的映射中
                Call basebandEqmIdManager.modifyReferencedBoardNoOfEqmId(basebandEqmId, newAddedReferenceBoardNo, "+")
            End If
        Next eachListIndex
    End With
    
    '重新显示
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
        '如果没有可填加的，则直接退出
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '依次遍历所有ListIndex，全部加入
            newAddedReferenceBoardNo = .List(eachListIndex)
            '将新增的单板编号加入到管理类的映射中
            Call basebandEqmIdManager.modifyReferencedBoardNoOfEqmId(basebandEqmId, newAddedReferenceBoardNo, "+")
        Next eachListIndex
    End With
    
    '重新显示
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
        '如果没有可填加的，则直接退出
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '依次遍历所有ListIndex，如果选定了，则加入
            If .Selected(eachListIndex) = True Then
                deletedAddedReferenceBoardNo = .List(eachListIndex)
                '将要删除的单板编号加入到管理类的映射中
                Call basebandEqmIdManager.modifyReferencedBoardNoOfEqmId(basebandEqmId, deletedAddedReferenceBoardNo, "-")
            End If
        Next eachListIndex
    End With
    
    '重新显示
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
        '如果没有可填加的，则直接退出
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '依次遍历所有ListIndex，全部加入
            deletedAddedReferenceBoardNo = .List(eachListIndex)
            '将要删除的单板编号加入到管理类的映射中
            Call basebandEqmIdManager.modifyReferencedBoardNoOfEqmId(basebandEqmId, deletedAddedReferenceBoardNo, "-")
        Next eachListIndex
    End With
    
    '重新显示
    Call BasebandEqmIdComboBox_Change
End Sub

Private Sub FinishCommandButton_Click()
    basebandEqmIdManager.writeNewBasebandEqmBoardNo
    Unload Me
End Sub

Private Sub CancelCommandButton_Click()
    Unload Me
End Sub

'根据可选的BoardNo和当前的BoardNo个数进行增加删除按钮的显示与灰化
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
    '初始化管理类
    If initCurrentSheetBaseBandEqmIdManager = False Then
        Unload Me
        Exit Sub
    End If
    
    Dim basebandeqmIdCol As Collection
    Set basebandeqmIdCol = basebandEqmIdManager.getBasebandEqmIdCol
    '添加EqmId
    Call addItemOfListBox(Me.BasebandEqmIdComboBox, basebandeqmIdCol)
End Sub

Private Sub BasebandEqmIdComboBox_Change()
    Call initReferenceBoardNoListBox
    Call controlAddDeleteButtons
End Sub

Private Sub initReferenceBoardNoListBox()
    '先得到EqmId值
    Dim basebandEqmId As String
    basebandEqmId = Me.BasebandEqmIdComboBox.List(Me.BasebandEqmIdComboBox.listIndex)
    
    Dim optionalBoardNoCol As New Collection, currentBoardNoCol As New Collection
    Call basebandEqmIdManager.getOptionalAndCurrentBoardNoCols(basebandEqmId, optionalBoardNoCol, currentBoardNoCol)
    
    '添加可选的BoardNo
    Call addItemOfListBox(Me.OptionalBoardNoListBox, optionalBoardNoCol)
    '添加当前的BoardNo
    Call addItemOfListBox(Me.CurrentBoardNoListBox, currentBoardNoCol)
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
    
    lb.listIndex = 0
End Sub
