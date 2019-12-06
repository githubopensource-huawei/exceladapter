VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDeleteBoardStyleButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private deleteBoardStyleMoiButton_ As CommandBarButton
Private addBoardStyleMoiRefButton_ As CommandBarButton
Private delBoardStyleMoiRefButton_ As CommandBarButton

Public Sub init()
    Dim capacityExpansionExist As Boolean
    capacityExpansionExist = containsAToolBar(CapactiyExpansionBarNameDeleteMoi)
    Dim bar As CommandBar
    If capacityExpansionExist Then
        Set bar = CommandBars(CapactiyExpansionBarNameDeleteMoi)
        '1是删除行按钮,2是调整调整基带处理单板编号按钮,3是添加引用按钮,4是删除引用按钮.
        Call initDeleteBoardStyleButtons(bar.Controls(1), bar.Controls(3), bar.Controls(4))
    End If
End Sub

Public Sub initDeleteBoardStyleButtons(ByRef deleteButton As CommandBarButton, ByRef addRefButton As CommandBarButton, ByRef delRefButton As CommandBarButton)
    Set deleteBoardStyleMoiButton_ = deleteButton
    Set addBoardStyleMoiRefButton_ = addRefButton
    Set delBoardStyleMoiRefButton_ = delRefButton
End Sub

Public Sub setDeleteBoardStyleButtons()
    deleteBoardStyleMoiButton_.Enabled = False
    addBoardStyleMoiRefButton_.Enabled = False
    delBoardStyleMoiRefButton_.Enabled = False
End Sub

Public Sub resetDeleteBoardStyleButtons()
    deleteBoardStyleMoiButton_.Enabled = True
    addBoardStyleMoiRefButton_.Enabled = True
    delBoardStyleMoiRefButton_.Enabled = True
End Sub
