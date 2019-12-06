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

Public Sub init()
    Dim capacityExpansionExist As Boolean
    capacityExpansionExist = containsAToolBar(CapactiyExpansionBarNameDeleteMoi)
    Dim bar As CommandBar
    If capacityExpansionExist Then
        Set bar = CommandBars(CapactiyExpansionBarNameDeleteMoi)
        Call initDeleteBoardStyleButtons(bar.Controls(1))
    End If
End Sub

Public Sub initDeleteBoardStyleButtons(ByRef deleteButton As CommandBarButton)
    Set deleteBoardStyleMoiButton_ = deleteButton
End Sub

Public Sub setDeleteBoardStyleButtons()
    deleteBoardStyleMoiButton_.Enabled = False
End Sub

Public Sub resetDeleteBoardStyleButtons()
    deleteBoardStyleMoiButton_.Enabled = True
End Sub
