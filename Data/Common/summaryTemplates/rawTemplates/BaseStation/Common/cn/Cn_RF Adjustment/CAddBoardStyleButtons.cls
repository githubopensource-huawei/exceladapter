VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAddBoardStyleButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private addBoardStyleMoiBarCol_ As Collection
Private addBoardStyleMoiButton_ As CommandBarButton
Private addBoardStyleMoiFinishButton_ As CommandBarButton
Private addBoardStyleMoiCancelButton_ As CommandBarButton

Public Sub init()
    Dim capacityExpansionExist As Boolean
    capacityExpansionExist = containsAToolBar(CapactiyExpansionBarNameAddMoi)
    Dim bar As CommandBar
'    capacityExpansionExist = False
'    Dim bar As CommandBar
'    For Each bar In CommandBars
'        If bar.name = CapactiyExpansionBarNameAddMoi Then
'            capacityExpansionExist = True
'            Exit For
'        End If
'    Next
    If capacityExpansionExist Then
        Set bar = CommandBars(CapactiyExpansionBarNameAddMoi)
        Call initAddBoardStyleButtons(bar.Controls(1), bar.Controls(2), bar.Controls(3))
    End If
End Sub

Public Sub initAddBoardStyleButtons(ByRef AddButton As CommandBarButton, ByRef FinishButton As CommandBarButton, ByRef CancelButton As CommandBarButton)
    Set addBoardStyleMoiButton_ = AddButton
    Set addBoardStyleMoiFinishButton_ = FinishButton
    Set addBoardStyleMoiCancelButton_ = CancelButton
End Sub

Public Sub setAddBoardStyleButtons()
    addBoardStyleMoiButton_.Enabled = False
    addBoardStyleMoiFinishButton_.Enabled = True
    addBoardStyleMoiCancelButton_.Enabled = True
End Sub

Public Sub resetAddBoardStyleButtons()
    addBoardStyleMoiButton_.Enabled = True
    addBoardStyleMoiFinishButton_.Enabled = False
    addBoardStyleMoiCancelButton_.Enabled = False
End Sub
