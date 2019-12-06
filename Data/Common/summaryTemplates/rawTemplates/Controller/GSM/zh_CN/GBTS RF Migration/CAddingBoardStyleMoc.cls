VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAddingBoardStyleMoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private addBoardStyleButtons_ As CAddBoardStyleButtons

Public Sub init()

End Sub
'
'Private Sub setAddingBoardStyleMoi()
'    Dim button As CommandBarButton
'    Set button = addBoardStyleMoiBarCol("add")
'    button.Enabled = False
'    Set button = addBoardStyleMoiBarCol("finish")
'    button.Enabled = True
'    Set button = addBoardStyleMoiBarCol("cancel")
'    button.Enabled = True
'End Sub
'
'Private Sub resetAddingBoardStyleMoi()
'    Dim button As CommandBarButton
'    Set button = addBoardStyleMoiBarCol("add")
'    button.Enabled = True
'    Set button = addBoardStyleMoiBarCol("finish")
'    button.Enabled = False
'    Set button = addBoardStyleMoiBarCol("cancel")
'    button.Enabled = False
'End Sub

Private Sub Class_Initialize()
    inAddProcessFlag = True
    'Set addBoardStyleButtons_ = New CAddBoardStyleButtons
    If addBoardStyleButtons Is Nothing Then Call initAddBoardStyleButtonsPublic
    If deleteBoardStyleButtons Is Nothing Then Call initDeleteBoardStyleButtonsPublic
    
    Call addBoardStyleButtons.setAddBoardStyleButtons
    Call deleteBoardStyleButtons.setDeleteBoardStyleButtons
End Sub

Private Sub Class_Terminate()
    inAddProcessFlag = False
    Call addBoardStyleButtons.resetAddBoardStyleButtons
    Call deleteBoardStyleButtons.resetDeleteBoardStyleButtons
End Sub

