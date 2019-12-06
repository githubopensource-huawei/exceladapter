VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CApplicationFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub init()

End Sub

Private Sub changeFlags(ByRef flag As Boolean)
    Application.DisplayAlerts = flag
    Application.ScreenUpdating = flag
    Application.EnableEvents = flag '关闭事件的触发
End Sub

Private Sub Class_Initialize()
    Call changeFlags(False)
End Sub

Private Sub Class_Terminate()
    Call changeFlags(True)
End Sub
