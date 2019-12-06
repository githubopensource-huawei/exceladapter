VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Running,please wait..."
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12780
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim fullPercentBoxWidth As Integer


Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    Call initProgressBar
    fullPercentBoxWidth = Me.FullPercentBox.Width
ErrorHandler:
End Sub

Public Sub initProgressBar()
    updateProgress (0)
End Sub

'percent[0 - 100]
Public Function updateProgress(percent As Integer)
    Dim curProgress As Integer
    If percent > 0 Then
        curProgress = fullPercentBoxWidth * (percent / 100)
        Me.progressLabel.Width = curProgress
        Me.Caption = getResByKey("RunWait") & CStr(percent) & "%"
        'DoEvents
        Me.Repaint
        If percent = 100 Then Me.Hide
    End If
End Function

