VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FreqInfo 
   Caption         =   "FreqInfo"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3705
   OleObjectBlob   =   "FreqInfo.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FreqInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub OK_Click()
    Dim resultStr As String

    Dim idx As Integer
    For idx = 0 To Me.Controls.count - 1
        If Me.Controls(idx) = True Then
            resultStr = resultStr & Me.Controls(idx).Caption & ","
        End If
    Next
    
    If resultStr <> "" Then resultStr = Left(resultStr, Len(resultStr) - 1)
    g_CurrentRange.value = resultStr

    Unload Me
End Sub

Private Sub UserForm_Activate()
    Me.Caption = getResByKey("SelectFreqsToSplit")
    Me.OK.Caption = getResByKey("OK")
    Me.cancel.Caption = getResByKey("Cancel")
    
    Const minWidth As Integer = 190
    Const fixedHeight As Integer = 180
    Me.Width = minWidth
    Me.Height = fixedHeight
    
    Call getFreqInfoCol
    
    Dim freqInfo As String
    freqInfo = Trim(g_CurrentSheet.Cells(g_CurrentRange.row, g_FreqInfoCol))
    
    Dim freqs() As String
    Dim freqCounter As Integer
    freqs = Split(freqInfo, ",")
    freqCounter = UBound(freqs)
    
    Dim idx, pos As Integer
    pos = 10
    For idx = 0 To freqCounter
        If idx Mod 5 = 0 Then pos = 10
        
        With Me.Controls.Add("Forms.CheckBox.1", str(idx))
            .Caption = freqs(idx)
            .Left = 10 + (idx \ 5) * 50
            .Top = pos
        End With
        pos = pos + 20
    Next
    
    Dim computeWidth As Integer
    computeWidth = (freqCounter \ 5 + 1) * 50 + 20
    If computeWidth > minWidth Then Me.Width = computeWidth
End Sub

