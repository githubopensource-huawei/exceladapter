VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SectorInfo 
   Caption         =   "SectorInfo"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3705
   OleObjectBlob   =   "SectorInfo.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "SectorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub OK_Click()
    Dim resultStr As String
    Call getSectorInfoCol
    
    resultStr = g_CurrentSheet.Cells(g_CurrentRange.row, g_SectorInfoCol) & "("
    
    Dim idx As Integer
    For idx = 0 To Me.Controls.count - 1
        If InStr(Me.Controls(idx).Caption, getResByKey("Sector")) <> 0 Then
            If Me.Controls(idx) = True Then
                resultStr = resultStr & "Y"
            ElseIf Me.Controls(idx) = False Then
                resultStr = resultStr & "N"
            End If
        End If
    Next
    
    resultStr = resultStr & ")"
    g_CurrentRange.value = resultStr
    
    Unload Me
End Sub

Private Sub UserForm_Activate()
    Me.Caption = getResByKey("SelectSectorsToSplit")
    Me.OK.Caption = getResByKey("OK")
    Me.cancel.Caption = getResByKey("Cancel")
    
    Const minWidth As Integer = 190
    Const fixedHeight As Integer = 180
    Me.Width = minWidth
    Me.Height = fixedHeight

    Call getSectorInfoCol
    
    Dim SectorInfo As String
    SectorInfo = Trim(g_CurrentSheet.Cells(g_CurrentRange.row, g_SectorInfoCol))
    
    Dim sectorCounter As Integer
    sectorCounter = UBound(Split(SectorInfo, "/"))
    
    Dim idx, pos As Integer
    pos = 10
    For idx = 0 To sectorCounter
        If idx Mod 5 = 0 Then pos = 10
        
        With Me.Controls.Add("Forms.CheckBox.1", str(idx))
            .Caption = getResByKey("sector") & str(idx + 1)
            .Left = 10 + (idx \ 5) * 60
            .Top = pos
        End With
        
        pos = pos + 20
    Next
    
    Dim computeWidth As Integer
    computeWidth = (sectorCounter \ 5 + 1) * 60 + 20
    If computeWidth > minWidth Then Me.Width = computeWidth
End Sub
