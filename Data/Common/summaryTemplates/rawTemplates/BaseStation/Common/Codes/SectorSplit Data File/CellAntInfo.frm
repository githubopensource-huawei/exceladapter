VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CellAntInfo 
   Caption         =   "Cell Ant"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2955
   OleObjectBlob   =   "CellAntInfo.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "CellAntInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    Me.Caption = getResByKey("SelectCellAntNo")
    Me.Button_OK.Caption = getResByKey("OK")
    Me.Button_Cancel.Caption = getResByKey("Cancel")
End Sub

Private Sub Button_Cancel_Click()
    Unload Me
End Sub

Private Sub Button_OK_Click()
    Dim result As String
    Dim pos As Integer
    
    Dim seperator As String
    seperator = ""
    
    Dim chkbox As Control
    For Each chkbox In Me.Controls
        If chkbox = True Then 'button = false
            result = result & seperator & chkbox.Caption
            seperator = ","
        End If
    Next
        
    If result <> "" Then
        g_CurrentRange.value = result
    End If
    
    Unload Me
End Sub
