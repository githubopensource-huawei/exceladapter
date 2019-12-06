VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectFreq 
   Caption         =   "UserForm1"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   OleObjectBlob   =   "frmSelectFreq.frx":0000
   StartUpPosition =   1  '所有者中心
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmSelectFreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    Application.StatusBar = RSC_STR_CANCELLED
    End
End Sub

Private Sub cmdOK_Click()
    If cmbFreq.ListIndex < 0 Then
        MsgBox RSC_STR_SELECT_ONE_FREQ
        Exit Sub
    Else
        If Not GetSectorMappingCell(cmbFreq.Text) Then
            If g_df_emConfigureMode = cmSingleSelected Then
                Exit Sub
            End If
        End If
    End If

    frmSelectFreq.Hide
End Sub

Private Sub UserForm_Terminate()
    Application.StatusBar = RSC_STR_CANCELLED
    End
End Sub
