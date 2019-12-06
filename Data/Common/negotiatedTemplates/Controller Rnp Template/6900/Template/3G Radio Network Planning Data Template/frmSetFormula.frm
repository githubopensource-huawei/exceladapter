VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetFormula 
   Caption         =   "Caculate Formula"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   OleObjectBlob   =   "frmSetFormula.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmSetFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim s As String, r As Range
    Set r = Sheets(MOC_DOUBLE_FREQ_CELL).Range("C3")
    If Not r.Comment Is Nothing Then
        s = r.Comment.Text
    Else
        s = ""
    End If
    Dim i As Integer
    i = InStr(s, Chr(10) + RSC_STR_FORMULA)
    If i > 0 Then
        s = Mid(s, 1, i - 1)
    End If
    s = s + Chr(10) + RSC_STR_FORMULA + Chr(10) + txtFormula.Text
    r.ClearComments
    r.AddComment s
    Unload Me
End Sub

Private Sub cmdTest_Click()
    txtResult.Text = GetSectorID(txtInputValue.Text, txtFormula.Text)
End Sub

Private Sub UserForm_Terminate()
    End
End Sub
