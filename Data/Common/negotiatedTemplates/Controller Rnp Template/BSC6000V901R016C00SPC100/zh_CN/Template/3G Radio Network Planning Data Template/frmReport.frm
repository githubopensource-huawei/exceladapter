VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "UserForm2"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim isVisible As Boolean, bk2 As Workbook, sht2 As Worksheet
    Set bk2 = Application.ActiveWorkbook
    Set sht2 = Application.ActiveSheet
    isVisible = Application.Visible
    Application.Visible = True

    Dim i As Long, s As String, strFileName As String, wk As Workbook, sht As Worksheet
    strFileName = bk2.Path + "\" + "Report_" + Format(Date, "yyyy-mm-dd") + "_" + Format(Time, "hh-mm-ss")
    Set wk = Application.Workbooks.Add()
    Set sht = wk.Sheets(1)
    For i = 0 To Me.lstReportInfos.ListCount - 1
        s = Me.lstReportInfos.List(i, 0)
        sht.Cells(i + 1, 1).Value = s
    Next i
    wk.SaveAs FileName:=strFileName
    Application.StatusBar = FormatStr(RSC_STR_FILE_SAVED, strFileName)

    Application.Visible = isVisible
    sht2.Activate
    Me.Hide
    Me.Show
End Sub
