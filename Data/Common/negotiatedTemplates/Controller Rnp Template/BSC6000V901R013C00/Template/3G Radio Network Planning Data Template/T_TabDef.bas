Attribute VB_Name = "T_TabDef"
Public Sub cmdGenNegotiatedFile_Click()
    Call T_DefineNegotiatedFile.GenNegotiatedFile
    'added by z00102652 at 2010-04-08, begin
    Call DoCmdSetCTClick

    Call DoCmdSetDFClick

    Call DoCmdExportCodeClick
    'added by z00102652 at 2010-04-08, end
End Sub

Private Sub DoCmdExportCodeClick()
    Call ExportCode
    MsgBox "Finished to export code."
End Sub

Private Sub DoCmdSetCTClick()
    Dim sht As Worksheet, sht2 As Worksheet
    Set sht2 = Application.ActiveSheet
    Set sht = Sheets(SHT_CONVERT_TEMPLATE)
    sht.Activate
    sht.Tab.ColorIndex = 6
    Call SetTemplate_CT
    MsgBox "Finished to set sheet '" + SHT_CONVERT_TEMPLATE + "'."

    sht2.Activate
End Sub

Private Sub DoCmdSetDFClick()
    Dim sht As Worksheet, sht2 As Worksheet
    Set sht2 = Application.ActiveSheet
    Set sht = Sheets(SHT_DOUBLE_FREQ_CELL_SETTING)
    sht.Visible = True
    sht.Activate

    Call SetTemplate_DF
    MsgBox "Finished to set sheet '" + SHT_DOUBLE_FREQ_CELL_SETTING + "'."

    sht.Visible = False
    sht2.Activate
End Sub

