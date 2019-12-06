VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateCMForm 
   Caption         =   "Customize Template"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   OleObjectBlob   =   "TemplateCMForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "TemplateCMForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit
Public CheckReport As String

Private Sub CancelBtn_Click()
    TemplateCMForm.Hide
End Sub

Private Sub cbDev_Click()
    cbTRXBRD.Value = cbDev.Value
    cbRXU.Value = cbDev.Value
    cbANT.Value = cbDev.Value
    cbRETANT.Value = cbDev.Value
    
    cbTRXBRD.Enabled = cbDev.Value
    cbRXU.Enabled = cbDev.Value
    cbANT.Enabled = cbDev.Value
    cbRETANT.Enabled = cbDev.Value
    frmDev.Enabled = cbDev.Value
End Sub


Private Sub refreshTs()
    cbMONITORTS.Value = (cbIPOE.Value Or cbTDM.Value) And cbTimeslot.Value
    cbIdle.Value = (cbTDM.Value And cbTimeslot.Value)
    cbMONITORTS.Enabled = cbMONITORTS.Value
    cbIdle.Enabled = cbIdle.Value
End Sub

Private Sub cbGTRXAdvance_Click()

End Sub

Private Sub cbIPFE_Click()
    Call refreshTs
End Sub

Private Sub cbIPOE_Click()
    cbMP.Value = cbIPOE.Value
    cbPPP.Value = cbIPOE.Value
    cbMP.Enabled = cbIPOE.Value
    cbPPP.Enabled = cbIPOE.Value
    frmIPOE.Enabled = cbIPOE.Value
    Call refreshTs
End Sub



Private Sub cbDXX_Click()

End Sub

Private Sub cbTDM_Click()
    Call refreshTs
End Sub

Private Sub cbTimeslot_Click()
    Call refreshTs
    frmTS.Enabled = cbTimeslot.Value
End Sub

Private Sub cbTRXBRD_Click()
  cbANT.Value = cbTRXBRD.Value
  cbANT.Enabled = cbTRXBRD.Value
End Sub

Private Sub cbGbIP_Click()

End Sub

Private Sub OKBtn_Click()
  ThisWorkbook.Unprotect
  
  Dim WkSht As Worksheet
    
  Set WkSht = ThisWorkbook.ActiveSheet

  Sheets("ATERE1T1").Visible = cbATERTDM.Value
  Sheets("ATEROML").Visible = cbATERTDM.Value
  Sheets("ATERSL").Visible = cbATERTDM.Value
  Sheets("DEVIP").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value Or cbGbIP.Value)
  Sheets("PPPLNK").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("MPGRP").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("MPLNK").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("IPLOGICPORT").Visible = (cbAIP.Value Or cbATDMIP.Value)
  Sheets("ADJNODE").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("TRMMAP").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("TRMFACTOR").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("ADJMAP").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("IPPATH").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("IPMUX").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  'Sheets("IPPM").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("N7DPC").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbATDM.Value)
  Sheets("GCNNODE").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbATDM.Value)
  Sheets("NRIMSCMAP").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbATDM.Value)
  Sheets("AITFREV").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbATDM.Value)
  Sheets("RSCGRP").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
  Sheets("M3LE").Visible = (cbAIP.Value Or cbATDMIP.Value)
  
  Sheets("AE1T1").Visible = (cbATDM.Value Or cbATDMIP.Value)
  Sheets("MTP3LKS").Visible = (cbATDM.Value Or cbATDMIP.Value)
  Sheets("MTP3LNK").Visible = (cbATDM.Value Or cbATDMIP.Value)
  Sheets("MTP3RT").Visible = (cbATDM.Value Or cbATDMIP.Value)
  Sheets("MTP3TMR").Visible = (cbATDM.Value Or cbATDMIP.Value)
  Sheets("ETHPORT").Visible = (cbAIP.Value Or cbATDMIP.Value)
  Sheets("ETHIP").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbGbIP.Value)
  Sheets("ETHTRK").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbGbIP.Value)
  Sheets("ETHTRKLNK").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbGbIP.Value)
  Sheets("ETHTRKIP").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbGbIP.Value)
  Sheets("IPCHK").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbGbIP.Value)
  Sheets("SCTPLNK").Visible = (cbAIP.Value Or cbATDMIP.Value)
  Sheets("M3DE").Visible = (cbAIP.Value Or cbATDMIP.Value)
  Sheets("M3LKS").Visible = (cbAIP.Value Or cbATDMIP.Value)
  Sheets("M3RT").Visible = (cbAIP.Value Or cbATDMIP.Value)
  Sheets("M3LNK").Visible = (cbAIP.Value Or cbATDMIP.Value)
  Sheets("VLANID").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbGbIP.Value)
  Sheets("IPRT").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbGbIP.Value)
  Sheets("IPPATHBIND").Visible = (cbAIP.Value Or cbATDMIP.Value)
  Sheets("IPRTBIND").Visible = (cbAIP.Value Or cbATDMIP.Value Or cbGbIP.Value)
  Sheets("SGSN").Visible = (cbGbFR.Value Or cbGbIP.Value)
  Sheets("SGSNNODE").Visible = (cbGbFR.Value Or cbGbIP.Value)
  Sheets("NRISGSNMAP").Visible = (cbGbFR.Value Or cbGbIP.Value)
  Sheets("SGSNROUTEINFO").Visible = (cbGbFR.Value Or cbGbIP.Value)
  Sheets("NSE").Visible = (cbGbFR.Value Or cbGbIP.Value)
  Sheets("BC").Visible = cbGbFR.Value
  Sheets("NSVC").Visible = cbGbFR.Value
  Sheets("PTPBVC").Visible = (cbGbFR.Value Or cbGbIP.Value)
  Sheets("NSVLLOCAL").Visible = cbGbIP.Value
  Sheets("NSVLREMOTE").Visible = cbGbIP.Value

  
  
  If WkSht.Visible Then
    WkSht.Activate
  End If
  TemplateCMForm.Hide
  
  ThisWorkbook.Protect Structure:=True, Windows:=False
End Sub

Private Sub ToolFrame_Click()

End Sub

Private Sub UserForm_Click()

End Sub
