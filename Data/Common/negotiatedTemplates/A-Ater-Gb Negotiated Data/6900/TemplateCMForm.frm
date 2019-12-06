VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateCMForm 
   Caption         =   "Customize Template"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   OleObjectBlob   =   "TemplateCMForm.frx":0000
   StartUpPosition =   1  'CenterOwner
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
    cbTRXBRD.value = cbDev.value
    cbRXU.value = cbDev.value
    cbANT.value = cbDev.value
    cbRETANT.value = cbDev.value
    
    cbTRXBRD.Enabled = cbDev.value
    cbRXU.Enabled = cbDev.value
    cbANT.Enabled = cbDev.value
    cbRETANT.Enabled = cbDev.value
    frmDev.Enabled = cbDev.value
End Sub


Private Sub refreshTs()
    cbMONITORTS.value = (cbIPOE.value Or cbTDM.value) And cbTimeslot.value
    cbIdle.value = (cbTDM.value And cbTimeslot.value)
    cbMONITORTS.Enabled = cbMONITORTS.value
    cbIdle.Enabled = cbIdle.value
End Sub

Private Sub cbGTRXAdvance_Click()

End Sub

Private Sub cbIPFE_Click()
    Call refreshTs
End Sub

Private Sub cbIPOE_Click()
    cbMP.value = cbIPOE.value
    cbPPP.value = cbIPOE.value
    cbMP.Enabled = cbIPOE.value
    cbPPP.Enabled = cbIPOE.value
    frmIPOE.Enabled = cbIPOE.value
    Call refreshTs
End Sub



Private Sub cbDXX_Click()

End Sub

Private Sub cbTDM_Click()
    Call refreshTs
End Sub

Private Sub cbTimeslot_Click()
    Call refreshTs
    frmTS.Enabled = cbTimeslot.value
End Sub

Private Sub cbTRXBRD_Click()
  cbANT.value = cbTRXBRD.value
  cbANT.Enabled = cbTRXBRD.value
End Sub

Private Sub cbGbIP_Click()

End Sub

Private Sub OKBtn_Click()
    ThisWorkbook.Unprotect
  
    ThisWorkbook.Sheets("category").Cells(1, 2).value = TemplateCMForm.cbATDM.value
    ThisWorkbook.Sheets("category").Cells(2, 2).value = TemplateCMForm.cbAIP.value
    ThisWorkbook.Sheets("category").Cells(3, 2).value = TemplateCMForm.cbATDMIP.value
    ThisWorkbook.Sheets("category").Cells(4, 2).value = TemplateCMForm.cbATERTDM.value
    ThisWorkbook.Sheets("category").Cells(5, 2).value = TemplateCMForm.cbATERIP.value
    ThisWorkbook.Sheets("category").Cells(6, 2).value = TemplateCMForm.cbGbFR.value
    ThisWorkbook.Sheets("category").Cells(7, 2).value = TemplateCMForm.cbGbIP.value
    
    Dim WkSht As Worksheet
    
    Set WkSht = ThisWorkbook.ActiveSheet
    
    
    Sheets("ATERE1T1").Visible = cbATERTDM.value
    Sheets("ATEROML").Visible = cbATERTDM.value
    Sheets("ATERSL").Visible = cbATERTDM.value
    Sheets("DEVIP").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value Or cbGbIP.value)
    Sheets("PPPLNK").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("MPGRP").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("MPLNK").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("IPLOGICPORT").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("ADJNODE").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("ADJNODEDIP").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("TRMMAP").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("TRMFACTOR").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("ADJMAP").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("IPPATH").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("IPMUX").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    'Sheets("IPPM").Visible = (cbATERIP.Value Or cbAIP.Value Or cbATDMIP.Value)
    Sheets("N7DPC").Visible = (cbAIP.value Or cbATDMIP.value Or cbATDM.value)
    Sheets("GCNNODE").Visible = (cbAIP.value Or cbATDMIP.value Or cbATDM.value)
    Sheets("NRIMSCMAP").Visible = (cbAIP.value Or cbATDMIP.value Or cbATDM.value)
    Sheets("AITFREV").Visible = (cbAIP.value Or cbATDMIP.value Or cbATDM.value)
    Sheets("RSCGRP").Visible = (cbATERIP.value Or cbAIP.value Or cbATDMIP.value)
    Sheets("M3LE").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("IPPOOL").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("IPPOOLIP").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("SRCIPRT").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("IPPOOLMUX").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("AE1T1").Visible = (cbATDM.value Or cbATDMIP.value)
    Sheets("MTP3LKS").Visible = (cbATDM.value Or cbATDMIP.value)
    Sheets("MTP3LNK").Visible = (cbATDM.value Or cbATDMIP.value)
    Sheets("MTP3RT").Visible = (cbATDM.value Or cbATDMIP.value)
    Sheets("MTP3TMR").Visible = (cbATDM.value Or cbATDMIP.value)
    Sheets("ETHPORT").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("ETHIP").Visible = (cbAIP.value Or cbATDMIP.value Or cbGbIP.value)
    Sheets("ETHTRK").Visible = (cbAIP.value Or cbATDMIP.value Or cbGbIP.value)
    Sheets("ETHTRKLNK").Visible = (cbAIP.value Or cbATDMIP.value Or cbGbIP.value)
    Sheets("ETHTRKIP").Visible = (cbAIP.value Or cbATDMIP.value Or cbGbIP.value)
    Sheets("IPCHK").Visible = (cbAIP.value Or cbATDMIP.value Or cbGbIP.value)
    Sheets("SCTPLNK").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("M3DE").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("M3LKS").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("M3RT").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("M3LNK").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("VLANID").Visible = (cbAIP.value Or cbATDMIP.value Or cbGbIP.value)
    Sheets("IPRT").Visible = (cbAIP.value Or cbATDMIP.value Or cbGbIP.value)
    Sheets("IPPATHBIND").Visible = (cbAIP.value Or cbATDMIP.value)
    Sheets("IPRTBIND").Visible = (cbAIP.value Or cbATDMIP.value Or cbGbIP.value)
    Sheets("SGSN").Visible = (cbGbFR.value Or cbGbIP.value)
    Sheets("SGSNNODE").Visible = (cbGbFR.value Or cbGbIP.value)
    Sheets("NRISGSNMAP").Visible = (cbGbFR.value Or cbGbIP.value)
    Sheets("SGSNROUTEINFO").Visible = (cbGbFR.value Or cbGbIP.value)
    Sheets("NSE").Visible = (cbGbFR.value Or cbGbIP.value)
    Sheets("BC").Visible = cbGbFR.value
    Sheets("NSVC").Visible = cbGbFR.value
    Sheets("PTPBVC").Visible = (cbGbFR.value Or cbGbIP.value)
    Sheets("NSVLLOCAL").Visible = cbGbIP.value
    Sheets("NSVLREMOTE").Visible = cbGbIP.value
   
    
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
