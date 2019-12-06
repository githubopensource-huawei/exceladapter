VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateCMForm 
   Caption         =   "Customize Template"
   ClientHeight    =   6180
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
    cbFor.Value = (cbIPOE.Value Or cbTDM.Value) And cbTimeslot.Value
    cbIdle.Value = (cbTDM.Value And cbTimeslot.Value)
    cbBTSTRANSTS.Value = (cbIPOE.Value Or cbTDM.Value) And cbTimeslot.Value
    cbBTSOMLTS.Value = (cbTDM.Value And cbTimeslot.Value)
    cbMONITORTS.Enabled = cbMONITORTS.Value
    cbFor.Enabled = cbFor.Value
    cbIdle.Enabled = cbIdle.Value
    cbBTSTRANSTS.Enabled = cbBTSTRANSTS.Value
    cbBTSOMLTS.Enabled = cbBTSOMLTS.Value
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

Private Sub OKBtn_Click()
  ThisWorkbook.Unprotect
  
  Dim WkSht As Worksheet
    
  Set WkSht = ThisWorkbook.ActiveSheet
  

  Sheets("BTS").Visible = True
  Sheets("BTSTRXBRD").Visible = cbTRXBRD.Value
  Sheets("BTSRXUCHAIN").Visible = cbRXU.Value
  Sheets("BTSRXUBRD").Visible = cbRXU.Value
  Sheets("BTSRXUBP").Visible = cbRXU.Value
  Sheets("BTSCONNECT").Visible = (cbTDM.Value Or cbIPOE.Value)
  Sheets("BTSTDM").Visible = cbTDM.Value
  Sheets("BTSTOPCONFIG").Visible = cbBtsTopConfig.Value '
  Sheets("BTSIDLETS").Visible = cbIdle.Value
  Sheets("BTSMONITORTS").Visible = cbMONITORTS.Value
  Sheets("BTSFORBIDTS").Visible = cbFor.Value
  Sheets("BTSANTFEEDERBRD").Visible = cbANT.Value
  Sheets("BTSANTFEEDERCONNECT").Visible = cbANT.Value
  Sheets("BTSRET").Visible = cbRETANT.Value
  Sheets("BTSRETSUBUNIT").Visible = cbRETANT.Value 'wy
  Sheets("BTSRETDEVICEDATA").Visible = cbRETANT.Value 'wy
  Sheets("BTSTMA").Visible = cbRETANT.Value
  Sheets("BTSTMASUBUNIT").Visible = cbRETANT.Value 'wy
  Sheets("BTSTMADEVICEDATA").Visible = cbRETANT.Value 'wy
  Sheets("GCELL").Visible = True
  Sheets("GTRX").Visible = True
  Sheets("GTRXDEV").Visible = cbGTRXAdvance.Value
  Sheets("GTRXCHAN").Visible = cbGTRXAdvance.Value
  Sheets("GTRXHOP").Visible = cbCellHOP.Value
  Sheets("GCELLMAGRP").Visible = cbCellHOP.Value
  Sheets("GCELLMAGRP_FREQ").Visible = cbCellHOP.Value
  Sheets("GTRXCHANHOP").Visible = cbCellHOP.Value
  Sheets("ADJNODE").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("ADJMAP").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("BTSIP").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("BTSDEVIP").Visible = (cbIPOE.Value Or cbIPFE.Value) 'wy 2010/03/05
  Sheets("IPLOGICPORT").Visible = (cbIPOE.Value Or cbIPFE.Value) '
  Sheets("BTSETHPORT").Visible = (cbIPOE.Value Or cbIPFE.Value) 'wy, 2010/12/16
  Sheets("BTSIPCLKPARA").Visible = cbIPFE.Value
  Sheets("BTSPPPLNK").Visible = cbPPP.Value
  Sheets("BTSMPGRP").Visible = cbMP.Value
  Sheets("BTSMPLNK").Visible = cbMP.Value
  Sheets("BTSESN").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("MPGRP").Visible = cbMP.Value
  Sheets("BTSBFD").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("BTSVLAN").Visible = (cbIPOE.Value Or cbIPFE.Value) '
  Sheets("BTSVLANCLASS").Visible = (cbIPOE.Value Or cbIPFE.Value) '
  Sheets("BTSVLANMAP").Visible = (cbIPOE.Value Or cbIPFE.Value) '
  Sheets("BTSDHCPSVRIP").Visible = (cbIPOE.Value Or cbIPFE.Value) '
  Sheets("BTSIPRT").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("BTSIPRTBIND").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("IPPATH").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("IPRT").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("DEVIP").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("PPPLNK").Visible = cbPPP.Value
  Sheets("MPLNK").Visible = cbMP.Value
  Sheets("ETHIP").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("DXX").Visible = cbDXX.Value
  Sheets("DXXCONNECT").Visible = cbDXX.Value
  Sheets("DXXTSEXGRELATION").Visible = cbDXX.Value
  Sheets("Frequency Tool").Visible = cbTool.Value
  Sheets("BTSARPSESSION").Visible = (cbIPOE.Value Or cbIPFE.Value)
  Sheets("TableDef").Visible = False
  Sheets("TableList").Visible = False
  Sheets("ValidDef").Visible = False
  Sheets("FieldMapDef").Visible = False
  
  Sheets("Frequency Tool").Move after:=Sheets("BTSARPSESSION")
  
  If WkSht.Visible Then
    WkSht.Activate
  End If
  TemplateCMForm.Hide
  
  ThisWorkbook.Protect Structure:=True, Windows:=False
End Sub

