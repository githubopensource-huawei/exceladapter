VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateCMForm 
   Caption         =   "Customize Template"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   OleObjectBlob   =   "TemplateCMForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "TemplateCMForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub CancelBtn_Click()
    Unload Me
End Sub


Private Sub OKBtn_Click()
    ThisWorkbook.Unprotect
  
    ThisWorkbook.Sheets("category").Cells(2, 2).value = TemplateCMForm.cbIPOE.value
    ThisWorkbook.Sheets("category").Cells(3, 2).value = TemplateCMForm.cbIPFE.value
    ThisWorkbook.Sheets("category").Cells(4, 2).value = TemplateCMForm.cbIPFEandE1T1.value
    
    Dim WkSht As Worksheet
    Set WkSht = ThisWorkbook.ActiveSheet
    
    Call ShowSheet
    
    If WkSht.Visible Then
        WkSht.Activate
    End If
    Unload Me
    
    ThisWorkbook.Protect Structure:=True, Windows:=False
End Sub

Private Sub ShowSheet()

  ThisWorkbook.Unprotect
  
  Sheets("MPGRP").Visible = False
  Sheets("MPLNK").Visible = False
  Sheets("ETHIP").Visible = False
  Sheets("PPPLNK").Visible = False
  Sheets("BTS").Visible = False
  Sheets("ADJNODE").Visible = False
  Sheets("IPPATH").Visible = False
  Sheets("IPRT").Visible = False
  Sheets("BTSIP").Visible = False
  Sheets("BTSETHPORT").Visible = False
  Sheets("BTSIPCLKPARA").Visible = False
  Sheets("BTSPPPLNK").Visible = False
  Sheets("BTSMPGRP").Visible = False
  Sheets("BTSMPLNK").Visible = False
  Sheets("BTSBFD").Visible = False
  Sheets("BTSIPRT").Visible = False
  Sheets("BTSIPRTBIND").Visible = False
  Sheets("BTSCONNECT").Visible = False
  Sheets("BTSMONITORTS").Visible = False
  Sheets("BTSDHCPSVRIP").Visible = False
  Sheets("BTSDEVIP").Visible = False
  Sheets("RSCGRP").Visible = False
  Sheets("IPLOGICPORT").Visible = False
  Sheets("DEVIP").Visible = False
  Sheets("ADJMAP").Visible = False
  Sheets("BTSVLAN").Visible = False
  Sheets("BTSVLANCLASS").Visible = False
  Sheets("BTSVLANMAP").Visible = False
  Sheets("BTSFORBIDTS").Visible = False
 
  If cbIPOE Then
    Sheets("BTS").Visible = True
    Sheets("ADJNODE").Visible = True
    Sheets("IPPATH").Visible = True
    Sheets("IPRT").Visible = True
    Sheets("BTSIP").Visible = True
    Sheets("BTSPPPLNK").Visible = True
    Sheets("BTSMPGRP").Visible = True
    Sheets("BTSMPLNK").Visible = True
    Sheets("BTSBFD").Visible = True
    Sheets("BTSIPRT").Visible = True
    Sheets("BTSIPRTBIND").Visible = True
    Sheets("BTSCONNECT").Visible = True
    Sheets("BTSMONITORTS").Visible = True
    Sheets("BTSDEVIP").Visible = True
    Sheets("RSCGRP").Visible = True
    Sheets("IPLOGICPORT").Visible = True
    Sheets("DEVIP").Visible = True
    Sheets("BTSDHCPSVRIP").Visible = True
    Sheets("BTSFORBIDTS").Visible = True
    Sheets("ADJMAP").Visible = True
    Sheets("MPGRP").Visible = True
    Sheets("MPLNK").Visible = True
    Sheets("PPPLNK").Visible = True
  End If
  
  If cbIPFE Then
    Sheets("ETHIP").Visible = True
    Sheets("BTS").Visible = True
    Sheets("ADJNODE").Visible = True
    Sheets("IPPATH").Visible = True
    Sheets("IPRT").Visible = True
    Sheets("BTSIP").Visible = True
    Sheets("BTSETHPORT").Visible = True
    Sheets("BTSIPCLKPARA").Visible = True
    Sheets("BTSBFD").Visible = True
    Sheets("BTSIPRT").Visible = True
    Sheets("BTSIPRTBIND").Visible = True
    Sheets("BTSDHCPSVRIP").Visible = True
    Sheets("BTSDEVIP").Visible = True
    Sheets("RSCGRP").Visible = True
    Sheets("IPLOGICPORT").Visible = True
    Sheets("DEVIP").Visible = True
    Sheets("ADJMAP").Visible = True
    Sheets("BTSVLAN").Visible = True
    Sheets("BTSVLANCLASS").Visible = True
    Sheets("BTSVLANMAP").Visible = True
  End If
  
  ThisWorkbook.Protect Structure:=True, Windows:=False

End Sub


Private Sub UserForm_Initialize()
    TemplateCMForm.Caption = getResByKey("FormCaption_CustomTemplate")
    TemplateCMForm.ToolFrame.Caption = getResByKey("ToolFrameCaption_Summary")
    
    TemplateCMForm.cbIPOE.Caption = getResByKey("CheckBoxCaption_IPOE")
    TemplateCMForm.cbIPFE.Caption = getResByKey("CheckBoxCaption_IPFE")
    TemplateCMForm.cbIPFEandE1T1.Caption = getResByKey("CheckBoxCaption_IPFEandE1T1")
    
    TemplateCMForm.cbIPOE.value = ThisWorkbook.Sheets("category").Cells(2, 2).value
    TemplateCMForm.cbIPFE.value = ThisWorkbook.Sheets("category").Cells(3, 2).value
    TemplateCMForm.cbIPFEandE1T1.value = ThisWorkbook.Sheets("category").Cells(4, 2).value
    
    TemplateCMForm.cbIPOE.Enabled = ThisWorkbook.Sheets("category").Cells(2, 3).value
    TemplateCMForm.cbIPFE.Enabled = ThisWorkbook.Sheets("category").Cells(3, 3).value
    TemplateCMForm.cbIPFEandE1T1.Enabled = ThisWorkbook.Sheets("category").Cells(4, 3).value

End Sub







