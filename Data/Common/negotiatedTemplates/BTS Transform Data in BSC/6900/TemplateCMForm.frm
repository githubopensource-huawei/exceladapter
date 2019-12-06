VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateCMForm 
   Caption         =   "Customize Template"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   OleObjectBlob   =   "TemplateCMForm.frx":0000
   StartUpPosition =   1  'CenterOwner
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
  
    ThisWorkbook.Sheets("category").Cells(2, 2).value = TemplateCMForm.cbBSCIPBrd.value
    ThisWorkbook.Sheets("category").Cells(3, 2).value = TemplateCMForm.cbBSCIPOE.value
    ThisWorkbook.Sheets("category").Cells(4, 2).value = TemplateCMForm.cbBSCIPFE.value
    ThisWorkbook.Sheets("category").Cells(5, 2).value = TemplateCMForm.cbBTSAttr.value
    ThisWorkbook.Sheets("category").Cells(6, 2).value = TemplateCMForm.cbIPOE.value
    ThisWorkbook.Sheets("category").Cells(7, 2).value = TemplateCMForm.cbIPFE.value
    ThisWorkbook.Sheets("category").Cells(8, 2).value = TemplateCMForm.cbIPFE_E1.value
    ThisWorkbook.Sheets("category").Cells(9, 2).value = TemplateCMForm.cbBTSIPSec.value
    
    
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
   
  Call ShowBSCIPBrd(cbBSCIPBrd.value)
  Call ShowBSCIP(cbBSCIPOE.value, cbBSCIPFE.value)
  Call ShowBTSAttr(cbBTSAttr.value)
  Call ShowIPOEIPFE(cbIPOE.value, cbIPFE.value, cbIPFE_E1.value)
  Call ShowBTSIPSec(cbBTSIPSec.value)
  
  Sheets("FieldMapDef").Visible = False
End Sub



Private Sub UserForm_Initialize()
    TemplateCMForm.Caption = getResByKey("FormCaption_CustomTemplate")
    TemplateCMForm.ToolFrame.Caption = getResByKey("ToolFrameCaption_Summary")
    
    TemplateCMForm.cbBSCIPBrd.Caption = getResByKey("CheckBoxCaption_BSCIPBrd")
    TemplateCMForm.cbBSCIPOE.Caption = getResByKey("CheckBoxCaption_BSCIPOE")
    TemplateCMForm.cbBSCIPFE.Caption = getResByKey("CheckBoxCaption_BSCIPFE")
    TemplateCMForm.cbBTSAttr.Caption = getResByKey("CheckBoxCaption_BTSAttr")
    
    TemplateCMForm.cbIPOE.Caption = getResByKey("CheckBoxCaption_IPOE")
    TemplateCMForm.cbIPFE.Caption = getResByKey("CheckBoxCaption_IPFE")
    TemplateCMForm.cbIPFE_E1.Caption = getResByKey("CheckBoxCaption_IPFE_E1")
    TemplateCMForm.cbBTSIPSec.Caption = getResByKey("CheckBoxCaption_BTSIPSec")
    
    
    TemplateCMForm.cbBSCIPBrd.value = ThisWorkbook.Sheets("category").Cells(2, 2).value
    TemplateCMForm.cbBSCIPOE.value = ThisWorkbook.Sheets("category").Cells(3, 2).value
    TemplateCMForm.cbBSCIPFE.value = ThisWorkbook.Sheets("category").Cells(4, 2).value
    TemplateCMForm.cbBTSAttr.value = ThisWorkbook.Sheets("category").Cells(5, 2).value
    
    TemplateCMForm.cbIPOE.value = ThisWorkbook.Sheets("category").Cells(6, 2).value
    TemplateCMForm.cbIPFE.value = ThisWorkbook.Sheets("category").Cells(7, 2).value
    TemplateCMForm.cbIPFE_E1.value = ThisWorkbook.Sheets("category").Cells(8, 2).value
    TemplateCMForm.cbBTSIPSec.value = ThisWorkbook.Sheets("category").Cells(9, 2).value
    
    
    TemplateCMForm.cbBSCIPBrd.Enabled = ThisWorkbook.Sheets("category").Cells(2, 3).value
    TemplateCMForm.cbBSCIPOE.Enabled = ThisWorkbook.Sheets("category").Cells(3, 3).value
    TemplateCMForm.cbBSCIPFE.Enabled = ThisWorkbook.Sheets("category").Cells(4, 3).value
    TemplateCMForm.cbBTSAttr.Enabled = ThisWorkbook.Sheets("category").Cells(5, 3).value
    
    TemplateCMForm.cbIPOE.Enabled = ThisWorkbook.Sheets("category").Cells(6, 3).value
    TemplateCMForm.cbIPFE.Enabled = ThisWorkbook.Sheets("category").Cells(7, 3).value
    TemplateCMForm.cbIPFE_E1.Enabled = ThisWorkbook.Sheets("category").Cells(8, 3).value
    TemplateCMForm.cbBTSIPSec.Enabled = ThisWorkbook.Sheets("category").Cells(9, 3).value
    

End Sub







