VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateCMForm 
   Caption         =   "Customize Template"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
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
  
    ThisWorkbook.Sheets("category").Cells(2, 2).value = TemplateCMForm.cbTDM.value
    ThisWorkbook.Sheets("category").Cells(3, 2).value = TemplateCMForm.cbTDMDXX.value
    ThisWorkbook.Sheets("category").Cells(4, 2).value = TemplateCMForm.cbIPOE.value
    ThisWorkbook.Sheets("category").Cells(5, 2).value = TemplateCMForm.cbIPFE.value
    ThisWorkbook.Sheets("category").Cells(6, 2).value = TemplateCMForm.cbIPFEandE1T1.value
    
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
  Call ShowSelectSheets(cbTDM.value, cbTDMDXX.value, cbIPOE.value, cbIPFE.value)
End Sub


Private Sub UserForm_Initialize()
    TemplateCMForm.Caption = getResByKey("FormCaption_CustomTemplate")
    TemplateCMForm.ToolFrame.Caption = getResByKey("ToolFrameCaption_Summary")

    TemplateCMForm.cbTDM.Caption = getResByKey("CheckBoxCaption_TDM")
    TemplateCMForm.cbTDMDXX.Caption = getResByKey("CheckBoxCaption_TMDDXX")
    TemplateCMForm.cbIPOE.Caption = getResByKey("CheckBoxCaption_IPOE")
    TemplateCMForm.cbIPFE.Caption = getResByKey("CheckBoxCaption_IPFE")
    TemplateCMForm.cbIPFEandE1T1.Caption = getResByKey("CheckBoxCaption_IPFEandE1T1")
    
    TemplateCMForm.cbTDM.value = ThisWorkbook.Sheets("category").Cells(2, 2).value
    TemplateCMForm.cbTDMDXX.value = ThisWorkbook.Sheets("category").Cells(3, 2).value
    TemplateCMForm.cbIPOE.value = ThisWorkbook.Sheets("category").Cells(4, 2).value
    TemplateCMForm.cbIPFE.value = ThisWorkbook.Sheets("category").Cells(5, 2).value
    TemplateCMForm.cbIPFEandE1T1.value = ThisWorkbook.Sheets("category").Cells(6, 2).value
    
    TemplateCMForm.cbTDM.Enabled = ThisWorkbook.Sheets("category").Cells(2, 3).value
    TemplateCMForm.cbTDMDXX.Enabled = ThisWorkbook.Sheets("category").Cells(3, 3).value
    TemplateCMForm.cbIPOE.Enabled = ThisWorkbook.Sheets("category").Cells(4, 3).value
    TemplateCMForm.cbIPFE.Enabled = ThisWorkbook.Sheets("category").Cells(5, 3).value
    TemplateCMForm.cbIPFEandE1T1.Enabled = ThisWorkbook.Sheets("category").Cells(6, 3).value

End Sub
